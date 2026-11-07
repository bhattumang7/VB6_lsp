use super::*;
use super::decl::*;
use super::expr::*;
use super::intrinsics::*;
use super::stmt::*;

pub(super) fn lower_assign(
    ctx: &LowerCtx,
    target_id: NodeId,
    value_id: NodeId,
    expr_arena: &ExprArena,
    out: &mut Vec<u8>,
) -> Result<(), LowerError> {
    // Array element store: `a(i) = v` (target is a subscript `Call`).
    if let ExprNode::Call { func, args } = expr_arena.get(target_id) {
        return lower_array_store(ctx, *func, *args, value_id, expr_arena, out);
    }
    // Member store: `t.X = v` (UDT field) or `o.F = v` (class field) — target
    // is a `MemberAccess`. Both go through the real resolver/value-emitter or
    // vtable-dispatch chain, not the scalar bypass store this function
    // otherwise builds — neither has a frame slot of its own to feed
    // `emit_var_store`.
    if let ExprNode::MemberAccess { base, member, bang, .. } = expr_arena.get(target_id) {
        let (base, member, bang) = (*base, *member, *bang);
        if member_access_base_is_class(ctx.module, base) {
            return lower_class_field_store(ctx, base, target_id, bang, value_id, expr_arena, out);
        }
        return lower_udt_field_store(ctx, base, member, bang, value_id, expr_arena, out);
    }
    // Resolve the target first so its type can be used to coerce integer literals
    // in the value expression (e.g. `r = 1` where r is Long → Long literal).
    let resolution = ctx
        .module
        .resolutions
        .get(&target_id.0)
        .ok_or(LowerError::Unresolved)?;
    let site = AssignSite { ctx, target_id, value_id, resolution, expr_arena };

    // Each special-case shape is tried in order; the first that recognizes the
    // statement lowers it completely.
    for handler in ASSIGN_HANDLERS {
        if handler(&site, out)? {
            return Ok(());
        }
    }
    lower_plain_assign(&site, out)
}

/// One scalar assignment statement: its target (already resolved) and value.
struct AssignSite<'a> {
    ctx: &'a LowerCtx<'a>,
    target_id: NodeId,
    value_id: NodeId,
    resolution: &'a NameResolution,
    expr_arena: &'a ExprArena,
}

/// A special-case lowering: `Ok(true)` when it recognized the statement and
/// emitted it, `Ok(false)` when the statement is not of its shape.
type AssignHandler = fn(&AssignSite<'_>, &mut Vec<u8>) -> Result<bool, LowerError>;

const ASSIGN_HANDLERS: [AssignHandler; 12] = [
    assign_rtc_string,
    assign_owned_concat,
    assign_owned_len,
    assign_owned_compare,
    assign_static_local,
    assign_fixed_string_source,
    assign_variant_target,
    assign_variant_source,
    assign_class_method_result,
    assign_proc_result,
    assign_array_element,
    assign_concat_chain,
];

impl AssignSite<'_> {
    /// The target's local index when it resolves to a local.
    fn local_idx(&self) -> Option<usize> {
        match self.resolution {
            NameResolution::Local { local_idx, .. } => Some(*local_idx),
            _ => None,
        }
    }

    /// The target's local index; any other target kind is not supported here.
    fn require_local(&self) -> Result<usize, LowerError> {
        self.local_idx().ok_or(LowerError::UnsupportedNode)
    }

    /// The frame offset of a local target; any other target kind is not
    /// supported here.
    fn require_local_offset(&self) -> Result<i16, LowerError> {
        Ok(self.ctx.local_slots[self.require_local()?].frame_offset)
    }
}

/// Emit a plain variable store into a local's frame slot with the given store
/// context.
fn emit_local_var_store(ctx: &LowerCtx, local_idx: usize, store_ctx: usize, out: &mut Vec<u8>) {
    let arena = NodeArena::new();
    let mut em = Emitter::new(&arena);
    em.emit_var_store(store_ctx, ctx.local_slots[local_idx].frame_offset);
    out.extend(em.into_bytes());
}

/// Release the staged String and Object argument temps of a class-method call
/// after the target store. A call staging both kinds at once has no confirmed
/// relative release order, so that combination is gated.
fn emit_pending_temp_releases(
    string_temps: &[i16],
    object_temps: &[i16],
    out: &mut Vec<u8>,
) -> Result<(), LowerError> {
    if !string_temps.is_empty() && !object_temps.is_empty() {
        return Err(LowerError::UnsupportedNode);
    }
    emit_temp_release_list(string_temps, out);
    emit_object_temp_release_list(object_temps, out);
    Ok(())
}

/// String-returning runtime intrinsic into a String target. Each argument is
/// pushed (right-to-left) either by value (size-load, coerced) or boxed into a
/// hidden 16-byte temp tagged with its VARTYPE (`04 <var> 4d <temp> <vt> 40`).
/// A final hidden result temp is passed by address; after the runtime call (0x0a)
/// the result is loaded (0x60), moved into the target (0x31), and the result temp
/// freed (0x35). The runtime-call reference index and total argument-byte count
/// follow the same convention as user calls.
fn assign_rtc_string(site: &AssignSite, out: &mut Vec<u8>) -> Result<bool, LowerError> {
    let ctx = site.ctx;
    if !matches!(ctx.module.builtins.get(&site.value_id.0), Some(BuiltinCall::RtcString { .. })) {
        return Ok(false);
    }
    let tgt_off = site.require_local_offset()?;
    // Emit the runtime call (result into a hidden temp), then load the result
    // (0x60), move it into the target (0x31), and free the owned variant temps.
    let (result_temp, free_offsets) =
        emit_rtc_string_call(ctx, site.value_id, site.expr_arena, out)?;
    out.push(0x04);
    out.extend_from_slice(&result_temp.to_le_bytes());
    out.push(0x60);
    out.push(0x31);
    out.extend_from_slice(&tgt_off.to_le_bytes());
    emit_variant_temp_free(out, &free_offsets);
    Ok(true)
}

/// An owned-temp string concatenation (a `&` chain with a runtime-string operand)
/// is lowered with the cleanup-aware concat opcode family.
fn assign_owned_concat(site: &AssignSite, out: &mut Vec<u8>) -> Result<bool, LowerError> {
    let ctx = site.ctx;
    if !is_owned_concat(ctx.module, site.value_id, site.expr_arena) {
        return Ok(false);
    }
    let Some(local_idx) = site.local_idx() else {
        return Ok(false);
    };
    let tgt_off = ctx.local_slots[local_idx].frame_offset;
    lower_owned_concat(ctx, tgt_off, site.value_id, site.expr_arena, out)?;
    Ok(true)
}

/// `Len` of a runtime-string result (the release-aware Len form) into a Long.
fn assign_owned_len(site: &AssignSite, out: &mut Vec<u8>) -> Result<bool, LowerError> {
    let ctx = site.ctx;
    if !is_owned_len(ctx.module, site.value_id, site.expr_arena) {
        return Ok(false);
    }
    let Some(local_idx) = site.local_idx() else {
        return Ok(false);
    };
    if !matches!(ctx.local_type(local_idx), VbaType::Long) {
        return Ok(false);
    }
    let tgt_off = ctx.local_slots[local_idx].frame_offset;
    lower_owned_len(ctx, tgt_off, site.value_id, site.expr_arena, out)?;
    Ok(true)
}

/// A String comparison with a runtime-string operand into an Integer/Boolean.
fn assign_owned_compare(site: &AssignSite, out: &mut Vec<u8>) -> Result<bool, LowerError> {
    let ctx = site.ctx;
    let cmp_value = unwrap_parens(site.value_id, site.expr_arena);
    if !is_owned_compare(ctx.module, cmp_value, site.expr_arena) {
        return Ok(false);
    }
    let Some(local_idx) = site.local_idx() else {
        return Ok(false);
    };
    if !matches!(ctx.local_type(local_idx), VbaType::Integer | VbaType::Boolean) {
        return Ok(false);
    }
    let tgt_off = ctx.local_slots[local_idx].frame_offset;
    lower_owned_compare(ctx, tgt_off, cmp_value, site.expr_arena, out)?;
    Ok(true)
}

/// Static local target: the value is pushed, then stored into the procedure's
/// static block (0x5f-addressed) rather than the frame.
fn assign_static_local(site: &AssignSite, out: &mut Vec<u8>) -> Result<bool, LowerError> {
    let ctx = site.ctx;
    let Some(local_idx) = site.local_idx() else {
        return Ok(false);
    };
    if !ctx.proc.locals[local_idx].is_static {
        return Ok(false);
    }
    let ty = ctx.local_type(local_idx).clone();
    let store_op = static_store_op(&ty).ok_or(LowerError::UnsupportedType)?;
    let off = ctx.static_offsets[local_idx];
    let coerce_tag = vba_type_to_node_tag(&ty);
    let mut arena = NodeArena::new();
    let root = lower_expr_coerced(ctx, site.value_id, site.expr_arena, &mut arena, coerce_tag)?;
    let root = coerce_assign_value(ctx, site.value_id, root, coerce_tag, &mut arena);
    let mut emitter = Emitter::new(&arena);
    emitter.emit_expr(root, 2);
    out.extend(emitter.into_bytes());
    emit_static_access(ctx, out, &store_op, off);
    Ok(true)
}

/// Fixed-length-string source: `s = fx` where `fx` is `As String * n` reads the
/// fixed buffer length-aware (LdAddr `fx` + 0x33<len>) and moves the resulting
/// BSTR temp into the target (0x31). The source isn't a plain BSTR-pointer load.
fn assign_fixed_string_source(site: &AssignSite, out: &mut Vec<u8>) -> Result<bool, LowerError> {
    let ctx = site.ctx;
    if !matches!(site.expr_arena.get(site.value_id), ExprNode::NameRef { .. }) {
        return Ok(false);
    }
    let Some(NameResolution::Local { local_idx: src_idx, .. }) =
        ctx.module.resolutions.get(&site.value_id.0)
    else {
        return Ok(false);
    };
    let Some(len) = ctx.proc.locals[*src_idx].fixed_string_len else {
        return Ok(false);
    };
    let Some(tgt_idx) = site.local_idx() else {
        return Ok(false);
    };
    let src_off = ctx.local_slots[*src_idx].frame_offset;
    let tgt_off = ctx.local_slots[tgt_idx].frame_offset;
    out.push(0x04);
    out.extend_from_slice(&src_off.to_le_bytes());
    out.push(0x33);
    out.extend_from_slice(&len.to_le_bytes());
    out.push(0x31);
    out.extend_from_slice(&tgt_off.to_le_bytes());
    Ok(true)
}

/// Variant target: the RHS is converted into a hidden 16-byte Variant temp,
/// then variant-stored (0xfc 0xf6 = value-emitter index 0x1f6) into the target.
fn assign_variant_target(site: &AssignSite, out: &mut Vec<u8>) -> Result<bool, LowerError> {
    let ctx = site.ctx;
    if !matches!(ctx.module.types.get(&site.target_id.0), Some(VbaType::Variant)) {
        return Ok(false);
    }
    let v_off = site.require_local_offset()?;
    let vi = ctx.variant_next.get();
    ctx.variant_next.set(vi + 1);
    let temp_off = ctx.local_slots[ctx.variant_base + vi].frame_offset;
    emit_variant_temp_init(ctx, site.value_id, temp_off, site.expr_arena, out)?;
    out.push(0xfc);
    out.push(0xf6);
    out.extend_from_slice(&v_off.to_le_bytes());
    Ok(true)
}

/// Initialize a hidden Variant temp from the assigned value.
fn emit_variant_temp_init(
    ctx: &LowerCtx,
    value_id: NodeId,
    temp_off: i16,
    expr_arena: &ExprArena,
    out: &mut Vec<u8>,
) -> Result<(), LowerError> {
    match expr_arena.get(value_id) {
        // Integer literal: init the temp from the inline literal (0x28 index
        // 0x3c0): 0x28 <temp> <2-byte int>.
        ExprNode::Literal { lit: AstLit::Int(n) } => {
            out.push(0x28);
            out.extend_from_slice(&temp_off.to_le_bytes());
            out.extend_from_slice(&(*n as i16).to_le_bytes());
        }
        // Long variable: load it, then convert Long->Variant into the temp
        // (0xfd 0x69 = value-emitter index 0x269).
        ExprNode::NameRef { .. }
            if matches!(ctx.module.types.get(&value_id.0), Some(VbaType::Long)) =>
        {
            out.extend_from_slice(&lower_expr_to_bytes(ctx, value_id, expr_arena)?);
            out.push(0xfd);
            out.push(0x69);
            out.extend_from_slice(&temp_off.to_le_bytes());
        }
        // `Empty` / `Null` initialize the temp directly: 0xfc 0x67 / 0xfc 0x64.
        ExprNode::Literal { lit: AstLit::Empty } => {
            out.extend_from_slice(&[0xfc, 0x67]);
            out.extend_from_slice(&temp_off.to_le_bytes());
        }
        ExprNode::Literal { lit: AstLit::Null } => {
            out.extend_from_slice(&[0xfc, 0x64]);
            out.extend_from_slice(&temp_off.to_le_bytes());
        }
        _ => return Err(LowerError::UnsupportedNode),
    }
    Ok(())
}

/// Variant source into a non-Variant scalar target: a Variant is read by
/// address (0x04), converted to the target type, then stored.
fn assign_variant_source(site: &AssignSite, out: &mut Vec<u8>) -> Result<bool, LowerError> {
    let ctx = site.ctx;
    let value_id = site.value_id;
    if !matches!(site.expr_arena.get(value_id), ExprNode::NameRef { .. })
        || !matches!(ctx.module.types.get(&value_id.0), Some(VbaType::Variant))
    {
        return Ok(false);
    }
    let v_off = match ctx.module.resolutions.get(&value_id.0) {
        Some(NameResolution::Local { local_idx, .. }) => ctx.local_slots[*local_idx].frame_offset,
        _ => return Err(LowerError::UnsupportedNode),
    };
    let local_idx = site.require_local()?;
    let ty = ctx.local_type(local_idx);
    let target_tag = vba_type_to_node_tag(ty).ok_or(LowerError::UnsupportedType)?;
    let store_ctx = load_store_ctx(ty).ok_or(LowerError::UnsupportedType)?;
    let store_off = ctx.local_slots[local_idx].frame_offset;
    let arena = NodeArena::new();
    let mut em = Emitter::new(&arena);
    em.emit_ldaddr(v_off);
    // Variant (tag 0xf) → target conversion, then the plain target store.
    em.emit_conversion(target_tag as i32, 0xf);
    em.emit_var_store(store_ctx, store_off);
    out.extend(em.into_bytes());
    Ok(true)
}

/// Class-member Function call as the RHS: `r = o.Method(args)` — emit the
/// vtable-dispatch call with a result temp (`lower_class_method_call`,
/// `is_value = true`), then store the loaded-back result into the target.
/// An `Object`-returning method belongs to the `Set` spelling (real VB6
/// requires `Set r = o.Method(...)` for an Object result) — that shape is
/// handled by `lower_set_assign`/`lower_set_from_class_method_call`
/// instead, so it's excluded here rather than guessed.
fn assign_class_method_result(site: &AssignSite, out: &mut Vec<u8>) -> Result<bool, LowerError> {
    let ctx = site.ctx;
    let expr_arena = site.expr_arena;
    let ExprNode::Call { func, args } = expr_arena.get(site.value_id) else {
        return Ok(false);
    };
    let ExprNode::MemberAccess { base, .. } = expr_arena.get(*func) else {
        return Ok(false);
    };
    let (base, func, args) = (*base, *func, *args);
    if !member_access_base_is_class(ctx.module, base) {
        return Ok(false);
    }
    // Target type drives the store — real VB6 requires the
    // target's declared type to match the function's return type,
    // so a plain-`Assign` `Object`-typed target here means the
    // source wrote it (`r = o.Method()` with `r As Object`),
    // which is invalid VB6 (`Set` is required); reject rather
    // than guess. `module.types` has no entry for a `Call` node
    // (the binder never records one), so the RHS's type is read
    // off the resolved TARGET local instead.
    let local_idx = site.require_local()?;
    let target_ty = ctx.local_type(local_idx).clone();
    if matches!(target_ty, VbaType::Object) {
        return Err(LowerError::UnsupportedNode);
    }
    let (pending_releases, pending_object_releases, _) =
        lower_class_method_call(ctx, base, func, args, true, false, expr_arena, out)?;
    // A String result's temp is read back via the `0x3e`
    // steal-load (see `lower_class_method_call`) — it must be
    // consumed with the move-store (ctx 9), the same rule
    // already grounded for a class-member Get returning
    // String (`value_is_class_get_string` below). Oracle-
    // confirmed: `oracle_bank/c5_func_string`.
    let sctx = if matches!(target_ty, VbaType::String) {
        9
    } else {
        load_store_ctx(&target_ty).ok_or(LowerError::UnsupportedType)?
    };
    emit_local_var_store(ctx, local_idx, sctx, out);
    // The staged String/Object argument temps' release comes
    // AFTER the target store — see `lower_class_method_call`'s
    // own doc comment on its returned temps. Oracle-confirmed:
    // `oracle_bank/c5_func_string` (String), `oracle_bank/
    // c15_func_with_obj_arg_release` (Object).
    emit_pending_temp_releases(&pending_releases, &pending_object_releases, out)?;
    Ok(true)
}

/// Function call as the RHS: `r = F(args)` — emit the result-producing call
/// (0x5e, result left on the stack), then store the result into the target.
fn assign_proc_result(site: &AssignSite, out: &mut Vec<u8>) -> Result<bool, LowerError> {
    let ctx = site.ctx;
    let ExprNode::Call { func, args } = site.expr_arena.get(site.value_id) else {
        return Ok(false);
    };
    if !matches!(ctx.module.resolutions.get(&func.0), Some(NameResolution::Proc(_))) {
        return Ok(false);
    }
    lower_call(ctx, *func, *args, true, site.expr_arena, out)?;
    let local_idx = site.require_local()?;
    let sctx = load_store_ctx(ctx.local_type(local_idx)).ok_or(LowerError::UnsupportedType)?;
    emit_local_var_store(ctx, local_idx, sctx, out);
    Ok(true)
}

/// Array element load as the RHS: `r = a(i)` — push the Long index, LdAddr the
/// array descriptor, the element-load opcode (0x9e for Long), then store to the
/// target.
fn assign_array_element(site: &AssignSite, out: &mut Vec<u8>) -> Result<bool, LowerError> {
    let ctx = site.ctx;
    let ExprNode::Call { func, args } = site.expr_arena.get(site.value_id) else {
        return Ok(false);
    };
    let Some((arr_off, elem, _dims)) = array_local_info(ctx, *func) else {
        return Ok(false);
    };
    let idx = single_index(*args, site.expr_arena).ok_or(LowerError::UnsupportedNode)?;
    let load_op: u8 = match elem {
        VbaType::Long => 0x9e,
        VbaType::Integer => 0x9d,
        _ => return Err(LowerError::UnsupportedNode),
    };
    out.extend_from_slice(&lower_expr_to_bytes_coerced(ctx, idx, site.expr_arena, Some(8))?);
    out.push(0x04);
    out.extend_from_slice(&arr_off.to_le_bytes());
    out.push(load_op);
    let local_idx = site.require_local()?;
    let sctx = load_store_ctx(ctx.local_type(local_idx)).ok_or(LowerError::UnsupportedType)?;
    out.push(RT_STORE_BY_CTX[sctx]);
    out.extend_from_slice(&ctx.local_slots[local_idx].frame_offset.to_le_bytes());
    Ok(true)
}

/// String concatenation chain `s = a & b & c`: emit the first concat, then for
/// each further operand materialize the running result to a hidden temp
/// (store-keep 0x23) and concat. The final result is moved into the target
/// (0x31); the intermediate temps are then freed (0x2f) for BSTR cleanup.
fn assign_concat_chain(site: &AssignSite, out: &mut Vec<u8>) -> Result<bool, LowerError> {
    let ctx = site.ctx;
    let is_string_cat = matches!(ctx.module.types.get(&site.target_id.0), Some(VbaType::String))
        && matches!(
            site.expr_arena.get(site.value_id),
            ExprNode::BinOp { op: BinOpKind::Cat, .. }
        );
    if !is_string_cat {
        return Ok(false);
    }
    let s_off = site.require_local_offset()?;
    let mut ops = Vec::new();
    flatten_concat(site.value_id, site.expr_arena, &mut ops);
    let temps = emit_concat_operands(ctx, &ops, site.expr_arena, out)?;
    out.push(0x31);
    out.extend_from_slice(&s_off.to_le_bytes());
    // Release the tracked temps: a single 0x2f, or 0x32 <byte-len> <offsets…>.
    emit_temp_release_list(&temps, out);
    Ok(true)
}

/// Emit a concatenation chain's operands and `&` steps, returning the hidden
/// temps to release. Every operand that materializes a fresh BSTR (a converted
/// numeric or a fixed-length string) is store-kept (0x23) to a temp for later
/// cleanup; so is every intermediate concat result except the last.
fn emit_concat_operands(
    ctx: &LowerCtx,
    ops: &[NodeId],
    expr_arena: &ExprArena,
    out: &mut Vec<u8>,
) -> Result<Vec<i16>, LowerError> {
    let mut temps: Vec<i16> = Vec::new();
    for (i, &op) in ops.iter().enumerate() {
        emit_concat_operand(ctx, op, expr_arena, out)?;
        if concat_operand_is_fresh(ctx.module, ctx.proc, op) {
            temps.push(alloc_concat_temp(ctx, out));
        }
        if i >= 1 {
            out.push(0x2a);
            if i < ops.len() - 1 {
                temps.push(alloc_concat_temp(ctx, out));
            }
        }
    }
    Ok(temps)
}

/// The generic scalar assignment: the value is lowered coerced to the target's
/// type and stored into the target's local, parameter or module-variable slot.
fn lower_plain_assign(site: &AssignSite, out: &mut Vec<u8>) -> Result<(), LowerError> {
    let ctx = site.ctx;
    let (value_id, expr_arena) = (site.value_id, site.expr_arena);
    let coerce_tag = match site.resolution {
        NameResolution::Local { local_idx, .. } => vba_type_to_node_tag(ctx.local_type(*local_idx)),
        NameResolution::Param { param_idx, .. } => vba_type_to_node_tag(ctx.param_type(*param_idx)),
        NameResolution::ModuleVar(idx) => vba_type_to_node_tag(ctx.global_type(*idx)),
        _ => None,
    };

    let mut arena = NodeArena::new();
    let value_root = lower_expr_coerced(ctx, value_id, expr_arena, &mut arena, coerce_tag)?;
    let value_root = coerce_assign_value(ctx, value_id, value_root, coerce_tag, &mut arena);
    let value_is_fresh_string = value_is_fresh_string(ctx, value_id, expr_arena, &arena, value_root);

    let mut emitter = Emitter::new(&arena);
    // A UDT field reference on the RHS (`y = t.X`) resolved through the
    // real resolver chain and left its `SymbolContext` in this side channel
    // (see `lower_udt_field_access`) — attach it before emitting.
    if let Some(sym) = ctx.member_symbol.borrow_mut().take() {
        emitter = emitter.with_symbol_context(sym);
    }
    // The assigned value is an rvalue: emit it in value context (2). This is the
    // context that selects the typed floating-point push forms (Double/Date push
    // 0xfa, Single push 0xf9); context 0 would fall to the untyped push (0xf6/0xf5).
    emitter.emit_expr(value_root, 2);
    emit_target_store(site, &mut emitter, value_is_fresh_string)?;
    out.extend(emitter.into_bytes());
    Ok(())
}

/// A String target receives a *move* store (0x31, ctx 9) when the value is a
/// freshly-produced temp — a `&` concatenation, a numeric→String coercion (0x78),
/// a `CStr` explicit conversion (0x7c), or a class-member Property Get/field
/// returning String (its own steal-load, `0x3e`, already zeroed the source
/// temp — see `ClassFieldRef::is_string`) — and a *copy* store (0x43, ctx 8)
/// when it is a plain string variable. The move avoids an extra BSTR
/// allocation (and, for the class-Get case, is the ONLY correct store: the
/// steal-load leaves nothing else owning the value to release later).
fn value_is_fresh_string(
    ctx: &LowerCtx,
    value_id: NodeId,
    expr_arena: &ExprArena,
    arena: &NodeArena,
    value_root: NodeRef,
) -> bool {
    let value_is_class_get_string = matches!(
        expr_arena.get(value_id),
        ExprNode::MemberAccess { base, .. } if member_access_base_is_class(ctx.module, *base)
    ) && matches!(ctx.module.types.get(&value_id.0), Some(VbaType::String));
    matches!(
        expr_arena.get(value_id),
        ExprNode::BinOp { op: BinOpKind::Cat, .. }
    ) || matches!(arena.get(value_root).opcode(), 0x78 | 0x7c)
        || value_is_class_get_string
}

/// Append the typed store of the already-emitted value into the target slot.
fn emit_target_store(
    site: &AssignSite,
    emitter: &mut Emitter,
    value_is_fresh_string: bool,
) -> Result<(), LowerError> {
    let ctx = site.ctx;
    match site.resolution {
        NameResolution::Local { local_idx, .. } => {
            let ty = ctx.local_type(*local_idx);
            let slot = &ctx.local_slots[*local_idx];
            // A String assigned a freshly-produced temp (a `&` concat result) is
            // moved, not copied: store opcode 0x31 (ctx 9) instead of 0x43 (ctx 8).
            let sctx = if matches!(ty, VbaType::String) && value_is_fresh_string {
                9
            } else {
                load_store_ctx(ty).ok_or(LowerError::UnsupportedType)?
            };
            emitter.emit_var_store(sctx, slot.frame_offset);
        }
        NameResolution::Param { param_idx, .. } => {
            let ty = ctx.param_type(*param_idx);
            let slot = &ctx.param_slots[*param_idx];
            let sctx = load_store_ctx(ty).ok_or(LowerError::UnsupportedType)?;
            if slot.byref {
                emitter.emit_byref_store(sctx, slot.frame_offset);
            } else {
                emitter.emit_var_store(sctx, slot.frame_offset);
            }
        }
        NameResolution::ModuleVar(idx) => {
            let ty = ctx.global_type(*idx);
            let slot = &ctx.global_slots[*idx];
            let sctx = load_store_ctx(ty).ok_or(LowerError::UnsupportedType)?;
            emitter.emit_global_store(sctx, slot.module_desc, slot.field_offset);
        }
        res if super::expr::is_project_level(res) => return Err(LowerError::UnsupportedNode),
        _ => return Err(LowerError::Unresolved),
    }
    Ok(())
}

/// `Set o.P = v` (an explicit `Property Set` call). Unlike plain `Assign`,
/// `Set`'s target is restricted here to a `MemberAccess` resolving to a class
/// Property Set — `Set localVar = v` (assigning a bare object variable, no
/// vtable call) is a distinct, unimplemented case gated below.
pub(super) fn lower_set_assign(
    ctx: &LowerCtx,
    target_id: NodeId,
    value_id: NodeId,
    expr_arena: &ExprArena,
    out: &mut Vec<u8>,
) -> Result<(), LowerError> {
    if let ExprNode::MemberAccess { base, bang, .. } = expr_arena.get(target_id) {
        let (base, bang) = (*base, *bang);
        if member_access_base_is_class(ctx.module, base) {
            return lower_class_field_set(ctx, base, target_id, bang, value_id, expr_arena, out);
        }
    }
    if let Some((target_class, dest_off, _)) = plain_object_local(ctx, target_id, expr_arena) {
        return lower_set_plain_object_local(ctx, target_class, dest_off, value_id, expr_arena, out);
    }
    // `Set x = o.P` where `x` is declared a plain `Object` (not a specific
    // class type) — only the class-member-Get value shape is grounded for
    // this target kind (see `lower_set_from_class_get`); every other RHS
    // shape against an `Object`-typed target is ungrounded and gated there.
    if let Some(dest_off) = object_typed_local(ctx, target_id, expr_arena) {
        return lower_set_from_class_get(ctx, dest_off, value_id, expr_arena, out);
    }
    Err(LowerError::UnsupportedNode)
}

/// A plain `Object`-typed local (`Dim x As Object`) — distinct from
/// `plain_object_local`, which only matches a specific-class-typed
/// (`UserDefined`) local.
fn object_typed_local(ctx: &LowerCtx, node_id: NodeId, expr_arena: &ExprArena) -> Option<i16> {
    let ExprNode::NameRef { .. } = expr_arena.get(node_id) else {
        return None;
    };
    match ctx.module.resolutions.get(&node_id.0) {
        Some(NameResolution::Local { local_idx, .. }) => {
            let local = ctx.proc.locals.get(*local_idx)?;
            if !matches!(local.vba_type, VbaType::Object) {
                return None;
            }
            Some(ctx.local_slots[*local_idx].frame_offset)
        }
        _ => None,
    }
}

/// `Set localVar = v` where `localVar` is a plain object-typed local (no
/// vtable dispatch — `o.Field`/`o.Property` route through
/// `lower_class_field_set` instead). Three source forms are oracle-confirmed
/// (`c7_set_new_reassign_nothing`, one `Sub Main` exercising all three in
/// sequence against a single class):
///
/// - `Set o = New ClassName`: `fd f4 <create-idx>` (construct, unref'd —
///   ownership is established by the store) then `19 <dest>` (pop, AddRef,
///   release-old-and-store — flag `-1`/AddRef, confirmed from the runtime
///   handler at table index 0x19: `push -1, push &dest, push TOS, call` the
///   shared `__vbaObjSet`/`__vbaObjSetAddref` routine).
/// - `Set o = otherObjLocal` where `otherObjLocal` was declared `As New`:
///   `04 <src>` (LdAddr) then `56 <create-idx>` (the lazy-fetch opcode: read
///   `*src`; if null, construct via the SAME create call as `New` and store
///   back; push the result — already-owned, per the runtime handler sharing
///   its post-construct tail with the non-null fast path) then `fc f8
///   <dest>` (pop, release-old-and-store, flag `0`/no-AddRef — steals the
///   reference `56` already owns, table index 0x2f8, confirmed sharing case
///   0x19's tail with `EBX=0` instead of `-1`).
/// - `Set o = Nothing`: `fc 63` (push literal 0, table index 0x263, no
///   operand of its own) then `3d <type-idx>` (table index 0x3d — pops the 0,
///   coerces it to `o`'s declared class via a SEPARATE `class_const_table`
///   entry than the create-idx above: index 1, not a fresh 0, proving the two
///   entry kinds share one table) then `19 <dest>` (AddRef-store; safe on a
///   null pointer — the shared routine no-ops the AddRef when the value it's
///   given is null, same call as the `New` case above).
///
/// A plain (non-`As New`) object local read as a `Set` source, a `New` of a
/// class other than the target's declared type, and a mismatched-class
/// var-to-var copy are all UNGROUNDED (no oracle capture distinguishes them
/// from the shapes above) and gated rather than guessed.
pub(super) fn lower_set_plain_object_local(
    ctx: &LowerCtx,
    target_class: u32,
    dest_off: i16,
    value_id: NodeId,
    expr_arena: &ExprArena,
    out: &mut Vec<u8>,
) -> Result<(), LowerError> {
    match expr_arena.get(value_id) {
        ExprNode::New { type_spec } => {
            let ExprNode::UserType { name, .. } = expr_arena.get(*type_spec) else {
                return Err(LowerError::UnsupportedNode);
            };
            if *name != target_class {
                return Err(LowerError::UnsupportedType);
            }
            let idx = ctx.intern_class_const(ClassConstKind::Create, target_class);
            out.push(0xfd);
            out.push(0xf4);
            out.extend_from_slice(&idx.to_le_bytes());
            out.push(0x19);
            out.extend_from_slice(&(dest_off as u16).to_le_bytes());
            Ok(())
        }
        ExprNode::Nothing => {
            // `fc 63` (push literal 0, no operand of its own) then `3d
            // <type-idx>` (the typed-value coercion opcode — pops the 0,
            // reads the class's type-descriptor const-table entry, pushes a
            // typed-Nothing value) then the AddRef-store.
            let idx = ctx.intern_class_const(ClassConstKind::TypeDesc, target_class);
            out.push(0xfc);
            out.push(0x63);
            out.push(0x3d);
            out.extend_from_slice(&idx.to_le_bytes());
            out.push(0x19);
            out.extend_from_slice(&(dest_off as u16).to_le_bytes());
            Ok(())
        }
        ExprNode::NameRef { .. } => {
            let Some((src_class, src_off, src_is_new)) = plain_object_local(ctx, value_id, expr_arena)
            else {
                return Err(LowerError::UnsupportedNode);
            };
            if !src_is_new || src_class != target_class {
                return Err(LowerError::UnsupportedType);
            }
            let idx = ctx.intern_class_const(ClassConstKind::Create, src_class);
            out.push(0x04);
            out.extend_from_slice(&(src_off as u16).to_le_bytes());
            out.push(0x56);
            out.extend_from_slice(&idx.to_le_bytes());
            out.push(0xfc);
            out.push(0xf8);
            out.extend_from_slice(&(dest_off as u16).to_le_bytes());
            Ok(())
        }
        ExprNode::MemberAccess { .. } | ExprNode::Call { .. } => {
            lower_set_class_get_to_specific_class_target(
                ctx,
                target_class,
                dest_off,
                value_id,
                expr_arena,
                out,
            )
        }
        _ => Err(LowerError::UnsupportedNode),
    }
}

/// `o.Method(args)` on a class-typed base whose declared return type is
/// `Object`: the `(base, callee, args)` node ids of the call, else `None`.
fn class_object_method_call(
    ctx: &LowerCtx,
    value_id: NodeId,
    expr_arena: &ExprArena,
) -> Option<(NodeId, NodeId, NodeId)> {
    let ExprNode::Call { func, args } = expr_arena.get(value_id) else {
        return None;
    };
    let ExprNode::MemberAccess { base, .. } = expr_arena.get(*func) else {
        return None;
    };
    let ret_is_object = ctx
        .module
        .class_member_slots
        .get(&func.0)
        .is_some_and(|r| matches!(r.method_ret_type, Some(VbaType::Object)));
    (member_access_base_is_class(ctx.module, *base) && ret_is_object).then_some((*base, *func, *args))
}

/// `Set dest = o.P` where `dest` is a SPECIFIC-class-typed local (`Dim x As
/// Class1`, not plain `Object`) and `o.P` is an Object-returning class-member
/// Get. A DIFFERENT byte sequence from the plain-`Object`-typed-target case
/// (`lower_set_from_class_get`, `oracle_bank/c1_get_object`'s direct `0x51`
/// read + `0x19` store, no coercion) — assigning into a specific class needs
/// an explicit type coercion/check first: the Get's result is read back with
/// a PLAIN generic load (`0x6c`, not `0x51`) into a temp, then `0x3d
/// <type-idx>` coerces/verifies it against the TARGET's class, THEN the
/// AddRef-store (`0x19`) and a release (`0x1a`) — the same four-opcode tail
/// already grounded for slice #15's ByRef-argument write-back.
///
/// **RESOLVED this session** (previously flagged as an open question): the
/// `0x3d` operand is `intern_member_type_const(dest's OWN declared class)`,
/// NOT the callee's class — a fresh two-DISTINCT-class capture
/// (`oracle_bank/c16_two_class_set_coercion`, `o As Class1` calling a method
/// that returns a `Class2` instance, `x As Class2`) shows the coercion
/// operand is a FRESH index, DIFFERENT from the call's own member-type
/// operand (which is keyed by `Class1`, the callee) — proving the two are
/// independently keyed, by DIFFERENT classes, not aliased. Every prior
/// single-class capture (this function's own `c11`/`c13`, and slice #15's
/// write-back) is STILL consistent with this corrected rule, since the
/// target/destination and the callee were the SAME class in all of them —
/// they just never had a second, DIFFERENT class to reveal the distinction.
/// `Set o = Nothing`'s own `ClassConstKind::TypeDesc(target_class)` (a
/// genuinely separate mechanism, no vtable call in that context at all)
/// remains unaffected and uncontradicted by this finding.
fn lower_set_class_get_to_specific_class_target(
    ctx: &LowerCtx,
    target_class: u32,
    dest_off: i16,
    value_id: NodeId,
    expr_arena: &ExprArena,
    out: &mut Vec<u8>,
) -> Result<(), LowerError> {
    // `Set x = o.Method()` (a Call, not a bare Get) against the same
    // SPECIFIC-class-typed target kind — byte-IDENTICAL to the bare-Get case
    // below (same `0x6c`/`0x3d`/`0x19`/`0x1a` tail after the call), proving
    // this Set-target convention depends only on the TARGET's type, not on
    // whether the source is a Get or a method call. Oracle-confirmed:
    // `oracle_bank/c13_set_call_to_specific_class_target`.
    if let Some((base, func, args)) = class_object_method_call(ctx, value_id, expr_arena) {
        let (pending_releases, pending_object_releases, result_temp) =
            lower_class_method_call(ctx, base, func, args, true, true, expr_arena, out)?;
        let temp_off = result_temp.expect("is_value implies a result temp");
        out.push(0x3d);
        out.extend_from_slice(&ctx.intern_member_type_const(target_class).to_le_bytes());
        out.push(0x19);
        out.extend_from_slice(&(dest_off as u16).to_le_bytes());
        out.push(0x1a);
        out.extend_from_slice(&temp_off.to_le_bytes());
        return emit_pending_temp_releases(&pending_releases, &pending_object_releases, out);
    }
    let ExprNode::MemberAccess { base, bang, .. } = expr_arena.get(value_id) else {
        return Err(LowerError::UnsupportedNode);
    };
    let (base, bang) = (*base, *bang);
    if bang || !member_access_base_is_class(ctx.module, base) {
        return Err(LowerError::UnsupportedNode);
    }
    let field = resolve_class_field(ctx, base, value_id)?;
    if !field.is_object {
        return Err(LowerError::UnsupportedType);
    }
    let get_slot = field.get_slot.ok_or(LowerError::UnsupportedNode)?;
    let frame_ctx = field.frame_ctx.ok_or(LowerError::UnsupportedType)?;
    let temp_off = ctx.class_member_slot(frame_ctx, 0);

    out.push(0x04);
    out.extend_from_slice(&temp_off.to_le_bytes());
    out.push(0x04);
    out.extend_from_slice(&(field.obj_offset as u16).to_le_bytes());
    out.push(0x24);
    out.extend_from_slice(&ctx.intern_class_const(ClassConstKind::Create, field.class_sym).to_le_bytes());
    out.push(0x0d);
    out.extend_from_slice(&get_slot.to_le_bytes());
    out.extend_from_slice(&ctx.intern_member_type_const(field.class_sym).to_le_bytes());
    out.push(0x6c);
    out.extend_from_slice(&temp_off.to_le_bytes());
    out.push(0x3d);
    out.extend_from_slice(&ctx.intern_member_type_const(target_class).to_le_bytes());
    out.push(0x19);
    out.extend_from_slice(&(dest_off as u16).to_le_bytes());
    out.push(0x1a);
    out.extend_from_slice(&temp_off.to_le_bytes());
    Ok(())
}

/// `Set dest = o.P` where `o.P` is an Object-returning class-member Get: the
/// same vtable-Get + temp-read-back sequence as any other Get-typed access
/// (`lower_class_field_get`), landing here because the client spelling is
/// `Set`, not a plain `Assign` — followed by the AddRef-store `0x19`, the
/// same store used for `Set o = New`/`Set o = Nothing`
/// (`lower_set_plain_object_local`). Shared by both a class-typed and a
/// plain `Object`-typed `dest` local — the store itself doesn't distinguish
/// them. Oracle-confirmed: `oracle_bank/c1_get_object`.
///
/// Also handles `Set dest = o.Method(args)`, an Object-returning class
/// method CALL (not a bare Get) — `lower_class_method_call` (`is_value =
/// true`) already reads the result temp back via the same `0x51` opcode
/// (see its own doc comment), so only the trailing AddRef-store needs
/// adding here, same as the Get case. Oracle-confirmed (recaptured this
/// session, HIGH confidence): `oracle_bank/c5_func_object`.
fn lower_set_from_class_get(
    ctx: &LowerCtx,
    dest_off: i16,
    value_id: NodeId,
    expr_arena: &ExprArena,
    out: &mut Vec<u8>,
) -> Result<(), LowerError> {
    // `module.types` has no entry for a `Call` node (the binder never
    // records one) — the method's return type comes from the resolved
    // class-member slot instead.
    if let Some((base, func, args)) = class_object_method_call(ctx, value_id, expr_arena) {
        let (pending_releases, pending_object_releases, _) =
            lower_class_method_call(ctx, base, func, args, true, false, expr_arena, out)?;
        out.push(0x19);
        out.extend_from_slice(&(dest_off as u16).to_le_bytes());
        return emit_pending_temp_releases(&pending_releases, &pending_object_releases, out);
    }
    let ExprNode::MemberAccess { base, bang, .. } = expr_arena.get(value_id) else {
        return Err(LowerError::UnsupportedNode);
    };
    let (base, bang) = (*base, *bang);
    if bang || !member_access_base_is_class(ctx.module, base) {
        return Err(LowerError::UnsupportedNode);
    }
    if !matches!(ctx.module.types.get(&value_id.0), Some(VbaType::Object)) {
        return Err(LowerError::UnsupportedType);
    }
    let mut arena = NodeArena::new();
    let value_root = lower_class_field_get(ctx, base, value_id, bang, &mut arena)?;
    let mut emitter = Emitter::new(&arena);
    emitter.emit_expr(value_root, 2);
    out.extend(emitter.into_bytes());
    out.push(0x19);
    out.extend_from_slice(&(dest_off as u16).to_le_bytes());
    Ok(())
}

/// Lower `t.X = v` (a UDT field assignment): build a real bound `0x2c`
/// assignment-statement node — LHS a `0x60` field reference (offset baked
/// into `word[7]`), RHS the coerced value — and emit it through the actual
/// resolver/value-emitter chain (`Emitter::emit_expr`'s `0x2c` case), the same
/// path a plain scalar assignment bypasses via the direct-store opcodes.
fn lower_udt_field_store(
    ctx: &LowerCtx,
    base: NodeId,
    member: u32,
    bang: bool,
    value_id: NodeId,
    expr_arena: &ExprArena,
    out: &mut Vec<u8>,
) -> Result<(), LowerError> {
    if bang {
        return Err(LowerError::UnsupportedNode);
    }
    let field = resolve_udt_field(ctx, base, member)?;

    let mut arena = NodeArena::new();
    // LHS: the 0x60 member-reference node (offset in word[7]).
    let size_desc = arena.alloc(NodeArena::node(4, 0, field.field_size as u32, 0, 0, 0));
    let mut lhs_n = NodeArena::node(0x60, field.type_tag, 0, 0, 0, field.offset as u32);
    lhs_n.w[5] = size_desc.0;
    let lhs = arena.alloc(lhs_n);

    // RHS: the value in its own natural type, then an explicit conversion
    // node inserted when it differs from the field's type (same two-step
    // pattern as every other store path — `coerce_assign_value` handles the
    // literal-retype-in-place vs. explicit-conversion distinction; skipping
    // it here was a bug: an Integer literal into a Double field must emit
    // `eb` (Int->Double) before the store, not fold the literal in place).
    let rhs = lower_expr_coerced(ctx, value_id, expr_arena, &mut arena, Some(field.type_tag))?;
    let rhs = coerce_assign_value(ctx, value_id, rhs, Some(field.type_tag), &mut arena);

    // 0x2c assignment statement node: w4 = LHS, w5 = RHS, region 0 (not the
    // 0x20000 array/special-LHS region).
    let mut asn = NodeArena::node(0x2c, 0, 0, 0, 0, 0);
    asn.w[4] = lhs.0;
    asn.w[5] = rhs.0;
    let node = arena.alloc(asn);

    let mut emitter = Emitter::new(&arena).with_symbol_context(field.sym);
    emitter.emit_expr(node, 0);
    out.extend(emitter.into_bytes());
    Ok(())
}
