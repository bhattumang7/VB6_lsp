use super::*;
use super::expr::*;
use super::stmt::*;


/// Flatten a left-associative `&` concatenation chain into its operands in order
/// (`a & b & c` → `[a, b, c]`).
pub(super) fn flatten_concat(node_id: NodeId, expr_arena: &ExprArena, out: &mut Vec<NodeId>) {
    if let ExprNode::BinOp { op: BinOpKind::Cat, lhs, rhs } = expr_arena.get(node_id) {
        flatten_concat(*lhs, expr_arena, out);
        out.push(*rhs);
    } else {
        out.push(node_id);
    }
}

/// Count the hidden string temps needed for concat chains: a chain of N operands
/// materializes its N-2 intermediate results to temps (for BSTR cleanup).
/// Allocate the next concat temp slot, emit a store-keep (0x23) of the top-of-stack
/// BSTR into it, and return its frame offset (recorded for later release).
pub(super) fn alloc_concat_temp(ctx: &LowerCtx, out: &mut Vec<u8>) -> i16 {
    let ti = ctx.concat_next.get();
    ctx.concat_next.set(ti + 1);
    let t_off = ctx.local_slots[ctx.concat_base + ti].frame_offset;
    out.push(0x23);
    out.extend_from_slice(&t_off.to_le_bytes());
    t_off
}

/// Emit one concatenation operand's value bytes: a fixed-length string loads its
/// inline buffer length-aware (LdAddr + 0x33<len>); a non-String operand loads and
/// converts to String (the numeric→String conversion); a plain String operand or
/// literal loads directly.
pub(super) fn emit_concat_operand(
    ctx: &LowerCtx,
    op_id: NodeId,
    expr_arena: &ExprArena,
    out: &mut Vec<u8>,
) -> Result<(), LowerError> {
    match ctx.module.types.get(&op_id.0) {
        Some(VbaType::String) => {
            if let Some(NameResolution::Local { local_idx, .. }) =
                ctx.module.resolutions.get(&op_id.0)
            {
                if let Some(len) = ctx.proc.locals[*local_idx].fixed_string_len {
                    let off = ctx.local_slots[*local_idx].frame_offset;
                    out.push(0x04);
                    out.extend_from_slice(&off.to_le_bytes());
                    out.push(0x33);
                    out.extend_from_slice(&len.to_le_bytes());
                    return Ok(());
                }
            }
            out.extend_from_slice(&lower_expr_to_bytes(ctx, op_id, expr_arena)?);
            Ok(())
        }
        Some(_) => {
            // Numeric operand: load it, then convert to String (e.g. Long → 0xfb 0xfe).
            let mut arena = NodeArena::new();
            let root = lower_expr(ctx, op_id, expr_arena, &mut arena)?;
            let src_tag = arena.get(root).type_tag();
            let mut emitter = Emitter::new(&arena);
            emitter.emit_expr(root, 2);
            emitter.emit_conversion(0x10, src_tag);
            out.extend(emitter.into_bytes());
            Ok(())
        }
        None => {
            out.extend_from_slice(&lower_expr_to_bytes(ctx, op_id, expr_arena)?);
            Ok(())
        }
    }
}

/// Whether a concat operand materializes a *fresh* BSTR temp that needs tracking
/// for cleanup: any non-String operand (it is converted to a string) and any
/// fixed-length-string operand (its inline buffer is copied into a BSTR). A plain
/// variable-length String operand and a string literal are not fresh.
pub(super) fn concat_operand_is_fresh(module: &BoundModule, proc: &BoundProc, op_id: NodeId) -> bool {
    match module.types.get(&op_id.0) {
        Some(VbaType::String) => {
            if let Some(NameResolution::Local { local_idx, .. }) = module.resolutions.get(&op_id.0) {
                proc.locals[*local_idx].fixed_string_len.is_some()
            } else {
                false
            }
        }
        Some(_) => true,
        None => false,
    }
}

/// Number of concat temps for one concatenation: one per fresh operand plus one
/// per intermediate result (every concat but the last keeps its accumulator).
pub(super) fn concat_chain_temps(module: &BoundModule, proc: &BoundProc, ops: &[NodeId]) -> usize {
    let fresh = ops.iter().filter(|&&o| concat_operand_is_fresh(module, proc, o)).count();
    fresh + ops.len().saturating_sub(2)
}

pub(super) fn count_concat_temps(module: &BoundModule, proc: &BoundProc, node_id: NodeId, expr_arena: &ExprArena) -> usize {
    let c = |id: NodeId| count_concat_temps(module, proc, id, expr_arena);
    match expr_arena.get(node_id) {
        ExprNode::Assign { value, .. } => {
            // An owned-temp concat uses the 16-byte string-rtc pool instead, so it is
            // not counted here.
            if matches!(expr_arena.get(*value), ExprNode::BinOp { op: BinOpKind::Cat, .. })
                && !is_owned_concat(module, *value, expr_arena)
            {
                let mut ops = Vec::new();
                flatten_concat(*value, expr_arena, &mut ops);
                concat_chain_temps(module, proc, &ops)
            } else {
                0
            }
        }
        ExprNode::Block { stmts } => stmts.iter().map(|&id| c(id)).sum(),
        ExprNode::If { then_body, else_body, .. } => {
            c(*then_body) + else_body.map(c).unwrap_or(0)
        }
        ExprNode::While { body, .. } | ExprNode::Do { body, .. } | ExprNode::For { body, .. } => {
            c(*body)
        }
        ExprNode::SelectCase { cases, .. } => cases.iter().map(|&id| c(id)).sum(),
        ExprNode::CaseBlock { body, .. } | ExprNode::CaseElse { body } => c(*body),
        _ => 0,
    }
}

/// Count assignments whose target is a `Variant`. Each needs a hidden 16-byte
/// Variant temp to hold the converted value before the variant store.
/// Count hidden string-result temps needed: one per String-returning runtime
/// intrinsic call whose result is stored (the numeric-input `Chr`/`Space` family).
pub(super) fn count_string_rtc_temps(module: &BoundModule, node_id: NodeId, expr_arena: &ExprArena) -> usize {
    let c = |id: NodeId| count_string_rtc_temps(module, id, expr_arena);
    match expr_arena.get(node_id) {
        ExprNode::Assign { value, .. } => match module.builtins.get(&value.0) {
            // 16-byte temps: one per boxed argument; two per omitted-Optional
            // (a Missing-variant literal needs a separate value buffer alongside its
            // temp); plus one result temp.
            Some(BuiltinCall::RtcString { args }) => rtc_call_temp_count(args),
            // An owned-temp concat uses the 16-byte pool for every operand
            // materialization and intermediate concat result.
            _ if is_owned_concat(module, *value, expr_arena) => {
                let mut ops = Vec::new();
                flatten_concat(*value, expr_arena, &mut ops);
                count_owned_concat_temps(module, &ops)
            }
            // `Len` of a runtime-string result: the rtc call's temps plus a Len
            // scratch temp.
            _ if is_owned_len(module, *value, expr_arena) => {
                if let ExprNode::Call { args, .. } = expr_arena.get(*value) {
                    if let Some(arg) = single_index(*args, expr_arena) {
                        if let Some(BuiltinCall::RtcString { args: sig }) =
                            module.builtins.get(&arg.0)
                        {
                            return rtc_call_temp_count(sig) + 1;
                        }
                    }
                }
                0
            }
            // A String comparison with a runtime-string operand: both operands'
            // materialization temps plus one compare temp.
            _ if is_owned_compare(module, unwrap_parens(*value, expr_arena), expr_arena) => {
                if let ExprNode::BinOp { lhs, rhs, .. } =
                    expr_arena.get(unwrap_parens(*value, expr_arena))
                {
                    count_owned_compare_temps(module, *lhs, *rhs)
                } else {
                    0
                }
            }
            _ => 0,
        },
        // A statement-level call whose argument is a runtime-string result needs the
        // rtc call's temps plus a copy temp for the owned BSTR.
        ExprNode::CallStmt { callee, args } => {
            let args_node = match expr_arena.get(*callee) {
                ExprNode::Call { args: inner, .. } => *inner,
                _ => *args,
            };
            count_call_owned_temps(module, args_node, expr_arena)
        }
        ExprNode::Call { args, .. } => count_call_owned_temps(module, *args, expr_arena),
        ExprNode::Block { stmts } => stmts.iter().map(|&id| c(id)).sum(),
        ExprNode::If { then_body, else_body, .. } => c(*then_body) + else_body.map(c).unwrap_or(0),
        ExprNode::While { body, .. } | ExprNode::Do { body, .. } | ExprNode::For { body, .. } => c(*body),
        ExprNode::SelectCase { cases, .. } => cases.iter().map(|&id| c(id)).sum(),
        ExprNode::CaseBlock { body, .. } | ExprNode::CaseElse { body } => c(*body),
        _ => 0,
    }
}

/// Count the hidden 16-byte temps a call's runtime-string arguments need (the rtc
/// call's own temps; the 4-byte owned-copy temp is counted separately).
pub(super) fn count_call_owned_temps(module: &BoundModule, args_node: NodeId, expr_arena: &ExprArena) -> usize {
    let args = match expr_arena.get(args_node) {
        ExprNode::ArgList { args } => args,
        _ => return 0,
    };
    args.iter()
        .map(|&a| match module.builtins.get(&a.0) {
            Some(BuiltinCall::RtcString { args: sig }) => rtc_call_temp_count(sig),
            _ => 0,
        })
        .sum()
}

/// Count the 4-byte owned-copy temps needed: one per runtime-string call argument.
pub(super) fn count_owned_copy_temps(module: &BoundModule, node_id: NodeId, expr_arena: &ExprArena) -> usize {
    let c = |id: NodeId| count_owned_copy_temps(module, id, expr_arena);
    let count_args = |args_node: NodeId| -> usize {
        match expr_arena.get(args_node) {
            ExprNode::ArgList { args } => args
                .iter()
                .filter(|&&a| matches!(module.builtins.get(&a.0), Some(BuiltinCall::RtcString { .. })))
                .count(),
            _ => 0,
        }
    };
    match expr_arena.get(node_id) {
        ExprNode::CallStmt { callee, args } => {
            let args_node = match expr_arena.get(*callee) {
                ExprNode::Call { args: inner, .. } => *inner,
                _ => *args,
            };
            count_args(args_node)
        }
        ExprNode::Call { args, .. } => count_args(*args),
        ExprNode::Block { stmts } => stmts.iter().map(|&id| c(id)).sum(),
        ExprNode::If { then_body, else_body, .. } => c(*then_body) + else_body.map(c).unwrap_or(0),
        ExprNode::While { body, .. } | ExprNode::Do { body, .. } | ExprNode::For { body, .. } => c(*body),
        ExprNode::SelectCase { cases, .. } => cases.iter().map(|&id| c(id)).sum(),
        ExprNode::CaseBlock { body, .. } | ExprNode::CaseElse { body } => c(*body),
        _ => 0,
    }
}

/// Allocate the next 4-byte owned-string call-argument copy temp.
pub(super) fn alloc_owned_copy_temp(ctx: &LowerCtx) -> i16 {
    let i = ctx.owned_copy_next.get();
    ctx.owned_copy_next.set(i + 1);
    ctx.local_slots[ctx.owned_copy_base + i].frame_offset
}

/// True if `base` (a `MemberAccess`'s qualifier) resolves to a local whose
/// type is a known external class (`Dim o As New ClassName`) rather than a
/// same-module `Type` (UDT) — the two route through entirely different
/// mechanisms (vtable dispatch vs. a flat frame offset).
pub(super) fn member_access_base_is_class(module: &BoundModule, base: NodeId) -> bool {
    match module.resolutions.get(&base.0) {
        Some(NameResolution::Local { proc_idx, local_idx }) => {
            match module.procs.get(*proc_idx).and_then(|p| p.locals.get(*local_idx)) {
                Some(v) => match &v.vba_type {
                    VbaType::UserDefined(sym) => module.class_field_info.contains_key(sym),
                    _ => false,
                },
                None => false,
            }
        }
        _ => false,
    }
}

/// Count the hidden 4-byte temps a proc's class-member Get accesses need
/// (`x = o.F`, field or property alike): the vtable Get call writes its
/// result through an out-parameter address into a temp, which is then
/// loaded as the expression's value. Only the `Assign` RHS position is
/// scanned — a class member access nested inside a larger expression is
/// gated elsewhere (see `resolve_class_field`).
pub(super) fn count_class_get_temps(module: &BoundModule, node_id: NodeId, expr_arena: &ExprArena) -> usize {
    let c = |id: NodeId| count_class_get_temps(module, id, expr_arena);
    match expr_arena.get(node_id) {
        ExprNode::Assign { value, .. } => match expr_arena.get(*value) {
            ExprNode::MemberAccess { base, .. } => member_access_base_is_class(module, *base) as usize,
            _ => 0,
        },
        ExprNode::Block { stmts } => stmts.iter().map(|&id| c(id)).sum(),
        ExprNode::If { then_body, else_body, .. } => {
            c(*then_body) + else_body.map(c).unwrap_or(0)
        }
        ExprNode::While { body, .. } | ExprNode::Do { body, .. } | ExprNode::For { body, .. } => {
            c(*body)
        }
        ExprNode::SelectCase { cases, .. } => cases.iter().map(|&id| c(id)).sum(),
        ExprNode::CaseBlock { body, .. } | ExprNode::CaseElse { body } => c(*body),
        _ => 0,
    }
}

/// Count the hidden 4-byte temps a proc's class-member Property-Let writes
/// need (`o.P = v`): the vtable Let call's argument is staged into an
/// addressable frame slot (`0x59 <offset>`) before the call — unlike a plain
/// field store, which passes the value directly. Only triggers when the
/// assignment TARGET's specific resolved class member (see
/// Whether a class-method argument at `arg_id` needs materializing into a
/// pool temp before the call, given its parameter's declared type/mode.
/// Oracle-confirmed via three probes this pass (`argtype_probe`,
/// `argmix_probe`, `argbyval_lit_probe` — see the `vb6-class-vtable-slot-
/// rule` memory note), refining the earlier "every argument always stages"
/// finding (which only ever tested all-literal calls):
/// - ByRef + a plain same-type local variable → its own address is pushed
///   directly (`04 <offset>`), no staging (`argtype_probe`'s `TakeInt`/
///   `TakeStr`, `argmix_probe`'s first argument).
/// - ByVal + a plain same-type local variable, scalar type → its own value
///   is loaded directly (a sized value load), no staging (`argtype_probe`'s
///   `TakeIntByVal`).
/// - ByVal + a literal/expression, scalar type → the value is pushed
///   directly, no staging (`argbyval_lit_probe`: a ByVal literal emits no
///   `0x59` at all, unlike the ByRef-literal case).
/// - ByRef + a literal/expression (not independently addressable) → staged
///   (`0x59`) — the original `argcount_probe`/`funcarg_probe` finding, now
///   understood to apply only to this narrower case.
/// - `Object` type, ANY mode → ALWAYS staged (`fd 9c`, refcount safety) even
///   when the source is a plain variable (`argtype_probe`'s `TakeObj(y)`,
///   consistent with `Set`'s own always-stages behavior).
pub(super) fn class_method_arg_needs_staging(
    module: &BoundModule,
    arg_id: NodeId,
    param_ty: &VbaType,
    by_val: bool,
    expr_arena: &ExprArena,
) -> bool {
    if matches!(param_ty, VbaType::Object) {
        return true;
    }
    let is_plain_same_type_var = matches!(expr_arena.get(arg_id), ExprNode::NameRef { .. })
        && matches!(module.resolutions.get(&arg_id.0), Some(NameResolution::Local { .. }))
        && module.types.get(&arg_id.0) == Some(param_ty);
    if is_plain_same_type_var {
        return false;
    }
    // A non-addressable source (literal/expression): ByRef needs a temp to
    // point to; ByVal just pushes the value, no address needed at all.
    !by_val
}

/// The set of class-method parameter types whose staging/addressing
/// convention is actually grounded (see `class_method_arg_needs_staging`'s
/// doc comment) — `Integer`/`Long` (both modes), `Object` (both modes,
/// always-staged), and `String` (both modes: ByRef via `argtype_probe`'s
/// `TakeStr`, a plain address, no staging, matching the type-agnostic ByRef
/// rule; ByVal via a dedicated `argstrbyval_probe` oracle capture —
/// `re_lab/pcode_lab/argstrbyval_probe/`, compiled directly with VB6.EXE and
/// read via `capture_pcode.extract_pcode`, no TTD needed — showing `o.
/// TakeStrByVal s` emits a PLAIN VALUE LOAD (`6c <offset>`, matching
/// `emit_sized_value_load`'s existing 4-byte-size fallback exactly) with NO
/// staging at all — overturning the earlier caution that String, being
/// refcounted like `Object`, might need `Object`-style `fd 9c` staging; the
/// real compiler does not do that for a plain-variable ByVal String
/// argument). `Variant` (both modes, plain-variable source ONLY): ByRef via
/// `argvariant_probe`'s `TakeVar v` — plain `LdAddr`, no staging, matching
/// the type-agnostic ByRef rule; ByVal via the SAME probe's `TakeVarByVal
/// v` capture, fully decoded (not merely observed) — the complete byte
/// sequence is `fc ed <offset:i16 LE>`, exactly the same "opcode(s) + 2-byte
/// offset" shape `emit_sized_value_load` already models for every other
/// size, now with an explicit `16 =>` arm rather than silently falling
/// through to the wrong `0x6c` default. What was previously described as an
/// "unexplored opcode" turned out to have no further structure to explore —
/// it's a complete, self-contained 4-byte sequence, not a prefix of
/// something larger, confirmed by the next bytes in the same oracle capture
/// being the START of the following call's own address-push. A non-
/// addressable Variant argument (literal/expression) is still NOT covered
/// for EITHER mode — VB6's Variant-boxing machinery for that case is
/// unverified, so `lower_class_method_call` additionally requires a plain-
/// variable source whenever a parameter is `Variant` (checked separately,
/// since this function only sees the type/mode, not the argument
/// expression). Anything else (Single/Double/Currency/Byte/
/// Boolean/UDT/Array) is gated too: the addressability mechanism plausibly
/// generalizes to the scalars, but is not oracle-confirmed, and
/// `emit_sized_value_load`'s existing 8-byte case is itself
/// known to be imprecise (hardcodes `0x6d`/Currency's opcode for ALL 8-byte
/// types, including Double, which is actually `0x6f`) — extending onto that
/// uncertainty would trade one gate for a hidden one.
pub(super) fn class_method_param_is_grounded(ty: &VbaType, _by_val: bool) -> bool {
    matches!(
        ty,
        VbaType::Integer | VbaType::Long | VbaType::Object | VbaType::String | VbaType::Variant
    )
}

/// Whether `func_id` (a call target expression) is a class-member
/// `Sub`/`Function` MemberAccess resolved for vtable dispatch, and if so, how
/// many argument-staging temps the call at `args_id` needs — only arguments
/// that actually need materializing (see `class_method_arg_needs_staging`)
/// consume a pool slot; a plain-variable argument needing no staging costs
/// nothing. `is_value` adds one more temp for the result (an expression-
/// position call's out-param temp) — a statement-position call (result
/// discarded) needs none.
fn class_method_call_temps(
    module: &BoundModule,
    func_id: NodeId,
    args_id: NodeId,
    expr_arena: &ExprArena,
    is_value: bool,
) -> usize {
    let ExprNode::MemberAccess { base, .. } = expr_arena.get(func_id) else {
        return 0;
    };
    if !member_access_base_is_class(module, *base) {
        return 0;
    }
    let Some(resolved) = module.class_member_slots.get(&func_id.0) else {
        return 0;
    };
    if resolved.method_slot.is_none() {
        return 0;
    }
    let args: Vec<NodeId> = match expr_arena.get(args_id) {
        ExprNode::ArgList { args } => args.clone(),
        _ => Vec::new(),
    };
    if args.len() != resolved.method_params.len() {
        return 0;
    }
    let staged = args
        .iter()
        .zip(&resolved.method_params)
        .filter(|(&arg, (ty, by_val))| {
            class_method_arg_needs_staging(module, arg, ty, *by_val, expr_arena)
        })
        .count();
    staged + is_value as usize
}

/// Max class-method-call argument-staging temps needed by any single call
/// site in the proc (the pool is reused per call, not accumulated — see
/// `class_member_base`), scanning both statement-position (`CallStmt`, no
/// result temp) and expression-position (`Assign` RHS, needs a result temp)
/// call shapes.
pub(super) fn max_class_method_temps(module: &BoundModule, node_id: NodeId, expr_arena: &ExprArena) -> usize {
    let m = |id: NodeId| max_class_method_temps(module, id, expr_arena);
    match expr_arena.get(node_id) {
        ExprNode::CallStmt { callee, args } => {
            let (func_id, args_id) = match expr_arena.get(*callee) {
                ExprNode::Call { func, args: inner } => (*func, *inner),
                _ => (*callee, *args),
            };
            class_method_call_temps(module, func_id, args_id, expr_arena, false)
        }
        ExprNode::Assign { value, .. } => match expr_arena.get(*value) {
            ExprNode::Call { func, args } => {
                class_method_call_temps(module, *func, *args, expr_arena, true)
            }
            _ => 0,
        },
        ExprNode::Block { stmts } => stmts.iter().map(|&id| m(id)).max().unwrap_or(0),
        ExprNode::If { then_body, else_body, .. } => {
            m(*then_body).max(else_body.map(m).unwrap_or(0))
        }
        ExprNode::While { body, .. } | ExprNode::Do { body, .. } | ExprNode::For { body, .. } => {
            m(*body)
        }
        ExprNode::SelectCase { cases, .. } => cases.iter().map(|&id| m(id)).max().unwrap_or(0),
        ExprNode::CaseBlock { body, .. } | ExprNode::CaseElse { body } => m(*body),
        _ => 0,
    }
}

pub(super) fn count_class_let_temps(module: &BoundModule, node_id: NodeId, expr_arena: &ExprArena) -> usize {
    let c = |id: NodeId| count_class_let_temps(module, id, expr_arena);
    match expr_arena.get(node_id) {
        ExprNode::Assign { target, .. } => match expr_arena.get(*target) {
            ExprNode::MemberAccess { base, .. } if member_access_base_is_class(module, *base) => {
                module
                    .class_member_slots
                    .get(&target.0)
                    .map(|r| r.is_property)
                    .unwrap_or(false) as usize
            }
            _ => 0,
        },
        // `Set o.P = v`: the Property-Set call's argument is staged into the
        // same reusable class-member temp as Let (`fd 9c <offset>` instead of
        // `0x59 <offset>`) — needs the slot reserved just the same.
        ExprNode::SetAssign { target, .. } => match expr_arena.get(*target) {
            ExprNode::MemberAccess { base, .. } if member_access_base_is_class(module, *base) => {
                module
                    .class_member_slots
                    .get(&target.0)
                    .map(|r| r.set_slot.is_some())
                    .unwrap_or(false) as usize
            }
            _ => 0,
        },
        ExprNode::Block { stmts } => stmts.iter().map(|&id| c(id)).sum(),
        ExprNode::If { then_body, else_body, .. } => {
            c(*then_body) + else_body.map(c).unwrap_or(0)
        }
        ExprNode::While { body, .. } | ExprNode::Do { body, .. } | ExprNode::For { body, .. } => {
            c(*body)
        }
        ExprNode::SelectCase { cases, .. } => cases.iter().map(|&id| c(id)).sum(),
        ExprNode::CaseBlock { body, .. } | ExprNode::CaseElse { body } => c(*body),
        _ => 0,
    }
}

pub(super) fn count_variant_assigns(module: &BoundModule, node_id: NodeId, expr_arena: &ExprArena) -> usize {
    let c = |id: NodeId| count_variant_assigns(module, id, expr_arena);
    match expr_arena.get(node_id) {
        ExprNode::Assign { target, .. } => {
            matches!(module.types.get(&target.0), Some(VbaType::Variant)) as usize
        }
        ExprNode::Block { stmts } => stmts.iter().map(|&id| c(id)).sum(),
        ExprNode::If { then_body, else_body, .. } => {
            c(*then_body) + else_body.map(c).unwrap_or(0)
        }
        ExprNode::While { body, .. } | ExprNode::Do { body, .. } | ExprNode::For { body, .. } => {
            c(*body)
        }
        ExprNode::SelectCase { cases, .. } => cases.iter().map(|&id| c(id)).sum(),
        ExprNode::CaseBlock { body, .. } | ExprNode::CaseElse { body } => c(*body),
        _ => 0,
    }
}

/// If `func_id` resolves to a local array, return its LdAddr frame offset, element
/// type, and declared dimension count.
pub(super) fn array_local_info(ctx: &LowerCtx, func_id: NodeId) -> Option<(i16, VbaType, u16)> {
    if let Some(NameResolution::Local { local_idx, .. }) = ctx.module.resolutions.get(&func_id.0) {
        let var = &ctx.proc.locals[*local_idx];
        if let VbaType::Array(elem) = &var.vba_type {
            let dims = var.array_dims.unwrap_or(1);
            return Some((ctx.local_slots[*local_idx].frame_offset, (**elem).clone(), dims));
        }
    }
    None
}

/// The single subscript index of a 1-D array `Call`'s argument list.
pub(super) fn single_index(args_id: NodeId, expr_arena: &ExprArena) -> Option<NodeId> {
    if let ExprNode::ArgList { args } = expr_arena.get(args_id) {
        if args.len() == 1 {
            return Some(args[0]);
        }
    }
    None
}

/// Byte size of a Static local within its procedure's static block.
pub(super) fn static_var_size(ty: &VbaType) -> u16 {
    match ty {
        VbaType::Byte => 1,
        VbaType::Integer | VbaType::Boolean => 2,
        VbaType::Long | VbaType::Single | VbaType::String => 4,
        VbaType::Double | VbaType::Currency | VbaType::Date => 8,
        VbaType::Variant => 16,
        _ => 4,
    }
}

/// Store-access opcode bytes for a Static local of the given type. The numeric/
/// String classes use a 1- or 2-byte opcode through the per-procedure static block.
pub(super) fn static_store_op(ty: &VbaType) -> Option<Vec<u8>> {
    Some(match ty {
        VbaType::Integer | VbaType::Boolean => vec![0x8e],
        VbaType::Long => vec![0x8f],
        VbaType::Single => vec![0x91],
        VbaType::Currency => vec![0x90],
        VbaType::Double | VbaType::Date => vec![0x92],
        VbaType::Byte => vec![0xfd, 0x80],
        VbaType::String => vec![0xfd, 0x91],
        _ => return None,
    })
}

/// Load-access opcode bytes for a Static local of the given type.
pub(super) fn static_load_op(ty: &VbaType) -> Option<Vec<u8>> {
    Some(match ty {
        VbaType::Integer | VbaType::Boolean => vec![0x89],
        VbaType::Long | VbaType::String => vec![0x8a],
        VbaType::Currency => vec![0x8b],
        VbaType::Single => vec![0x8c],
        VbaType::Double | VbaType::Date => vec![0x8d],
        VbaType::Byte => vec![0xfd, 0x70],
        _ => return None,
    })
}

/// Emit a per-procedure static-block access: `0x5f <module_desc> 0x0004 <op>
/// <var_offset>`. The `0x0004` is the module field holding this single
/// procedure's static-block handle.
pub(super) fn emit_static_access(ctx: &LowerCtx, out: &mut Vec<u8>, op: &[u8], var_off: u16) {
    out.push(0x5f);
    out.extend_from_slice(&ctx.module_desc.to_le_bytes());
    out.extend_from_slice(&0x0004u16.to_le_bytes());
    out.extend_from_slice(op);
    out.extend_from_slice(&var_off.to_le_bytes());
}

/// Total pushed byte size of a ByVal argument of `ty`: rounded up to a multiple
/// of 4 (Integer/Byte → 4), except 8- and 16-byte types keep their size.
pub(super) fn call_arg_bytes(ty: &VbaType) -> u16 {
    let s = static_var_size(ty);
    if s <= 4 { 4 } else { s }
}

/// The OLE Automation VARTYPE used to tag a value boxed into a runtime-call temp
/// (the byte before the fixed `0x40` marker). `None` for types whose boxed form is
/// not yet confirmed (gated rather than guessed).
pub(super) fn vba_type_to_vartype(ty: &VbaType) -> Option<u8> {
    Some(match ty {
        VbaType::Integer => 2,
        VbaType::Long => 3,
        VbaType::Single => 4,
        VbaType::Double => 5,
        VbaType::Currency => 6,
        VbaType::Date => 7,
        VbaType::String => 8,
        VbaType::Boolean => 11,
        VbaType::Byte => 17,
        _ => return None,
    })
}

/// Emit a size-based value load (load `size` bytes from a frame offset), the form
/// used to push a same-typed ByVal argument: 1→0xfc 0xe0, 2→0x6b, 4→0x6c, 8→0x6d,
/// 16→0xfc 0xed (Variant — oracle-captured via `argvariant_probe`'s
/// `TakeVarByVal v`, `re_lab/pcode_lab/argvariant_probe/`: the full byte
/// sequence is `fc ed <offset:i16 LE>`, the same "opcode(s) + 2-byte offset"
/// shape as every other case here, not a special form).
pub(super) fn emit_sized_value_load(size: u16, off: i16, buf: &mut Vec<u8>) {
    match size {
        1 => buf.extend_from_slice(&[0xfc, 0xe0]),
        2 => buf.push(0x6b),
        8 => buf.push(0x6d),
        16 => buf.extend_from_slice(&[0xfc, 0xed]),
        _ => buf.push(0x6c),
    }
    buf.extend_from_slice(&off.to_le_bytes());
}

/// Frame offset of a local-variable argument (for a ByRef argument's LdAddr).
pub(super) fn arg_var_offset(ctx: &LowerCtx, arg_id: NodeId) -> Option<i16> {
    match ctx.module.resolutions.get(&arg_id.0) {
        Some(NameResolution::Local { local_idx, .. }) => Some(ctx.local_slots[*local_idx].frame_offset),
        _ => None,
    }
}

/// Number of hidden 16-byte temps a String-returning runtime call needs: one per
/// boxed argument, two per omitted-Optional argument, plus the result temp.
pub(super) fn rtc_call_temp_count(sig: &[RtcArg]) -> usize {
    sig.iter()
        .map(|a| match a {
            RtcArg::Boxed => 1,
            RtcArg::MissingVariant => 2,
            RtcArg::ByVal(_) => 0,
        })
        .sum::<usize>()
        + 1
}

/// Allocate the next hidden 16-byte string/variant temp (from the `string_rtc`
/// pool) and return its frame offset.
pub(super) fn alloc_string_rtc_temp(ctx: &LowerCtx) -> i16 {
    let i = ctx.string_rtc_next.get();
    ctx.string_rtc_next.set(i + 1);
    ctx.local_slots[ctx.string_rtc_base + i].frame_offset
}

/// Emit the variant-temp free for a set of owned temps: a single temp uses the
/// short free (0x35); two or more use the combined free (0x36 <count*2> <offsets…>),
/// offsets in allocation order.
pub(super) fn emit_variant_temp_free(out: &mut Vec<u8>, offsets: &[i16]) {
    if offsets.len() == 1 {
        out.push(0x35);
        out.extend_from_slice(&offsets[0].to_le_bytes());
    } else {
        out.push(0x36);
        out.extend_from_slice(&((offsets.len() * 2) as u16).to_le_bytes());
        for off in offsets {
            out.extend_from_slice(&off.to_le_bytes());
        }
    }
}

/// Reserve the hidden temps for one String-returning runtime call, returning the
/// base slot index (relative to `string_rtc_base`). Lets a caller pre-allocate an
/// rtc operand's temps before emitting unrelated operands (the compiler assigns all
/// of an expression's rtc-call temps ahead of its materialization/result temps).
pub(super) fn reserve_rtc_temps(ctx: &LowerCtx, sig: &[RtcArg]) -> usize {
    let base = ctx.string_rtc_next.get();
    ctx.string_rtc_next.set(base + rtc_call_temp_count(sig));
    base
}

/// True if `node` is a `&` concatenation at least one of whose operands is a
/// String-returning runtime intrinsic call — an owned-temp concat, which uses the
/// cleanup-aware concat opcode (0xfb 0xef) over the plain one (0x2a).
pub(super) fn is_owned_concat(module: &BoundModule, node: NodeId, expr_arena: &ExprArena) -> bool {
    if !matches!(expr_arena.get(node), ExprNode::BinOp { op: BinOpKind::Cat, .. }) {
        return false;
    }
    let mut ops = Vec::new();
    flatten_concat(node, expr_arena, &mut ops);
    ops.iter()
        .any(|&o| matches!(module.builtins.get(&o.0), Some(BuiltinCall::RtcString { .. })))
}

/// Number of hidden 16-byte temps an owned-temp concat needs: each runtime-string
/// operand contributes its call's temps, each plain String var or string literal
/// contributes one materialization temp, plus one temp per intermediate concat.
pub(super) fn count_owned_concat_temps(module: &BoundModule, ops: &[NodeId]) -> usize {
    let mut n = 0;
    for &op in ops {
        match module.builtins.get(&op.0) {
            Some(BuiltinCall::RtcString { args }) => n += rtc_call_temp_count(args),
            _ => n += 1,
        }
    }
    n + ops.len().saturating_sub(1)
}

/// True if `node` is `Len(<runtime-string call>)` — a `Len` whose argument is a
/// freshly-produced runtime-string result, which uses the release-aware Len opcode
/// (0xfb 0xeb) over an owned BSTR temp.
pub(super) fn is_owned_len(module: &BoundModule, node: NodeId, expr_arena: &ExprArena) -> bool {
    if !matches!(
        module.builtins.get(&node.0),
        Some(BuiltinCall::Unary(UnaryIntrinsic::Len))
    ) {
        return false;
    }
    if let ExprNode::Call { args, .. } = expr_arena.get(node) {
        if let Some(arg) = single_index(*args, expr_arena) {
            return matches!(module.builtins.get(&arg.0), Some(BuiltinCall::RtcString { .. }));
        }
    }
    false
}

/// Follow `Paren` wrappers to the inner expression node.
pub(super) fn unwrap_parens(mut node: NodeId, expr_arena: &ExprArena) -> NodeId {
    while let ExprNode::Paren { inner } = expr_arena.get(node) {
        node = *inner;
    }
    node
}

/// The release-aware string-comparison opcode (`0xfb <op>`) for an owned-operand
/// comparison — one less than the plain string-compare second byte for each operator.
pub(super) fn owned_string_compare_op(op: BinOpKind) -> Option<u8> {
    Some(match op {
        BinOpKind::Eq => 0x2f,
        BinOpKind::Ne => 0x3c,
        BinOpKind::Lt => 0x63,
        BinOpKind::Gt => 0x70,
        BinOpKind::Le => 0x49,
        BinOpKind::Ge => 0x56,
        _ => return None,
    })
}

/// True if `node` is a String relational comparison at least one of whose operands
/// is a runtime-string result — an owned-operand compare, which uses the
/// compare-into-temp opcode family wrapped by the `0x5d`/`0x55` owned markers.
pub(super) fn is_owned_compare(module: &BoundModule, node: NodeId, expr_arena: &ExprArena) -> bool {
    if let ExprNode::BinOp { op, lhs, rhs } = expr_arena.get(node) {
        if owned_string_compare_op(*op).is_some()
            && matches!(module.types.get(&lhs.0), Some(VbaType::String))
            && matches!(module.types.get(&rhs.0), Some(VbaType::String))
        {
            return matches!(module.builtins.get(&lhs.0), Some(BuiltinCall::RtcString { .. }))
                || matches!(module.builtins.get(&rhs.0), Some(BuiltinCall::RtcString { .. }));
        }
    }
    false
}

/// Number of hidden 16-byte temps an owned-operand compare needs: each operand's
/// materialization temps (an rtc call's temps, or one copy/literal temp) plus one
/// compare temp.
pub(super) fn count_owned_compare_temps(module: &BoundModule, lhs: NodeId, rhs: NodeId) -> usize {
    let one = |op: NodeId| match module.builtins.get(&op.0) {
        Some(BuiltinCall::RtcString { args }) => rtc_call_temp_count(args),
        _ => 1,
    };
    one(lhs) + one(rhs) + 1
}
