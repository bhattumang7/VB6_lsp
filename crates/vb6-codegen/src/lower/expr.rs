use super::*;
use super::decl::*;
use super::intrinsics::*;
use super::stmt::*;


// ── Expression lowering ───────────────────────────────────────────────────────

/// Returns the "wider" of two VBA types for numeric literal promotion.
/// VB6 widens integer literals to match the type of the wider operand.
pub(super) fn wider_numeric_tag(a: Option<&VbaType>, b: Option<&VbaType>) -> Option<u16> {
    // VB6 numeric promotion order, matching vb6-sema's `numeric_promote`.
    fn rank(t: &VbaType) -> u8 {
        match t {
            VbaType::Boolean => 0,
            VbaType::Byte    => 1,
            VbaType::Integer => 2,
            VbaType::Long    => 3,
            VbaType::Single  => 4,
            VbaType::Double | VbaType::Date => 5,
            VbaType::Currency => 6,
            _ => 7,
        }
    }
    let ta = a.map(rank).unwrap_or(0);
    let tb = b.map(rank).unwrap_or(0);
    let wider = if ta >= tb { a } else { b };
    // Date operands are computed in Double (the runtime widens the OLE serial to a
    // Double; the result is converted back to Date on store).
    wider.and_then(|t| {
        vba_type_to_node_tag(if matches!(t, VbaType::Date) { &VbaType::Double } else { t })
    })
}

pub(super) fn lower_expr(
    ctx: &LowerCtx,
    node_id: NodeId,
    expr_arena: &ExprArena,
    arena: &mut NodeArena,
) -> Result<NodeRef, LowerError> {
    lower_expr_coerced(ctx, node_id, expr_arena, arena, None)
}

/// Like `lower_expr` but when `coerce_tag` is `Some(tag)`, integer literals are
/// emitted with that type tag instead of their natural type.  This implements
/// VB6's implicit widening of integer literals to match their context type.
pub(super) fn lower_expr_coerced(
    ctx: &LowerCtx,
    node_id: NodeId,
    expr_arena: &ExprArena,
    arena: &mut NodeArena,
    coerce_tag: Option<u16>,
) -> Result<NodeRef, LowerError> {
    match expr_arena.get(node_id) {
        ExprNode::Literal { lit } => {
            // A string literal interns into the module string pool and emits
            // `0x1b <pool index>` (synthetic node 0x79); other literals fold in place.
            // Each string-literal reference also consumes one per-procedure
            // reference slot (the same counter a call's 2-byte operand reports).
            if let AstLit::Str(s) = lit {
                let idx = ctx.intern_string(s);
                ctx.call_next.set(ctx.call_next.get() + 1);
                return Ok(arena.alloc(NodeArena::node(0x79, 0x10, idx as u32, 0, 0, 0)));
            }
            lower_lit_coerced(lit, coerce_tag, arena)
        }
        ExprNode::NameRef { .. } => lower_name_ref(ctx, node_id, arena),
        ExprNode::MemberAccess { base, member, bang } => {
            let (base, member, bang) = (*base, *member, *bang);
            if member_access_base_is_class(ctx.module, base) {
                lower_class_field_get(ctx, base, node_id, bang, arena)
            } else {
                lower_udt_field_access(ctx, base, member, bang, arena)
            }
        }
        ExprNode::BinOp { op, lhs, rhs } => {
            let (op, lhs_id, rhs_id) = (*op, *lhs, *rhs);
            lower_binop(ctx, node_id, op, lhs_id, rhs_id, expr_arena, arena)
        }
        ExprNode::UnOp { op, operand } => {
            let (op, operand_id) = (*op, *operand);
            // VB6 constant-folds a negated integer literal (`-5`) into a single
            // push of the negated value (coerced to the context type), rather than
            // push-then-negate.
            if matches!(op, UnOpKind::Neg) {
                if let ExprNode::Literal { lit: AstLit::Int(v) } = expr_arena.get(operand_id) {
                    return lower_lit_coerced(&AstLit::Int(-*v), coerce_tag, arena);
                }
            }
            lower_unop(ctx, node_id, op, operand_id, expr_arena, arena)
        }
        ExprNode::Paren { inner } => {
            let inner_id = *inner;
            lower_expr_coerced(ctx, inner_id, expr_arena, arena, coerce_tag)
        }
        // A type-conversion intrinsic (CInt/CLng/CSng/CDbl/CCur/CStr): convert its
        // single argument to the target type via the explicit-conversion node
        // (0x7c), so it composes inside any expression. Sources and destinations
        // outside the supported scalar set (Byte, Boolean/Date/Variant) are gated.
        ExprNode::Call { func, args } => {
            let (func, args_id) = (*func, *args);
            // A user Function call used as an expression operand: emit the call
            // (0x5e, result left on the stack) into a byte blob wrapped in a 0x7e
            // node, so it composes with the enclosing operator.
            if ctx.module.builtins.get(&node_id.0).is_none()
                && matches!(ctx.module.resolutions.get(&func.0), Some(NameResolution::Proc(_)))
            {
                let mut bytes = Vec::new();
                lower_call(ctx, func, args_id, true, expr_arena, &mut bytes)?;
                let off = arena.alloc_blob(&bytes);
                let ret = ctx
                    .module
                    .types
                    .get(&node_id.0)
                    .and_then(vba_type_to_node_tag)
                    .unwrap_or(8);
                return Ok(arena.alloc(NodeArena::node(0x7e, ret, off, bytes.len() as u32, 0, 0)));
            }
            // `InStr` (2- or 3-argument): a dedicated opcode (`fe fd`) with four
            // operands pushed in order — start (Long), string1, string2, compare-mode
            // (Long) — leaving a Long result on the stack (blob node 0x7e). An omitted
            // leading start defaults to literal 1; the compare-mode is literal 0
            // (Option Compare Binary).
            if let Some(BuiltinCall::Instr { three_arg }) = ctx.module.builtins.get(&node_id.0) {
                let three_arg = *three_arg;
                let arg_ids: Vec<NodeId> = match expr_arena.get(args_id) {
                    ExprNode::ArgList { args } => args.clone(),
                    _ => return Err(LowerError::UnsupportedNode),
                };
                if arg_ids.len() != if three_arg { 3 } else { 2 } {
                    return Err(LowerError::UnsupportedNode);
                }
                let (start, s1, s2) = if three_arg {
                    (Some(arg_ids[0]), arg_ids[1], arg_ids[2])
                } else {
                    (None, arg_ids[0], arg_ids[1])
                };
                // The two searched operands must be Strings.
                for s in [s1, s2] {
                    if !matches!(ctx.module.types.get(&s.0), Some(VbaType::String)) {
                        return Err(LowerError::UnsupportedNode);
                    }
                }
                let mut bytes = Vec::new();
                match start {
                    Some(st) => bytes.extend(lower_expr_to_bytes_coerced(ctx, st, expr_arena, Some(8))?),
                    None => bytes.extend_from_slice(&[0xf5, 0x01, 0x00, 0x00, 0x00]),
                }
                bytes.extend(lower_expr_to_bytes(ctx, s1, expr_arena)?);
                bytes.extend(lower_expr_to_bytes(ctx, s2, expr_arena)?);
                bytes.extend_from_slice(&[0xf5, 0x00, 0x00, 0x00, 0x00]);
                bytes.extend_from_slice(&[0xfe, 0xfd]);
                let off = arena.alloc_blob(&bytes);
                return Ok(arena.alloc(NodeArena::node(0x7e, 8, off, bytes.len() as u32, 0, 0)));
            }
            let arg = single_index(args_id, expr_arena).ok_or(LowerError::UnsupportedNode)?;
            match ctx.module.builtins.get(&node_id.0) {
                // Type-conversion intrinsic → explicit-conversion node (0x7c).
                Some(BuiltinCall::Convert(t)) => {
                    let dest = vba_type_to_node_tag(t).ok_or(LowerError::UnsupportedType)?;
                    let src = ctx
                        .module
                        .types
                        .get(&arg.0)
                        .and_then(vba_type_to_node_tag)
                        .ok_or(LowerError::UnsupportedType)?;
                    const OK: &[u16] = &[6, 8, 0xa, 0xb, 0xd, 0x10];
                    if !OK.contains(&dest) || !matches!(src, 6 | 8 | 0xa | 0xb | 0xd) {
                        return Err(LowerError::UnsupportedNode);
                    }
                    let arg_root = lower_expr_coerced(ctx, arg, expr_arena, arena, None)?;
                    Ok(arena.alloc(NodeArena::node(0x7c, dest, arg_root.0, 0, 0, 0)))
                }
                // Dedicated-opcode unary intrinsic → node 0x7d (kind in w[5]); the
                // opcode is selected by the argument type at emit time. The node's
                // type tag is the intrinsic's result type.
                Some(BuiltinCall::Unary(k)) => {
                    let result = ctx
                        .module
                        .types
                        .get(&node_id.0)
                        .and_then(vba_type_to_node_tag)
                        .ok_or(LowerError::UnsupportedType)?;
                    let arg_tag = ctx
                        .module
                        .types
                        .get(&arg.0)
                        .and_then(vba_type_to_node_tag)
                        .ok_or(LowerError::UnsupportedType)?;
                    // Supported argument types: Len takes a String; the numeric
                    // intrinsics take Integer/Long/Single/Double/Currency.
                    let supported = match k {
                        UnaryIntrinsic::Len => arg_tag == 0x10,
                        _ => matches!(arg_tag, 6 | 8 | 0xa | 0xb | 0xd),
                    };
                    if !supported {
                        return Err(LowerError::UnsupportedNode);
                    }
                    let arg_root = lower_expr_coerced(ctx, arg, expr_arena, arena, None)?;
                    let kind = *k as u32;
                    Ok(arena.alloc(NodeArena::node(0x7d, result, arg_root.0, kind, 0, 0)))
                }
                // Single-argument numeric-result runtime call (Asc/Sqr/Val): the
                // argument is size-loaded, then a runtime call whose opcode follows
                // the result type (Integer 0x0b, Double 0x0a) with the per-proc
                // reference index and argument bytes. Result left on the stack
                // (blob node 0x7e). The argument must be a local variable.
                Some(BuiltinCall::RtcNumeric { ret, .. }) => {
                    let ret_tag = vba_type_to_node_tag(ret).ok_or(LowerError::UnsupportedType)?;
                    let opcode: u8 = match ret_tag {
                        6 => 0x0b,
                        0xb => 0x0a,
                        _ => return Err(LowerError::UnsupportedNode),
                    };
                    let off = arg_var_offset(ctx, arg).ok_or(LowerError::UnsupportedNode)?;
                    let arg_ty = ctx.module.types.get(&arg.0).ok_or(LowerError::UnsupportedType)?;
                    let mut bytes = Vec::new();
                    emit_sized_value_load(static_var_size(arg_ty), off, &mut bytes);
                    let r = ctx.call_next.get();
                    ctx.call_next.set(r + 1);
                    bytes.push(opcode);
                    bytes.extend_from_slice(&(r as u16).to_le_bytes());
                    bytes.extend_from_slice(&call_arg_bytes(arg_ty).to_le_bytes());
                    let blob = arena.alloc_blob(&bytes);
                    Ok(arena.alloc(NodeArena::node(0x7e, ret_tag, blob, bytes.len() as u32, 0, 0)))
                }
                // A String-returning runtime intrinsic is handled as an assignment
                // (it writes a result temp); using one as an expression operand is
                // not yet supported.
                Some(BuiltinCall::RtcString { .. }) => Err(LowerError::UnsupportedNode),
                // `InStr` is handled by the dedicated-opcode path above (early return).
                Some(BuiltinCall::Instr { .. }) => Err(LowerError::UnsupportedNode),
                None => Err(LowerError::UnsupportedNode),
            }
        }
        _ => Err(LowerError::UnsupportedNode),
    }
}

pub(super) fn lower_lit(lit: &AstLit, arena: &mut NodeArena) -> Result<NodeRef, LowerError> {
    match lit {
        AstLit::Int(v) => Ok(arena.alloc(NodeArena::node(1, 6, *v as u32, 0, 0, 0))),
        AstLit::Long(v) => Ok(arena.alloc(NodeArena::node(1, 8, *v as u32, 0, 0, 0))),
        AstLit::Bool(v) => {
            let val: u32 = if *v { (-1i32) as u32 } else { 0 };
            Ok(arena.alloc(NodeArena::node(1, 6, val, 0, 0, 0)))
        }
        AstLit::Single(v) => Ok(arena.alloc(NodeArena::node(3, 10, (*v).to_bits(), 0, 0, 0))),
        AstLit::Double(v) => {
            let bits = v.to_bits();
            Ok(arena.alloc(NodeArena::node(3, 11, bits as u32, (bits >> 32) as u32, 0, 0)))
        }
        AstLit::Currency(v) => {
            let bits = *v as u64;
            Ok(arena.alloc(NodeArena::node(2, 0, bits as u32, (bits >> 32) as u32, 0, 0)))
        }
        // A Date literal carries the OLE date serial as an f64; it pushes as an
        // 8-byte literal (the case-2 path) tagged Date (0xc) → push opcode 0xfa.
        AstLit::Date(v) => {
            let bits = v.to_bits();
            Ok(arena.alloc(NodeArena::node(2, 0xc, bits as u32, (bits >> 32) as u32, 0, 0)))
        }
        _ => Err(LowerError::UnsupportedNode),
    }
}

/// Like `lower_lit` but promotes an integer literal to `coerce_tag` when the
/// context type is a wider *integer* type (e.g. an integer literal in a Long
/// expression is pushed directly as Long). Promotion to a floating-point or
/// Currency context is *not* done in place: VB6 pushes the integer literal in its
/// natural type and emits a runtime conversion (handled by the operand/assignment
/// coercion), so for those targets the literal keeps its natural type here.
pub(super) fn lower_lit_coerced(
    lit: &AstLit,
    coerce_tag: Option<u16>,
    arena: &mut NodeArena,
) -> Result<NodeRef, LowerError> {
    match coerce_tag {
        // Integer tags: Byte=5, Integer=6, Long=8.
        Some(tag @ (5 | 6 | 8)) => match lit {
            AstLit::Int(v) => Ok(arena.alloc(NodeArena::node(1, tag, *v as u32, 0, 0, 0))),
            _ => lower_lit(lit, arena),
        },
        _ => lower_lit(lit, arena),
    }
}

pub(super) fn lower_name_ref(
    ctx: &LowerCtx,
    node_id: NodeId,
    arena: &mut NodeArena,
) -> Result<NodeRef, LowerError> {
    let resolution = ctx
        .module
        .resolutions
        .get(&node_id.0)
        .ok_or(LowerError::Unresolved)?;

    match resolution {
        NameResolution::Local { local_idx, .. } => {
            let local = &ctx.proc.locals[*local_idx];
            // A `Const` is folded to its literal value at the use site (it has no
            // frame slot). Integer-valued consts emit an integer literal node.
            if local.is_const {
                // Integer-valued consts emit an integer literal node in the
                // declared type. Non-integer consts (String/Double/etc.) fold the
                // carried literal directly: a String interns into the pool, every
                // other literal lowers in its own type.
                if let Some(v) = local.const_value {
                    let tag = vba_type_to_node_tag(&local.vba_type)
                        .ok_or(LowerError::UnsupportedType)?;
                    return Ok(arena.alloc(NodeArena::node(1, tag, v as u32, 0, 0, 0)));
                }
                let lit = local.const_lit.as_ref().ok_or(LowerError::UnsupportedType)?;
                if let AstLit::Str(s) = lit {
                    let idx = ctx.intern_string(s);
                    ctx.call_next.set(ctx.call_next.get() + 1);
                    return Ok(arena.alloc(NodeArena::node(0x79, 0x10, idx as u32, 0, 0, 0)));
                }
                return lower_lit(lit, arena);
            }
            // A Static local is read from the procedure's static block: synthetic
            // node 0x7b carries the module descriptor (low 16 of w[4]) and the
            // static offset (high 16); the load opcode follows from the type tag.
            if local.is_static {
                if static_load_op(&local.vba_type).is_none() {
                    return Err(LowerError::UnsupportedType);
                }
                let tag = vba_type_to_node_tag(&local.vba_type).ok_or(LowerError::UnsupportedType)?;
                let off = ctx.static_offsets[*local_idx] as u32;
                let packed = (ctx.module_desc as u32) | (off << 16);
                return Ok(arena.alloc(NodeArena::node(0x7b, tag, packed, 0, 0, 0)));
            }
            // Fixed-length strings (`As String * n`) copy via a length-aware
            // sequence (LdAddr + 0x33<len> + store 0x31), not the plain BSTR load;
            // gated until that path is ported.
            if local.fixed_string_len.is_some() {
                return Err(LowerError::UnsupportedType);
            }
            let ty = &local.vba_type;
            let slot = &ctx.local_slots[*local_idx];
            let tag = vba_type_to_node_tag(ty).ok_or(LowerError::UnsupportedType)?;
            let lctx = load_store_ctx(ty).ok_or(LowerError::UnsupportedType)?;
            Ok(build_frame_load_node(arena, 0x74, tag, lctx, slot.frame_offset))
        }
        NameResolution::Param { param_idx, .. } => {
            let ty = ctx.param_type(*param_idx);
            let slot = &ctx.param_slots[*param_idx];
            let tag = vba_type_to_node_tag(ty).ok_or(LowerError::UnsupportedType)?;
            let lctx = load_store_ctx(ty).ok_or(LowerError::UnsupportedType)?;
            let opcode = if slot.byref { 0x75u16 } else { 0x74u16 };
            Ok(build_frame_load_node(arena, opcode, tag, lctx, slot.frame_offset))
        }
        NameResolution::ModuleVar(idx) => {
            let ty = ctx.global_type(*idx);
            let slot = &ctx.global_slots[*idx];
            let lctx = load_store_ctx(ty).ok_or(LowerError::UnsupportedType)?;
            Ok(build_global_load_node(arena, lctx, slot.module_desc, slot.field_offset))
        }
        _ => Err(LowerError::Unresolved),
    }
}

/// A resolved `t.X` field reference: everything both the read path
/// ([`lower_udt_field_access`]) and the assignment-LHS store path
/// (`lower_udt_field_store` in `lower::assign`) need to build the `0x60`
/// reference node and the `Emitter`'s `SymbolContext`.
pub(super) struct UdtFieldRef {
    /// The records-heap context the resolver needs to classify the field
    /// (built the same way a real `Type...End Type` field's declaration
    /// would be — see `crate::decl::build_property_slot_scalar`).
    pub sym: crate::emit::SymbolContext,
    /// The field's combined absolute frame offset: the UDT local's own frame
    /// base plus the field's byte offset within it (uniform-size fields
    /// only — see `crate::bind::UdtLocal`). This is what reaches the
    /// emitted bytecode (`resolver.rs::init_expr_descriptor`'s operand),
    /// not the type size.
    pub offset: i16,
    /// The field's frame size — feeds the `0x60` node's size descriptor,
    /// which only drives `init_expr_descriptor`'s kind selection.
    pub field_size: i16,
    /// The field's VBA6 type tag (`vba_type_to_node_tag`), e.g. `8` for Long.
    pub type_tag: u16,
}

/// Resolve `base.member` to its combined frame offset and a `SymbolContext`
/// classifying it, for a plain local `Type...End Type`-typed `base`.
///
/// Only a directly-resolved local UDT base is supported (module-level /
/// nested / object-typed bases are out of scope for this milestone); anything
/// else is `LowerError::UnsupportedNode`.
pub(super) fn resolve_udt_field(
    ctx: &LowerCtx,
    base: NodeId,
    member: u32,
) -> Result<UdtFieldRef, LowerError> {
    let local_idx = match ctx.module.resolutions.get(&base.0) {
        Some(NameResolution::Local { local_idx, .. }) => *local_idx,
        Some(_) => return Err(LowerError::UnsupportedNode),
        None => return Err(LowerError::Unresolved),
    };
    let VbaType::UserDefined(type_sym) = ctx.local_type(local_idx) else {
        return Err(LowerError::UnsupportedNode);
    };
    let decl = ctx
        .module
        .type_decls
        .iter()
        .find(|d| d.sym_id == *type_sym)
        .ok_or(LowerError::Unresolved)?;
    let field_index = decl
        .members
        .iter()
        .position(|m| m.sym_id == member)
        .ok_or(LowerError::Unresolved)?;
    let type_tag = vba_type_to_node_tag(&decl.members[field_index].vba_type)
        .ok_or(LowerError::UnsupportedType)?;

    let udt = ctx.local_udt(local_idx).ok_or(LowerError::UnsupportedNode)?;
    let offset = udt.field_offset(field_index);
    let field_size = udt.field_size(field_index);

    // Build the field's declaration-compiler record: a scratch method-bag
    // "container" (never read back; it only exists so the field's
    // property-bag record has somewhere real to link under) plus the
    // field's own property-bag record, exactly as `Type...End Type`
    // declaration lowering will build them once that front-end exists.
    let mut heap = crate::heap::HeapContext::new(true, false);
    let parent = heap
        .allocate_method_bag()
        .map_err(|_| LowerError::UnsupportedNode)?;
    let mut tail = crate::heap::NIL;
    let field_rec =
        crate::decl::build_property_slot_scalar(&mut heap, type_tag, 0, parent, &mut tail)
            .map_err(|_| LowerError::UnsupportedNode)?;

    // The binder-supplied (kind, byref) fixed to the plain-scalar convention
    // already proven (unit-tested in resolver_tests/emit_tests) to resolve a
    // property-bag scalar record to a local-style (kind 1/2) descriptor.
    let sym = crate::emit::SymbolContext {
        heap: heap.mem,
        member_off: field_rec as usize,
        ctx_flag_c: 0,
        binding: Some((4, 0)),
    };

    Ok(UdtFieldRef { sym, offset, field_size, type_tag })
}

/// Lower `t.X` (read position): a `0x60` reference node whose `word[7]`
/// carries the field's combined resolved frame offset directly (this front
/// end's own convention — see [`crate::resolver::init_expr_descriptor`]). No
/// qualifier sub-node is built (`word[4]` stays 0): the offset is already
/// fully resolved here, so nothing downstream needs to re-walk `t`.
///
/// Sets `ctx.member_symbol` so the statement-level caller can attach the
/// resolved `SymbolContext` to its `Emitter` before calling `emit_expr`.
pub(super) fn lower_udt_field_access(
    ctx: &LowerCtx,
    base: NodeId,
    member: u32,
    bang: bool,
    arena: &mut NodeArena,
) -> Result<NodeRef, LowerError> {
    if bang {
        // `!` (bang/default-member) access is a late-bound Object/Recordset
        // path, not applicable to a plain UDT — out of scope.
        return Err(LowerError::UnsupportedNode);
    }
    let field = resolve_udt_field(ctx, base, member)?;
    ctx.member_symbol.replace(Some(field.sym));

    let size_desc = arena.alloc(NodeArena::node(4, 0, field.field_size as u32, 0, 0, 0));
    let mut n = NodeArena::node(0x60, field.type_tag, 0, 0, 0, field.offset as u32);
    n.w[5] = size_desc.0;
    Ok(arena.alloc(n))
}

/// A resolved class-instance field/property reference: everything the
/// vtable-dispatch (`0x24` resolve-object / `0x0d` call) emission needs.
pub(super) struct ClassFieldRef {
    /// The class-instance local's own frame offset (its object-reference
    /// slot — LdAddr'd before every vtable call, per the oracle capture).
    pub obj_offset: i16,
    /// Vtable byte offset of the Property Get accessor, when present.
    pub get_slot: Option<u16>,
    /// Vtable byte offset of the Property Let accessor, when present.
    pub let_slot: Option<u16>,
    /// Vtable byte offset of the Property Set accessor, when present.
    pub set_slot: Option<u16>,
    /// The field's/property's VBA6 type tag (`vba_type_to_node_tag`). `None`
    /// for types with no confirmed node tag (e.g. Object) — fine for `Set`,
    /// which never numerically coerces its value; Get/Let callers require it.
    pub type_tag: Option<u16>,
    /// `true` for an explicit `Property Get`/`Let` member; `false` for a
    /// plain Public field. A property's Let call stages its argument into a
    /// temp frame slot first (`0x59 <offset>`) — a field's store does not.
    pub is_property: bool,
}

/// Resolve `base.<member>` for a plain local class-instance base (`Dim o As
/// New ClassName`), given `access_id` — the `MemberAccess` expression node's
/// OWN id (not `member`'s sym_id, which repeats across every occurrence of
/// that name in source). Codegen has no scanner/interner to turn a sym_id
/// back into text, so it never compares member names itself: sema resolves
/// each access SITE individually (by walking the class's ordered member list
/// and summing slot widths — ordinary name comparison, but done once, ahead
/// of time, where a scanner is available) and hands codegen the final byte
/// offsets directly via `BoundModule::class_member_slots`, keyed by
/// `access_id.0`. See `ResolvedClassMember` and the `vb6-class-vtable-slot-
/// rule` memory note for the full derivation (live-TTD-traced against
/// VBA6.DLL: a per-class-compile-context counter at a fixed struct offset,
/// initialized to the IDispatch prefix 0x1c, advanced by 8/12/4 bytes per
/// field/object-field/single-accessor-or-proc in source-declaration order).
pub(super) fn resolve_class_field(
    ctx: &LowerCtx,
    base: NodeId,
    access_id: NodeId,
) -> Result<ClassFieldRef, LowerError> {
    let local_idx = match ctx.module.resolutions.get(&base.0) {
        Some(NameResolution::Local { local_idx, .. }) => *local_idx,
        Some(_) => return Err(LowerError::UnsupportedNode),
        None => return Err(LowerError::Unresolved),
    };
    // Only used to confirm `base` really is a known class-instance local;
    // the resolved slots themselves come from `class_member_slots` below.
    ctx.local_class(local_idx).ok_or(LowerError::UnsupportedNode)?;
    let obj_offset = ctx.local_slots[local_idx].frame_offset;
    let resolved = ctx
        .module
        .class_member_slots
        .get(&access_id.0)
        .ok_or(LowerError::Unresolved)?;
    let member_ty = ctx
        .module
        .types
        .get(&access_id.0)
        .ok_or(LowerError::Unresolved)?;
    let type_tag = vba_type_to_node_tag(member_ty);
    Ok(ClassFieldRef {
        obj_offset,
        get_slot: resolved.get_slot,
        let_slot: resolved.let_slot,
        set_slot: resolved.set_slot,
        type_tag,
        is_property: resolved.is_property,
    })
}

/// Lower `o.F` (read position): the vtable-dispatch Property-Get idiom —
/// `LdAddr(out-temp)`, `LdAddr(o)`, resolve-object (`0x24`, module index 0),
/// vtable-call (`0x0d`) at the field's Get slot, then load the temp as the
/// expression's value. Byte-for-byte from a live oracle capture (see
/// `re_lab`'s Class1.F recon) — this front end has no scanner/interner, so
/// the emitted bytes are pre-built here (bypassing `NodeArena`'s per-node
/// dispatch, which has no shape for "several fixed instructions then a
/// value") and wrapped in a synthetic `0x7e` blob node, the same mechanism
/// already used for pre-emitted call sequences.
pub(super) fn lower_class_field_get(
    ctx: &LowerCtx,
    base: NodeId,
    access_id: NodeId,
    bang: bool,
    arena: &mut NodeArena,
) -> Result<NodeRef, LowerError> {
    if bang {
        return Err(LowerError::UnsupportedNode);
    }
    let field = resolve_class_field(ctx, base, access_id)?;
    let get_slot = field.get_slot.ok_or(LowerError::UnsupportedNode)?;
    let type_tag = field.type_tag.ok_or(LowerError::UnsupportedType)?;
    let temp_offset = ctx.local_slots[ctx.class_member_base].frame_offset;

    let mut bytes = Vec::new();
    bytes.push(0x04);
    bytes.extend_from_slice(&(temp_offset as u16).to_le_bytes());
    bytes.push(0x04);
    bytes.extend_from_slice(&(field.obj_offset as u16).to_le_bytes());
    bytes.extend_from_slice(&[0x24, 0x00, 0x00]);
    bytes.push(0x0d);
    bytes.extend_from_slice(&get_slot.to_le_bytes());
    bytes.extend_from_slice(&[0x01, 0x00]);
    bytes.push(0x6c);
    bytes.extend_from_slice(&(temp_offset as u16).to_le_bytes());

    let off = arena.alloc_blob(&bytes);
    Ok(arena.alloc(NodeArena::node(0x7e, type_tag, off, bytes.len() as u32, 0, 0)))
}

/// Lower `o.F = v` (a class-field assignment) or `o.P = v` (a Property-Let
/// call): the vtable-dispatch idiom — emit the coerced value, then (for a
/// property only) stage it into an addressable temp (`0x59 <offset>`, the
/// runtime handler for opcode 0x59 stores the top-of-stack value at
/// `[EBP+offset]` — confirmed from the decompiled runtime handler, matching
/// the same staging already used for `Erase`'s in-place reinit), `LdAddr(o)`,
/// resolve-object (`0x24`), vtable-call (`0x0d`) at the Let slot. A plain
/// field store passes the value directly with no staging — oracle-confirmed:
/// the field recon capture has no `0x59` between the pushed value and
/// `LdAddr(o)`, while every Property-Let recon capture does.
pub(super) fn lower_class_field_store(
    ctx: &LowerCtx,
    base: NodeId,
    access_id: NodeId,
    bang: bool,
    value_id: NodeId,
    expr_arena: &ExprArena,
    out: &mut Vec<u8>,
) -> Result<(), LowerError> {
    if bang {
        return Err(LowerError::UnsupportedNode);
    }
    let field = resolve_class_field(ctx, base, access_id)?;
    let let_slot = field.let_slot.ok_or(LowerError::UnsupportedNode)?;
    let type_tag = field.type_tag.ok_or(LowerError::UnsupportedType)?;

    let mut arena = NodeArena::new();
    let root = lower_expr_coerced(ctx, value_id, expr_arena, &mut arena, Some(type_tag))?;
    let root = coerce_assign_value(ctx, value_id, root, Some(type_tag), &mut arena);
    let mut emitter = Emitter::new(&arena);
    emitter.emit_expr(root, 2);
    out.extend(emitter.into_bytes());

    if field.is_property {
        let temp_offset = ctx.local_slots[ctx.class_member_base].frame_offset;
        out.push(0x59);
        out.extend_from_slice(&(temp_offset as u16).to_le_bytes());
    }

    out.push(0x04);
    out.extend_from_slice(&(field.obj_offset as u16).to_le_bytes());
    out.extend_from_slice(&[0x24, 0x00, 0x00]);
    out.push(0x0d);
    out.extend_from_slice(&let_slot.to_le_bytes());
    out.extend_from_slice(&[0x01, 0x00]);
    Ok(())
}

/// Lower `Set o.P = v` (an explicit `Property Set` call): the vtable-dispatch
/// idiom — emit the value, stage it into an addressable temp via the
/// refcounted Set-staging opcode (`fd 9c <offset>`, calling the runtime's
/// `__vbaObjSet`, flag=0/no-addref — confirmed identical for both `Nothing`
/// and a real object-reference source value via two independent oracle
/// captures, `set_probe`/`set_probe2`), then `LdAddr(o)`, resolve-object
/// (`0x24`), vtable-call (`0x0d`) at the Set slot. Unlike Let, staging is
/// unconditional here — a `Property Set` always takes its argument by
/// address for the runtime AddRef bookkeeping, never a plain pushed value.
/// Restricted to explicit `Property Set` members (`is_property`); a plain
/// object/Variant field's own synthesized Set accessor is a separate,
/// ungrounded case (no oracle capture yet shows whether it also stages).
pub(super) fn lower_class_field_set(
    ctx: &LowerCtx,
    base: NodeId,
    access_id: NodeId,
    bang: bool,
    value_id: NodeId,
    expr_arena: &ExprArena,
    out: &mut Vec<u8>,
) -> Result<(), LowerError> {
    if bang {
        return Err(LowerError::UnsupportedNode);
    }
    let field = resolve_class_field(ctx, base, access_id)?;
    let set_slot = field.set_slot.ok_or(LowerError::UnsupportedNode)?;
    if !field.is_property {
        return Err(LowerError::UnsupportedNode);
    }

    let mut arena = NodeArena::new();
    let root = lower_expr_coerced(ctx, value_id, expr_arena, &mut arena, field.type_tag)?;
    let mut emitter = Emitter::new(&arena);
    emitter.emit_expr(root, 2);
    out.extend(emitter.into_bytes());

    let temp_offset = ctx.local_slots[ctx.class_member_base].frame_offset;
    out.push(0xfd);
    out.push(0x9c);
    out.extend_from_slice(&(temp_offset as u16).to_le_bytes());

    out.push(0x04);
    out.extend_from_slice(&(field.obj_offset as u16).to_le_bytes());
    out.extend_from_slice(&[0x24, 0x00, 0x00]);
    out.push(0x0d);
    out.extend_from_slice(&set_slot.to_le_bytes());
    out.extend_from_slice(&[0x01, 0x00]);
    Ok(())
}

/// Build a typed local/param load node (opcode 0x74 = ByVal, 0x75 = ByRef).
/// Node layout: w[4] = sym-child (frame offset in high 16 bits), w[5] = type_ctx.
pub(super) fn build_frame_load_node(
    arena: &mut NodeArena,
    opcode: u16,
    type_tag: u16,
    ctx: usize,
    frame_offset: i16,
) -> NodeRef {
    let sym = arena.alloc(NodeArena::node(0, 0, (frame_offset as u16 as u32) << 16, 0, 0, 0));
    arena.alloc(NodeArena::node(opcode, type_tag, sym.0, ctx as u32, 0, 0))
}

/// Build a module-global load node (opcode 0x77).
/// Node layout: w[4] = (module_desc | (field_offset << 16)), w[5] = type_ctx.
pub(super) fn build_global_load_node(
    arena: &mut NodeArena,
    ctx: usize,
    module_desc: u16,
    field_offset: u16,
) -> NodeRef {
    let packed = (module_desc as u32) | ((field_offset as u32) << 16);
    arena.alloc(NodeArena::node(0x77, 0, packed, ctx as u32, 0, 0))
}

pub(super) fn lower_binop(
    ctx: &LowerCtx,
    node_id: NodeId,
    op: BinOpKind,
    lhs_id: NodeId,
    rhs_id: NodeId,
    expr_arena: &ExprArena,
    arena: &mut NodeArena,
) -> Result<NodeRef, LowerError> {
    // VB6 widens both operands to the wider of their types before emitting;
    // integer literals in Long expressions are promoted to Long.
    let lhs_ty = ctx.module.types.get(&lhs_id.0);
    let rhs_ty = ctx.module.types.get(&rhs_id.0);
    let mut operand_coerce = wider_numeric_tag(lhs_ty, rhs_ty);
    // Floating division (`/`) widens its operands to the operation (result) type.
    // For Currency operands that is Double (Currency is divided in Double), not
    // the operands' common Currency type. Exponentiation (`^`) always computes in
    // Double regardless of operand type, so it widens the same way.
    if matches!(op, BinOpKind::Div | BinOpKind::Pow) {
        if let Some(rt) = ctx.module.types.get(&node_id.0).and_then(vba_type_to_node_tag) {
            operand_coerce = Some(rt);
        }
    }

    let lhs_ref = lower_expr_coerced(ctx, lhs_id, expr_arena, arena, operand_coerce)?;
    let lhs_ref = coerce_operand(ctx, lhs_id, lhs_ref, operand_coerce, expr_arena, arena);
    let rhs_ref = lower_expr_coerced(ctx, rhs_id, expr_arena, arena, operand_coerce)?;
    let rhs_ref = coerce_operand(ctx, rhs_id, rhs_ref, operand_coerce, expr_arena, arena);

    // Power (`^`) is its own bound-node opcode (0x1a) and always yields Double;
    // both operands are already widened to Double above via `operand_coerce`.
    if op == BinOpKind::Pow {
        return Ok(arena.alloc(NodeArena::node(0x1a, 11, lhs_ref.0, rhs_ref.0, 0, 0)));
    }

    let opcode = binop_node_opcode(op).ok_or(LowerError::UnsupportedNode)?;

    let type_tag = if is_comparison_op(op) {
        0u16
    } else {
        let result_type = ctx.module.types.get(&node_id.0).ok_or(LowerError::Unresolved)?;
        vba_type_to_node_tag(result_type).ok_or(LowerError::UnsupportedType)?
    };

    Ok(arena.alloc(NodeArena::node(opcode, type_tag, lhs_ref.0, rhs_ref.0, 0, 0)))
}

/// Widen a binary-operation operand to the operation type when its own type is
/// narrower. VB6 wraps such an operand in a coercion node whose emission is the
/// operand load followed by a single type-conversion opcode (e.g. Integer→Long
/// emits 0xe7, Long→Double emits 0xec). We model that with the synthetic
/// coercion node (opcode 0x78): target type tag in the node, operand as word[4].
///
/// Integer *literals* are already widened in place by `lower_lit_coerced`, so
/// they need no conversion node (VB6 folds the literal to the wider type).
/// Insert an assignment-value coercion node when the source type differs from the
/// assignment target type. VB6 computes the rvalue in its natural type and then
/// converts it to the destination type before storing.
///
/// The conversion is omitted in two cases: when the types already match, and when
/// both source and destination are floating-point (Single/Double) — there the
/// typed load pushes the common float representation and the typed store
/// narrows/widens, so no separate conversion opcode appears. Every other cross-type
/// pair (integer<->integer, ->/<- Currency, ->/<- Date, ->/<- Byte) carries an
/// explicit conversion opcode, emitted by the value-emitter for the 0x78 node.
///
/// Literals are skipped: integer literals are already lowered in the destination
/// type by `lower_lit_coerced`, so wrapping them would double-convert.
pub(super) fn coerce_assign_value(
    ctx: &LowerCtx,
    value_id: NodeId,
    value_root: NodeRef,
    target_tag: Option<u16>,
    arena: &mut NodeArena,
) -> NodeRef {
    let Some(target) = target_tag else {
        return value_root;
    };
    match arena.get(value_root).opcode() {
        // 8-byte (Currency/Date), floating-point, and pooled-string literals are
        // already their final type — wrapping would double-convert.
        2 | 3 | 0x79 => return value_root,
        // An integer literal is retyped in place for an integer target (no
        // conversion); for a float/Currency target it keeps its natural type and
        // needs an explicit Int→T conversion, emitted below.
        1 if matches!(target, 5 | 6 | 8) => return value_root,
        _ => {}
    }
    let Some(src_tag) = ctx.module.types.get(&value_id.0).and_then(vba_type_to_node_tag) else {
        return value_root;
    };
    if src_tag == target {
        return value_root;
    }
    let is_float = |t: u16| t == 10 || t == 11;
    if is_float(src_tag) && is_float(target) {
        return value_root;
    }
    arena.alloc(NodeArena::node(0x78, target, value_root.0, 0, 0, 0))
}

pub(super) fn coerce_operand(
    ctx: &LowerCtx,
    operand_id: NodeId,
    operand_ref: NodeRef,
    target_tag: Option<u16>,
    expr_arena: &ExprArena,
    arena: &mut NodeArena,
) -> NodeRef {
    let Some(target) = target_tag else {
        return operand_ref;
    };
    // An integer literal is retyped in place by `lower_lit_coerced` for an integer
    // operation type (no conversion); for a float/Currency operation type it keeps
    // its natural type and needs an explicit conversion, emitted below. Every other
    // literal is already its final type.
    if let ExprNode::Literal { lit } = expr_arena.get(operand_id) {
        let is_int_lit = matches!(lit, AstLit::Int(_) | AstLit::Long(_) | AstLit::Bool(_));
        if !is_int_lit || matches!(target, 5 | 6 | 8) {
            return operand_ref;
        }
    }
    let Some(src_ty) = ctx.module.types.get(&operand_id.0) else {
        return operand_ref;
    };
    let Some(src_tag) = vba_type_to_node_tag(src_ty) else {
        return operand_ref;
    };
    if src_tag == target {
        return operand_ref;
    }
    // VB6 emits an explicit widening conversion opcode only for an integer-typed
    // operand (Integer / Long / Byte). A floating-point operand widened to a
    // wider float (e.g. Single -> Double) carries no conversion opcode — the
    // operation consumes the value directly. (The complete per-type-pair gate
    // lives in the value-emitter; for the reachable widening cases — where the
    // operand type is never wider than the operation type — this integer-source
    // rule is exact and oracle-confirmed.)
    if !matches!(
        src_ty,
        VbaType::Integer | VbaType::Long | VbaType::Byte | VbaType::Boolean | VbaType::Currency
    ) {
        return operand_ref;
    }
    arena.alloc(NodeArena::node(0x78, target, operand_ref.0, 0, 0, 0))
}

pub(super) fn lower_unop(
    ctx: &LowerCtx,
    node_id: NodeId,
    op: UnOpKind,
    operand_id: NodeId,
    expr_arena: &ExprArena,
    arena: &mut NodeArena,
) -> Result<NodeRef, LowerError> {
    let operand_ref = lower_expr(ctx, operand_id, expr_arena, arena)?;
    match op {
        // Unary plus is a no-op: the operand is emitted unchanged.
        UnOpKind::Pos => Ok(operand_ref),
        // Negate and Not are emitted through the generic operation emitter
        // (the single-operand arithmetic form: only `word[4]` is set, so no
        // right operand is emitted). The opcode byte is selected by that
        // emitter as RT_BINOP_BASE[op] + RT_TYPE_OFFSET[result type tag]:
        //   negate -> op 7 (base 0x00c6),  Not -> op 6 (base 0x00be).
        // Both use the arithmetic dispatch (RT_DISPATCH_FLAG & 0x10 == 0), so
        // the offset comes from the node's own (result) type tag.
        UnOpKind::Neg | UnOpKind::Not => {
            let result_type = ctx
                .module
                .types
                .get(&node_id.0)
                .ok_or(LowerError::Unresolved)?;
            let type_tag = vba_type_to_node_tag(result_type).ok_or(LowerError::UnsupportedType)?;
            let opcode: u16 = if matches!(op, UnOpKind::Neg) { 7 } else { 6 };
            Ok(arena.alloc(NodeArena::node(opcode, type_tag, operand_ref.0, 0, 0, 0)))
        }
    }
}
