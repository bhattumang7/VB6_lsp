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
    if let ExprNode::MemberAccess { base, member, bang } = expr_arena.get(target_id) {
        let (base, member, bang) = (*base, *member, *bang);
        if member_access_base_is_class(ctx.module, base) {
            return lower_class_field_store(ctx, base, bang, value_id, expr_arena, out);
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

    // String-returning runtime intrinsic into a String target. Each argument is
    // pushed (right-to-left) either by value (size-load, coerced) or boxed into a
    // hidden 16-byte temp tagged with its VARTYPE (`04 <var> 4d <temp> <vt> 40`).
    // A final hidden result temp is passed by address; after the runtime call (0x0a)
    // the result is loaded (0x60), moved into the target (0x31), and the result temp
    // freed (0x35). The runtime-call reference index and total argument-byte count
    // follow the same convention as user calls.
    if let Some(BuiltinCall::RtcString { .. }) = ctx.module.builtins.get(&value_id.0) {
        let tgt_off = match resolution {
            NameResolution::Local { local_idx, .. } => ctx.local_slots[*local_idx].frame_offset,
            _ => return Err(LowerError::UnsupportedNode),
        };
        // Emit the runtime call (result into a hidden temp), then load the result
        // (0x60), move it into the target (0x31), and free the owned variant temps.
        let (result_temp, free_offsets) = emit_rtc_string_call(ctx, value_id, expr_arena, out)?;
        out.push(0x04);
        out.extend_from_slice(&result_temp.to_le_bytes());
        out.push(0x60);
        out.push(0x31);
        out.extend_from_slice(&tgt_off.to_le_bytes());
        emit_variant_temp_free(out, &free_offsets);
        return Ok(());
    }
    // An owned-temp string concatenation (a `&` chain with a runtime-string operand)
    // is lowered with the cleanup-aware concat opcode family.
    if is_owned_concat(ctx.module, value_id, expr_arena) {
        if let NameResolution::Local { local_idx, .. } = resolution {
            let tgt_off = ctx.local_slots[*local_idx].frame_offset;
            return lower_owned_concat(ctx, tgt_off, value_id, expr_arena, out);
        }
    }
    // `Len` of a runtime-string result (the release-aware Len form) into a Long.
    if is_owned_len(ctx.module, value_id, expr_arena) {
        if let NameResolution::Local { local_idx, .. } = resolution {
            if matches!(ctx.local_type(*local_idx), VbaType::Long) {
                let tgt_off = ctx.local_slots[*local_idx].frame_offset;
                return lower_owned_len(ctx, tgt_off, value_id, expr_arena, out);
            }
        }
    }
    // A String comparison with a runtime-string operand into an Integer/Boolean.
    let cmp_value = unwrap_parens(value_id, expr_arena);
    if is_owned_compare(ctx.module, cmp_value, expr_arena) {
        if let NameResolution::Local { local_idx, .. } = resolution {
            if matches!(ctx.local_type(*local_idx), VbaType::Integer | VbaType::Boolean) {
                let tgt_off = ctx.local_slots[*local_idx].frame_offset;
                return lower_owned_compare(ctx, tgt_off, cmp_value, expr_arena, out);
            }
        }
    }

    // Static local target: the value is pushed, then stored into the procedure's
    // static block (0x5f-addressed) rather than the frame.
    if let NameResolution::Local { local_idx, .. } = resolution {
        if ctx.proc.locals[*local_idx].is_static {
            let ty = ctx.local_type(*local_idx).clone();
            let store_op = static_store_op(&ty).ok_or(LowerError::UnsupportedType)?;
            let off = ctx.static_offsets[*local_idx];
            let coerce_tag = vba_type_to_node_tag(&ty);
            let mut arena = NodeArena::new();
            let root = lower_expr_coerced(ctx, value_id, expr_arena, &mut arena, coerce_tag)?;
            let root = coerce_assign_value(ctx, value_id, root, coerce_tag, &mut arena);
            let mut emitter = Emitter::new(&arena);
            emitter.emit_expr(root, 2);
            out.extend(emitter.into_bytes());
            emit_static_access(ctx, out, &store_op, off);
            return Ok(());
        }
    }

    // Fixed-length-string source: `s = fx` where `fx` is `As String * n` reads the
    // fixed buffer length-aware (LdAddr `fx` + 0x33<len>) and moves the resulting
    // BSTR temp into the target (0x31). The source isn't a plain BSTR-pointer load.
    if let ExprNode::NameRef { .. } = expr_arena.get(value_id) {
        if let Some(NameResolution::Local { local_idx, .. }) =
            ctx.module.resolutions.get(&value_id.0)
        {
            if let Some(len) = ctx.proc.locals[*local_idx].fixed_string_len {
                if let NameResolution::Local { local_idx: t_idx, .. } = resolution {
                    let src_off = ctx.local_slots[*local_idx].frame_offset;
                    let tgt_off = ctx.local_slots[*t_idx].frame_offset;
                    out.push(0x04);
                    out.extend_from_slice(&src_off.to_le_bytes());
                    out.push(0x33);
                    out.extend_from_slice(&len.to_le_bytes());
                    out.push(0x31);
                    out.extend_from_slice(&tgt_off.to_le_bytes());
                    return Ok(());
                }
            }
        }
    }

    // Variant target: the RHS is converted into a hidden 16-byte Variant temp,
    // then variant-stored (0xfc 0xf6 = value-emitter index 0x1f6) into the target.
    if matches!(ctx.module.types.get(&target_id.0), Some(VbaType::Variant)) {
        let v_off = match resolution {
            NameResolution::Local { local_idx, .. } => ctx.local_slots[*local_idx].frame_offset,
            _ => return Err(LowerError::UnsupportedNode),
        };
        let vi = ctx.variant_next.get();
        ctx.variant_next.set(vi + 1);
        let temp_off = ctx.local_slots[ctx.variant_base + vi].frame_offset;
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
        out.push(0xfc);
        out.push(0xf6);
        out.extend_from_slice(&v_off.to_le_bytes());
        return Ok(());
    }

    // Variant source into a non-Variant scalar target: a Variant is read by
    // address (0x04), converted to the target type, then stored.
    if let ExprNode::NameRef { .. } = expr_arena.get(value_id) {
        if matches!(ctx.module.types.get(&value_id.0), Some(VbaType::Variant)) {
            let v_off = match ctx.module.resolutions.get(&value_id.0) {
                Some(NameResolution::Local { local_idx, .. }) => {
                    ctx.local_slots[*local_idx].frame_offset
                }
                _ => return Err(LowerError::UnsupportedNode),
            };
            let (target_tag, store_ctx, store_off) = match resolution {
                NameResolution::Local { local_idx, .. } => {
                    let ty = ctx.local_type(*local_idx);
                    (
                        vba_type_to_node_tag(ty).ok_or(LowerError::UnsupportedType)?,
                        load_store_ctx(ty).ok_or(LowerError::UnsupportedType)?,
                        ctx.local_slots[*local_idx].frame_offset,
                    )
                }
                _ => return Err(LowerError::UnsupportedNode),
            };
            let arena = NodeArena::new();
            let mut em = Emitter::new(&arena);
            em.emit_ldaddr(v_off);
            // Variant (tag 0xf) → target conversion, then the plain target store.
            em.emit_conversion(target_tag as i32, 0xf);
            em.emit_var_store(store_ctx, store_off);
            out.extend(em.into_bytes());
            return Ok(());
        }
    }

    // Function call as the RHS: `r = F(args)` — emit the result-producing call
    // (0x5e, result left on the stack), then store the result into the target.
    if let ExprNode::Call { func, args } = expr_arena.get(value_id) {
        if matches!(ctx.module.resolutions.get(&func.0), Some(NameResolution::Proc(_))) {
            let (func, args) = (*func, *args);
            lower_call(ctx, func, args, true, expr_arena, out)?;
            match resolution {
                NameResolution::Local { local_idx, .. } => {
                    let ty = ctx.local_type(*local_idx);
                    let sctx = load_store_ctx(ty).ok_or(LowerError::UnsupportedType)?;
                    let arena = NodeArena::new();
                    let mut em = Emitter::new(&arena);
                    em.emit_var_store(sctx, ctx.local_slots[*local_idx].frame_offset);
                    out.extend(em.into_bytes());
                }
                _ => return Err(LowerError::UnsupportedNode),
            }
            return Ok(());
        }
    }

    // Array element load as the RHS: `r = a(i)` — push the Long index, LdAddr the
    // array descriptor, the element-load opcode (0x9e for Long), then store to the
    // target.
    if let ExprNode::Call { func, args } = expr_arena.get(value_id) {
        if let Some((arr_off, elem, _dims)) = array_local_info(ctx, *func) {
            let idx = single_index(*args, expr_arena).ok_or(LowerError::UnsupportedNode)?;
            let load_op: u8 = match elem {
                VbaType::Long => 0x9e,
                VbaType::Integer => 0x9d,
                _ => return Err(LowerError::UnsupportedNode),
            };
            out.extend_from_slice(&lower_expr_to_bytes_coerced(ctx, idx, expr_arena, Some(8))?);
            out.push(0x04);
            out.extend_from_slice(&arr_off.to_le_bytes());
            out.push(load_op);
            match resolution {
                NameResolution::Local { local_idx, .. } => {
                    let ty = ctx.local_type(*local_idx);
                    let sctx = load_store_ctx(ty).ok_or(LowerError::UnsupportedType)?;
                    out.push(RT_STORE_BY_CTX[sctx]);
                    out.extend_from_slice(&ctx.local_slots[*local_idx].frame_offset.to_le_bytes());
                }
                _ => return Err(LowerError::UnsupportedNode),
            }
            return Ok(());
        }
    }

    // String concatenation chain `s = a & b & c`: emit the first concat, then for
    // each further operand materialize the running result to a hidden temp
    // (store-keep 0x23) and concat. The final result is moved into the target
    // (0x31); the intermediate temps are then freed (0x2f) for BSTR cleanup.
    if matches!(ctx.module.types.get(&target_id.0), Some(VbaType::String))
        && matches!(expr_arena.get(value_id), ExprNode::BinOp { op: BinOpKind::Cat, .. })
    {
        let s_off = match resolution {
            NameResolution::Local { local_idx, .. } => ctx.local_slots[*local_idx].frame_offset,
            _ => return Err(LowerError::UnsupportedNode),
        };
        let mut ops = Vec::new();
        flatten_concat(value_id, expr_arena, &mut ops);
        let mut temps: Vec<i16> = Vec::new();
        // Every operand that materializes a fresh BSTR (a converted numeric or a
        // fixed-length string) is store-kept (0x23) to a temp for later cleanup;
        // so is every intermediate concat result except the last.
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
        out.push(0x31);
        out.extend_from_slice(&s_off.to_le_bytes());
        // Release the tracked temps: a single 0x2f, or 0x32 <byte-len> <offsets…>.
        match temps.len() {
            0 => {}
            1 => {
                out.push(0x2f);
                out.extend_from_slice(&temps[0].to_le_bytes());
            }
            n => {
                out.push(0x32);
                out.extend_from_slice(&((n * 2) as u16).to_le_bytes());
                for t in &temps {
                    out.extend_from_slice(&t.to_le_bytes());
                }
            }
        }
        return Ok(());
    }

    let coerce_tag = match resolution {
        NameResolution::Local { local_idx, .. } => vba_type_to_node_tag(ctx.local_type(*local_idx)),
        NameResolution::Param { param_idx, .. } => vba_type_to_node_tag(ctx.param_type(*param_idx)),
        NameResolution::ModuleVar(idx) => vba_type_to_node_tag(ctx.global_type(*idx)),
        _ => None,
    };

    let mut arena = NodeArena::new();
    let value_root = lower_expr_coerced(ctx, value_id, expr_arena, &mut arena, coerce_tag)?;
    let value_root = coerce_assign_value(ctx, value_id, value_root, coerce_tag, &mut arena);

    // A String target receives a *move* store (0x31, ctx 9) when the value is a
    // freshly-produced temp — a `&` concatenation, a numeric→String coercion (0x78),
    // or a `CStr` explicit conversion (0x7c) — and a *copy* store (0x43, ctx 8) when
    // it is a plain string variable. The move avoids an extra BSTR allocation.
    let value_is_fresh_string = matches!(
        expr_arena.get(value_id),
        ExprNode::BinOp { op: BinOpKind::Cat, .. }
    ) || matches!(arena.get(value_root).opcode(), 0x78 | 0x7c);

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

    match resolution {
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
        _ => return Err(LowerError::Unresolved),
    }

    out.extend(emitter.into_bytes());
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
