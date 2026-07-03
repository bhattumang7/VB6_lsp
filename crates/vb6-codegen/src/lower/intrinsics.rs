use super::*;
use super::decl::*;
use super::expr::*;


/// Lower an intra-module Sub/Function call: push each argument (ByVal value or
/// ByRef address), then the call opcode and its `<callsite-index><arg-bytes>`
/// operand. The call site index is the call's emission-order position in this
/// procedure; arg-bytes is the total pushed-argument byte size (a ByRef argument
/// pushes a 4-byte pointer). `is_function` selects the result-producing form
/// (0x5e, result left on the stack) over the statement form (0x0a).
pub(super) fn lower_call(
    ctx: &LowerCtx,
    callee_id: NodeId,
    args_id: NodeId,
    is_function: bool,
    expr_arena: &ExprArena,
    out: &mut Vec<u8>,
) -> Result<(), LowerError> {
    let proc_idx = match ctx.module.resolutions.get(&callee_id.0) {
        Some(NameResolution::Proc(pi)) => *pi,
        _ => return Err(LowerError::UnsupportedNode),
    };
    let args: Vec<NodeId> = match expr_arena.get(args_id) {
        ExprNode::ArgList { args } => args.clone(),
        _ => Vec::new(),
    };
    // An omitted Optional argument must push the parameter's default value, which
    // needs the default-value expression carried on the bound parameter (not yet
    // modelled). Gate any call whose argument count doesn't match the parameter
    // count rather than emit a short argument list.
    if args.len() != ctx.module.procs[proc_idx].params.len() {
        return Err(LowerError::UnsupportedNode);
    }
    // A runtime-string result passed as a ByVal String argument: materialize the
    // rtc into a temp, copy that owned BSTR into a string temp (0xfd 0xfe), pass the
    // string temp, and after the call free the string temp (0x2f) and the rtc result
    // temp (0x35). Only the single-argument Sub-call form is handled.
    let has_owned_arg = args
        .iter()
        .any(|&a| matches!(ctx.module.builtins.get(&a.0), Some(BuiltinCall::RtcString { .. })));
    if has_owned_arg {
        if args.len() != 1 || is_function {
            return Err(LowerError::UnsupportedNode);
        }
        let param = &ctx.module.procs[proc_idx].params[0];
        if !param.flags.by_val || !matches!(param.vba_type, VbaType::String) {
            return Err(LowerError::UnsupportedNode);
        }
        let (result_temp, free) = emit_rtc_string_call(ctx, args[0], expr_arena, out)?;
        out.push(0x04);
        out.extend_from_slice(&result_temp.to_le_bytes());
        let copy_temp = alloc_owned_copy_temp(ctx);
        out.extend_from_slice(&[0xfd, 0xfe]);
        out.extend_from_slice(&copy_temp.to_le_bytes());
        let idx = ctx.call_next.get();
        ctx.call_next.set(idx + 1);
        out.push(0x0a);
        out.extend_from_slice(&(idx as u16).to_le_bytes());
        out.extend_from_slice(&4u16.to_le_bytes());
        out.push(0x2f);
        out.extend_from_slice(&copy_temp.to_le_bytes());
        emit_variant_temp_free(out, &free);
        return Ok(());
    }
    // Build each argument's push bytes (matched to its parameter by position),
    // then emit them right-to-left — VB6 pushes arguments in reverse order.
    let mut arg_bytes: u16 = 0;
    let mut pushes: Vec<Vec<u8>> = Vec::with_capacity(args.len());
    for (i, &arg) in args.iter().enumerate() {
        let param = ctx.module.procs[proc_idx].params.get(i);
        let by_ref = param.map(|p| !p.flags.by_val).unwrap_or(false);
        let mut buf = Vec::new();
        if by_ref {
            // ByRef: push the argument variable's address.
            let off = arg_var_offset(ctx, arg).ok_or(LowerError::UnsupportedNode)?;
            buf.push(0x04);
            buf.extend_from_slice(&off.to_le_bytes());
            arg_bytes += 4;
        } else {
            // ByVal. A same-typed local-variable argument (no conversion) is pushed
            // with a *size-based* value load (load N bytes), not the type-specific
            // load: 1→0xfc 0xe0, 2→0x6b, 4→0x6c, 8→0x6d. An argument needing a
            // conversion is loaded as its own type and converted (the store-style
            // coercion). A Variant ByVal argument needs the variant-copy path.
            let pty = param.map(|p| p.vba_type.clone());
            if matches!(pty, Some(VbaType::Variant)) {
                return Err(LowerError::UnsupportedNode);
            }
            let same_type = matches!(expr_arena.get(arg), ExprNode::NameRef { .. })
                && pty.as_ref() == ctx.module.types.get(&arg.0);
            let off = arg_var_offset(ctx, arg);
            if let (true, Some(off), Some(ty)) = (same_type, off, pty.as_ref()) {
                emit_sized_value_load(static_var_size(ty), off, &mut buf);
            } else {
                let coerce = pty.as_ref().and_then(vba_type_to_node_tag);
                let mut arena = NodeArena::new();
                let root = lower_expr_coerced(ctx, arg, expr_arena, &mut arena, coerce)?;
                let root = coerce_assign_value(ctx, arg, root, coerce, &mut arena);
                let mut emitter = Emitter::new(&arena);
                emitter.emit_expr(root, 2);
                buf.extend(emitter.into_bytes());
            }
            arg_bytes += pty.as_ref().map(call_arg_bytes).unwrap_or(4);
        }
        pushes.push(buf);
    }
    for buf in pushes.iter().rev() {
        out.extend_from_slice(buf);
    }
    let idx = ctx.call_next.get();
    ctx.call_next.set(idx + 1);
    out.push(if is_function { 0x5e } else { 0x0a });
    out.extend_from_slice(&(idx as u16).to_le_bytes());
    out.extend_from_slice(&arg_bytes.to_le_bytes());
    Ok(())
}

/// Emit a String-returning runtime-library call, producing its result into a hidden
/// 16-byte temp: the sequence `<args> 04<result> 0a<ref> <arg-bytes>`. Allocates the
/// call's temps and emits; see [`emit_rtc_string_call_at`] for the pre-reserved form.
pub(super) fn emit_rtc_string_call(
    ctx: &LowerCtx,
    call_id: NodeId,
    expr_arena: &ExprArena,
    out: &mut Vec<u8>,
) -> Result<(i16, Vec<i16>), LowerError> {
    let sig = match ctx.module.builtins.get(&call_id.0) {
        Some(BuiltinCall::RtcString { args }) => args.clone(),
        _ => return Err(LowerError::UnsupportedNode),
    };
    let base = reserve_rtc_temps(ctx, &sig);
    emit_rtc_string_call_at(ctx, call_id, base, expr_arena, out)
}

/// Emit a String-returning runtime call using temps pre-reserved at slot `base`
/// (relative to `string_rtc_base`). Leaves nothing on the stack. Returns
/// `(result_temp, variant_temps_to_free)` — the owned variant temps (any
/// Missing-variant temps and the result temp), in allocation order.
pub(super) fn emit_rtc_string_call_at(
    ctx: &LowerCtx,
    call_id: NodeId,
    base: usize,
    expr_arena: &ExprArena,
    out: &mut Vec<u8>,
) -> Result<(i16, Vec<i16>), LowerError> {
    let sig = match ctx.module.builtins.get(&call_id.0) {
        Some(BuiltinCall::RtcString { args }) => args.clone(),
        _ => return Err(LowerError::UnsupportedNode),
    };
    let args_id = match expr_arena.get(call_id) {
        ExprNode::Call { args, .. } => *args,
        _ => return Err(LowerError::UnsupportedNode),
    };
    let arg_ids: Vec<NodeId> = match expr_arena.get(args_id) {
        ExprNode::ArgList { args } => args.clone(),
        _ => return Err(LowerError::UnsupportedNode),
    };
    let supplied = sig
        .iter()
        .filter(|a| !matches!(a, RtcArg::MissingVariant))
        .count();
    if arg_ids.len() != supplied {
        return Err(LowerError::UnsupportedNode);
    }
    let n_temps = rtc_call_temp_count(&sig) - 1;
    let base = ctx.string_rtc_base + base;
    let result_temp = ctx.local_slots[base + n_temps].frame_offset;

    let mut pushes: Vec<Vec<u8>> = Vec::with_capacity(sig.len());
    let mut arg_bytes: u16 = 0;
    let mut slot_k = 0usize;
    let mut arg_cursor = 0usize;
    let mut missing_temps: Vec<i16> = Vec::new();
    for mode in sig.iter() {
        let mut p = Vec::new();
        match mode {
            RtcArg::ByVal(ty) => {
                let arg = arg_ids[arg_cursor];
                arg_cursor += 1;
                let coerce = vba_type_to_node_tag(ty);
                let mut a = NodeArena::new();
                let root = lower_expr_coerced(ctx, arg, expr_arena, &mut a, coerce)?;
                let root = coerce_assign_value(ctx, arg, root, coerce, &mut a);
                let mut em = Emitter::new(&a);
                em.emit_expr(root, 2);
                p.extend(em.into_bytes());
                arg_bytes += call_arg_bytes(ty);
            }
            RtcArg::Boxed => {
                let arg = arg_ids[arg_cursor];
                arg_cursor += 1;
                let temp = ctx.local_slots[base + slot_k].frame_offset;
                slot_k += 1;
                let src = arg_var_offset(ctx, arg).ok_or(LowerError::UnsupportedNode)?;
                let arg_ty = ctx.module.types.get(&arg.0).ok_or(LowerError::UnsupportedNode)?;
                let vt = vba_type_to_vartype(arg_ty).ok_or(LowerError::UnsupportedType)?;
                p.push(0x04);
                p.extend_from_slice(&src.to_le_bytes());
                p.push(0x4d);
                p.extend_from_slice(&temp.to_le_bytes());
                p.extend_from_slice(&[vt, 0x40]);
                arg_bytes += 4;
            }
            RtcArg::MissingVariant => {
                slot_k += 1;
                let temp = ctx.local_slots[base + slot_k].frame_offset;
                slot_k += 1;
                missing_temps.push(temp);
                p.push(0x27);
                p.extend_from_slice(&temp.to_le_bytes());
                arg_bytes += 4;
            }
        }
        pushes.push(p);
    }
    for p in pushes.iter().rev() {
        out.extend_from_slice(p);
    }
    out.push(0x04);
    out.extend_from_slice(&result_temp.to_le_bytes());
    let r = ctx.call_next.get();
    ctx.call_next.set(r + 1);
    out.push(0x0a);
    out.extend_from_slice(&(r as u16).to_le_bytes());
    out.extend_from_slice(&(arg_bytes + 4).to_le_bytes());
    let mut free = missing_temps;
    free.push(result_temp);
    Ok((result_temp, free))
}

/// Materialize one operand of an owned-temp concat as a BSTR address on the stack.
/// `rtc_base` is the pre-reserved temp base for an rtc operand (`None` for a plain
/// operand). A runtime-string result is produced into its temp then `04<temp>` (the
/// owned variant temps are appended to `rtc_owned`); a plain String variable is
/// value-loaded and copied into a fresh temp (`6c<var> 46<temp>`); a string literal
/// loads into a fresh temp (`3a<temp> <ref>`). The copy/literal temps are not owned.
pub(super) fn materialize_owned_concat_operand(
    ctx: &LowerCtx,
    op: NodeId,
    rtc_base: Option<usize>,
    expr_arena: &ExprArena,
    out: &mut Vec<u8>,
    rtc_owned: &mut Vec<i16>,
) -> Result<(), LowerError> {
    if let Some(base) = rtc_base {
        let (result_temp, free) = emit_rtc_string_call_at(ctx, op, base, expr_arena, out)?;
        rtc_owned.extend(free);
        out.push(0x04);
        out.extend_from_slice(&result_temp.to_le_bytes());
        return Ok(());
    }
    match expr_arena.get(op) {
        ExprNode::Literal { lit: AstLit::Str(s) } => {
            let temp = alloc_string_rtc_temp(ctx);
            let r = ctx.call_next.get();
            ctx.call_next.set(r + 1);
            ctx.intern_string(s);
            out.push(0x3a);
            out.extend_from_slice(&temp.to_le_bytes());
            out.extend_from_slice(&(r as u16).to_le_bytes());
            Ok(())
        }
        ExprNode::NameRef { .. }
            if matches!(ctx.module.types.get(&op.0), Some(VbaType::String)) =>
        {
            let off = arg_var_offset(ctx, op).ok_or(LowerError::UnsupportedNode)?;
            let temp = alloc_string_rtc_temp(ctx);
            out.push(0x6c);
            out.extend_from_slice(&off.to_le_bytes());
            out.push(0x46);
            out.extend_from_slice(&temp.to_le_bytes());
            Ok(())
        }
        _ => Err(LowerError::UnsupportedNode),
    }
}

/// Lower an owned-temp concatenation assignment `s = <chain>` where at least one
/// operand is a runtime-string result. The compiler reserves every rtc operand's
/// temps first (in source order), then the materialization/concat temps. Each
/// operand is materialized to a BSTR address (left-to-right); adjacent pairs
/// concatenate with `0xfb 0xef <temp>` (the result address stays on the stack and
/// chains). Finally the result is loaded (0x60), moved into the target (0x31), and
/// all owned temps (rtc results first, then concat results) freed.
pub(super) fn lower_owned_concat(
    ctx: &LowerCtx,
    tgt_off: i16,
    value_id: NodeId,
    expr_arena: &ExprArena,
    out: &mut Vec<u8>,
) -> Result<(), LowerError> {
    let mut ops = Vec::new();
    flatten_concat(value_id, expr_arena, &mut ops);
    if ops.len() < 2 {
        return Err(LowerError::UnsupportedNode);
    }
    // Reserve each rtc operand's temps up front, in source order.
    let rtc_bases: Vec<Option<usize>> = ops
        .iter()
        .map(|&op| match ctx.module.builtins.get(&op.0) {
            Some(BuiltinCall::RtcString { args }) => Some(reserve_rtc_temps(ctx, args)),
            _ => None,
        })
        .collect();

    let mut rtc_owned: Vec<i16> = Vec::new();
    let mut concat_owned: Vec<i16> = Vec::new();
    materialize_owned_concat_operand(ctx, ops[0], rtc_bases[0], expr_arena, out, &mut rtc_owned)?;
    for (i, &op) in ops.iter().enumerate().skip(1) {
        materialize_owned_concat_operand(ctx, op, rtc_bases[i], expr_arena, out, &mut rtc_owned)?;
        let result = alloc_string_rtc_temp(ctx);
        concat_owned.push(result);
        out.extend_from_slice(&[0xfb, 0xef]);
        out.extend_from_slice(&result.to_le_bytes());
    }
    out.push(0x60);
    out.push(0x31);
    out.extend_from_slice(&tgt_off.to_le_bytes());
    // Owned temps to free: the rtc result/Missing temps (allocation order) then the
    // concat result temps (allocation order).
    rtc_owned.append(&mut concat_owned);
    emit_variant_temp_free(out, &rtc_owned);
    Ok(())
}

/// Lower `r = Len(<runtime-string call>)` into a Long target: produce the rtc result
/// into a hidden temp, take its length with the release-aware Len opcode
/// (`0xfb 0xeb <scratch>` then `0xfc 0x22`), store the Long, and free the result
/// temp. The Len scratch temp is not freed.
pub(super) fn lower_owned_len(
    ctx: &LowerCtx,
    tgt_off: i16,
    len_node: NodeId,
    expr_arena: &ExprArena,
    out: &mut Vec<u8>,
) -> Result<(), LowerError> {
    let args_id = match expr_arena.get(len_node) {
        ExprNode::Call { args, .. } => *args,
        _ => return Err(LowerError::UnsupportedNode),
    };
    let arg = single_index(args_id, expr_arena).ok_or(LowerError::UnsupportedNode)?;
    let (result_temp, free) = emit_rtc_string_call(ctx, arg, expr_arena, out)?;
    out.push(0x04);
    out.extend_from_slice(&result_temp.to_le_bytes());
    let scratch = alloc_string_rtc_temp(ctx);
    out.extend_from_slice(&[0xfb, 0xeb]);
    out.extend_from_slice(&scratch.to_le_bytes());
    out.extend_from_slice(&[0xfc, 0x22]);
    out.push(0x71);
    out.extend_from_slice(&tgt_off.to_le_bytes());
    emit_variant_temp_free(out, &free);
    Ok(())
}

/// Lower `r = (<a> CMP <b>)` where at least one operand is a runtime-string result.
/// Operands are materialized to BSTR addresses (rtc temps reserved first); the
/// owned marker `0x5d` follows the non-owned operand (none if both are owned); then
/// the compare-into-temp opcode (`0xfb <op> <temp>`) and finalize (`0x55`). The
/// Integer result is stored; only the rtc result temps are freed.
pub(super) fn lower_owned_compare(
    ctx: &LowerCtx,
    tgt_off: i16,
    node: NodeId,
    expr_arena: &ExprArena,
    out: &mut Vec<u8>,
) -> Result<(), LowerError> {
    let (op, lhs, rhs) = match expr_arena.get(node) {
        ExprNode::BinOp { op, lhs, rhs } => (*op, *lhs, *rhs),
        _ => return Err(LowerError::UnsupportedNode),
    };
    let cmp = owned_string_compare_op(op).ok_or(LowerError::UnsupportedNode)?;
    let ops = [lhs, rhs];
    let rtc_bases: Vec<Option<usize>> = ops
        .iter()
        .map(|&o| match ctx.module.builtins.get(&o.0) {
            Some(BuiltinCall::RtcString { args }) => Some(reserve_rtc_temps(ctx, args)),
            _ => None,
        })
        .collect();
    let mut rtc_owned: Vec<i16> = Vec::new();
    for (i, &op_id) in ops.iter().enumerate() {
        materialize_owned_concat_operand(ctx, op_id, rtc_bases[i], expr_arena, out, &mut rtc_owned)?;
        // The owned marker follows the non-owned (plain) operand's materialization.
        if rtc_bases[i].is_none() {
            out.push(0x5d);
        }
    }
    let cmp_temp = alloc_string_rtc_temp(ctx);
    out.extend_from_slice(&[0xfb, cmp]);
    out.extend_from_slice(&cmp_temp.to_le_bytes());
    out.push(0x55);
    out.push(0x70);
    out.extend_from_slice(&tgt_off.to_le_bytes());
    emit_variant_temp_free(out, &rtc_owned);
    Ok(())
}
