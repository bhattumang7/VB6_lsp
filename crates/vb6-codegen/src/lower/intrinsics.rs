use super::*;
use super::argcoerce::{eb_emit_arg_coerce, known_local18_for_grounded_case, ArgCoerceOutcome};
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
    // A `Declare`d external DLL function resolves to `NameResolution::Proc`
    // exactly like an ordinary same-module Sub/Function (both are collected
    // into `module.procs` — see `collect_top_decl`'s `DeclareDecl` arm), but
    // it has a WHOLLY DIFFERENT, unported calling convention (no same-
    // module callsite index applies to an external DLL entry point) —
    // gated rather than silently falling through to the intra-module `0x0a`/
    // `0x5e` convention, which would emit plausible-looking but ungrounded
    // bytes. `body == u32::MAX` is the same "no body" sentinel `collect_
    // top_decl` uses to mark a Declare.
    if ctx.module.procs[proc_idx].body == u32::MAX {
        return Err(LowerError::UnsupportedNode);
    }
    let args: Vec<NodeId> = match expr_arena.get(args_id) {
        ExprNode::ArgList { args } => args.clone(),
        _ => Vec::new(),
    };
    // TODO(not implemented): an omitted Optional argument must push the
    // parameter's default value, which needs the default-value expression
    // carried on the bound parameter — this is a SEMA-LAYER gap (no such
    // field exists on `BoundParam` anywhere in this codebase yet), not
    // something this lowering pass alone can close. Gate any call whose
    // argument count doesn't match the parameter count rather than emit a
    // short argument list.
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
            // TODO(not implemented): UDT (`UserDefined`) and `Array` ByVal are NOT
            // a fixed-size value load — a UDT's size depends on its fields, and an
            // Array is a SAFEARRAY descriptor, neither of which `static_var_size`
            // models (it silently falls back to `4`, which is WRONG for both — a
            // real bug caught this session, not a hypothetical). No oracle capture
            // exists yet for either shape's ByVal argument-passing convention.
            // Gate explicitly here rather than let `static_var_size`'s fallback
            // produce a silently wrong `6c <offset>` load.
            if matches!(pty, Some(VbaType::Variant) | Some(VbaType::UserDefined(_)) | Some(VbaType::Array(_))) {
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

/// Lower a call to a class member's `Sub`/`Function` through the vtable
/// (`o.Method args` or `r = o.Method(args)`). Per-argument convention
/// (oracle-confirmed this pass — `argtype_probe`, `argmix_probe`,
/// `argbyval_lit_probe`, refining the original all-literal-only finding; see
/// the `vb6-class-vtable-slot-rule` memory note and
/// `class_method_arg_needs_staging`'s doc comment for the full derivation):
/// a plain same-type local-variable argument needing no staging pushes its
/// own address (ByRef) or loads its own value directly (ByVal, scalar) — NO
/// copy; anything else (a literal/expression, or any `Object`-typed
/// argument regardless of mode) is materialized into its own pool temp at
/// `class_member_base + <declaration-order stage index>`, staged via `0x59`
/// (scalar) or `fd 9c` (`Object`, refcount-safe). A `Function` in value
/// position (`is_value`) additionally `LdAddr`s a result temp — at
/// `class_member_base + <total staged count>`, i.e. right after every
/// argument's own stage slot — *before* any argument is pushed, and loads
/// the result back *after* the vtable call — via `0x6c` (`Long`), `0x3e`
/// (`String`, a steal-load) or `0x51` (`Object`, a plain pointer read),
/// matching the per-type split already grounded for a class-member Get.
/// Gated to the types whose convention is actually grounded
/// (`class_method_param_is_grounded`) — `Integer`/`Long`/`Object` (both
/// modes) and `String` (ByRef only); a `Function` return in value position is
/// additionally gated to `Long`/`String`/`Object` (oracle-confirmed:
/// `oracle_bank/c5_func_string`, `oracle_bank/c5_func_object`).
///
/// Returns `(pending_string_releases, result_temp_offset)` — both `None`/
/// empty when `is_value` is `false`. `result_temp_offset` lets a caller doing
/// its own post-call coercion (`lower_set_class_get_to_specific_class_target`'s
/// `0x3d`-coercion tail) release the SAME temp this function loaded the
/// result from, without re-deriving its bucket assignment. The string
/// releases are the staged String argument temps still awaiting release, in
/// PARAMETER DECLARATION order — NOT emitted here when `is_value` is `true`.
/// `oracle_bank/c5_func_string` (two `String` args, `String` result, all
/// sharing one region) shows the release must come AFTER the caller's own
/// store of the loaded-back result (`call → result-load(0x3e) → caller's
/// move-store(0x31) → release`), which this function cannot emit itself
/// since the store belongs to the assignment lowering that calls it. When
/// `is_value` is `false` (a `Sub` statement, no result to store) the
/// caller emits the returned releases immediately after this call returns,
/// which is byte-identical to emitting them here (oracle-confirmed:
/// `oracle_bank/c4_sub_2arg_mixed_call`, a single release right after the
/// call). Two or more temps share ONE bulk release (`0x32 <byte-len>
/// <offsets…>`, the same opcode already grounded for multi-temp concat
/// cleanup) rather than one `0x2f` apiece — oracle-confirmed:
/// `oracle_bank/c5_func_string`'s two-argument release.
/// `object_raw_load` selects the Object-return read-back opcode: `false`
/// (every existing caller) is the plain-pointer read `0x51` already grounded
/// for a plain-`Object`-typed consuming target (`oracle_bank/
/// c5_func_object`); `true` is the generic 4-byte load `0x6c`, needed ONLY
/// when the caller itself follows up with the `0x3d`-coercion tail for a
/// SPECIFIC-class-typed consuming target (`oracle_bank/
/// c13_set_call_to_specific_class_target` — byte-identical to the bare-Get
/// analogue, `c11_set_get_to_specific_class_target`, proving a `Set`
/// target's read-back convention depends on the TARGET's type, not on
/// whether the source is a Get or a method call).
pub(super) fn lower_class_method_call(
    ctx: &LowerCtx,
    base: NodeId,
    func_access_id: NodeId,
    args_id: NodeId,
    is_value: bool,
    object_raw_load: bool,
    expr_arena: &ExprArena,
    out: &mut Vec<u8>,
) -> Result<(Vec<i16>, Vec<i16>, Option<i16>), LowerError> {
    let local_idx = match ctx.module.resolutions.get(&base.0) {
        Some(NameResolution::Local { local_idx, .. }) => *local_idx,
        Some(_) => return Err(LowerError::UnsupportedNode),
        None => return Err(LowerError::Unresolved),
    };
    ctx.local_class(local_idx).ok_or(LowerError::UnsupportedNode)?;
    let obj_offset = ctx.local_slots[local_idx].frame_offset;
    let class_sym = match ctx.local_type(local_idx) {
        VbaType::UserDefined(sym) => *sym,
        _ => return Err(LowerError::UnsupportedNode),
    };
    let resolved = ctx
        .module
        .class_member_slots
        .get(&func_access_id.0)
        .ok_or(LowerError::Unresolved)?;
    let method_slot = resolved.method_slot.ok_or(LowerError::UnsupportedNode)?;
    let params = resolved.method_params.clone();

    let args: Vec<NodeId> = match expr_arena.get(args_id) {
        ExprNode::ArgList { args } => args.clone(),
        _ => Vec::new(),
    };
    // TODO(not implemented): same Optional gap as `lower_call` — an omitted
    // argument needs the parameter's default-value expression, which does
    // not exist on `BoundParam` yet (a sema-layer prerequisite, not
    // something this function can add on its own). Reject any arg-count
    // mismatch rather than guess a short/padded argument list.
    if args.len() != params.len() {
        return Err(LowerError::UnsupportedNode);
    }
    if params.iter().any(|(ty, by_val)| !class_method_param_is_grounded(ty, *by_val)) {
        return Err(LowerError::UnsupportedNode);
    }
    // TODO(not implemented): a Variant parameter is only grounded for a
    // plain same-type variable source (see `class_method_param_is_grounded`'s
    // doc comment) — a literal/expression Variant argument's boxing
    // sequence has no oracle capture yet, so it stays gated here rather
    // than guessed.
    for (i, (ty, _)) in params.iter().enumerate() {
        if matches!(ty, VbaType::Variant) {
            let is_plain_var = matches!(expr_arena.get(args[i]), ExprNode::NameRef { .. })
                && matches!(ctx.module.resolutions.get(&args[i].0), Some(NameResolution::Local { .. }))
                && ctx.module.types.get(&args[i].0) == Some(ty);
            if !is_plain_var {
                return Err(LowerError::UnsupportedNode);
            }
        }
    }
    let ret_type = resolved.method_ret_type.clone();
    if is_value
        && !matches!(
            ret_type,
            Some(VbaType::Long) | Some(VbaType::String) | Some(VbaType::Object) | Some(VbaType::Double)
        )
    {
        return Err(LowerError::UnsupportedNode);
    }

    // Assign each argument that actually needs a pool temp its own
    // `(type_ctx, index-within-that-context's-bucket)` — a plain-variable
    // argument needing no staging costs no slot. See `bucket_method_call_
    // slots`'s doc comment for the oracle evidence this per-context
    // bucketing (not one flat proc-wide numbered pool) is grounded on.
    let needs_staging: Vec<bool> = args
        .iter()
        .zip(&params)
        .map(|(&arg, (ty, by_val))| {
            class_method_arg_needs_staging(ctx.module, arg, ty, *by_val, expr_arena)
        })
        .collect();
    let result_ctx = if is_value {
        let ty = ret_type.as_ref().expect("is_value implies a return type (checked above)");
        Some(crate::bridge::type_ctx(ty).ok_or(LowerError::UnsupportedType)?)
    } else {
        None
    };
    let (stage_slot, _buckets, result_slot) =
        bucket_method_call_slots(&params, &needs_staging, result_ctx)?;

    let result_temp = if let Some((ctx_, idx)) = result_slot {
        let off = ctx.class_member_slot(ctx_, idx);
        out.push(0x04);
        out.extend_from_slice(&off.to_le_bytes());
        Some(off)
    } else {
        None
    };
    // Staged String arguments' temps, released after the vtable call — see
    // the staging loop below. Collected in PUSH order (right-to-left, the
    // last parameter first); reversed after the loop back to declaration
    // order, matching the release order oracle-confirmed by `oracle_bank/
    // c5_func_string`'s two-argument release (`x` released before `y`,
    // despite `y` having been staged first).
    let mut string_release_temps: Vec<i16> = Vec::new();
    // Staged `Object` arguments' temps — released via `0x1a <offset>` (the
    // same release opcode already grounded for a Property Set's staged
    // Object temp), one instruction per temp; only a SINGLE staged Object
    // argument is oracle-confirmed (`oracle_bank/c8_obj_byval_param`), so a
    // call staging two or more is gated below rather than guessing a bulk
    // form or a release order.
    let mut object_release_temps: Vec<i16> = Vec::new();
    // A ByRef `Object` argument sourced from a specific-class-typed `As New`
    // local: `(temp_off, class_sym, dest_off)` — the temp's FINAL value
    // (possibly reassigned by the callee through the ByRef parameter) is
    // written back into the source local after the call, then the temp is
    // released. Only a SINGLE such argument is oracle-confirmed
    // (`oracle_bank/c8_obj_byref_param`); two or more in one call is gated
    // (no confirmed relative order).
    let mut object_byref_writebacks: Vec<(i16, u32, i16)> = Vec::new();

    // Push each argument, right-to-left (VB6 evaluation order) — matching
    // the intra-module call convention's own push order. Three distinct
    // shapes (a plain-variable source needs no staging AND no `lower_expr`
    // call; a non-addressable source that needs staging computes its value
    // then stages it; a non-addressable ByVal-scalar source computes its
    // value and pushes it directly, no staging at all — a ByVal literal
    // costs neither a variable's own offset nor a pool temp).
    for i in (0..args.len()).rev() {
        let (ty, by_val) = &params[i];
        let is_plain_var = matches!(expr_arena.get(args[i]), ExprNode::NameRef { .. })
            && matches!(ctx.module.resolutions.get(&args[i].0), Some(NameResolution::Local { .. }))
            && ctx.module.types.get(&args[i].0) == Some(ty);
        if is_plain_var && !needs_staging[i] {
            let off = arg_var_offset(ctx, args[i]).ok_or(LowerError::UnsupportedNode)?;
            if *by_val {
                // `eb_emit_arg_coerce` (the `EbEmitArgCoerce` word-form
                // port, `argcoerce.rs`) is now the ACTUAL byte source for
                // Integer ByVal (`local_18 == 4`, delegation verified at the
                // byte level), not merely a parallel validator: its
                // `Value(node)` is rendered through the same `Emitter` and
                // used DIRECTLY. Scoped to `Integer` specifically, matching
                // `eb_emit_arg_coerce`'s own gate — `Variant` shares the
                // SAME `local_18 == 4` classification but does NOT share
                // this delegation (`lower_expr_coerced` errors on a bare
                // Variant reference; the real shipped bytes are `fc ed
                // <offset>`, a dedicated opcode, see `emit_sized_value_
                // load`'s `16 =>` arm) — this port's own doc comment on the
                // `local_18==4` branch has the full story of catching that
                // exact conflation live, via this cross-check. `emit_sized_
                // value_load` remains the implementation for every type
                // this port doesn't delegate for (`Long`, `Variant`) —
                // falling back keeps existing, independently oracle-
                // verified behavior for those.
                let ported = if matches!(ty, VbaType::Integer)
                    && known_local18_for_grounded_case(ty, true) == Some(4)
                {
                    let mut scratch = NodeArena::new();
                    match eb_emit_arg_coerce(ctx, args[i], ty, true, 0x10, expr_arena, &mut scratch) {
                        Ok(ArgCoerceOutcome::Value(node)) => {
                            let mut emitter = Emitter::new(&scratch);
                            emitter.emit_expr(node, 2);
                            Some(emitter.into_bytes())
                        }
                        other => {
                            debug_assert!(
                                false,
                                "eb_emit_arg_coerce disagreed with the shipped ByVal plain-variable path: {other:?}"
                            );
                            None
                        }
                    }
                } else {
                    None
                };
                match ported {
                    Some(bytes) => {
                        // Still cross-checked against the independently
                        // oracle-verified reference bytes in debug builds —
                        // a divergence here means the port itself is wrong,
                        // not merely unwired.
                        if cfg!(debug_assertions) {
                            let mut shipped_bytes = Vec::new();
                            emit_sized_value_load(static_var_size(ty), off, &mut shipped_bytes);
                            debug_assert_eq!(
                                bytes, shipped_bytes,
                                "eb_emit_arg_coerce's ByVal output diverged from the oracle-verified reference"
                            );
                        }
                        out.extend(bytes);
                    }
                    None => emit_sized_value_load(static_var_size(ty), off, out),
                }
            } else {
                // `eb_emit_arg_coerce` (the `EbEmitArgCoerce` word-form
                // port) is now the ACTUAL OPERAND SOURCE for every `(type,
                // mode)` pair traced end-to-end for the ByRef plain-
                // variable case (`local_18 == 7` — Integer/String/Variant):
                // its `AddressOfOriginal(offset)` carries the exact operand
                // this `04 <offset>` sequence needs, computed by the port
                // itself (via `arg_var_offset`, the SAME function this
                // fallback also calls) rather than merely validated against
                // a locally-recomputed value. Falls back to the locally-
                // computed `off` only for types this port hasn't traced
                // (`Long`) — unchanged, independently oracle-verified
                // behavior for those.
                let ported_off = if known_local18_for_grounded_case(ty, false) == Some(7) {
                    let mut scratch = NodeArena::new();
                    match eb_emit_arg_coerce(ctx, args[i], ty, false, 0x10, expr_arena, &mut scratch) {
                        Ok(ArgCoerceOutcome::AddressOfOriginal(port_off)) => Some(port_off),
                        other => {
                            debug_assert!(
                                false,
                                "eb_emit_arg_coerce disagreed with the shipped ByRef plain-variable path: {other:?}"
                            );
                            None
                        }
                    }
                } else {
                    None
                };
                let emit_off = ported_off.unwrap_or(off);
                debug_assert_eq!(
                    emit_off, off,
                    "eb_emit_arg_coerce's ByRef offset diverged from the oracle-verified reference"
                );
                out.push(0x04);
                out.extend_from_slice(&emit_off.to_le_bytes());
            }
            continue;
        }

        // An `Array` parameter's argument (always a plain array variable —
        // VB6 has no array literals) stages the ARRAY VARIABLE'S OWN
        // ADDRESS (`04 <src>`, the same `LdAddr` any ByRef argument's source
        // uses) into its scratch temp via the generic scalar store (`0x59`,
        // NOT a dedicated array opcode) — no further push follows, matching
        // every other non-`String` grounded type. Oracle-confirmed:
        // `oracle_bank/c19_array_arg_ref` (single Array argument) and
        // `c19_array_arg_mixed_with_long` (an Array argument sharing its
        // scratch region with an unrelated `Long` argument in the same
        // proc — proves the shared `type_ctx` mapping, not just the byte
        // shape). ByVal is ungrounded (no VB6 array-literal source exists
        // to probe it with) — gated via `class_method_param_is_grounded`.
        if let VbaType::Array(_) = ty {
            let src_off = arg_var_offset(ctx, args[i]).ok_or(LowerError::UnsupportedNode)?;
            out.push(0x04);
            out.extend_from_slice(&src_off.to_le_bytes());
            let (arg_ctx, arg_idx) = stage_slot[i].expect("Array args always need staging");
            let temp_off = ctx.class_member_slot(arg_ctx, arg_idx);
            out.push(0x59);
            out.extend_from_slice(&temp_off.to_le_bytes());
            continue;
        }

        // An `Object`-typed parameter whose argument is a SPECIFIC-class-typed
        // `As New` local (`Dim y As New Class1`, passed where the parameter
        // itself is declared plain `Object`) reads via the SAME lazy-fetch
        // sequence already grounded for `Set o = otherObjLocal`
        // (`plain_object_local`/`lower_set_plain_object_local`'s `NameRef`
        // arm): `04 <src>` (LdAddr) then `56 <create-idx>` (construct-if-null,
        // already-owned push) — NOT the generic `lower_expr_coerced` pipeline,
        // which has no notion of coercing a `UserDefined`-typed variable into
        // an `Object`-tagged node and would error `UnsupportedType`.
        if matches!(ty, VbaType::Object) {
            if let Some((src_class, src_off, true)) = plain_object_local(ctx, args[i], expr_arena) {
                let idx = ctx.intern_class_const(ClassConstKind::Create, src_class);
                out.push(0x04);
                out.extend_from_slice(&src_off.to_le_bytes());
                out.push(0x56);
                out.extend_from_slice(&idx.to_le_bytes());
                let (arg_ctx, arg_idx) = stage_slot[i].expect("Object args always need staging");
                let temp_off = ctx.class_member_slot(arg_ctx, arg_idx);
                if *by_val {
                    // The loaded value feeds the SAME `fd 9c` staging store as
                    // any other ByVal Object argument (below) — only the
                    // VALUE computation differs, not the consumer. Released
                    // afterward (`0x1a`, tracked in `object_release_temps`).
                    // Oracle-confirmed: `oracle_bank/c8_obj_byval_param`.
                    out.push(0xfd);
                    out.push(0x9c);
                    out.extend_from_slice(&temp_off.to_le_bytes());
                    object_release_temps.push(temp_off);
                } else {
                    // ByRef is a genuinely different shape: the lazy-fetched
                    // value is MOVE-STORED (`fc f8`, no AddRef — the same
                    // "steal" store already grounded for `Set o =
                    // otherObjLocal`) into the temp, and the temp's ADDRESS
                    // (not a staged value) is what's actually pushed as the
                    // argument — the callee can reassign through a ByRef
                    // parameter, so the temp's FINAL value must be written
                    // BACK into the source local after the call (`object_
                    // byref_writebacks`, emitted post-call below) rather than
                    // simply released. Oracle-confirmed: `oracle_bank/
                    // c8_obj_byref_param`.
                    out.push(0xfc);
                    out.push(0xf8);
                    out.extend_from_slice(&temp_off.to_le_bytes());
                    out.push(0x04);
                    out.extend_from_slice(&temp_off.to_le_bytes());
                    object_byref_writebacks.push((temp_off, src_class, src_off));
                }
                continue;
            }
        }

        let coerce = vba_type_to_node_tag(ty);
        let mut arena = NodeArena::new();
        let root = lower_expr_coerced(ctx, args[i], expr_arena, &mut arena, coerce)?;
        let root = coerce_assign_value(ctx, args[i], root, coerce, &mut arena);
        let mut emitter = Emitter::new(&arena);
        emitter.emit_expr(root, 2);
        let shipped_bytes = emitter.into_bytes();

        // `eb_emit_arg_coerce`, Object-ByVal side (`local_18 == 9`): scoped
        // exactly to a plain same-type variable source — the ONLY shape
        // this session traced `EbCheckSetBinding`/`EbEmitPropertyExpr`/
        // `EbNormalizeTypeReference` for (see `eb_normalize_type_reference_
        // object_case_is_noop_for_plain_var`'s doc comment) — matching
        // `class_method_arg_needs_staging`'s own unconditional `true` for
        // `Object`, this branch is where a plain Object variable ByVal
        // argument actually lands (staging always applies, so it never
        // hits the fast-path branch above). When grounded, its `Value`
        // node IS the bytes used — not merely cross-checked — with a
        // debug-only equality assertion against the independently oracle-
        // verified `shipped_bytes` kept as a safety net (any divergence
        // means the port itself regressed, not that it's simply unwired).
        // Every other source shape (non-plain-variable, non-Object, or a
        // literal/expression argument) keeps the existing `lower_expr_
        // coerced`+`coerce_assign_value` pipeline unchanged — this port
        // has not traced those shapes and must not silently substitute for
        // them.
        let ported_object_bytes = if matches!(ty, VbaType::Object) && *by_val && is_plain_var {
            let mut scratch = NodeArena::new();
            match eb_emit_arg_coerce(ctx, args[i], ty, true, 0x10, expr_arena, &mut scratch) {
                Ok(ArgCoerceOutcome::Value(node)) => {
                    let mut port_emitter = Emitter::new(&scratch);
                    port_emitter.emit_expr(node, 2);
                    Some(port_emitter.into_bytes())
                }
                other => {
                    debug_assert!(
                        false,
                        "eb_emit_arg_coerce disagreed with the shipped Object-ByVal plain-variable path: {other:?}"
                    );
                    None
                }
            }
        } else {
            None
        };
        match ported_object_bytes {
            Some(bytes) => {
                if cfg!(debug_assertions) {
                    debug_assert_eq!(
                        bytes, shipped_bytes,
                        "eb_emit_arg_coerce's Object-ByVal output diverged from the oracle-verified reference"
                    );
                }
                out.extend(bytes);
            }
            None => out.extend(shipped_bytes),
        }

        if !needs_staging[i] {
            // ByVal, non-addressable source, scalar type: push the computed
            // value directly — no staging (`argbyval_lit_probe`).
            continue;
        }
        let (arg_ctx, arg_idx) = stage_slot[i].expect("needs_staging[i] implies a stage slot");
        let temp_off = ctx.class_member_slot(arg_ctx, arg_idx);
        if matches!(ty, VbaType::Object) {
            // NO release afterward — unlike the `As New`-class-typed-source
            // lazy-fetch shape above, a plain `Object`-typed variable's
            // staged value needs no post-call cleanup. Oracle-confirmed:
            // `e2e_class_method_arg_variable` (`o.TakeObj y`, `y As Object`).
            out.push(0xfd);
            out.push(0x9c);
            out.extend_from_slice(&temp_off.to_le_bytes());
        } else if matches!(ty, VbaType::String) {
            // A staged String argument uses the SAME copy-store-then-by-
            // address convention already grounded for a Property Let's
            // String staging (`lower_class_field_store`'s `is_string`
            // branch): the value is COPY-STORED (`0x43`, properly owning/
            // addref'ing the BSTR) into the temp, and the temp's OWN
            // ADDRESS (`0x04`) is what's actually pushed as the argument —
            // not a staged value like every other type. The temp is
            // released after the call (`0x2f`, see below). Oracle-confirmed:
            // `oracle_bank/c4_sub_2arg_mixed_call`.
            out.push(0x43);
            out.extend_from_slice(&temp_off.to_le_bytes());
            out.push(0x04);
            out.extend_from_slice(&temp_off.to_le_bytes());
            string_release_temps.push(temp_off);
        } else if matches!(ty, VbaType::Double) {
            // A staged Double argument uses the SAME FPU-aware store already
            // grounded for a Property Let's Double staging
            // (`lower_class_field_store`'s `is_double` branch, `0xfd 0xc9`) —
            // and, like that branch, pushes NO separate address afterward
            // (unlike `String`'s explicit `0x04` follow-up push): the vtable
            // call's own instruction stream never references the temp's
            // address again, matching Property Let's identical no-follow-up
            // shape byte-for-byte. Oracle-confirmed: `oracle_bank/
            // c6_nonlong_arg_and_return` (recompiled and re-extracted fresh
            // this session after finding the originally-banked capture was
            // truncated by a wrong preamble length — see its own `notes.md`
            // for the corrected 34-byte capture).
            out.push(0xfd);
            out.push(0xc9);
            out.extend_from_slice(&temp_off.to_le_bytes());
        } else {
            out.push(0x59);
            out.extend_from_slice(&temp_off.to_le_bytes());
        }
    }

    out.push(0x04);
    out.extend_from_slice(&(obj_offset as u16).to_le_bytes());
    out.push(0x24);
    out.extend_from_slice(&ctx.intern_class_const(ClassConstKind::Create, class_sym).to_le_bytes());
    out.push(0x0d);
    out.extend_from_slice(&method_slot.to_le_bytes());
    // The vtable-call operand's own member-type-descriptor const-pool entry
    // (`ModuleConstEntry::MemberType`) — keyed by the CALLEE'S CLASS (see its
    // own doc comment for the full derivation, including the two-class
    // capture that corrected the earlier "single per-module shared entry"
    // model), extended here to method calls (`Sub` and `Function` alike,
    // regardless of return type or argument shape). Was previously
    // hardcoded `01 00`, which only ever coincidentally matched every
    // single-call-site fixture shipped before the #7/#9 slice (`c4_sub_
    // 2arg_mixed_call`'s oracle capture shows index 2, not 1, proving the
    // operand must be interned, not hardcoded).
    out.extend_from_slice(&ctx.intern_member_type_const(class_sym).to_le_bytes());
    string_release_temps.reverse();
    object_release_temps.reverse();

    // A `Sub` statement call (no result to store) releases its staged String
    // argument temps immediately — there is nothing else to interleave them
    // with, so emitting here is byte-identical to the caller doing it right
    // after this call returns. A `Function` call in value position (`is_value`)
    // defers release to its caller instead (see the returned `Vec`'s own doc
    // comment above) — the release must come AFTER the caller's store of the
    // loaded-back result, which this function has no way to emit itself.
    if !is_value {
        emit_temp_release_list(&string_release_temps, out);
    }

    // A staged `Object` argument's temp is released via `0x1a <offset>` for a
    // SINGLE temp — the same opcode already grounded for a Property Set's
    // staged Object temp — or a bulk `0x29` for two or more (a DIFFERENT
    // opcode from String's `0x32`, same declaration-order convention). For a
    // `Sub` statement (`!is_value`) this is emitted right here, immediately
    // after the call — matching the String release's placement. For a
    // `Function` in value position, the release is DEFERRED to the caller
    // instead, the same way `string_release_temps` already is — oracle-
    // confirmed for both a single Object release (`oracle_bank/
    // c15_func_with_obj_arg_release`) and two (`oracle_bank/
    // c17_two_obj_release_value`, a Long-returning `Function` with two
    // `ByVal Object` arguments): the byte order is `call → result-load →
    // caller's own store → release` either way, and the bulk-release form
    // (`0x29`) and declaration-order convention are UNCHANGED in value
    // position — the deferred release is byte-identical in shape to the
    // `!is_value` case, just moved later in the stream.
    if !is_value {
        emit_object_temp_release_list(&object_release_temps, out);
    }

    // A ByRef Object argument sourced from a specific-class-typed `As New`
    // local writes the temp's (possibly callee-reassigned) final value BACK
    // into the source local: `6c <temp>` (plain 4-byte load), `3d
    // <type-idx>` (coerce/type-check to the source's declared class — the
    // SAME opcode already grounded for `Set o = Nothing`'s typed-Nothing
    // coercion, here reused for a typed re-check instead), `19 <dest>`
    // (AddRef-store into the source local), then `1a <temp>` (release the
    // temp) — a different release point from the ByVal case (after the
    // write-back, not immediately after the call). Only a SINGLE such
    // argument, `is_value = false`, is oracle-confirmed: `oracle_bank/
    // c8_obj_byref_param`.
    if !object_byref_writebacks.is_empty() {
        if is_value || object_byref_writebacks.len() > 1 {
            return Err(LowerError::UnsupportedNode);
        }
        let (temp_off, class_sym, dest_off) = object_byref_writebacks[0];
        out.push(0x6c);
        out.extend_from_slice(&temp_off.to_le_bytes());
        // Reuses the SAME `MemberType`-const mechanism as the call's own
        // operand (`intern_member_type_const`, now confirmed class-keyed —
        // see its doc comment), keyed by the ARGUMENT's own class, not a
        // fresh `ClassConstKind::TypeDesc` entry. In this capture the
        // argument's class (`y As New Class1`) happens to equal the callee's
        // own class (`Class1.Use`), so this dedups to the SAME index as the
        // call's own operand — that's exactly why they matched, not a
        // coincidence or an open question. Oracle-confirmed: `oracle_bank/
        // c8_obj_byref_param`.
        let type_idx = ctx.intern_member_type_const(class_sym);
        out.push(0x3d);
        out.extend_from_slice(&type_idx.to_le_bytes());
        out.push(0x19);
        out.extend_from_slice(&dest_off.to_le_bytes());
        out.push(0x1a);
        out.extend_from_slice(&temp_off.to_le_bytes());
    }

    if let Some(off) = result_temp {
        // The result temp's read-back opcode follows the same per-type split
        // already grounded for a class-member Get (`ClassFieldRef::is_string`/
        // `is_object`): `String` steals the temp's BSTR pointer (`0x3e`, no
        // separate release — the temp is left zeroed) and `Object` reads a
        // plain 4-byte pointer (`0x51`); every other grounded return type
        // (`Long`'s `0x6c`, `Double`'s `0x6f`) uses the ordinary
        // `RT_LOAD_BY_CTX`-style load. Oracle-confirmed: `oracle_bank/
        // c5_func_string`, `oracle_bank/c5_func_object` (both recaptured this
        // session, HIGH confidence), `oracle_bank/c6_nonlong_arg_and_return`
        // (Double, also recaptured this session after finding the banked
        // capture truncated).
        let load_op = match ret_type {
            Some(VbaType::String) => 0x3e,
            Some(VbaType::Object) if object_raw_load => 0x6c,
            Some(VbaType::Object) => 0x51,
            Some(ref ty) => {
                let load_ctx = crate::bridge::load_store_ctx(ty).ok_or(LowerError::UnsupportedType)?;
                let op = *crate::tables::RT_LOAD_BY_CTX.get(load_ctx).ok_or(LowerError::UnsupportedType)?;
                if op == 0 {
                    return Err(LowerError::UnsupportedType);
                }
                op
            }
            None => return Err(LowerError::UnsupportedNode),
        };
        out.push(load_op);
        out.extend_from_slice(&off.to_le_bytes());
    }
    Ok((
        if is_value { string_release_temps } else { Vec::new() },
        if is_value { object_release_temps } else { Vec::new() },
        if is_value { result_temp } else { None },
    ))
}

/// Release a set of scratch temps: nothing for an empty list, a single
/// `0x2f <offset>` for exactly one, or one bulk `0x32 <byte-len> <offsets…>`
/// for two or more — the same convention already grounded for multi-temp
/// concat cleanup (`lower_concat`'s own release code). Oracle-confirmed for
/// the class-method-call case: `oracle_bank/c5_func_string`'s two-argument
/// release.
pub(super) fn emit_temp_release_list(temps: &[i16], out: &mut Vec<u8>) {
    match temps.len() {
        0 => {}
        1 => {
            out.push(0x2f);
            out.extend_from_slice(&temps[0].to_le_bytes());
        }
        n => {
            out.push(0x32);
            out.extend_from_slice(&((n * 2) as u16).to_le_bytes());
            for t in temps {
                out.extend_from_slice(&t.to_le_bytes());
            }
        }
    }
}

/// Release a set of staged `Object` argument temps: nothing for an empty
/// list, a single `0x1a <offset>` for exactly one, or one bulk `0x29
/// <byte-len> <offsets…>` for two or more — a DIFFERENT opcode from
/// `emit_temp_release_list`'s String-oriented `0x2f`/`0x32` (same operand
/// shape, dedicated opcode). Oracle-confirmed: `oracle_bank/
/// c14_two_object_args_release`.
pub(super) fn emit_object_temp_release_list(temps: &[i16], out: &mut Vec<u8>) {
    match temps.len() {
        0 => {}
        1 => {
            out.push(0x1a);
            out.extend_from_slice(&temps[0].to_le_bytes());
        }
        n => {
            out.push(0x29);
            out.extend_from_slice(&((n * 2) as u16).to_le_bytes());
            for t in temps {
                out.extend_from_slice(&t.to_le_bytes());
            }
        }
    }
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
