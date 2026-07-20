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

/// Resolve `node_id` (a `SetAssign` target or a bare-identifier value) to a
/// plain object-typed local — a `Dim x As ClassName` / `Dim x As New
/// ClassName` local referenced by its bare name, as opposed to `o.Field`
/// (routed through `lower_class_field_set`/`_store` instead). Returns the
/// local's class symbol, frame offset, and whether it was declared `As New`
/// (which changes how a READ of it — as a `Set` source — must be lowered:
/// see `lower_set_plain_object_local`'s `NameRef` arm).
pub(super) fn plain_object_local(
    ctx: &LowerCtx,
    node_id: NodeId,
    expr_arena: &ExprArena,
) -> Option<(u32, i16, bool)> {
    let ExprNode::NameRef { .. } = expr_arena.get(node_id) else {
        return None;
    };
    match ctx.module.resolutions.get(&node_id.0) {
        Some(NameResolution::Local { local_idx, .. }) => {
            let local = ctx.proc.locals.get(*local_idx)?;
            let VbaType::UserDefined(sym) = &local.vba_type else {
                return None;
            };
            if !ctx.module.class_field_info.contains_key(sym) {
                return None;
            }
            Some((*sym, ctx.local_slots[*local_idx].frame_offset, local.is_new))
        }
        _ => None,
    }
}

/// One class-member scratch region — a reusable slot area for exactly one
/// frame type-context (`crate::bridge::type_ctx`'s index space), sized to
/// the maximum number of CONCURRENT slots of that context ever needed at
/// once in this proc. Regions are allocated in FIRST-ENCOUNTERED order
/// (source order).
///
/// An earlier pass of this port modeled this as a SINGLE shared temp for the
/// whole proc, gating loudly the moment two distinct type-contexts appeared
/// — that turned out to be wrong in general, not merely an unfinished
/// generalization. Two fresh oracle captures this session directly disprove
/// it: a `Double`-returning Property `Get` followed by a no-return `Sub`
/// call (one proc) lands on TWO separate, non-overlapping frame offsets —
/// not one shared slot; a `Long` field `Get` followed by a `String`
/// property `Get` (one proc) ALSO lands on two separate offsets despite
/// both being 4-byte-wide contexts, proving the split is genuinely per
/// TYPE-CONTEXT, not per byte-width. `e2e_class_multi_field_and_property`'s
/// six SAME-typed (`Long`) accesses still correctly share ONE slot — that
/// finding stands, it just never had a second, DIFFERENTLY-typed access in
/// the same proc to reveal the type-independence.
pub(super) struct ClassMemberRegion {
    pub type_ctx: usize,
    pub slots: usize,
}

/// Assign each staged argument of one class-method call its own
/// `(type_ctx, index-within-that-context's-bucket)`, in PARAMETER
/// DECLARATION order (bucket-creation order — oracle-confirmed:
/// `oracle_bank/c4_sub_2arg_mixed_call`'s `DoIt(x As Long, s As String)`
/// creates the Long bucket before the String bucket, matching declaration
/// order, not right-to-left push order). `result_ctx` (a `Function` call in
/// value position) claims one more slot in ITS OWN return type's bucket,
/// positioned AFTER that bucket's own argument slots — oracle-confirmed:
/// `oracle_bank/c5_func_string`'s result temp and its two `String`
/// arguments all share the one String bucket; the result's frame offset is
/// SMALLER than either argument's (frame offsets decrement per allocation),
/// meaning it was allocated LAST, i.e. after the args. Returns the
/// per-argument assignment (`None` for an arg needing no staging), each
/// bucket's final size in creation order, and the result's own assignment.
pub(super) fn bucket_method_call_slots(
    params: &[(VbaType, bool)],
    needs_staging: &[bool],
    result_ctx: Option<usize>,
) -> Result<(Vec<Option<(usize, usize)>>, Vec<(usize, usize)>, Option<(usize, usize)>), LowerError> {
    fn claim(ctx: usize, order: &mut Vec<usize>, counts: &mut Vec<usize>) -> usize {
        let bucket = match order.iter().position(|&c| c == ctx) {
            Some(i) => i,
            None => {
                order.push(ctx);
                counts.push(0);
                order.len() - 1
            }
        };
        let idx = counts[bucket];
        counts[bucket] += 1;
        idx
    }
    let mut order: Vec<usize> = Vec::new();
    let mut counts: Vec<usize> = Vec::new();
    let mut assign = Vec::with_capacity(params.len());
    for ((ty, _), &needs) in params.iter().zip(needs_staging) {
        if !needs {
            assign.push(None);
            continue;
        }
        let ctx = crate::bridge::type_ctx(ty).ok_or(LowerError::UnsupportedType)?;
        assign.push(Some((ctx, claim(ctx, &mut order, &mut counts))));
    }
    let result = result_ctx.map(|ctx| (ctx, claim(ctx, &mut order, &mut counts)));
    Ok((assign, order.into_iter().zip(counts).collect(), result))
}

/// Compute the proc's whole set of class-member scratch regions (see
/// `ClassMemberRegion`) in one tree walk: a Get access (`x = o.F`/`x = o.P`/
/// `Set x = o.P`), a Property-Let staging target, a Property-Set staging
/// target (always the `Object` context — `crate::bridge::type_ctx`'s `0`),
/// and every class-method call's own argument/result buckets
/// (`bucket_method_call_slots`), each merged into the running per-context
/// max (not summed — the same physical region is reused across every call
/// site needing that context, exactly like the already-grounded same-type
/// repeated-Get case).
pub(super) fn class_member_regions(
    module: &BoundModule,
    node_id: NodeId,
    expr_arena: &ExprArena,
) -> Result<Vec<ClassMemberRegion>, LowerError> {
    fn merge(order: &mut Vec<usize>, counts: &mut Vec<usize>, ctx: usize, need: usize) {
        match order.iter().position(|&c| c == ctx) {
            Some(i) => {
                if need > counts[i] {
                    counts[i] = need;
                }
            }
            None => {
                order.push(ctx);
                counts.push(need);
            }
        }
    }
    fn note_single_slot(
        module: &BoundModule,
        ty_node: NodeId,
        order: &mut Vec<usize>,
        counts: &mut Vec<usize>,
    ) -> Result<(), LowerError> {
        let ty = module.types.get(&ty_node.0).ok_or(LowerError::Unresolved)?;
        let ctx = crate::bridge::type_ctx(ty).ok_or(LowerError::UnsupportedType)?;
        merge(order, counts, ctx, 1);
        Ok(())
    }
    fn note_call(
        module: &BoundModule,
        func_id: NodeId,
        args_id: NodeId,
        expr_arena: &ExprArena,
        is_value: bool,
        order: &mut Vec<usize>,
        counts: &mut Vec<usize>,
    ) -> Result<(), LowerError> {
        let ExprNode::MemberAccess { base, .. } = expr_arena.get(func_id) else {
            return Ok(());
        };
        if !member_access_base_is_class(module, *base) {
            return Ok(());
        }
        let Some(resolved) = module.class_member_slots.get(&func_id.0) else {
            return Ok(());
        };
        if resolved.method_slot.is_none() {
            return Ok(());
        }
        let args: Vec<NodeId> = match expr_arena.get(args_id) {
            ExprNode::ArgList { args } => args.clone(),
            _ => Vec::new(),
        };
        if args.len() != resolved.method_params.len() {
            return Ok(());
        }
        let needs_staging: Vec<bool> = args
            .iter()
            .zip(&resolved.method_params)
            .map(|(&a, (ty, by_val))| class_method_arg_needs_staging(module, a, ty, *by_val, expr_arena))
            .collect();
        let result_ctx = if is_value {
            match &resolved.method_ret_type {
                Some(ty) => Some(crate::bridge::type_ctx(ty).ok_or(LowerError::UnsupportedType)?),
                None => None,
            }
        } else {
            None
        };
        let (_, buckets, _) =
            bucket_method_call_slots(&resolved.method_params, &needs_staging, result_ctx)?;
        for (ctx, count) in buckets {
            merge(order, counts, ctx, count);
        }
        Ok(())
    }
    fn walk(
        module: &BoundModule,
        node_id: NodeId,
        expr_arena: &ExprArena,
        order: &mut Vec<usize>,
        counts: &mut Vec<usize>,
    ) -> Result<(), LowerError> {
        let mut recurse =
            |id: NodeId, order: &mut Vec<usize>, counts: &mut Vec<usize>| walk(module, id, expr_arena, order, counts);
        match expr_arena.get(node_id) {
            ExprNode::Assign { target, value } => {
                if let ExprNode::MemberAccess { base, .. } = expr_arena.get(*value) {
                    if member_access_base_is_class(module, *base) {
                        note_single_slot(module, *value, order, counts)?;
                    }
                }
                if let ExprNode::Call { func, args } = expr_arena.get(*value) {
                    note_call(module, *func, *args, expr_arena, true, order, counts)?;
                }
                if let ExprNode::MemberAccess { base, .. } = expr_arena.get(*target) {
                    if member_access_base_is_class(module, *base)
                        && module.class_member_slots.get(&target.0).map(|r| r.is_property).unwrap_or(false)
                    {
                        note_single_slot(module, *target, order, counts)?;
                    }
                }
                Ok(())
            }
            // `Set x = o.P`/`Set x = o.Method()`: same Get-temp sizing
            // requirement as a plain `Assign` RHS. `Set o.P = v`/
            // `Set o.Field = v`: the Property-Set call always stages an
            // `Object` (`fd 9c`), regardless of the target's own declared
            // type — its context is unconditionally `0`.
            ExprNode::SetAssign { target, value } => {
                if let ExprNode::MemberAccess { base, .. } = expr_arena.get(*value) {
                    if member_access_base_is_class(module, *base) {
                        note_single_slot(module, *value, order, counts)?;
                    }
                }
                if let ExprNode::Call { func, args } = expr_arena.get(*value) {
                    note_call(module, *func, *args, expr_arena, true, order, counts)?;
                }
                if let ExprNode::MemberAccess { base, .. } = expr_arena.get(*target) {
                    if member_access_base_is_class(module, *base)
                        && module.class_member_slots.get(&target.0).map(|r| r.set_slot.is_some()).unwrap_or(false)
                    {
                        merge(order, counts, 0, 1);
                    }
                }
                Ok(())
            }
            ExprNode::CallStmt { callee, args } => {
                let (func_id, args_id) = match expr_arena.get(*callee) {
                    ExprNode::Call { func, args: inner } => (*func, *inner),
                    _ => (*callee, *args),
                };
                note_call(module, func_id, args_id, expr_arena, false, order, counts)
            }
            ExprNode::Block { stmts } => {
                for &id in stmts {
                    recurse(id, order, counts)?;
                }
                Ok(())
            }
            ExprNode::If { then_body, else_body, .. } => {
                recurse(*then_body, order, counts)?;
                if let Some(id) = else_body {
                    recurse(*id, order, counts)?;
                }
                Ok(())
            }
            ExprNode::While { body, .. } | ExprNode::Do { body, .. } | ExprNode::For { body, .. } => {
                recurse(*body, order, counts)
            }
            ExprNode::SelectCase { cases, .. } => {
                for &id in cases {
                    recurse(id, order, counts)?;
                }
                Ok(())
            }
            ExprNode::CaseBlock { body, .. } | ExprNode::CaseElse { body } => recurse(*body, order, counts),
            _ => Ok(()),
        }
    }
    let mut order = Vec::new();
    let mut counts = Vec::new();
    walk(module, node_id, expr_arena, &mut order, &mut counts)?;
    Ok(order
        .into_iter()
        .zip(counts)
        .map(|(type_ctx, slots)| ClassMemberRegion { type_ctx, slots })
        .collect())
}

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
/// expression). `Byte`/`Boolean`/`Single`/`Double`/`Currency` (all five, both
/// modes, plain-variable source): oracle-captured directly, one dedicated
/// 2-call probe per type (`re_lab/pcode_lab/argbyte_probe/`, `argbool_
/// probe/`, `argsingle_probe/`, `argdouble_probe/`, `argcurrency_probe/` —
/// a SIX-call single probe was tried first and produced garbage, `00 14`;
/// `capture_pcode.py`'s backward-scan window is a fixed 128 bytes from
/// `ProcDsc`, too small for six calls' worth of body — splitting into
/// 2-call probes, matching every prior successful capture's shape, fixed
/// it). Each confirms the SAME already-established pattern with no new
/// opcodes: ByRef is a plain `04 <offset>` address push; ByVal is exactly
/// what `emit_sized_value_load(static_var_size(ty), ...)` already produces
/// (`Byte`→`fc e0` size 1; `Boolean`→`6b` size 2, same opcode as `Integer`;
/// `Single`→`6c` size 4, same opcode as `Long`/`String`; `Double`→`6d` size
/// 8, SAME opcode as `Currency`) — no staging in any of them.
///
/// **A specific claim in an earlier version of this comment was WRONG and
/// is corrected here, not just updated**: it asserted `emit_sized_value_
/// load`'s 8-byte case was a known bug (hardcoding Currency's `0x6d` for
/// Double too, which it claimed should be `0x6f`) — that assumption came
/// from `RT_LOAD_BY_CTX` (`tables.rs`), an already-oracle-confirmed table
/// for a DIFFERENT context (general expression loads, e.g. `r = d`
/// arithmetic/assignment) where Double genuinely does use `0x6f`, distinct
/// from Currency's `0x6d`. Assuming that distinction carried over to THIS
/// context (ByVal class-method-call argument passing) without checking was
/// exactly the kind of unverified extrapolation this project's discipline
/// exists to prevent — direct oracle capture of `argdouble_probe`/
/// `argcurrency_probe` settled it: BOTH emit `6d <offset>` here, no
/// distinction. The two contexts use genuinely different opcode schemes —
/// argument-passing is purely SIZE-based (the callee's own parameter
/// descriptor handles subtype interpretation), general expression loading
/// is type-based (arithmetic needs the exact subtype) — and conflating them
/// was the actual error, not a bug in `emit_sized_value_load`. `Date` (both
/// modes, plain-variable source): oracle-captured directly (`argdate_
/// probe`) rather than assumed from `Double`'s pattern despite sharing its
/// 8-byte size — ByVal emits `6d <offset>`, IDENTICAL to `Double`/
/// `Currency`, confirming (not assuming) the same size-based scheme applies.
/// TODO(not implemented): `UDT`/`Array` remain gated (excluded from the
/// `matches!` below, so `LowerError::UnsupportedNode`) — no addressability/
/// staging convention has been captured for either. Closing this needs a
/// dedicated oracle probe per shape (a UDT ByVal/ByRef pair, an Array ByVal/
/// ByRef pair) the same way every scalar type above was closed this
/// session, plus (for ByVal specifically) a real per-shape size/copy
/// convention — `static_var_size`'s fixed-size model does not apply to
/// either (see its own doc comment).
pub(super) fn class_method_param_is_grounded(ty: &VbaType, _by_val: bool) -> bool {
    matches!(
        ty,
        VbaType::Integer
            | VbaType::Long
            | VbaType::Object
            | VbaType::String
            | VbaType::Variant
            | VbaType::Byte
            | VbaType::Boolean
            | VbaType::Single
            | VbaType::Double
            | VbaType::Currency
            | VbaType::Date
    )
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
/// TODO(not implemented): the `_ => 4` fallback below is WRONG for
/// `VbaType::UserDefined` (a UDT's size depends on its fields; no oracle
/// capture exists for the actual layout convention) and `VbaType::Array`
/// (a SAFEARRAY descriptor, not a fixed 4 bytes either). Every
/// argument-passing call site that can reach an arbitrary `VbaType` gates
/// these two explicitly before calling this function (see `lower_call`'s
/// and the `RtcNumeric` intrinsic's own `UserDefined`/`Array` checks in
/// `intrinsics.rs`/`expr.rs`) — but this function itself has no `Result`
/// to signal the gap, so a `debug_assert` catches any NEW call site that
/// forgets to gate first, rather than staying silent.
pub(super) fn static_var_size(ty: &VbaType) -> u16 {
    debug_assert!(
        !matches!(ty, VbaType::UserDefined(_) | VbaType::Array(_)),
        "static_var_size called with an unsized type ({ty:?}) — the caller must gate \
         UDT/Array before reaching here, not rely on this function's fallback"
    );
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
