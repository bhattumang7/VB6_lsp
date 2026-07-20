use super::*;
use super::assign::*;
use super::decl::*;
use super::expr::*;
use super::intrinsics::*;


// ── Statement lowering ────────────────────────────────────────────────────────

pub(super) fn lower_block(
    ctx: &LowerCtx,
    node_id: NodeId,
    expr_arena: &ExprArena,
    out: &mut Vec<u8>,
) -> Result<(), LowerError> {
    match expr_arena.get(node_id) {
        ExprNode::Block { stmts } => {
            let ids: Vec<NodeId> = stmts.clone();
            for id in ids {
                lower_tracked_stmt(ctx, id, expr_arena, out)?;
            }
        }
        _ => lower_tracked_stmt(ctx, node_id, expr_arena, out)?,
    }
    Ok(())
}

/// Lower a statement, prefixing it with a line-table marker when the procedure
/// needs line tracking and the statement emits code. Declarations (`Dim`/`Const`)
/// and line labels emit no code and carry no marker; a numeric line label records
/// its byte offset as the position of the *next* statement's marker, which is why
/// the marker must be threaded here rather than inside the statement.
pub(super) fn lower_tracked_stmt(
    ctx: &LowerCtx,
    id: NodeId,
    expr_arena: &ExprArena,
    out: &mut Vec<u8>,
) -> Result<(), LowerError> {
    let no_code = matches!(
        expr_arena.get(id),
        ExprNode::DimItem { .. } | ExprNode::Label { .. } | ExprNode::Block { .. }
    );
    if ctx.line_tracking && !no_code {
        ctx.line_markers.borrow_mut().push(out.len());
        out.extend_from_slice(&[0x00, 0x00]);
    }
    lower_stmt(ctx, id, expr_arena, out)
}

/// Whether a procedure body needs the statement line-number table: it contains a
/// numeric line label, a `Resume`, or `On Error Resume Next` — the constructs that
/// require runtime line tracking (for `Erl` / `Resume`).
pub(super) fn proc_needs_line_tracking(body: NodeId, expr_arena: &ExprArena) -> bool {
    let mut found = false;
    fn walk(id: NodeId, expr_arena: &ExprArena, found: &mut bool) {
        if *found {
            return;
        }
        match expr_arena.get(id) {
            ExprNode::Label { target: LabelRef::Line(_) } => *found = true,
            ExprNode::Resume { .. } => *found = true,
            ExprNode::OnError { kind: OnErrorKind::ResumeNext } => *found = true,
            other => {
                let mut kids = Vec::new();
                other.for_each_child(&mut |c| kids.push(c));
                for c in kids {
                    walk(c, expr_arena, found);
                }
            }
        }
    }
    walk(body, expr_arena, &mut found);
    found
}

pub(super) fn lower_stmt(
    ctx: &LowerCtx,
    node_id: NodeId,
    expr_arena: &ExprArena,
    out: &mut Vec<u8>,
) -> Result<(), LowerError> {
    match expr_arena.get(node_id) {
        ExprNode::Assign { target, value } => {
            let (tgt, val) = (*target, *value);
            lower_assign(ctx, tgt, val, expr_arena, out)
        }
        ExprNode::SetAssign { target, value } => {
            let (tgt, val) = (*target, *value);
            lower_set_assign(ctx, tgt, val, expr_arena, out)
        }
        ExprNode::DimItem { .. } => Ok(()),
        ExprNode::Block { stmts } => {
            let ids: Vec<NodeId> = stmts.clone();
            for id in ids {
                lower_tracked_stmt(ctx, id, expr_arena, out)?;
            }
            Ok(())
        }
        ExprNode::If { cond, then_body, else_body } => {
            let (cond_id, then_id, else_id) = (*cond, *then_body, *else_body);
            lower_if(ctx, cond_id, then_id, else_id, expr_arena, out)
        }
        ExprNode::While { cond, body } => {
            let (cond_id, body_id) = (*cond, *body);
            lower_while(ctx, cond_id, body_id, expr_arena, out)
        }
        ExprNode::Do { kind, cond, body } => {
            let (kind, cond_id, body_id) = (*kind, *cond, *body);
            lower_do(ctx, kind, cond_id, body_id, expr_arena, out)
        }
        ExprNode::For { var, start, end, step, body } => {
            let (var_id, start_id, end_id, step_id, body_id) =
                (*var, *start, *end, *step, *body);
            lower_for(ctx, var_id, start_id, end_id, step_id, body_id, expr_arena, out)
        }
        ExprNode::SelectCase { subject, pre, cases } => {
            if !pre.is_empty() {
                return Err(LowerError::UnsupportedNode);
            }
            let (subject_id, cases) = (*subject, cases.clone());
            lower_select(ctx, subject_id, &cases, expr_arena, out)
        }
        // A line label emits no code; it records the current byte offset as the
        // jump target for any `GoTo`/`Exit` referencing it.
        ExprNode::Label { target } => {
            ctx.labels.borrow_mut().push((*target, out.len() as u16));
            Ok(())
        }
        // `GoTo label` = unconditional jump (0x1e) to the label's byte offset,
        // backpatched at proc end (the target may be a forward reference).
        ExprNode::GoTo { target } => {
            out.push(0x1e);
            let patch = out.len();
            out.push(0x00);
            out.push(0x00);
            ctx.goto_patches.borrow_mut().push((*target, patch));
            Ok(())
        }
        // `Exit For`/`Exit Do` = unconditional jump (0x1e) to the enclosing
        // loop's end offset, backpatched when the loop finishes emitting.
        ExprNode::ExitStmt { kind } => match kind {
            ExitKind::For | ExitKind::Do => {
                out.push(0x1e);
                let patch = out.len();
                out.push(0x00);
                out.push(0x00);
                ctx.exit_stack
                    .borrow_mut()
                    .last_mut()
                    .ok_or(LowerError::UnsupportedNode)?
                    .push(patch);
                Ok(())
            }
            // Exit Sub / Exit Function = the procedure-return opcode 0x14.
            ExitKind::Sub | ExitKind::Function => {
                out.push(0x14);
                Ok(())
            }
            _ => Err(LowerError::UnsupportedNode),
        },
        // `Mid(s, start[, len]) = value`: LdAddr the target string, push start and
        // (optional) len as Long, push the replacement string, then the Mid opcode
        // 0x4f. (The byte-oriented `MidB` / `$` spellings are gated.)
        ExprNode::MidAssign { byte_oriented, args, value, .. } => {
            let arg_ids = match expr_arena.get(*args) {
                ExprNode::ArgList { args } => args.clone(),
                _ => return Err(LowerError::UnsupportedNode),
            };
            if arg_ids.len() < 2 {
                return Err(LowerError::UnsupportedNode);
            }
            let s_off = match ctx.module.resolutions.get(&arg_ids[0].0) {
                Some(NameResolution::Local { local_idx, .. }) => {
                    ctx.local_slots[*local_idx].frame_offset
                }
                _ => return Err(LowerError::UnsupportedNode),
            };
            out.push(0x04);
            out.extend_from_slice(&s_off.to_le_bytes());
            out.extend_from_slice(&lower_expr_to_bytes_coerced(ctx, arg_ids[1], expr_arena, Some(8))?);
            if arg_ids.len() >= 3 {
                out.extend_from_slice(&lower_expr_to_bytes_coerced(ctx, arg_ids[2], expr_arena, Some(8))?);
            }
            out.extend_from_slice(&lower_expr_to_bytes(ctx, *value, expr_arena)?);
            // Character-oriented Mid/Mid$ use opcode 0x4f; the byte-oriented MidB
            // family uses the escaped opcode 0xfc 0xbe. The `$` spelling does not
            // change the opcode.
            if *byte_oriented {
                out.extend_from_slice(&[0xfc, 0xbe]);
            } else {
                out.push(0x4f);
            }
            out.push(0x00);
            out.push(0x00);
            Ok(())
        }
        // `LSet`/`RSet target = value` (range-copy assignment): load the value, load
        // the target, then the justify opcode — LSet 0x47, RSet 0xfe 0x1e.
        ExprNode::RangeAssign { right_justify, target, value } => {
            out.extend_from_slice(&lower_expr_to_bytes(ctx, *value, expr_arena)?);
            out.extend_from_slice(&lower_expr_to_bytes(ctx, *target, expr_arena)?);
            if *right_justify {
                out.extend_from_slice(&[0xfe, 0x1e]);
            } else {
                out.push(0x47);
            }
            out.push(0x00);
            out.push(0x00);
            Ok(())
        }
        // `ReDim a(bounds)`: push each dimension's lower and upper bound (Long),
        // LdAddr the array pointer, then the ReDim opcode (0xfe 0x8e) followed by
        // dim-count, element VARTYPE, element size, and flags.
        ExprNode::ReDimItem { name, bounds, preserve, .. } => {
            let local_idx = ctx
                .proc
                .locals
                .iter()
                .position(|v| v.sym_id == *name)
                .ok_or(LowerError::Unresolved)?;
            let elem = match &ctx.proc.locals[local_idx].vba_type {
                VbaType::Array(e) => (**e).clone(),
                _ => return Err(LowerError::UnsupportedNode),
            };
            let (vartype, size): (u16, u16) = match elem {
                VbaType::Long => (3, 4),
                VbaType::Integer => (2, 2),
                VbaType::Double => (5, 8),
                _ => return Err(LowerError::UnsupportedNode),
            };
            let arr_off = ctx.local_slots[local_idx].frame_offset;
            let bounds_id = bounds.ok_or(LowerError::UnsupportedNode)?;
            let dim_ids = match expr_arena.get(bounds_id) {
                ExprNode::ArgList { args } => args.clone(),
                _ => return Err(LowerError::UnsupportedNode),
            };
            // The FIRST `ReDim`/`ReDim Preserve` of a given array (textual
            // order) is tracked separately from every later one — see
            // `LowerCtx::redimmed_arrays`'s own doc comment for the oracle
            // evidence a lone `ReDim Preserve` (no earlier ReDim of that
            // array in this proc) omits the bare-bound's explicit
            // lower-bound-0 push that every other captured shape has.
            let is_first_redim = ctx.redimmed_arrays.borrow_mut().insert(local_idx);
            for &d in &dim_ids {
                match expr_arena.get(d) {
                    ExprNode::RangeTo { lo, hi } => {
                        out.extend_from_slice(&lower_expr_to_bytes_coerced(ctx, *lo, expr_arena, Some(8))?);
                        out.extend_from_slice(&lower_expr_to_bytes_coerced(ctx, *hi, expr_arena, Some(8))?);
                    }
                    // Bare upper bound — the lower bound defaults to 0, pushed
                    // explicitly EXCEPT for a `Preserve` that is this array's
                    // very first `ReDim` in the proc (oracle-confirmed:
                    // `oracle_bank/c20_redim_preserve_first_use`).
                    _ => {
                        if !(*preserve && is_first_redim) {
                            out.extend_from_slice(&[0xf5, 0x00, 0x00, 0x00, 0x00]);
                        }
                        out.extend_from_slice(&lower_expr_to_bytes_coerced(ctx, d, expr_arena, Some(8))?);
                    }
                }
            }
            out.push(0x04);
            out.extend_from_slice(&arr_off.to_le_bytes());
            out.push(0xfe);
            // `ReDim Preserve` reallocates while copying existing elements: opcode
            // 0x8f; a plain `ReDim` discards them: opcode 0x8e.
            out.push(if *preserve { 0x8f } else { 0x8e });
            out.extend_from_slice(&(dim_ids.len() as u16).to_le_bytes());
            out.extend_from_slice(&vartype.to_le_bytes());
            out.extend_from_slice(&size.to_le_bytes());
            out.extend_from_slice(&0x80u16.to_le_bytes());
            Ok(())
        }
        // `On Error GoTo label` = opcode 0x4b + the (backpatched) label offset.
        // `On Error GoTo 0` disables the handler: opcode 0x4b with the sentinel
        // target 0xfffe.
        ExprNode::OnError { kind } => match kind {
            OnErrorKind::Goto(target) => {
                out.push(0x4b);
                let patch = out.len();
                out.push(0x00);
                out.push(0x00);
                ctx.goto_patches.borrow_mut().push((*target, patch));
                Ok(())
            }
            OnErrorKind::Disable => {
                out.push(0x4b);
                out.extend_from_slice(&0xfffe_u16.to_le_bytes());
                Ok(())
            }
            // `On Error Resume Next` = opcode 0x4b with the sentinel target 0xffff.
            // (This statement makes the procedure need the line-number table.)
            OnErrorKind::ResumeNext => {
                out.push(0x4b);
                out.extend_from_slice(&0xffff_u16.to_le_bytes());
                Ok(())
            }
        },
        // `Resume` / `Resume Next` / `Resume label`: opcode 0xfd 0x0c + target.
        // Retry (`Resume`) = sentinel 0xfffe; `Resume Next` = 0xffff; a label/line
        // target is the (backpatched) byte offset.
        ExprNode::Resume { target } => {
            out.extend_from_slice(&[0xfd, 0x0c]);
            match target {
                ResumeTarget::Retry => out.extend_from_slice(&0xfffe_u16.to_le_bytes()),
                ResumeTarget::Next => out.extend_from_slice(&0xffff_u16.to_le_bytes()),
                ResumeTarget::At(lref) => {
                    let patch = out.len();
                    out.push(0x00);
                    out.push(0x00);
                    ctx.goto_patches.borrow_mut().push((*lref, patch));
                }
            }
            Ok(())
        }
        // `Stop` — break to the debugger: escaped opcode 0xfc 0xc2.
        // `Call Foo(args)` — an intra-module Sub call (or a Function call whose
        // result is discarded): statement-form call opcode 0x0a. `Call Foo(5)`
        // parses the arguments into a `Call` expression at `callee`; `Call Foo`
        // keeps the callee and (statement) arg list separate.
        ExprNode::CallStmt { callee, args } => {
            let (callee_ref, args_ref) = match expr_arena.get(*callee) {
                ExprNode::Call { func, args: inner } => (*func, *inner),
                _ => (*callee, *args),
            };
            if let ExprNode::MemberAccess { base, .. } = expr_arena.get(callee_ref) {
                let base = *base;
                if member_access_base_is_class(ctx.module, base) {
                    return lower_class_method_call(
                        ctx, base, callee_ref, args_ref, false, false, expr_arena, out,
                    )
                    .map(|_| ());
                }
            }
            lower_call(ctx, callee_ref, args_ref, false, expr_arena, out)
        }
        // Implicit Sub call with parenthesised args as a statement: `Foo(5)`
        // (parses to a bare `Call` node) — statement-form call, result discarded.
        ExprNode::Call { func, args }
            if matches!(ctx.module.resolutions.get(&func.0), Some(NameResolution::Proc(_))) =>
        {
            lower_call(ctx, *func, *args, false, expr_arena, out)
        }
        // A class-member `Sub`/`Function` called as a bare statement with
        // parenthesised args: `o.Method(args)` (result discarded).
        ExprNode::Call { func, args }
            if matches!(expr_arena.get(*func), ExprNode::MemberAccess { base, .. }
                if member_access_base_is_class(ctx.module, *base)) =>
        {
            let ExprNode::MemberAccess { base, .. } = expr_arena.get(*func) else { unreachable!() };
            lower_class_method_call(ctx, *base, *func, *args, false, false, expr_arena, out).map(|_| ())
        }
        // Bare implicit Sub call with no arguments as a statement: `Foo`
        // (parses to a bare `NameRef`). Passing the node itself as the arg list
        // yields no arguments (it is not an ArgList).
        ExprNode::NameRef { .. }
            if matches!(ctx.module.resolutions.get(&node_id.0), Some(NameResolution::Proc(_))) =>
        {
            lower_call(ctx, node_id, node_id, false, expr_arena, out)
        }
        // A class-member `Sub`/`Function` called as a bare statement with NO
        // arguments and no parens: `o.Ping` (result discarded) — the parser
        // returns the bare `MemberAccess` node directly for this shape
        // (`parse_ident_stmt`'s `args.is_empty()` fast path skips wrapping in
        // `CallStmt` entirely, unlike every other spelling above), so it
        // never reaches the `CallStmt`/`Call` arms. Passing the node itself
        // as the arg list (matching the bare-`Proc`-call convention just
        // above it) yields no arguments, since it is not an `ArgList`.
        // Oracle-confirmed: `oracle_bank/c10_bare_0arg_method_call`.
        ExprNode::MemberAccess { base, .. } if member_access_base_is_class(ctx.module, *base) => {
            let base = *base;
            lower_class_method_call(ctx, base, node_id, node_id, false, false, expr_arena, out).map(|_| ())
        }
        ExprNode::Stop => {
            out.extend_from_slice(&[0xfc, 0xc2]);
            Ok(())
        }
        // `End` — terminate the program: escaped opcode 0xfc 0xc8.
        ExprNode::EndStmt => {
            out.extend_from_slice(&[0xfc, 0xc8]);
            Ok(())
        }
        // `Return` — return from the most recent GoSub: escaped opcode 0xfc 0xc9.
        ExprNode::ReturnStmt => {
            out.extend_from_slice(&[0xfc, 0xc9]);
            Ok(())
        }
        // `Error number` — raise the given error: push the number (Long) then 0x45.
        ExprNode::ErrorStmt { expr } => {
            let expr_id = *expr;
            out.extend_from_slice(&lower_expr_to_bytes_coerced(ctx, expr_id, expr_arena, Some(8))?);
            out.push(0x45);
            Ok(())
        }
        // `Erase var1[, var2...]`: for each array, LdAddr the array slot then the
        // erase opcode. A dynamic array is freed (0x5a); a fixed array is
        // reinitialized in place, which needs its type-descriptor offset — gated.
        ExprNode::Erase { vars } => {
            for &v in vars {
                let (arr_off, _elem, _dims) =
                    array_local_info(ctx, v).ok_or(LowerError::UnsupportedNode)?;
                let is_dynamic = match ctx.module.resolutions.get(&v.0) {
                    Some(NameResolution::Local { local_idx, .. }) => {
                        ctx.proc.locals[*local_idx].array_dims.is_none()
                    }
                    _ => return Err(LowerError::UnsupportedNode),
                };
                out.push(0x04);
                out.extend_from_slice(&arr_off.to_le_bytes());
                if !is_dynamic {
                    // A fixed array is reinitialized in place: 0x59 with the array's
                    // data offset (the LdAddr target minus 8), then 0x5a.
                    out.push(0x59);
                    out.extend_from_slice(&(arr_off - 8).to_le_bytes());
                }
                out.push(0x5a);
            }
            Ok(())
        }
        // `On expr GoTo/GoSub label-list`: emit the selector coerced to Integer,
        // then the dispatch opcode (0xfe 0x96 GoTo / 0xfe 0x95 GoSub), the table
        // byte length (labels * 2), then one 2-byte (backpatched) offset per label.
        ExprNode::OnGo { is_gosub, expr, labels } => {
            let mut arena = NodeArena::new();
            let root = lower_expr(ctx, *expr, expr_arena, &mut arena)?;
            // The selector is converted to Integer (e.g. Long selector → 0xe4).
            let root = coerce_operand(ctx, *expr, root, Some(6), expr_arena, &mut arena);
            let mut emitter = Emitter::new(&arena);
            emitter.emit_expr(root, 2);
            out.extend(emitter.into_bytes());
            out.push(0xfe);
            out.push(if *is_gosub { 0x95 } else { 0x96 });
            out.extend_from_slice(&((labels.len() * 2) as u16).to_le_bytes());
            for &lbl in labels {
                let target = match expr_arena.get(lbl) {
                    ExprNode::NameRef { sym, .. } => LabelRef::Name(*sym),
                    ExprNode::Literal { lit: AstLit::Int(v) } => LabelRef::Line(*v),
                    _ => return Err(LowerError::UnsupportedNode),
                };
                let patch = out.len();
                out.push(0x00);
                out.push(0x00);
                ctx.goto_patches.borrow_mut().push((target, patch));
            }
            Ok(())
        }
        // `GoSub label` — push a return address and jump: opcode 0xfd 0x0a + the
        // (backpatched) label byte offset.
        ExprNode::GoSub { target } => {
            out.extend_from_slice(&[0xfd, 0x0a]);
            let patch = out.len();
            out.push(0x00);
            out.push(0x00);
            ctx.goto_patches.borrow_mut().push((*target, patch));
            Ok(())
        }
        _ => Err(LowerError::UnsupportedNode),
    }
}

// ── Control-flow helpers ──────────────────────────────────────────────────────

/// Emit the expression for `node_id` to a scratch byte vector and return it.
pub(super) fn lower_expr_to_bytes(
    ctx: &LowerCtx,
    node_id: NodeId,
    expr_arena: &ExprArena,
) -> Result<Vec<u8>, LowerError> {
    let mut arena = NodeArena::new();
    let root = lower_expr(ctx, node_id, expr_arena, &mut arena)?;
    let mut emitter = Emitter::new(&arena);
    emitter.emit_expr(root, 0);
    Ok(emitter.into_bytes())
}

/// Emit a 2-byte LE placeholder at the current position; return the patch offset.
pub(super) fn emit_branch_placeholder(out: &mut Vec<u8>, opcode: u8) -> usize {
    out.push(opcode);
    let patch = out.len();
    out.push(0x00);
    out.push(0x00);
    patch
}

/// Backpatch a 2-byte LE u16 at the given byte offset with the given value.
pub(super) fn patch_u16(out: &mut Vec<u8>, patch: usize, value: u16) {
    out[patch..patch + 2].copy_from_slice(&value.to_le_bytes());
}

/// `If cond Then then_body [Else else_body] End If`
///
/// Opcode layout (oracle-confirmed):
///   <cond bytes>
///   0x1c [2-byte LE absolute offset to else_body or end]   ; BranchFalse
///   <then_body bytes>
///   [0x1e [2-byte LE absolute offset to end]               ; Jump (only when else present)
///   <else_body bytes>]
pub(super) fn lower_if(
    ctx: &LowerCtx,
    cond_id: NodeId,
    then_id: NodeId,
    else_id: Option<NodeId>,
    expr_arena: &ExprArena,
    out: &mut Vec<u8>,
) -> Result<(), LowerError> {
    let cond_bytes = lower_expr_to_bytes(ctx, cond_id, expr_arena)?;
    out.extend_from_slice(&cond_bytes);

    // BranchFalse — opcode 0x1c + 2-byte absolute target (patched below)
    let branch_false_patch = emit_branch_placeholder(out, 0x1c);

    lower_block(ctx, then_id, expr_arena, out)?;

    if let Some(e_id) = else_id {
        // Jump over else — opcode 0x1e + 2-byte absolute target
        let jump_patch = emit_branch_placeholder(out, 0x1e);

        // BranchFalse target = start of else body
        patch_u16(out, branch_false_patch, out.len() as u16);

        lower_block(ctx, e_id, expr_arena, out)?;

        // Jump target = end
        patch_u16(out, jump_patch, out.len() as u16);
    } else {
        // BranchFalse target = end of If block
        patch_u16(out, branch_false_patch, out.len() as u16);
    }

    Ok(())
}

/// `While cond ... Wend`  and  `Do While cond ... Loop`
///
/// Opcode layout (oracle-confirmed):
///   [loop_start:]
///   <cond bytes>
///   0x1c [2-byte LE absolute offset to past loop]   ; BranchFalse
///   <body bytes>
///   0x1e [2-byte LE loop_start]                     ; Jump back
pub(super) fn lower_while(
    ctx: &LowerCtx,
    cond_id: NodeId,
    body_id: NodeId,
    expr_arena: &ExprArena,
    out: &mut Vec<u8>,
) -> Result<(), LowerError> {
    let loop_start = out.len() as u16;

    let cond_bytes = lower_expr_to_bytes(ctx, cond_id, expr_arena)?;
    out.extend_from_slice(&cond_bytes);

    let branch_false_patch = emit_branch_placeholder(out, 0x1c);

    lower_block(ctx, body_id, expr_arena, out)?;

    // Unconditional jump back to loop start
    out.push(0x1e);
    out.extend_from_slice(&loop_start.to_le_bytes());

    // BranchFalse target = past end of jump instruction (current position)
    patch_u16(out, branch_false_patch, out.len() as u16);

    Ok(())
}

/// `Do [While/Until cond] ... Loop [While/Until cond]`
///
/// Opcode layout variants (oracle-confirmed):
///
///   PreWhile:  [start:] cond BranchFalse[end] body Jump[start]
///   PreUntil:  [start:] cond BranchTrue[end]  body Jump[start]
///   PostWhile: [start:] body cond BranchTrue[start]
///   PostUntil: [start:] body cond BranchFalse[start]
///   Inf:       [start:] body Jump[start]
/// The negated comparison operator. VB6 compiles `Until cond` as `While Not cond`;
/// for a comparison condition the negation folds into the compare opcode (e.g.
/// `Until a > 9` emits the `<=` compare + branch-false, identical to `While a <= 9`).
pub(super) fn negate_comparison(op: BinOpKind) -> Option<BinOpKind> {
    Some(match op {
        BinOpKind::Eq => BinOpKind::Ne,
        BinOpKind::Ne => BinOpKind::Eq,
        BinOpKind::Lt => BinOpKind::Ge,
        BinOpKind::Ge => BinOpKind::Lt,
        BinOpKind::Gt => BinOpKind::Le,
        BinOpKind::Le => BinOpKind::Gt,
        _ => return None,
    })
}

/// Emit the bytes for a loop `Until` condition as its negation (`While Not cond`).
/// Returns `Some(bytes)` when the condition is a negatable comparison; `None`
/// otherwise (the caller then uses the non-negated condition + branch-on-true).
pub(super) fn lower_negated_condition_bytes(
    ctx: &LowerCtx,
    cond_id: NodeId,
    expr_arena: &ExprArena,
) -> Result<Option<Vec<u8>>, LowerError> {
    if let ExprNode::BinOp { op, lhs, rhs } = expr_arena.get(cond_id) {
        if let Some(neg) = negate_comparison(*op) {
            let (lhs_id, rhs_id) = (*lhs, *rhs);
            let mut arena = NodeArena::new();
            let root = lower_binop(ctx, cond_id, neg, lhs_id, rhs_id, expr_arena, &mut arena)?;
            let mut emitter = Emitter::new(&arena);
            emitter.emit_expr(root, 0);
            return Ok(Some(emitter.into_bytes()));
        }
    }
    Ok(None)
}

pub(super) fn lower_do(
    ctx: &LowerCtx,
    kind: DoKind,
    cond_id: Option<NodeId>,
    body_id: NodeId,
    expr_arena: &ExprArena,
    out: &mut Vec<u8>,
) -> Result<(), LowerError> {
    // Open an Exit-Do patch scope; every `Exit Do` in the body is backpatched to
    // the loop-end offset (the byte just past the loop's back-branch).
    ctx.exit_stack.borrow_mut().push(Vec::new());
    let result = lower_do_inner(ctx, kind, cond_id, body_id, expr_arena, out);
    let exits = ctx.exit_stack.borrow_mut().pop().unwrap_or_default();
    let end = out.len() as u16;
    for patch in exits {
        out[patch..patch + 2].copy_from_slice(&end.to_le_bytes());
    }
    result
}

pub(super) fn lower_do_inner(
    ctx: &LowerCtx,
    kind: DoKind,
    cond_id: Option<NodeId>,
    body_id: NodeId,
    expr_arena: &ExprArena,
    out: &mut Vec<u8>,
) -> Result<(), LowerError> {
    match kind {
        DoKind::PreWhile => {
            // Same structure as While/Wend
            let cond = cond_id.ok_or(LowerError::UnsupportedNode)?;
            lower_while(ctx, cond, body_id, expr_arena, out)
        }
        DoKind::PreUntil => {
            let cond = cond_id.ok_or(LowerError::UnsupportedNode)?;
            let loop_start = out.len() as u16;

            // VB6 compiles `Do Until cond` as `Do While Not cond`: for a
            // comparison the negation folds into the compare opcode and the exit
            // branch is BranchFalse (0x1c) — byte-identical to PreWhile.
            if let Some(neg_bytes) = lower_negated_condition_bytes(ctx, cond, expr_arena)? {
                out.extend_from_slice(&neg_bytes);
                let branch_false_patch = emit_branch_placeholder(out, 0x1c);
                lower_block(ctx, body_id, expr_arena, out)?;
                out.push(0x1e);
                out.extend_from_slice(&loop_start.to_le_bytes());
                patch_u16(out, branch_false_patch, out.len() as u16);
                return Ok(());
            }

            // Non-comparison condition: exit when the condition is true.
            let cond_bytes = lower_expr_to_bytes(ctx, cond, expr_arena)?;
            out.extend_from_slice(&cond_bytes);
            let branch_true_patch = emit_branch_placeholder(out, 0x1d);
            lower_block(ctx, body_id, expr_arena, out)?;
            out.push(0x1e);
            out.extend_from_slice(&loop_start.to_le_bytes());
            patch_u16(out, branch_true_patch, out.len() as u16);
            Ok(())
        }
        DoKind::PostWhile => {
            let cond = cond_id.ok_or(LowerError::UnsupportedNode)?;
            let loop_start = out.len() as u16;

            lower_block(ctx, body_id, expr_arena, out)?;

            let cond_bytes = lower_expr_to_bytes(ctx, cond, expr_arena)?;
            out.extend_from_slice(&cond_bytes);

            // BranchTrue back to start — "Loop While" continues when true
            out.push(0x1d);
            out.extend_from_slice(&loop_start.to_le_bytes());
            Ok(())
        }
        DoKind::PostUntil => {
            let cond = cond_id.ok_or(LowerError::UnsupportedNode)?;
            let loop_start = out.len() as u16;

            lower_block(ctx, body_id, expr_arena, out)?;

            // `Loop Until cond` = `Loop While Not cond`: negated comparison +
            // BranchTrue back to start (mirrors PostWhile).
            if let Some(neg_bytes) = lower_negated_condition_bytes(ctx, cond, expr_arena)? {
                out.extend_from_slice(&neg_bytes);
                out.push(0x1d);
                out.extend_from_slice(&loop_start.to_le_bytes());
                return Ok(());
            }

            let cond_bytes = lower_expr_to_bytes(ctx, cond, expr_arena)?;
            out.extend_from_slice(&cond_bytes);
            // BranchFalse back to start — "Loop Until" continues when still false
            out.push(0x1c);
            out.extend_from_slice(&loop_start.to_le_bytes());
            Ok(())
        }
        DoKind::Inf => {
            let loop_start = out.len() as u16;
            lower_block(ctx, body_id, expr_arena, out)?;
            out.push(0x1e);
            out.extend_from_slice(&loop_start.to_le_bytes());
            Ok(())
        }
    }
}

/// Emit expression bytes with optional integer-literal coercion.
pub(super) fn lower_expr_to_bytes_coerced(
    ctx: &LowerCtx,
    node_id: NodeId,
    expr_arena: &ExprArena,
    coerce_tag: Option<u16>,
) -> Result<Vec<u8>, LowerError> {
    let mut arena = NodeArena::new();
    let root = lower_expr_coerced(ctx, node_id, expr_arena, &mut arena, coerce_tag)?;
    let mut emitter = Emitter::new(&arena);
    emitter.emit_expr(root, 0);
    Ok(emitter.into_bytes())
}

/// `For var = start To end [Step step] ... Next`
///
/// Opcode layout (oracle-confirmed for Long counter, no-step):
///   <start bytes>
///   0x04 [2-byte frame_var]              ; LdAddr: push address of counter var
///   <end bytes>
///   0xfe 0x64 [frame_hidden] [exit_off]  ; ForInit no-step
///   <body bytes>
///   0x04 [2-byte frame_var]              ; LdAddr again for ForNext
///   0x66 [frame_hidden] [body_start]     ; ForNext no-step; back_offset = body start
///
/// With-step replaces 0xfe 0x64 with 0xfe 0x6c and 0x66 with 0x67, and pushes
/// the step value between end and ForInit.
///
/// frame_hidden: the second of two Long hidden frame slots pre-allocated for
/// this For loop (user_local_count + 2*pair + 1).
pub(super) fn lower_for(
    ctx: &LowerCtx,
    var_id: NodeId,
    start_id: NodeId,
    end_id: NodeId,
    step_id: Option<NodeId>,
    body_id: NodeId,
    expr_arena: &ExprArena,
    out: &mut Vec<u8>,
) -> Result<(), LowerError> {
    // Resolve counter variable to its frame offset and type.
    let resolution = ctx
        .module
        .resolutions
        .get(&var_id.0)
        .ok_or(LowerError::Unresolved)?;
    let (frame_var, coerce_tag) = match resolution {
        NameResolution::Local { local_idx, .. } => {
            let slot = &ctx.local_slots[*local_idx];
            let tag = vba_type_to_node_tag(ctx.local_type(*local_idx));
            (slot.frame_offset, tag)
        }
        _ => return Err(LowerError::UnsupportedNode),
    };

    // Claim the hidden-slot pair for this For loop.
    let pair = ctx.for_next_pair.get();
    ctx.for_next_pair.set(pair + 1);
    let hidden_idx = ctx.user_local_count + 2 * pair + 1;
    let frame_hidden = ctx.local_slots[hidden_idx].frame_offset;

    let has_step = step_id.is_some();

    // Emit start value (coerced to counter type).
    out.extend_from_slice(&lower_expr_to_bytes_coerced(ctx, start_id, expr_arena, coerce_tag)?);

    // LdAddr: push address of counter variable (opcode 0x04 + 2-byte frame offset).
    out.push(0x04);
    out.extend_from_slice(&frame_var.to_le_bytes());

    // Emit limit value (coerced to counter type).
    out.extend_from_slice(&lower_expr_to_bytes_coerced(ctx, end_id, expr_arena, coerce_tag)?);

    // If with-step: emit step value.
    if let Some(s_id) = step_id {
        out.extend_from_slice(&lower_expr_to_bytes_coerced(ctx, s_id, expr_arena, coerce_tag)?);
    }

    // ForInit: 2-byte opcode + 2-byte frame_hidden + 2-byte exit_offset placeholder.
    let forinit_byte2: u8 = if has_step { 0x6c } else { 0x64 };
    out.push(0xfe);
    out.push(forinit_byte2);
    out.extend_from_slice(&frame_hidden.to_le_bytes());
    let exit_patch = out.len();
    out.push(0x00);
    out.push(0x00);

    // Body starts immediately after ForInit.
    let body_start = out.len() as u16;

    // Open an Exit-For patch scope for this loop's body.
    ctx.exit_stack.borrow_mut().push(Vec::new());
    lower_block(ctx, body_id, expr_arena, out)?;

    // LdAddr counter again before ForNext.
    out.push(0x04);
    out.extend_from_slice(&frame_var.to_le_bytes());

    // ForNext: 1-byte opcode + 2-byte frame_hidden + 2-byte back_offset.
    let fornext_opcode: u8 = if has_step { 0x67 } else { 0x66 };
    out.push(fornext_opcode);
    out.extend_from_slice(&frame_hidden.to_le_bytes());
    out.extend_from_slice(&body_start.to_le_bytes());

    // Backpatch ForInit exit_offset to current position, and every `Exit For`
    // jump emitted inside the body to the same loop-end offset.
    let exit_offset = out.len() as u16;
    out[exit_patch..exit_patch + 2].copy_from_slice(&exit_offset.to_le_bytes());
    let exits = ctx.exit_stack.borrow_mut().pop().unwrap_or_default();
    for patch in exits {
        out[patch..patch + 2].copy_from_slice(&exit_offset.to_le_bytes());
    }

    Ok(())
}

/// Emit one `Case` clause's comparison against the subject temp, producing a
/// boolean on the stack: a bare value uses `=`; `lo To hi` uses the range test
/// (0xfb 0x86); `Is <op> <expr>` uses that operator's comparison.
pub(super) fn emit_clause_test(
    ctx: &LowerCtx,
    clause_id: NodeId,
    subj_tag: u16,
    lctx: usize,
    temp_offset: i16,
    expr_arena: &ExprArena,
) -> Result<Vec<u8>, LowerError> {
    let mut arena = NodeArena::new();
    let temp_load = build_frame_load_node(&mut arena, 0x74, subj_tag, lctx, temp_offset);
    match expr_arena.get(clause_id) {
        ExprNode::RangeTo { lo, hi } => {
            let lo_r = lower_expr_coerced(ctx, *lo, expr_arena, &mut arena, Some(subj_tag))?;
            let hi_r = lower_expr_coerced(ctx, *hi, expr_arena, &mut arena, Some(subj_tag))?;
            let mut em = Emitter::new(&arena);
            em.emit_expr(temp_load, 0);
            em.emit_expr(lo_r, 0);
            em.emit_expr(hi_r, 0);
            let mut bytes = em.into_bytes();
            bytes.push(0xfb); // range-test opcode (value-emitter index 0x86)
            bytes.push(0x86);
            Ok(bytes)
        }
        ExprNode::CaseIs { op, rhs } => {
            let rhs_r = lower_expr_coerced(ctx, *rhs, expr_arena, &mut arena, Some(subj_tag))?;
            let opcode = binop_node_opcode(*op).ok_or(LowerError::UnsupportedNode)?;
            let cmp = arena.alloc(NodeArena::node(opcode, 0, temp_load.0, rhs_r.0, 0, 0));
            let mut em = Emitter::new(&arena);
            em.emit_expr(cmp, 0);
            Ok(em.into_bytes())
        }
        _ => {
            let v = lower_expr_coerced(ctx, clause_id, expr_arena, &mut arena, Some(subj_tag))?;
            let eq = arena.alloc(NodeArena::node(0x26, 0, temp_load.0, v.0, 0, 0));
            let mut em = Emitter::new(&arena);
            em.emit_expr(eq, 0);
            Ok(em.into_bytes())
        }
    }
}

/// `Select Case subject ... End Select`. VB6 evaluates the subject once into a
/// hidden temp slot, then for each `Case` loads the temp, compares it (`=`)
/// against the case value, and branches past the body (BranchFalse) when not
/// equal. Each matched body jumps to the Select end; `Case Else` falls through.
pub(super) fn lower_select(
    ctx: &LowerCtx,
    subject_id: NodeId,
    cases: &[NodeId],
    expr_arena: &ExprArena,
    out: &mut Vec<u8>,
) -> Result<(), LowerError> {
    let sel = ctx.select_next.get();
    ctx.select_next.set(sel + 1);
    let temp_offset = ctx.local_slots[ctx.select_base + sel].frame_offset;

    let subj_ty = ctx
        .module
        .types
        .get(&subject_id.0)
        .ok_or(LowerError::Unresolved)?;
    let subj_tag = vba_type_to_node_tag(subj_ty).ok_or(LowerError::UnsupportedType)?;
    let lctx = load_store_ctx(subj_ty).ok_or(LowerError::UnsupportedType)?;

    // Evaluate the subject once and store it to the hidden temp.
    {
        let mut arena = NodeArena::new();
        let root = lower_expr(ctx, subject_id, expr_arena, &mut arena)?;
        let mut emitter = Emitter::new(&arena);
        emitter.emit_expr(root, 0);
        emitter.emit_var_store(lctx, temp_offset);
        out.extend_from_slice(&emitter.into_bytes());
    }

    let mut end_patches = Vec::new();
    let case_count = cases.len();
    for (case_idx, &case_id) in cases.iter().enumerate() {
        match expr_arena.get(case_id) {
            ExprNode::CaseBlock { test, body } => {
                // `test` is an ArgList of clauses (values / ranges / `Is` tests).
                // A match on ANY clause enters the body: every clause but the last
                // branches TRUE to the body; the last branches FALSE to the next
                // case (so the single-clause form is just compare + BranchFalse).
                let clauses = match expr_arena.get(*test) {
                    ExprNode::ArgList { args } => args.clone(),
                    _ => return Err(LowerError::UnsupportedNode),
                };
                if clauses.is_empty() {
                    return Err(LowerError::UnsupportedNode);
                }
                let mut body_patches = Vec::new();
                let mut next_patch = 0usize;
                let last = clauses.len() - 1;
                for (i, &clause) in clauses.iter().enumerate() {
                    out.extend_from_slice(&emit_clause_test(
                        ctx, clause, subj_tag, lctx, temp_offset, expr_arena,
                    )?);
                    if i < last {
                        body_patches.push(emit_branch_placeholder(out, 0x1d));
                    } else {
                        next_patch = emit_branch_placeholder(out, 0x1c);
                    }
                }
                let body_start = out.len() as u16;
                for p in body_patches {
                    patch_u16(out, p, body_start);
                }
                lower_block(ctx, *body, expr_arena, out)?;
                // A matched body jumps to the Select end only when a later case
                // follows it; the last case-arm falls through to the end.
                if case_idx + 1 < case_count {
                    out.push(0x1e);
                    let end_patch = out.len();
                    out.push(0x00);
                    out.push(0x00);
                    end_patches.push(end_patch);
                }
                // The final BranchFalse lands at the next case (or the end).
                patch_u16(out, next_patch, out.len() as u16);
            }
            ExprNode::CaseElse { body } => {
                lower_block(ctx, *body, expr_arena, out)?;
            }
            _ => return Err(LowerError::UnsupportedNode),
        }
    }
    let end = out.len() as u16;
    for patch in end_patches {
        patch_u16(out, patch, end);
    }
    Ok(())
}

/// Array element store `a(i…) = v`: push the value (element-typed), push each Long
/// index, LdAddr the array descriptor (0x04), then the element-store. A 1-D array
/// uses the direct store opcode (0xa3 Long / 0xa2 Integer); a multi-dimensional
/// array uses the indexed-store sequence (0xa7 <dims> 0x8f).
pub(super) fn lower_array_store(
    ctx: &LowerCtx,
    func_id: NodeId,
    args_id: NodeId,
    value_id: NodeId,
    expr_arena: &ExprArena,
    out: &mut Vec<u8>,
) -> Result<(), LowerError> {
    let (arr_off, elem, dims) =
        array_local_info(ctx, func_id).ok_or(LowerError::UnsupportedNode)?;
    let indices = match expr_arena.get(args_id) {
        ExprNode::ArgList { args } => args.clone(),
        _ => return Err(LowerError::UnsupportedNode),
    };
    if indices.len() != dims as usize {
        return Err(LowerError::UnsupportedNode);
    }
    let elem_tag = vba_type_to_node_tag(&elem).ok_or(LowerError::UnsupportedType)?;
    if matches!(elem, VbaType::Variant) {
        // A Variant element store pushes the source variant's address, not value.
        let v_off = match (expr_arena.get(value_id), ctx.module.resolutions.get(&value_id.0)) {
            (ExprNode::NameRef { .. }, Some(NameResolution::Local { local_idx, .. }))
                if matches!(ctx.module.types.get(&value_id.0), Some(VbaType::Variant)) =>
            {
                ctx.local_slots[*local_idx].frame_offset
            }
            _ => return Err(LowerError::UnsupportedNode),
        };
        out.push(0x04);
        out.extend_from_slice(&v_off.to_le_bytes());
    } else {
        out.extend_from_slice(&lower_expr_to_bytes_coerced(ctx, value_id, expr_arena, Some(elem_tag))?);
    }
    for &idx in &indices {
        out.extend_from_slice(&lower_expr_to_bytes_coerced(ctx, idx, expr_arena, Some(8))?);
    }
    out.push(0x04);
    out.extend_from_slice(&arr_off.to_le_bytes());
    if dims == 1 {
        // 1-D element store, by element type. Byte uses a 2-byte escaped opcode.
        match elem {
            VbaType::Integer | VbaType::Boolean => out.push(0xa2),
            VbaType::Long => out.push(0xa3),
            VbaType::Currency => out.push(0xa4),
            VbaType::Single => out.push(0xa5),
            VbaType::Double | VbaType::Date => out.push(0xa6),
            VbaType::String => out.push(0x3b),
            VbaType::Byte => out.extend_from_slice(&[0xfc, 0xa0]),
            // Variant element store: move the source variant (loaded by address
            // above) into the element.
            VbaType::Variant => out.extend_from_slice(&[0xfc, 0xb0]),
            _ => return Err(LowerError::UnsupportedNode),
        }
    } else {
        // Multi-dimensional indexed store: compute the element address (0xa7 +
        // dimension count), then store through it (0x8f for Long).
        if !matches!(elem, VbaType::Long) {
            return Err(LowerError::UnsupportedNode);
        }
        out.push(0xa7);
        out.extend_from_slice(&dims.to_le_bytes());
        out.push(0x8f);
        out.push(0x00);
        out.push(0x00);
    }
    Ok(())
}
