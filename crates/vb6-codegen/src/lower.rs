//! AST-to-P-code lowering: walks a `BoundModule`/`BoundProc` from vb6-sema and
//! emits the runtime P-code byte stream via the reference emitter.
//!
//! The entry point is [`lower_proc`].  It builds local/param/global frame
//! layouts (matching VB6's exact slot ordering), walks the body
//! [`ExprNode::Block`] recursively, and for each statement that can be lowered
//! builds a [`NodeArena`] sub-tree and calls [`Emitter::emit_expr`] followed
//! by the appropriate typed store.
//!
//! Only the code paths whose P-code bytes have been oracle-confirmed are
//! implemented.  Unhandled constructs return [`LowerError::UnsupportedNode`]
//! or [`LowerError::UnsupportedType`] — never a silently wrong byte.

use std::cell::Cell;

use vb6_sema::sema::{BoundModule, BoundProc, VbaType, NameResolution};
use vb6_syntax::frontend::ast::{ExprArena, ExprNode, AstLit, BinOpKind, UnOpKind, DoKind};
use vb6_syntax::support::arena::NodeId;

use crate::bind::{GlobalFrame, GlobalVar, LocalVar, ParamVar};
use crate::bridge::{frame_from_local_types, load_store_ctx, param_frame_from_types, type_ctx, UnsupportedType};
use crate::emit::Emitter;
use crate::node::{NodeArena, NodeRef};

/// Errors that can arise while lowering a proc.
#[derive(Debug, Clone, PartialEq, Eq)]
pub enum LowerError {
    /// A variable or literal has a type whose P-code emission is not yet
    /// oracle-confirmed (e.g. String, Variant, user-defined type).
    UnsupportedType,
    /// A name reference has no entry in `BoundModule.resolutions`.
    Unresolved,
    /// An AST node kind the lowering pass cannot yet handle.
    UnsupportedNode,
    /// `proc_idx` exceeds `module.procs.len()`.
    ProcIndexOutOfRange,
}

impl From<UnsupportedType> for LowerError {
    fn from(_: UnsupportedType) -> Self {
        LowerError::UnsupportedType
    }
}

/// Map a `VbaType` to the VB6 internal type tag stored in the high 16 bits of
/// an expression node's `word[0]`.
///
/// Oracle-confirmed constants (from oracle_pcode.rs constants and emit.rs
/// float-literal comments): T_INTEGER/Boolean = 6, T_LONG = 8, T_SINGLE = 10,
/// T_DOUBLE = 11.
fn vba_type_to_node_tag(ty: &VbaType) -> Option<u16> {
    match ty {
        VbaType::Integer | VbaType::Boolean => Some(6),
        VbaType::Long => Some(8),
        VbaType::Single => Some(10),
        VbaType::Double => Some(11),
        _ => None,
    }
}

/// Map a `BinOpKind` to the oracle-confirmed bound-node opcode.
/// Returns `None` for operators whose P-code bytes are not yet confirmed.
fn binop_node_opcode(op: BinOpKind) -> Option<u16> {
    Some(match op {
        BinOpKind::Add => 0x16,
        BinOpKind::Sub => 0x17,
        BinOpKind::Mul => 0x18,
        // Div (`/`) is the arithmetic binop occupying 0x19 — the gap in the
        // contiguous +/-/*//^ block (0x16..0x1a): RT_BINOP_BASE[0x19]=0xb6 is a
        // valid base and the stmt jump table routes 0x19 to the generic-binop
        // emitter (stmt_case_0fab1da9), like Add/Sub.
        BinOpKind::Div => 0x19,
        // The multiplicative and logical operators the front-end operator table
        // assigns by precedence (consumed by the generic operation emitter):
        //   `\`  (integer divide) -> 0x1e   Mod -> 0x1d
        //   Eqv -> 0x20                      Imp -> 0x1f
        // (precedence ladder * / > \ > Mod > + - and And > Or > Xor > Eqv > Imp).
        BinOpKind::IDiv => 0x1e,
        BinOpKind::Mod => 0x1d,
        BinOpKind::Or  => 0x21,
        BinOpKind::Xor => 0x22,
        BinOpKind::And => 0x23,
        BinOpKind::Eqv => 0x20,
        BinOpKind::Imp => 0x1f,
        BinOpKind::Eq  => 0x26,
        BinOpKind::Ne  => 0x27,
        BinOpKind::Le  => 0x28,
        BinOpKind::Ge  => 0x29,
        BinOpKind::Lt  => 0x2a,
        BinOpKind::Gt  => 0x2b,
        _ => return None,
    })
}

fn is_comparison_op(op: BinOpKind) -> bool {
    matches!(
        op,
        BinOpKind::Eq | BinOpKind::Ne | BinOpKind::Lt | BinOpKind::Le | BinOpKind::Gt | BinOpKind::Ge
    )
}

/// Allocate global slots from module-level variable types in declaration order.
/// Mirrors [`frame_from_local_types`] for the global data block.
pub fn global_frame_from_types(
    types: &[VbaType],
    module_desc: u16,
) -> Result<Vec<GlobalVar>, UnsupportedType> {
    let mut frame = GlobalFrame::new(module_desc);
    let mut out = Vec::with_capacity(types.len());
    for ty in types {
        let ctx = type_ctx(ty).ok_or(UnsupportedType)?;
        out.push(frame.declare_anon_global(ctx));
    }
    Ok(out)
}

/// Count the number of For loops directly or indirectly in an AST subtree.
/// Each For loop needs 2 hidden Long slots in the frame.
fn count_for_loops(node_id: NodeId, expr_arena: &ExprArena) -> usize {
    match expr_arena.get(node_id) {
        ExprNode::For { body, .. } => 1 + count_for_loops(*body, expr_arena),
        ExprNode::Block { stmts } => {
            stmts.iter().map(|&id| count_for_loops(id, expr_arena)).sum()
        }
        ExprNode::If { then_body, else_body, .. } => {
            count_for_loops(*then_body, expr_arena)
                + else_body.map(|id| count_for_loops(id, expr_arena)).unwrap_or(0)
        }
        ExprNode::While { body, .. } => count_for_loops(*body, expr_arena),
        ExprNode::Do { body, .. } => count_for_loops(*body, expr_arena),
        _ => 0,
    }
}

/// Lower a single `BoundProc` to its P-code byte vector.
///
/// Frame layout follows VB6's exact convention: locals at negative offsets
/// from -136 downward (4 bytes per Integer/Long/Single/Object, 8 bytes per
/// Double/Currency), params at positive offsets from +12 upward.
///
/// For loops each need 2 hidden Long slots allocated below all user locals.
/// These are pre-allocated here by scanning the body first.
///
/// `module_desc` is the compiled module-object descriptor word — `0x0008` for
/// the primary module in a single-module project (oracle-confirmed).
pub fn lower_proc(
    module: &BoundModule,
    proc_idx: usize,
    expr_arena: &ExprArena,
    module_desc: u16,
) -> Result<Vec<u8>, LowerError> {
    let proc = module.procs.get(proc_idx).ok_or(LowerError::ProcIndexOutOfRange)?;

    let mut local_types: Vec<VbaType> = proc.locals.iter().map(|v| v.vba_type.clone()).collect();
    let user_local_count = local_types.len();

    // Pre-allocate 2 Long hidden slots per For loop.
    let for_count = count_for_loops(NodeId(proc.body), expr_arena);
    for _ in 0..(for_count * 2) {
        local_types.push(VbaType::Long);
    }

    let param_types: Vec<VbaType> = proc.params.iter().map(|p| p.vba_type.clone()).collect();
    let param_byref: Vec<bool> = proc.params.iter().map(|p| !p.flags.by_val).collect();
    let global_types: Vec<VbaType> =
        module.module_vars.iter().map(|v| v.vba_type.clone()).collect();

    let local_slots = frame_from_local_types(&local_types)?;
    let param_slots = param_frame_from_types(&param_types, &param_byref)?;
    let global_slots = global_frame_from_types(&global_types, module_desc)?;

    let ctx = LowerCtx {
        module,
        proc,
        local_slots,
        param_slots,
        global_slots,
        user_local_count,
        for_next_pair: Cell::new(0),
    };

    let mut out = Vec::new();
    lower_block(&ctx, NodeId(proc.body), expr_arena, &mut out)?;
    Ok(out)
}

// ── Internal lowering context ─────────────────────────────────────────────────

struct LowerCtx<'m> {
    module: &'m BoundModule,
    proc: &'m BoundProc,
    local_slots: Vec<LocalVar>,
    param_slots: Vec<ParamVar>,
    global_slots: Vec<GlobalVar>,
    /// Number of user-declared locals (hidden For-loop slots come after).
    user_local_count: usize,
    /// Which hidden-slot pair the next For loop should use.
    for_next_pair: Cell<usize>,
}

impl<'m> LowerCtx<'m> {
    fn local_type(&self, idx: usize) -> &VbaType {
        &self.proc.locals[idx].vba_type
    }
    fn param_type(&self, idx: usize) -> &VbaType {
        &self.proc.params[idx].vba_type
    }
    fn global_type(&self, idx: usize) -> &VbaType {
        &self.module.module_vars[idx].vba_type
    }
}

// ── Statement lowering ────────────────────────────────────────────────────────

fn lower_block(
    ctx: &LowerCtx,
    node_id: NodeId,
    expr_arena: &ExprArena,
    out: &mut Vec<u8>,
) -> Result<(), LowerError> {
    match expr_arena.get(node_id) {
        ExprNode::Block { stmts } => {
            let ids: Vec<NodeId> = stmts.clone();
            for id in ids {
                lower_stmt(ctx, id, expr_arena, out)?;
            }
        }
        _ => lower_stmt(ctx, node_id, expr_arena, out)?,
    }
    Ok(())
}

fn lower_stmt(
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
        ExprNode::DimItem { .. } => Ok(()),
        ExprNode::Block { stmts } => {
            let ids: Vec<NodeId> = stmts.clone();
            for id in ids {
                lower_stmt(ctx, id, expr_arena, out)?;
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
        _ => Err(LowerError::UnsupportedNode),
    }
}

// ── Control-flow helpers ──────────────────────────────────────────────────────

/// Emit the expression for `node_id` to a scratch byte vector and return it.
fn lower_expr_to_bytes(
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
fn emit_branch_placeholder(out: &mut Vec<u8>, opcode: u8) -> usize {
    out.push(opcode);
    let patch = out.len();
    out.push(0x00);
    out.push(0x00);
    patch
}

/// Backpatch a 2-byte LE u16 at the given byte offset with the given value.
fn patch_u16(out: &mut Vec<u8>, patch: usize, value: u16) {
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
fn lower_if(
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
fn lower_while(
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
fn lower_do(
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

            let cond_bytes = lower_expr_to_bytes(ctx, cond, expr_arena)?;
            out.extend_from_slice(&cond_bytes);

            // BranchTrue to exit — Until means "exit when condition IS true"
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
fn lower_expr_to_bytes_coerced(
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
fn lower_for(
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

    lower_block(ctx, body_id, expr_arena, out)?;

    // LdAddr counter again before ForNext.
    out.push(0x04);
    out.extend_from_slice(&frame_var.to_le_bytes());

    // ForNext: 1-byte opcode + 2-byte frame_hidden + 2-byte back_offset.
    let fornext_opcode: u8 = if has_step { 0x67 } else { 0x66 };
    out.push(fornext_opcode);
    out.extend_from_slice(&frame_hidden.to_le_bytes());
    out.extend_from_slice(&body_start.to_le_bytes());

    // Backpatch ForInit exit_offset to current position.
    let exit_offset = out.len() as u16;
    out[exit_patch..exit_patch + 2].copy_from_slice(&exit_offset.to_le_bytes());

    Ok(())
}

fn lower_assign(
    ctx: &LowerCtx,
    target_id: NodeId,
    value_id: NodeId,
    expr_arena: &ExprArena,
    out: &mut Vec<u8>,
) -> Result<(), LowerError> {
    // Resolve the target first so its type can be used to coerce integer literals
    // in the value expression (e.g. `r = 1` where r is Long → Long literal).
    let resolution = ctx
        .module
        .resolutions
        .get(&target_id.0)
        .ok_or(LowerError::Unresolved)?;

    let coerce_tag = match resolution {
        NameResolution::Local { local_idx, .. } => vba_type_to_node_tag(ctx.local_type(*local_idx)),
        NameResolution::Param { param_idx, .. } => vba_type_to_node_tag(ctx.param_type(*param_idx)),
        NameResolution::ModuleVar(idx) => vba_type_to_node_tag(ctx.global_type(*idx)),
        _ => None,
    };

    let mut arena = NodeArena::new();
    let value_root = lower_expr_coerced(ctx, value_id, expr_arena, &mut arena, coerce_tag)?;

    let mut emitter = Emitter::new(&arena);
    emitter.emit_expr(value_root, 0);

    match resolution {
        NameResolution::Local { local_idx, .. } => {
            let ty = ctx.local_type(*local_idx);
            let slot = &ctx.local_slots[*local_idx];
            let sctx = load_store_ctx(ty).ok_or(LowerError::UnsupportedType)?;
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

// ── Expression lowering ───────────────────────────────────────────────────────

/// Returns the "wider" of two VBA types for numeric literal promotion.
/// VB6 widens integer literals to match the type of the wider operand.
fn wider_numeric_tag(a: Option<&VbaType>, b: Option<&VbaType>) -> Option<u16> {
    fn rank(t: &VbaType) -> u8 {
        match t {
            VbaType::Integer | VbaType::Boolean => 1,
            VbaType::Long   => 2,
            VbaType::Single => 3,
            VbaType::Double => 4,
            _ => 0,
        }
    }
    let ta = a.map(rank).unwrap_or(0);
    let tb = b.map(rank).unwrap_or(0);
    let wider = if ta >= tb { a } else { b };
    wider.and_then(|t| vba_type_to_node_tag(t))
}

fn lower_expr(
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
fn lower_expr_coerced(
    ctx: &LowerCtx,
    node_id: NodeId,
    expr_arena: &ExprArena,
    arena: &mut NodeArena,
    coerce_tag: Option<u16>,
) -> Result<NodeRef, LowerError> {
    match expr_arena.get(node_id) {
        ExprNode::Literal { lit } => lower_lit_coerced(lit, coerce_tag, arena),
        ExprNode::NameRef { .. } => lower_name_ref(ctx, node_id, arena),
        ExprNode::BinOp { op, lhs, rhs } => {
            let (op, lhs_id, rhs_id) = (*op, *lhs, *rhs);
            lower_binop(ctx, node_id, op, lhs_id, rhs_id, expr_arena, arena)
        }
        ExprNode::UnOp { op, operand } => {
            let (op, operand_id) = (*op, *operand);
            lower_unop(ctx, node_id, op, operand_id, expr_arena, arena)
        }
        ExprNode::Paren { inner } => {
            let inner_id = *inner;
            lower_expr_coerced(ctx, inner_id, expr_arena, arena, coerce_tag)
        }
        _ => Err(LowerError::UnsupportedNode),
    }
}

fn lower_lit(lit: &AstLit, arena: &mut NodeArena) -> Result<NodeRef, LowerError> {
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
        _ => Err(LowerError::UnsupportedNode),
    }
}

/// Like `lower_lit` but promotes an integer literal to `coerce_tag` when the
/// context type is wider (e.g. integer literal in a Long expression).
fn lower_lit_coerced(
    lit: &AstLit,
    coerce_tag: Option<u16>,
    arena: &mut NodeArena,
) -> Result<NodeRef, LowerError> {
    match coerce_tag {
        Some(tag) => match lit {
            AstLit::Int(v) => Ok(arena.alloc(NodeArena::node(1, tag, *v as u32, 0, 0, 0))),
            _ => lower_lit(lit, arena),
        },
        None => lower_lit(lit, arena),
    }
}

fn lower_name_ref(
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
            let ty = ctx.local_type(*local_idx);
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

/// Build a typed local/param load node (opcode 0x74 = ByVal, 0x75 = ByRef).
/// Node layout: w[4] = sym-child (frame offset in high 16 bits), w[5] = type_ctx.
fn build_frame_load_node(
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
fn build_global_load_node(
    arena: &mut NodeArena,
    ctx: usize,
    module_desc: u16,
    field_offset: u16,
) -> NodeRef {
    let packed = (module_desc as u32) | ((field_offset as u32) << 16);
    arena.alloc(NodeArena::node(0x77, 0, packed, ctx as u32, 0, 0))
}

fn lower_binop(
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
    let operand_coerce = wider_numeric_tag(lhs_ty, rhs_ty);

    let lhs_ref = lower_expr_coerced(ctx, lhs_id, expr_arena, arena, operand_coerce)?;
    let lhs_ref = coerce_operand(ctx, lhs_id, lhs_ref, operand_coerce, expr_arena, arena);
    let rhs_ref = lower_expr_coerced(ctx, rhs_id, expr_arena, arena, operand_coerce)?;
    let rhs_ref = coerce_operand(ctx, rhs_id, rhs_ref, operand_coerce, expr_arena, arena);

    // Power (`^`) is its own bound-node opcode (0x1a) and always yields Double.
    // It is byte-exact only when both operands are already Double (no operand
    // coercion to emit); other operand types need the coercion machinery, so
    // they are left unsupported rather than mis-emitted.
    if op == BinOpKind::Pow {
        let both_double = matches!(lhs_ty, Some(VbaType::Double))
            && matches!(rhs_ty, Some(VbaType::Double));
        if !both_double {
            return Err(LowerError::UnsupportedType);
        }
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
fn coerce_operand(
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
    if matches!(expr_arena.get(operand_id), ExprNode::Literal { .. }) {
        return operand_ref;
    }
    let Some(src_tag) = ctx
        .module
        .types
        .get(&operand_id.0)
        .and_then(vba_type_to_node_tag)
    else {
        return operand_ref;
    };
    if src_tag == target {
        return operand_ref;
    }
    arena.alloc(NodeArena::node(0x78, target, operand_ref.0, 0, 0, 0))
}

fn lower_unop(
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

#[cfg(test)]
#[path = "tests/lower_tests.rs"]
mod tests;
