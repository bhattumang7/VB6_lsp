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

use std::collections::HashMap;

use vb6_sema::sema::{BoundModule, BoundProc, VbaType, NameResolution};
use vb6_syntax::frontend::ast::{ExprArena, ExprNode, AstLit, BinOpKind, UnOpKind};
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
        BinOpKind::Or  => 0x21,
        BinOpKind::Xor => 0x22,
        BinOpKind::And => 0x23,
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

/// Lower a single `BoundProc` to its P-code byte vector.
///
/// Frame layout follows VB6's exact convention: locals at negative offsets
/// from -136 downward (4 bytes per Integer/Long/Single/Object, 8 bytes per
/// Double/Currency), params at positive offsets from +12 upward.
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

    let local_types: Vec<VbaType> = proc.locals.iter().map(|v| v.vba_type.clone()).collect();
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
        _ => Err(LowerError::UnsupportedNode),
    }
}

fn lower_assign(
    ctx: &LowerCtx,
    target_id: NodeId,
    value_id: NodeId,
    expr_arena: &ExprArena,
    out: &mut Vec<u8>,
) -> Result<(), LowerError> {
    let mut arena = NodeArena::new();
    let value_root = lower_expr(ctx, value_id, expr_arena, &mut arena)?;

    let resolution = ctx
        .module
        .resolutions
        .get(&target_id.0)
        .ok_or(LowerError::Unresolved)?;

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

fn lower_expr(
    ctx: &LowerCtx,
    node_id: NodeId,
    expr_arena: &ExprArena,
    arena: &mut NodeArena,
) -> Result<NodeRef, LowerError> {
    match expr_arena.get(node_id) {
        ExprNode::Literal { lit } => lower_lit(lit, arena),
        ExprNode::NameRef { .. } => lower_name_ref(ctx, node_id, arena),
        ExprNode::BinOp { op, lhs, rhs } => {
            let (op, lhs_id, rhs_id) = (*op, *lhs, *rhs);
            lower_binop(ctx, node_id, op, lhs_id, rhs_id, expr_arena, arena)
        }
        ExprNode::UnOp { op, operand } => {
            let (op, operand_id) = (*op, *operand);
            lower_unop(ctx, op, operand_id, expr_arena, arena)
        }
        ExprNode::Paren { inner } => {
            let inner_id = *inner;
            lower_expr(ctx, inner_id, expr_arena, arena)
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
    let lhs_ref = lower_expr(ctx, lhs_id, expr_arena, arena)?;
    let rhs_ref = lower_expr(ctx, rhs_id, expr_arena, arena)?;

    let opcode = binop_node_opcode(op).ok_or(LowerError::UnsupportedNode)?;

    let type_tag = if is_comparison_op(op) {
        0u16
    } else {
        let result_type = ctx.module.types.get(&node_id.0).ok_or(LowerError::Unresolved)?;
        vba_type_to_node_tag(result_type).ok_or(LowerError::UnsupportedType)?
    };

    Ok(arena.alloc(NodeArena::node(opcode, type_tag, lhs_ref.0, rhs_ref.0, 0, 0)))
}

fn lower_unop(
    ctx: &LowerCtx,
    op: UnOpKind,
    operand_id: NodeId,
    expr_arena: &ExprArena,
    arena: &mut NodeArena,
) -> Result<NodeRef, LowerError> {
    let operand_ref = lower_expr(ctx, operand_id, expr_arena, arena)?;
    match op {
        UnOpKind::Pos => Ok(operand_ref),
        UnOpKind::Neg => Ok(arena.alloc(NodeArena::node(0x0b, 0, operand_ref.0, 0, 0, 0))),
        UnOpKind::Not => Ok(arena.alloc(NodeArena::node(0x10, 0, operand_ref.0, 0, 0, 0))),
    }
}

#[cfg(test)]
#[path = "tests/lower_tests.rs"]
mod tests;
