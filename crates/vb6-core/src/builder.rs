//! Binder-phase node builder context and builder functions.
//!
//! Provides the context and logic for building binder-phase AST nodes,
//! managing scope nodes and declaration linkages.
//!
//! This module implements the internal node building logic used during
//! semantic analysis to construct the resolved AST.

use crate::context::CompilerContext;
use crate::frontend::ast::{
    alloc_decl_node, ast_node_create, ast_node_link, DeclSymKind, NodeId, ScopeNodeKind,
};

/// Access-specifier / field-descriptor struct.
///
/// Contains metadata for field and member access, used by builder functions
/// to configure declaration and scope nodes.
#[derive(Clone, Default)]
pub struct AccessSpec {
    /// Internal flags and data words.
    pub words: [u32; 4],

    /// Dispatch discriminant for member access.
    pub discriminant: u32,

    /// Call or access kind.
    pub call_kind: u32,

    /// Internal flags.
    pub short_30: u16,

    /// Property-accessor and visibility flags.
    pub flags: u16,

    /// Type reference identifier.
    pub type_ref_3c: u32,
}

/// Builder context shared by node builder functions.
///
/// Tracks input parameters and output results during the construction
/// of declaration and scope nodes.
pub struct BuilderCtx {
    /// OUTPUT — resulting DeclNode id (set during build).
    pub result_decl: Option<NodeId>,
    /// OUTPUT — resulting ScopeNode id (set during build).
    pub result_scope: Option<NodeId>,
    /// INPUT — field/slot index within the parent's member list.
    pub field_index: u32,
    /// INPUT — type-expression identifier.
    pub type_ref: u32,
    /// INPUT — access specifier / field descriptor.
    pub access_spec: AccessSpec,
    /// INPUT — parent DeclNode id.
    pub parent_decl: NodeId,
    /// INPUT — scope block identifier, or `None`.
    pub scope_block: Option<NodeId>,
}

// ---------------------------------------------------------------------------
// Field Access Node
// ---------------------------------------------------------------------------

/// Builds a field-access declaration.
///
/// Allocates a K06 DeclNode and a K1a ScopeNode, links them, and sets
/// property flags from the access specifier.
pub fn build_field_access_node(ctx: &mut CompilerContext, bctx: &mut BuilderCtx) -> i32 {
    let spec = &bctx.access_spec;

    // Map discriminant to slot index (1→4, 2→6, 4→8).
    let slot: u32 = match spec.discriminant {
        1 => 4,
        2 => 6,
        4 => 8,
        _ => 0,
    };

    // The K1a node's slot_info is the parent's field-type-table entry for this
    // slot: parent[0x38][0x34 + slot*4] → table[0xd + slot] (dword index). The
    // table is empty until the type-layout pass populates it, in which case the
    // entry is absent and slot_info is 0.
    let slot_info = ctx
        .decls
        .get(bctx.parent_decl)
        .field_type_table
        .get(0xd + slot as usize)
        .copied()
        .unwrap_or(0);

    // Allocate K06 DeclNode.
    let decl_id = alloc_decl_node(
        &mut ctx.decls,
        &mut ctx.scopes,
        DeclSymKind::K06,
        bctx.scope_block,
        3,
        Some(bctx.parent_decl),
    );
    bctx.result_decl = Some(decl_id);

    // Allocate K1a ScopeNode.
    let scope_id = ast_node_create(&mut ctx.scope_nodes, ScopeNodeKind::K1a);
    bctx.result_scope = Some(scope_id);

    // Wire scope → decl.
    ctx.scope_nodes.get_mut(scope_id).linked_decl = Some(decl_id);

    // Copy AccessSpec words into scope node extra fields.
    ctx.scope_nodes.get_mut(scope_id).extra = spec.words;
    // Slot/type entry from the parent's field-type table (offset 0x1c).
    ctx.scope_nodes.get_mut(scope_id).slot_info = slot_info;

    // Wire decl → scope.
    ctx.decls.get_mut(decl_id).scope_parent = Some(scope_id);

    // Set field index.
    ctx.decls.get_mut(decl_id).field_34 = bctx.field_index as u16;

    // Set property flags.
    ctx.decls.get_mut(decl_id).flags_10 |= 8 | 2;
    if spec.flags & 0x210 != 0 {
        ctx.decls.get_mut(decl_id).flags_10 |= 0x10;
    }

    0
}

// ---------------------------------------------------------------------------
// Let Accessor Node
// ---------------------------------------------------------------------------

/// Builds a property-let accessor.
pub fn build_let_accessor_node(ctx: &mut CompilerContext, bctx: &mut BuilderCtx) -> i32 {
    let decl_id = alloc_decl_node(
        &mut ctx.decls,
        &mut ctx.scopes,
        DeclSymKind::K07,
        bctx.scope_block,
        2,
        Some(bctx.parent_decl),
    );
    bctx.result_decl = Some(decl_id);

    let scope_id = ast_node_create(&mut ctx.scope_nodes, ScopeNodeKind::K14);
    bctx.result_scope = Some(scope_id);

    ctx.scope_nodes.get_mut(scope_id).linked_decl = Some(decl_id);

    // Set field index.
    ctx.decls.get_mut(decl_id).field_38 = bctx.field_index as u16;

    // Set flags.
    ctx.decls.get_mut(decl_id).flags_10 |= 8 | 2;
    if bctx.access_spec.flags & 0x10 != 0 {
        ctx.decls.get_mut(decl_id).flags_10 |= 0x10;
    }

    // Wire decl → scope.
    ctx.decls.get_mut(decl_id).scope_parent = Some(scope_id);

    0
}

// ---------------------------------------------------------------------------
// Member Reference Node
// ---------------------------------------------------------------------------

/// Builds a member-reference declaration.
///
/// Called after the member's type is resolved.
pub fn build_member_ref_node(
    ctx: &mut CompilerContext,
    bctx: &mut BuilderCtx,
    type_result: u32,
) -> i32 {
    let decl_id = alloc_decl_node(
        &mut ctx.decls,
        &mut ctx.scopes,
        DeclSymKind::K03,
        bctx.scope_block,
        3,
        Some(bctx.parent_decl),
    );
    bctx.result_decl = Some(decl_id);

    let scope_id = ast_node_create(&mut ctx.scope_nodes, ScopeNodeKind::K18);
    bctx.result_scope = Some(scope_id);

    ctx.scope_nodes.get_mut(scope_id).linked_decl = Some(decl_id);
    ctx.scope_nodes.get_mut(scope_id).extra[0] = type_result;

    // Wire decl → scope.
    ctx.decls.get_mut(decl_id).scope_parent = Some(scope_id);

    // Set flags.
    ctx.decls.get_mut(decl_id).flags_10 |= 8 | 2;
    if bctx.access_spec.flags & 0x210 != 0 {
        ctx.decls.get_mut(decl_id).flags_10 |= 0x10;
    }

    // Set field index.
    ctx.decls.get_mut(decl_id).field_2c = bctx.field_index as u16;

    0
}

// ---------------------------------------------------------------------------
// Binding Node
// ---------------------------------------------------------------------------

/// Creates a binding node for a named member in a module scope.
pub fn create_binding_node(ctx: &mut CompilerContext, bctx: &mut BuilderCtx) -> i32 {
    let scope_id = ast_node_create(&mut ctx.scope_nodes, ScopeNodeKind::Module);
    bctx.result_scope = Some(scope_id);

    let decl_id = ast_node_link(
        &mut ctx.scope_nodes,
        &mut ctx.decls,
        &mut ctx.scopes,
        scope_id,
        bctx.scope_block,
        Some(bctx.parent_decl),
        2,
    );
    bctx.result_decl = Some(decl_id);

    // Adjust flags.
    ctx.decls.get_mut(decl_id).flags = (ctx.decls.get(decl_id).flags & 0xc2) | 2;
    ctx.decls.get_mut(decl_id).flags_10 |= 0x0a;
    ctx.decls.get_mut(decl_id).sec_flags |= 8;
    ctx.decls.get_mut(decl_id).type_info = 2;
    ctx.decls.get_mut(decl_id).field_68 = bctx.field_index as u16;
    ctx.decls.get_mut(decl_id).flags_10 |= 0x10;

    // Wire decl → scope.
    ctx.decls.get_mut(decl_id).scope_parent = Some(scope_id);

    0
}

// ---------------------------------------------------------------------------
// Procedure Call Expression Node
// ---------------------------------------------------------------------------

/// Builds a procedure-call expression node.
pub fn build_proc_call_expr(ctx: &mut CompilerContext, bctx: &mut BuilderCtx) -> i32 {
    let bvar6 = bctx.access_spec.call_kind == 2;
    let link_kind: u32 = if bvar6 { 0 } else { 2 };

    let scope_id = ast_node_create(&mut ctx.scope_nodes, ScopeNodeKind::Module);
    bctx.result_scope = Some(scope_id);

    let decl_id = ast_node_link(
        &mut ctx.scope_nodes,
        &mut ctx.decls,
        &mut ctx.scopes,
        scope_id,
        bctx.scope_block,
        Some(bctx.parent_decl),
        link_kind,
    );
    bctx.result_decl = Some(decl_id);

    // Set flags.
    let new_flags: u8 = if bvar6 { 3 } else { 2 };
    let old_flags = ctx.decls.get(decl_id).flags;
    ctx.decls.get_mut(decl_id).flags = (new_flags ^ old_flags) & 0x3f ^ old_flags;

    // Copy access_spec words into scope node extra fields.
    ctx.scope_nodes.get_mut(scope_id).extra = bctx.access_spec.words;

    // Set decl fields.
    ctx.decls.get_mut(decl_id).flags_10 |= 0x0a;
    ctx.decls.get_mut(decl_id).sec_flags |= 8;
    ctx.decls.get_mut(decl_id).type_info = 2;
    ctx.decls.get_mut(decl_id).field_68 = bctx.field_index as u16;

    if bctx.access_spec.flags & 0x210 != 0 {
        ctx.decls.get_mut(decl_id).flags_10 |= 0x10;
    }

    // Wire decl → scope.
    ctx.decls.get_mut(decl_id).scope_parent = Some(scope_id);

    0
}

// ---------------------------------------------------------------------------
// Implements Procedure Node
// ---------------------------------------------------------------------------

/// Builds an implements-procedure node.
///
/// Sets multiple flag bytes and copies access specifier data into the new nodes.
pub fn build_impls_proc_node(ctx: &mut CompilerContext, bctx: &mut BuilderCtx) -> i32 {
    let scope_id = ast_node_create(&mut ctx.scope_nodes, ScopeNodeKind::Module);
    bctx.result_scope = Some(scope_id);

    let decl_id = ast_node_link(
        &mut ctx.scope_nodes,
        &mut ctx.decls,
        &mut ctx.scopes,
        scope_id,
        bctx.scope_block,
        Some(bctx.parent_decl),
        2,
    );
    bctx.result_decl = Some(decl_id);

    // Set flags.
    {
        let decl = ctx.decls.get_mut(decl_id);
        decl.flags = (decl.flags & 0xc2) | 2;
        decl.flags_10 |= 0x0a;
        decl.sec_flags |= 8;
        decl.flags_13 |= 8;
        decl.type_info = 2;
        decl.field_68 = bctx.field_index as u16;
    }

    if bctx.access_spec.flags & 0x210 != 0 {
        ctx.decls.get_mut(decl_id).flags_10 |= 0x10;
    }
    if bctx.access_spec.flags & 2 != 0 {
        ctx.decls.get_mut(decl_id).flags_13 |= 0x20;
    }
    if bctx.access_spec.flags & 1 != 0 {
        ctx.decls.get_mut(decl_id).ext_flags |= 4;
    }
    let ext4 = ctx.decls.get(decl_id).ext_flags & 4 != 0;
    let flags8 = bctx.access_spec.flags & 8 != 0;
    if ext4 || flags8 {
        let decl = ctx.decls.get_mut(decl_id);
        decl.ext_flags |= 8;
        decl.flags = (decl.flags & 0xc3) | 3;
    }

    ctx.scope_nodes.get_mut(scope_id).extra = bctx.access_spec.words;
    ctx.scope_nodes.get_mut(scope_id).flags_b6 |= 0x20;

    // Wire decl → scope.
    ctx.decls.get_mut(decl_id).scope_parent = Some(scope_id);

    0
}
