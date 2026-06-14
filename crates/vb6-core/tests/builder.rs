//! Structural tests for binder-phase builder functions.
//!
//! Functions under test:
//!   build_field_access_node
//!   build_let_accessor_node
//!   build_member_ref_node
//!   create_binding_node
//!   build_proc_call_expr
//!   build_impls_proc_node

use vb6_core::context::CompilerContext;
use vb6_core::frontend::ast::{alloc_decl_node, DeclSymKind, ScopeNodeKind};
use vb6_core::builder::{AccessSpec, BuilderCtx, build_field_access_node,
    build_impls_proc_node, build_let_accessor_node, build_member_ref_node, build_proc_call_expr,
    create_binding_node};

fn make_ctx_with_parent() -> (CompilerContext, vb6_core::frontend::ast::NodeId) {
    let mut ctx = CompilerContext::new();
    let parent = alloc_decl_node(&mut ctx.decls, &mut ctx.scopes, DeclSymKind::K09, None, 0, None);
    (ctx, parent)
}

fn default_bctx(parent_decl: vb6_core::frontend::ast::NodeId) -> BuilderCtx {
    BuilderCtx {
        result_decl: None,
        result_scope: None,
        field_index: 3,
        type_ref: 0,
        access_spec: AccessSpec::default(),
        parent_decl,
        scope_block: None,
    }
}

// ── build_field_access_node ───────────────────────────────────────────────────

#[test]
fn build_field_access_node_creates_k06_decl() {
    let (mut ctx, parent) = make_ctx_with_parent();
    let mut bctx = default_bctx(parent);
    build_field_access_node(&mut ctx, &mut bctx);
    let decl_id = bctx.result_decl.expect("result_decl should be set");
    assert_eq!(ctx.decls.get(decl_id).kind, DeclSymKind::K06);
}

#[test]
fn build_field_access_node_creates_k1a_scope() {
    let (mut ctx, parent) = make_ctx_with_parent();
    let mut bctx = default_bctx(parent);
    build_field_access_node(&mut ctx, &mut bctx);
    let scope_id = bctx.result_scope.expect("result_scope should be set");
    assert_eq!(ctx.scope_nodes.get(scope_id).kind, ScopeNodeKind::K1a);
}

#[test]
fn build_field_access_node_links_scope_to_decl() {
    let (mut ctx, parent) = make_ctx_with_parent();
    let mut bctx = default_bctx(parent);
    build_field_access_node(&mut ctx, &mut bctx);
    let scope_id = bctx.result_scope.unwrap();
    let decl_id = bctx.result_decl.unwrap();
    assert_eq!(ctx.scope_nodes.get(scope_id).linked_decl, Some(decl_id));
    assert_eq!(ctx.decls.get(decl_id).scope_parent, Some(scope_id));
}

#[test]
fn build_field_access_node_sets_field_34() {
    let (mut ctx, parent) = make_ctx_with_parent();
    let mut bctx = default_bctx(parent);
    bctx.field_index = 7;
    build_field_access_node(&mut ctx, &mut bctx);
    let decl_id = bctx.result_decl.unwrap();
    assert_eq!(ctx.decls.get(decl_id).field_34, 7);
}

#[test]
fn build_field_access_node_copies_access_spec_words_into_scope_extra() {
    let (mut ctx, parent) = make_ctx_with_parent();
    let mut bctx = default_bctx(parent);
    bctx.access_spec.words = [0x11, 0x22, 0x33, 0x44];
    build_field_access_node(&mut ctx, &mut bctx);
    let scope_id = bctx.result_scope.unwrap();
    assert_eq!(ctx.scope_nodes.get(scope_id).extra, [0x11, 0x22, 0x33, 0x44]);
}

#[test]
fn build_field_access_node_reads_slot_info_from_parent_field_type_table() {
    // A field-access K1a node's slot_info (+0x1c) comes from the parent decl's
    // field-type-table entry at index 0xd + slot, where the discriminant selects
    // the slot (1→4, 2→6, 4→8).
    let (mut ctx, parent) = make_ctx_with_parent();
    // Populate the parent's field-type table so slot 6 (discriminant 2) maps to
    // index 0xd + 6 = 0x13.
    let mut table = vec![0u32; 0x14];
    table[0x13] = 0xCAFE;
    ctx.decls.get_mut(parent).field_type_table = table;

    let mut bctx = default_bctx(parent);
    bctx.access_spec.discriminant = 2; // slot 6
    build_field_access_node(&mut ctx, &mut bctx);

    let scope_id = bctx.result_scope.unwrap();
    assert_eq!(ctx.scope_nodes.get(scope_id).slot_info, 0xCAFE,
        "slot_info must come from parent.field_type_table[0xd + slot]");
}

#[test]
fn build_field_access_node_slot_info_zero_when_table_empty() {
    // With no field-type table populated, slot_info reads as 0 (no entry) —
    // not a fabricated value.
    let (mut ctx, parent) = make_ctx_with_parent();
    let mut bctx = default_bctx(parent);
    bctx.access_spec.discriminant = 1;
    build_field_access_node(&mut ctx, &mut bctx);
    let scope_id = bctx.result_scope.unwrap();
    assert_eq!(ctx.scope_nodes.get(scope_id).slot_info, 0);
}

#[test]
fn build_field_access_node_sets_flags_10_8_and_2() {
    let (mut ctx, parent) = make_ctx_with_parent();
    let mut bctx = default_bctx(parent);
    build_field_access_node(&mut ctx, &mut bctx);
    let decl_id = bctx.result_decl.unwrap();
    assert_eq!(ctx.decls.get(decl_id).flags_10 & (8 | 2), 8 | 2);
}

#[test]
fn build_field_access_node_flags_0x210_sets_flags_10_bit_4() {
    let (mut ctx, parent) = make_ctx_with_parent();
    let mut bctx = default_bctx(parent);
    bctx.access_spec.flags = 0x210;
    build_field_access_node(&mut ctx, &mut bctx);
    let decl_id = bctx.result_decl.unwrap();
    assert_eq!(ctx.decls.get(decl_id).flags_10 & 0x10, 0x10);
}

// ── build_let_accessor_node ───────────────────────────────────────────────────

#[test]
fn build_let_accessor_node_creates_k07_decl_and_k14_scope() {
    let (mut ctx, parent) = make_ctx_with_parent();
    let mut bctx = default_bctx(parent);
    build_let_accessor_node(&mut ctx, &mut bctx);
    let decl_id = bctx.result_decl.unwrap();
    let scope_id = bctx.result_scope.unwrap();
    assert_eq!(ctx.decls.get(decl_id).kind, DeclSymKind::K07);
    assert_eq!(ctx.scope_nodes.get(scope_id).kind, ScopeNodeKind::K14);
}

#[test]
fn build_let_accessor_node_sets_field_38() {
    let (mut ctx, parent) = make_ctx_with_parent();
    let mut bctx = default_bctx(parent);
    bctx.field_index = 5;
    build_let_accessor_node(&mut ctx, &mut bctx);
    let decl_id = bctx.result_decl.unwrap();
    assert_eq!(ctx.decls.get(decl_id).field_38, 5);
}

#[test]
fn build_let_accessor_node_setter_flag_sets_flags_10_bit_4() {
    let (mut ctx, parent) = make_ctx_with_parent();
    let mut bctx = default_bctx(parent);
    bctx.access_spec.flags = 0x10;
    build_let_accessor_node(&mut ctx, &mut bctx);
    let decl_id = bctx.result_decl.unwrap();
    assert_eq!(ctx.decls.get(decl_id).flags_10 & 0x10, 0x10);
}

// ── build_member_ref_node ─────────────────────────────────────────────────────

#[test]
fn build_member_ref_node_creates_k03_decl_and_k18_scope() {
    let (mut ctx, parent) = make_ctx_with_parent();
    let mut bctx = default_bctx(parent);
    build_member_ref_node(&mut ctx, &mut bctx, 0);
    let decl_id = bctx.result_decl.unwrap();
    let scope_id = bctx.result_scope.unwrap();
    assert_eq!(ctx.decls.get(decl_id).kind, DeclSymKind::K03);
    assert_eq!(ctx.scope_nodes.get(scope_id).kind, ScopeNodeKind::K18);
}

#[test]
fn build_member_ref_node_stores_type_result_in_extra0() {
    let (mut ctx, parent) = make_ctx_with_parent();
    let mut bctx = default_bctx(parent);
    build_member_ref_node(&mut ctx, &mut bctx, 0xdeadbeef);
    let scope_id = bctx.result_scope.unwrap();
    assert_eq!(ctx.scope_nodes.get(scope_id).extra[0], 0xdeadbeef);
}

#[test]
fn build_member_ref_node_sets_field_2c() {
    let (mut ctx, parent) = make_ctx_with_parent();
    let mut bctx = default_bctx(parent);
    bctx.field_index = 2;
    build_member_ref_node(&mut ctx, &mut bctx, 0);
    let decl_id = bctx.result_decl.unwrap();
    assert_eq!(ctx.decls.get(decl_id).field_2c, 2);
}

// ── create_binding_node ───────────────────────────────────────────────────────

#[test]
fn create_binding_node_creates_module_scope_and_k08_decl() {
    let (mut ctx, parent) = make_ctx_with_parent();
    let mut bctx = default_bctx(parent);
    create_binding_node(&mut ctx, &mut bctx);
    let scope_id = bctx.result_scope.unwrap();
    let decl_id = bctx.result_decl.unwrap();
    assert_eq!(ctx.scope_nodes.get(scope_id).kind, ScopeNodeKind::Module);
    assert_eq!(ctx.decls.get(decl_id).kind, DeclSymKind::K08);
}

#[test]
fn create_binding_node_sets_type_info_2() {
    let (mut ctx, parent) = make_ctx_with_parent();
    let mut bctx = default_bctx(parent);
    create_binding_node(&mut ctx, &mut bctx);
    let decl_id = bctx.result_decl.unwrap();
    assert_eq!(ctx.decls.get(decl_id).type_info, 2);
}

#[test]
fn create_binding_node_sets_field_68() {
    let (mut ctx, parent) = make_ctx_with_parent();
    let mut bctx = default_bctx(parent);
    bctx.field_index = 4;
    create_binding_node(&mut ctx, &mut bctx);
    let decl_id = bctx.result_decl.unwrap();
    assert_eq!(ctx.decls.get(decl_id).field_68, 4);
}

#[test]
fn create_binding_node_sets_flags_2() {
    let (mut ctx, parent) = make_ctx_with_parent();
    let mut bctx = default_bctx(parent);
    create_binding_node(&mut ctx, &mut bctx);
    let decl_id = bctx.result_decl.unwrap();
    assert_eq!(ctx.decls.get(decl_id).flags & 0x3f, 2);
}

// ── build_proc_call_expr ──────────────────────────────────────────────────────

#[test]
fn build_proc_call_expr_creates_module_scope_and_k08_decl() {
    let (mut ctx, parent) = make_ctx_with_parent();
    let mut bctx = default_bctx(parent);
    build_proc_call_expr(&mut ctx, &mut bctx);
    let scope_id = bctx.result_scope.unwrap();
    let decl_id = bctx.result_decl.unwrap();
    assert_eq!(ctx.scope_nodes.get(scope_id).kind, ScopeNodeKind::Module);
    assert_eq!(ctx.decls.get(decl_id).kind, DeclSymKind::K08);
}

#[test]
fn build_proc_call_expr_call_kind_2_sets_flags_3() {
    // call_kind == 2 → new_flags = 3
    let (mut ctx, parent) = make_ctx_with_parent();
    let mut bctx = default_bctx(parent);
    bctx.access_spec.call_kind = 2;
    build_proc_call_expr(&mut ctx, &mut bctx);
    let decl_id = bctx.result_decl.unwrap();
    assert_eq!(ctx.decls.get(decl_id).flags & 0x3f, 3);
}

#[test]
fn build_proc_call_expr_call_kind_not_2_sets_flags_2() {
    // call_kind != 2 → new_flags = 2
    let (mut ctx, parent) = make_ctx_with_parent();
    let mut bctx = default_bctx(parent);
    bctx.access_spec.call_kind = 0;
    build_proc_call_expr(&mut ctx, &mut bctx);
    let decl_id = bctx.result_decl.unwrap();
    assert_eq!(ctx.decls.get(decl_id).flags & 0x3f, 2);
}

#[test]
fn build_proc_call_expr_copies_access_spec_words_to_scope_extra() {
    let (mut ctx, parent) = make_ctx_with_parent();
    let mut bctx = default_bctx(parent);
    bctx.access_spec.words = [1, 2, 3, 4];
    build_proc_call_expr(&mut ctx, &mut bctx);
    let scope_id = bctx.result_scope.unwrap();
    assert_eq!(ctx.scope_nodes.get(scope_id).extra, [1, 2, 3, 4]);
}

#[test]
fn build_proc_call_expr_sets_field_68() {
    let (mut ctx, parent) = make_ctx_with_parent();
    let mut bctx = default_bctx(parent);
    bctx.field_index = 6;
    build_proc_call_expr(&mut ctx, &mut bctx);
    let decl_id = bctx.result_decl.unwrap();
    assert_eq!(ctx.decls.get(decl_id).field_68, 6);
}

// ── build_impls_proc_node ─────────────────────────────────────────────────────

#[test]
fn build_impls_proc_node_creates_module_scope_and_k08_decl() {
    let (mut ctx, parent) = make_ctx_with_parent();
    let mut bctx = default_bctx(parent);
    build_impls_proc_node(&mut ctx, &mut bctx);
    let scope_id = bctx.result_scope.unwrap();
    let decl_id = bctx.result_decl.unwrap();
    assert_eq!(ctx.scope_nodes.get(scope_id).kind, ScopeNodeKind::Module);
    assert_eq!(ctx.decls.get(decl_id).kind, DeclSymKind::K08);
}

#[test]
fn build_impls_proc_node_sets_flags_10_0a_and_sec_flags_8() {
    let (mut ctx, parent) = make_ctx_with_parent();
    let mut bctx = default_bctx(parent);
    build_impls_proc_node(&mut ctx, &mut bctx);
    let decl_id = bctx.result_decl.unwrap();
    assert_eq!(ctx.decls.get(decl_id).flags_10 & 0x0a, 0x0a);
    assert_eq!(ctx.decls.get(decl_id).sec_flags & 8, 8);
}

#[test]
fn build_impls_proc_node_sets_flags_13_bit3() {
    let (mut ctx, parent) = make_ctx_with_parent();
    let mut bctx = default_bctx(parent);
    build_impls_proc_node(&mut ctx, &mut bctx);
    let decl_id = bctx.result_decl.unwrap();
    assert_eq!(ctx.decls.get(decl_id).flags_13 & 8, 8);
}

#[test]
fn build_impls_proc_node_sets_type_info_2_and_field_68() {
    let (mut ctx, parent) = make_ctx_with_parent();
    let mut bctx = default_bctx(parent);
    bctx.field_index = 5;
    build_impls_proc_node(&mut ctx, &mut bctx);
    let decl_id = bctx.result_decl.unwrap();
    assert_eq!(ctx.decls.get(decl_id).type_info, 2);
    assert_eq!(ctx.decls.get(decl_id).field_68, 5);
}

#[test]
fn build_impls_proc_node_access_flags_2_sets_flags_13_bit5() {
    // access_spec.flags & 2 → flags_13 |= 0x20
    let (mut ctx, parent) = make_ctx_with_parent();
    let mut bctx = default_bctx(parent);
    bctx.access_spec.flags = 2;
    build_impls_proc_node(&mut ctx, &mut bctx);
    let decl_id = bctx.result_decl.unwrap();
    assert_eq!(ctx.decls.get(decl_id).flags_13 & 0x20, 0x20);
}

#[test]
fn build_impls_proc_node_access_flags_1_then_ext_flags_has_bits_4_and_8() {
    // access_spec.flags & 1 → ext_flags |= 4; then (ext_flags&4) → ext_flags |= 8, flags → 3
    let (mut ctx, parent) = make_ctx_with_parent();
    let mut bctx = default_bctx(parent);
    bctx.access_spec.flags = 1;
    build_impls_proc_node(&mut ctx, &mut bctx);
    let decl_id = bctx.result_decl.unwrap();
    // ext_flags should have bits 1 (from link_kind=2) | 4 | 8
    assert_eq!(ctx.decls.get(decl_id).ext_flags & (4 | 8), 4 | 8);
    assert_eq!(ctx.decls.get(decl_id).flags & 0x3f, 3);
}

#[test]
fn build_impls_proc_node_copies_access_spec_words_to_scope_extra() {
    let (mut ctx, parent) = make_ctx_with_parent();
    let mut bctx = default_bctx(parent);
    bctx.access_spec.words = [0xa, 0xb, 0xc, 0xd];
    build_impls_proc_node(&mut ctx, &mut bctx);
    let scope_id = bctx.result_scope.unwrap();
    assert_eq!(ctx.scope_nodes.get(scope_id).extra, [0xa, 0xb, 0xc, 0xd]);
}

#[test]
fn build_impls_proc_node_sets_scope_flags_b6_bit5() {
    let (mut ctx, parent) = make_ctx_with_parent();
    let mut bctx = default_bctx(parent);
    build_impls_proc_node(&mut ctx, &mut bctx);
    let scope_id = bctx.result_scope.unwrap();
    assert_eq!(ctx.scope_nodes.get(scope_id).flags_b6 & 0x20, 0x20);
}
