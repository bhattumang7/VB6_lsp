//! Structural tests for AST node constructors. These tests verify the
//! constructors build well-formed nodes.
//!
//! Constructors under test:
//!   make_type_spec_node
//!   make_udt_type_node
//!   append_list_node
//!   ast_node_create
//!   alloc_scope_block
//!   ast_node_link

use vb6_core::frontend::ast::{
    alloc_ast_node, alloc_decl_node, alloc_scope_block, append_list_node, ast_node_create,
    ast_node_link, build_abstract_member_node, build_array_bounds_node, build_for_node,
    build_on_node, emit_simple_node, make_type_node, make_type_spec_node, make_udt_type_node,
    DeclArena, DeclSymKind, Diagnostics, ExprArena, ExprNode, NodeId, ScopeArena,
    ScopeBlockArena, ScopeNodeKind, Span, VARIANT_TYPE_DESC_MARKER,
};

// ── make_type_spec_node ──────────────────────────────────────────────────────

#[test]
fn type_spec_no_child_round_trips() {
    let mut arena = ExprArena::new();
    let id = make_type_spec_node(&mut arena, 0, 0, 2 /* Integer */, None);

    match arena.get(id) {
        ExprNode::TypeSpec { type_flags, parent_scope, type_kind, child } => {
            assert_eq!(*type_flags, 0);
            assert_eq!(*parent_scope, 0);
            assert_eq!(*type_kind, 2);
            assert!(child.is_none());
        }
        _ => panic!("expected TypeSpec"),
    }
}

#[test]
fn type_spec_with_child_round_trips() {
    let mut arena = ExprArena::new();
    // First allocate a dummy child node.
    let child_id = make_type_spec_node(&mut arena, 0, 0, 0x10 /* Object */, None);
    let parent_id =
        make_type_spec_node(&mut arena, 0, 5, 0x11 /* qualified Object */, Some(child_id));

    match arena.get(parent_id) {
        ExprNode::TypeSpec { type_flags, parent_scope, type_kind, child } => {
            assert_eq!(*type_flags, 0);
            assert_eq!(*parent_scope, 5);
            assert_eq!(*type_kind, 0x11);
            assert_eq!(*child, Some(child_id));
        }
        _ => panic!("expected TypeSpec"),
    }
}

#[test]
fn type_spec_fixed_len_flag() {
    // 0x4000 = fixed-length String * n flag
    let mut arena = ExprArena::new();
    let id = make_type_spec_node(&mut arena, 0x4000, 0, 8 /* String */, None);

    match arena.get(id) {
        ExprNode::TypeSpec { type_flags, .. } => assert_eq!(*type_flags, 0x4000),
        _ => panic!("expected TypeSpec"),
    }
}

#[test]
fn type_spec_udt_flag() {
    // 0x8000 = UDT path (the UDT constructor sets this flag)
    let mut arena = ExprArena::new();
    let id = make_type_spec_node(&mut arena, 0x8000, 3, 0, None);

    match arena.get(id) {
        ExprNode::TypeSpec { type_flags, parent_scope, .. } => {
            assert_eq!(*type_flags, 0x8000);
            assert_eq!(*parent_scope, 3);
        }
        _ => panic!("expected TypeSpec"),
    }
}

// ── make_udt_type_node ───────────────────────────────────────────────────────

#[test]
fn udt_type_spec_always_has_0x8000_flag() {
    let mut arena = ExprArena::new();
    let id = make_udt_type_node(&mut arena, 0, 0, 3, None);
    match arena.get(id) {
        ExprNode::UdtTypeSpec { flags, .. } => assert_eq!(*flags & 0x8000, 0x8000),
        _ => panic!("expected UdtTypeSpec"),
    }
}

#[test]
fn udt_type_spec_fixed_len_and_udt_flags_combine() {
    // 0x4000 (fixed-length String) ORed with 0x8000 (UDT) = 0xc000
    let mut arena = ExprArena::new();
    let id = make_udt_type_node(&mut arena, 0x4000, 0, 1, None);
    match arena.get(id) {
        ExprNode::UdtTypeSpec { flags, .. } => assert_eq!(*flags, 0xc000),
        _ => panic!("expected UdtTypeSpec"),
    }
}

#[test]
fn udt_type_spec_fields_round_trip() {
    let mut arena = ExprArena::new();
    let list_id = make_udt_type_node(&mut arena, 0, 0, 0, None);
    let id = make_udt_type_node(&mut arena, 0, 7, 4, Some(list_id));
    match arena.get(id) {
        ExprNode::UdtTypeSpec { flags, parent_scope, udt_count, type_list } => {
            assert_eq!(*flags, 0x8000);
            assert_eq!(*parent_scope, 7);
            assert_eq!(*udt_count, 4);
            assert_eq!(*type_list, Some(list_id));
        }
        _ => panic!("expected UdtTypeSpec"),
    }
}

// ── append_list_node ─────────────────────────────────────────────────────────

#[test]
fn append_list_node_builds_ordered_list() {
    let mut arena = ExprArena::new();
    let a = make_type_spec_node(&mut arena, 0, 0, 2, None);
    let b = make_type_spec_node(&mut arena, 0, 0, 3, None);
    let c = make_type_spec_node(&mut arena, 0, 0, 8, None);

    let mut list: Vec<NodeId> = Vec::new();
    append_list_node(&mut list, a);
    append_list_node(&mut list, b);
    append_list_node(&mut list, c);

    assert_eq!(list, vec![a, b, c]);
}

#[test]
fn append_list_node_empty_list_stays_empty_on_no_calls() {
    let list: Vec<NodeId> = Vec::new();
    assert!(list.is_empty());
}

#[test]
fn arena_ids_are_sequential() {
    let mut arena = ExprArena::new();
    let id0 = make_type_spec_node(&mut arena, 0, 0, 2, None);
    let id1 = make_type_spec_node(&mut arena, 0, 0, 3, None);
    let id2 = make_type_spec_node(&mut arena, 0, 0, 8, None);

    assert_eq!(id0, NodeId(0));
    assert_eq!(id1, NodeId(1));
    assert_eq!(id2, NodeId(2));
    assert_eq!(arena.len(), 3);
}

// ── ast_node_create ───────────────────────────────────────────────────────────

#[test]
fn ast_node_create_module_allocates_two_nodes() {
    // kind 0x13 (Module) must create a paired 0x16 (ModuleList) child.
    // The parent references its child, and the child back-references its parent.
    let mut arena = ScopeArena::new();
    let parent_id = ast_node_create(&mut arena, ScopeNodeKind::Module);
    assert_eq!(arena.len(), 2, "Module should allocate parent + child");
    let parent = arena.get(parent_id);
    assert_eq!(parent.kind, ScopeNodeKind::Module);
    assert!(parent.child.is_some(), "Module parent must reference its ModuleList child");
}

#[test]
fn ast_node_create_module_child_is_module_list_kind() {
    let mut arena = ScopeArena::new();
    let parent_id = ast_node_create(&mut arena, ScopeNodeKind::Module);
    let child_id = arena.get(parent_id).child.unwrap();
    assert_eq!(arena.get(child_id).kind, ScopeNodeKind::ModuleList);
}

#[test]
fn ast_node_create_module_list_child_has_flag_bit_set() {
    // For kind 0x16, flags bit 0 is set on the child.
    let mut arena = ScopeArena::new();
    let parent_id = ast_node_create(&mut arena, ScopeNodeKind::Module);
    let child_id = arena.get(parent_id).child.unwrap();
    assert_eq!(arena.get(child_id).flags_b7 & 1, 1, "ModuleList flags bit 0 must be set");
}

#[test]
fn ast_node_create_module_list_standalone_has_flag_bit_set() {
    // ModuleList created directly (not as Module child) also gets flags=1.
    let mut arena = ScopeArena::new();
    let id = ast_node_create(&mut arena, ScopeNodeKind::ModuleList);
    assert_eq!(arena.len(), 1);
    assert_eq!(arena.get(id).flags_b7 & 1, 1);
}

#[test]
fn ast_node_create_other_kind_allocates_one_node_no_child() {
    let mut arena = ScopeArena::new();
    let id = ast_node_create(&mut arena, ScopeNodeKind::K12);
    assert_eq!(arena.len(), 1);
    assert_eq!(arena.get(id).kind, ScopeNodeKind::K12);
    assert!(arena.get(id).child.is_none());
    assert_eq!(arena.get(id).flags_b7, 0);
}

// ── alloc_scope_block ─────────────────────────────────────────────────────────

#[test]
fn alloc_scope_block_head_is_none() {
    let mut arena = ScopeBlockArena::new();
    let id = alloc_scope_block(&mut arena);
    assert!(arena.get(id).head.is_none());
}

#[test]
fn alloc_scope_block_sentinel_is_u32_max() {
    // The hash sentinel (offset 12) is initialised to 0xffffffff.
    let mut arena = ScopeBlockArena::new();
    let id = alloc_scope_block(&mut arena);
    assert_eq!(arena.get(id).hash_sentinel, u32::MAX);
}

#[test]
fn alloc_scope_block_allocates_one_node() {
    let mut arena = ScopeBlockArena::new();
    alloc_scope_block(&mut arena);
    assert_eq!(arena.len(), 1);
}

// ── alloc_decl_node ───────────────────────────────────────────────────────────

#[test]
fn alloc_decl_node_fields_round_trip() {
    let mut decls = DeclArena::new();
    let mut scopes = ScopeBlockArena::new();
    let id = alloc_decl_node(&mut decls, &mut scopes, DeclSymKind::K07, None, 0x15, None);
    let n = decls.get(id);
    assert_eq!(n.kind, DeclSymKind::K07);
    assert_eq!(n.flags, 0x15 & 0x3f);
    assert!(n.parent.is_none());
    assert!(n.scope.is_none());
}

#[test]
fn alloc_decl_node_flags_masked_to_6_bits() {
    // flags are masked to 6 bits. 0xff & 0x3f = 0x3f.
    let mut decls = DeclArena::new();
    let mut scopes = ScopeBlockArena::new();
    let id = alloc_decl_node(&mut decls, &mut scopes, DeclSymKind::K07, None, 0xff, None);
    assert_eq!(decls.get(id).flags, 0x3f);
}

#[test]
fn alloc_decl_node_prepends_into_scope_block() {
    // The node becomes the new head; its scope chain points at the previous head.
    let mut decls = DeclArena::new();
    let mut scopes = ScopeBlockArena::new();
    let scope_id = alloc_scope_block(&mut scopes);

    let id1 = alloc_decl_node(&mut decls, &mut scopes, DeclSymKind::K07, Some(scope_id), 0, None);
    let id2 = alloc_decl_node(&mut decls, &mut scopes, DeclSymKind::K07, Some(scope_id), 0, None);

    // id2 is the new head; id2.scope_chain == id1 (prepend order).
    assert_eq!(scopes.get(scope_id).head, Some(id2));
    assert_eq!(decls.get(id2).scope_chain, Some(id1));
    assert!(decls.get(id1).scope_chain.is_none());
}

#[test]
fn alloc_decl_node_resets_scope_block_sentinel() {
    let mut decls = DeclArena::new();
    let mut scopes = ScopeBlockArena::new();
    let scope_id = alloc_scope_block(&mut scopes);
    scopes.get_mut(scope_id).hash_sentinel = 0; // clear it first
    alloc_decl_node(&mut decls, &mut scopes, DeclSymKind::K07, Some(scope_id), 0, None);
    assert_eq!(scopes.get(scope_id).hash_sentinel, u32::MAX);
}

#[test]
fn alloc_decl_node_appends_to_parent_children() {
    let mut decls = DeclArena::new();
    let mut scopes = ScopeBlockArena::new();
    let parent_id = alloc_decl_node(&mut decls, &mut scopes, DeclSymKind::K07, None, 0, None);
    let child_id = alloc_decl_node(&mut decls, &mut scopes, DeclSymKind::K07, None, 0, Some(parent_id));
    assert_eq!(decls.get(parent_id).children, vec![child_id]);
}

#[test]
fn alloc_decl_node_class_module_parent_adds_to_sec_children() {
    // parent kind==8 → secondary list also gets the child.
    let mut decls = DeclArena::new();
    let mut scopes = ScopeBlockArena::new();
    let parent_id = alloc_decl_node(&mut decls, &mut scopes, DeclSymKind::K08, None, 0, None);
    let child_id = alloc_decl_node(&mut decls, &mut scopes, DeclSymKind::K07, None, 0, Some(parent_id));
    assert_eq!(decls.get(parent_id).children, vec![child_id]);
    assert_eq!(decls.get(parent_id).sec_children, vec![child_id]);
}

#[test]
fn alloc_decl_node_std_module_parent_sets_module_level_flag() {
    // parent kind==9 → child sec_flags bit 2 set.
    let mut decls = DeclArena::new();
    let mut scopes = ScopeBlockArena::new();
    let parent_id = alloc_decl_node(&mut decls, &mut scopes, DeclSymKind::K09, None, 0, None);
    let child_id = alloc_decl_node(&mut decls, &mut scopes, DeclSymKind::K07, None, 0, Some(parent_id));
    assert_eq!(decls.get(child_id).sec_flags & 4, 4);
}

// ── ast_node_link ─────────────────────────────────────────────────────────────

fn make_link_arenas() -> (ScopeArena, DeclArena, ScopeBlockArena) {
    (ScopeArena::new(), DeclArena::new(), ScopeBlockArena::new())
}

#[test]
fn ast_node_link_module_scope_creates_k08_decl() {
    let (mut sa, mut da, mut sba) = make_link_arenas();
    let scope_id = ast_node_create(&mut sa, ScopeNodeKind::Module);
    let decl_id = ast_node_link(&mut sa, &mut da, &mut sba, scope_id, None, None, 0);
    assert_eq!(da.get(decl_id).kind, DeclSymKind::K08);
    assert_eq!(da.get(decl_id).flags, 3 & 0x3f);
}

#[test]
fn ast_node_link_non_module_scope_creates_k07_decl() {
    let (mut sa, mut da, mut sba) = make_link_arenas();
    let scope_id = ast_node_create(&mut sa, ScopeNodeKind::K12);
    let decl_id = ast_node_link(&mut sa, &mut da, &mut sba, scope_id, None, None, 0);
    assert_eq!(da.get(decl_id).kind, DeclSymKind::K07);
    assert_eq!(da.get(decl_id).flags, 1);
}

#[test]
fn ast_node_link_stores_decl_in_scope_node() {
    // The linked decl is stored on the scope node (byte offset 8).
    let (mut sa, mut da, mut sba) = make_link_arenas();
    let scope_id = ast_node_create(&mut sa, ScopeNodeKind::Module);
    let decl_id = ast_node_link(&mut sa, &mut da, &mut sba, scope_id, None, None, 0);
    assert_eq!(sa.get(scope_id).linked_decl, Some(decl_id));
}

#[test]
fn ast_node_link_sets_scope_parent_on_decl() {
    // The scope parent is stored on the decl (offset 0x20).
    let (mut sa, mut da, mut sba) = make_link_arenas();
    let scope_id = ast_node_create(&mut sa, ScopeNodeKind::K12);
    let decl_id = ast_node_link(&mut sa, &mut da, &mut sba, scope_id, None, None, 0);
    assert_eq!(da.get(decl_id).scope_parent, Some(scope_id));
}

#[test]
fn ast_node_link_kind1_sets_ext_flags_0x0d() {
    let (mut sa, mut da, mut sba) = make_link_arenas();
    let scope_id = ast_node_create(&mut sa, ScopeNodeKind::K12);
    let decl_id = ast_node_link(&mut sa, &mut da, &mut sba, scope_id, None, None, 1);
    assert_eq!(da.get(decl_id).ext_flags & 0x0d, 0x0d);
}

#[test]
fn ast_node_link_kind2_sets_ext_flags_0x01() {
    let (mut sa, mut da, mut sba) = make_link_arenas();
    let scope_id = ast_node_create(&mut sa, ScopeNodeKind::K12);
    let decl_id = ast_node_link(&mut sa, &mut da, &mut sba, scope_id, None, None, 2);
    assert_eq!(da.get(decl_id).ext_flags & 0x01, 0x01);
}

#[test]
fn ast_node_link_stores_link_param() {
    // The link param is stored on the decl (offset 0x6c).
    let (mut sa, mut da, mut sba) = make_link_arenas();
    let scope_id = ast_node_create(&mut sa, ScopeNodeKind::K12);
    let decl_id = ast_node_link(&mut sa, &mut da, &mut sba, scope_id, None, None, 42);
    assert_eq!(da.get(decl_id).link_param, 42);
}

// ── alloc_ast_node ───────────────────────────────────────────────────────────

#[test]
fn alloc_ast_node_stores_opcode() {
    let mut arena = ExprArena::new();
    let id = alloc_ast_node(&mut arena, 0x86, 0, 0, 0);
    match arena.get(id) {
        ExprNode::Generic { opcode, .. } => assert_eq!(*opcode, 0x86),
        _ => panic!("expected Generic"),
    }
}

#[test]
fn alloc_ast_node_stores_flags() {
    let mut arena = ExprArena::new();
    let id = alloc_ast_node(&mut arena, 0, 0x0f, 0, 0);
    match arena.get(id) {
        ExprNode::Generic { flags, .. } => assert_eq!(*flags, 0x0f),
        _ => panic!("expected Generic"),
    }
}

#[test]
fn alloc_ast_node_stores_lhs_and_rhs() {
    let mut arena = ExprArena::new();
    let id = alloc_ast_node(&mut arena, 0x7b, 0, 42, 99);
    match arena.get(id) {
        ExprNode::Generic { lhs, rhs, .. } => {
            assert_eq!(*lhs, 42);
            assert_eq!(*rhs, 99);
        }
        _ => panic!("expected Generic"),
    }
}

#[test]
fn alloc_ast_node_zero_children_for_leaf() {
    // Leaf node: both lhs and rhs are 0 (no children).
    let mut arena = ExprArena::new();
    let id = alloc_ast_node(&mut arena, 0x82, 0, 0, 0);
    match arena.get(id) {
        ExprNode::Generic { lhs, rhs, .. } => {
            assert_eq!(*lhs, 0);
            assert_eq!(*rhs, 0);
        }
        _ => panic!("expected Generic"),
    }
}

// ── build_for_node ───────────────────────────────────────────────────────────

#[test]
fn build_for_node_creates_for_range_variant() {
    let mut arena = ExprArena::new();
    let id = build_for_node(&mut arena, 1, 2, 3, 0);
    match arena.get(id) {
        ExprNode::ForRange { loop_var, step, .. } => {
            assert_eq!(*loop_var, 1);
            assert_eq!(*step, 0);
        }
        _ => panic!("expected ForRange"),
    }
}

#[test]
fn build_for_node_creates_range_sub_node() {
    // The range (start To end) is a nested 0x7b Generic node.
    let mut arena = ExprArena::new();
    let id = build_for_node(&mut arena, 0, 10, 20, 0);
    let range_id = match arena.get(id) {
        ExprNode::ForRange { range, .. } => vb6_core::frontend::ast::NodeId(*range),
        _ => panic!("expected ForRange"),
    };
    match arena.get(range_id) {
        ExprNode::Generic { opcode, lhs, rhs, .. } => {
            assert_eq!(*opcode, 0x7b);
            assert_eq!(*lhs, 10);
            assert_eq!(*rhs, 20);
        }
        _ => panic!("expected Generic range sub-node"),
    }
}

// ── build_on_node ────────────────────────────────────────────────────────────

#[test]
fn build_on_node_opcode_0x82() {
    let mut arena = ExprArena::new();
    let id = build_on_node(&mut arena, 1);
    match arena.get(id) {
        ExprNode::Generic { opcode, lhs, rhs, .. } => {
            assert_eq!(*opcode, 0x82);
            assert_eq!(*lhs, 1);
            assert_eq!(*rhs, 0);
        }
        _ => panic!("expected Generic"),
    }
}

// ── build_array_bounds_node ──────────────────────────────────────────────────

#[test]
fn build_array_bounds_node_opcode_0x7f() {
    let mut arena = ExprArena::new();
    let id = build_array_bounds_node(&mut arena, 0, 5, 0x42);
    match arena.get(id) {
        ExprNode::Generic { opcode, flags, lhs, rhs } => {
            assert_eq!(*opcode, 0x7f);
            assert_eq!(*flags, 0);
            assert_eq!(*lhs, 5);
            assert_eq!(*rhs, 0x42);
        }
        _ => panic!("expected Generic"),
    }
}

#[test]
fn build_array_bounds_node_stores_node_flags() {
    let mut arena = ExprArena::new();
    let id = build_array_bounds_node(&mut arena, 0x0c, 0, 0);
    match arena.get(id) {
        ExprNode::Generic { flags, .. } => assert_eq!(*flags, 0x0c),
        _ => panic!("expected Generic"),
    }
}

// ── emit_simple_node ──────────────────────────────────────────────────────────

#[test]
fn emit_simple_node_creates_leaf_with_opcode() {
    let mut arena = ExprArena::new();
    let id = emit_simple_node(&mut arena, 0xab);
    match arena.get(id) {
        ExprNode::Generic { opcode, flags, lhs, rhs } => {
            assert_eq!(*opcode, 0xab);
            assert_eq!(*flags, 0);
            assert_eq!(*lhs, 0);
            assert_eq!(*rhs, 0);
        }
        _ => panic!("expected Generic"),
    }
}

// ── build_abstract_member_node ────────────────────────────────────────────────

#[test]
fn build_abstract_member_node_qual_flag_0_sets_bit2() {
    // qual_flag == 0 → flags |= 0x04  (byte +4 |= 4)
    let mut arena = ExprArena::new();
    let id = build_abstract_member_node(&mut arena, 0x00, 0, 0, 0, &mut Diagnostics::new(), Span::DUMMY);
    match arena.get(id) {
        ExprNode::Generic { opcode, flags, .. } => {
            assert_eq!(*opcode, 0xbb);
            assert_eq!(*flags & 0x04, 0x04);
        }
        _ => panic!("expected Generic"),
    }
}

#[test]
fn build_abstract_member_node_qual_flag_4_sets_bit15() {
    // qual_flag == 4 → flags |= 0x8000  (byte +5 |= 0x80)
    let mut arena = ExprArena::new();
    let id = build_abstract_member_node(&mut arena, 0x00, 0, 0, 4, &mut Diagnostics::new(), Span::DUMMY);
    match arena.get(id) {
        ExprNode::Generic { opcode, flags, .. } => {
            assert_eq!(*opcode, 0xbb);
            assert_eq!(*flags & 0x8000, 0x8000);
        }
        _ => panic!("expected Generic"),
    }
}

#[test]
fn build_abstract_member_node_other_qual_flag_leaves_flags_unchanged() {
    let mut arena = ExprArena::new();
    let id = build_abstract_member_node(&mut arena, 0x0001, 0, 0, 99, &mut Diagnostics::new(), Span::DUMMY);
    match arena.get(id) {
        ExprNode::Generic { opcode, flags, .. } => {
            assert_eq!(*opcode, 0xbb);
            assert_eq!(*flags, 0x0001);
        }
        _ => panic!("expected Generic"),
    }
}

#[test]
fn build_abstract_member_node_passes_lhs_rhs_through() {
    let mut arena = ExprArena::new();
    let id = build_abstract_member_node(&mut arena, 0, 0xdead, 0xbeef, 0, &mut Diagnostics::new(), Span::DUMMY);
    match arena.get(id) {
        ExprNode::Generic { lhs, rhs, .. } => {
            assert_eq!(*lhs, 0xdead);
            assert_eq!(*rhs, 0xbeef);
        }
        _ => panic!("expected Generic"),
    }
}

// ── make_type_node ────────────────────────────────────────────────────────────

#[test]
fn make_type_node_not_qualified_passes_flags_unchanged() {
    let mut arena = ExprArena::new();
    let id = make_type_node(&mut arena, 0xb3, 0x0001, 0x1000, 7, false);
    match arena.get(id) {
        ExprNode::Generic { opcode, flags, lhs, rhs } => {
            assert_eq!(*opcode, 0xb3);
            assert_eq!(*flags, 0x0001);
            assert_eq!(*lhs, 0x1000);
            assert_eq!(*rhs, 7);
        }
        _ => panic!("expected Generic"),
    }
}

#[test]
fn make_type_node_qualified_sets_bit15() {
    let mut arena = ExprArena::new();
    let id = make_type_node(&mut arena, 0xb3, 0x0001, 0, 0, true);
    match arena.get(id) {
        ExprNode::Generic { flags, .. } => {
            assert_eq!(*flags & 0x8000, 0x8000);
            assert_eq!(*flags & 0x0001, 0x0001); // original bit preserved
        }
        _ => panic!("expected Generic"),
    }
}

#[test]
fn make_type_node_zero_flags_qualified_gives_0x8000() {
    let mut arena = ExprArena::new();
    let id = make_type_node(&mut arena, 0x99, 0, 0, 0, true);
    match arena.get(id) {
        ExprNode::Generic { opcode, flags, .. } => {
            assert_eq!(*opcode, 0x99);
            assert_eq!(*flags, 0x8000);
        }
        _ => panic!("expected Generic"),
    }
}

#[test]
fn make_type_node_variant_marker_normalizes_type_aux_to_5() {
    let mut arena = ExprArena::new();
    // Pass the Variant type descriptor marker ('?' = 0x3f) with type_aux=0.
    // The function should normalize type_aux to 5 (canonical Variant).
    let id = make_type_node(&mut arena, 0xb3, 0, VARIANT_TYPE_DESC_MARKER, 0, false);
    match arena.get(id) {
        ExprNode::Generic { lhs, rhs, .. } => {
            assert_eq!(*lhs, VARIANT_TYPE_DESC_MARKER);
            assert_eq!(*rhs, 5, "type_aux must be normalized to 5 for Variant");
        }
        _ => panic!("expected Generic"),
    }
}

#[test]
fn make_type_node_non_variant_preserves_type_aux() {
    let mut arena = ExprArena::new();
    let id = make_type_node(&mut arena, 0xb3, 0, 0x1234, 3, false);
    match arena.get(id) {
        ExprNode::Generic { rhs, .. } => {
            assert_eq!(*rhs, 3, "type_aux must be unchanged for non-Variant");
        }
        _ => panic!("expected Generic"),
    }
}
