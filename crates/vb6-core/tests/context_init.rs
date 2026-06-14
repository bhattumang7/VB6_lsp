//! Structural tests for CompilerContext initialisation helpers.
//!
//! Functions under test:
//!   build_predefined_nodes

use vb6_core::context::CompilerContext;
use vb6_core::frontend::ast::{alloc_decl_node, DeclSymKind, ScopeNodeKind};

fn make_parent_decl(ctx: &mut CompilerContext) -> vb6_core::frontend::ast::NodeId {
    // A K09 (standard module) DeclNode — the typical parent_decl.
    alloc_decl_node(&mut ctx.decls, &mut ctx.scopes, DeclSymKind::K09, None, 0, None)
}

#[test]
fn build_predefined_nodes_creates_two_scope_nodes() {
    let mut ctx = CompilerContext::new();
    let parent = make_parent_decl(&mut ctx);
    ctx.build_predefined_nodes(None, parent);
    // Module kind creates 2 nodes each (parent + ModuleList child), so 2*2 = 4.
    assert_eq!(ctx.scope_nodes.len(), 4);
}

#[test]
fn build_predefined_nodes_sets_first_predefined_scope() {
    let mut ctx = CompilerContext::new();
    let parent = make_parent_decl(&mut ctx);
    ctx.build_predefined_nodes(None, parent);
    assert!(ctx.first_predefined_scope.is_some());
    let scope = ctx.scope_nodes.get(ctx.first_predefined_scope.unwrap());
    assert_eq!(scope.kind, ScopeNodeKind::Module);
}

#[test]
fn build_predefined_nodes_first_scope_flags_b6_has_bits_4_and_80() {
    // scope1 flags_b6: bit 4 set before link, bit 0x80 set after.
    let mut ctx = CompilerContext::new();
    let parent = make_parent_decl(&mut ctx);
    ctx.build_predefined_nodes(None, parent);
    let scope = ctx.scope_nodes.get(ctx.first_predefined_scope.unwrap());
    assert_eq!(scope.flags_b6 & 0x84, 0x84);
}

#[test]
fn build_predefined_nodes_second_scope_flags_b6_has_bit_8() {
    // scope2 flags_b6: bit 8 set.
    let mut ctx = CompilerContext::new();
    let parent = make_parent_decl(&mut ctx);
    ctx.build_predefined_nodes(None, parent);
    // scope2 is the second module node. ast_node_create(Module) allocates the
    // child first (even index) and the parent second.
    // scope1: child=NodeId(0), parent=NodeId(1). scope2: child=NodeId(2), parent=NodeId(3).
    // first_predefined_scope = NodeId(1), so scope2 = NodeId(3).
    use vb6_core::frontend::ast::NodeId;
    let scope2_id = NodeId(3);
    let scope2 = ctx.scope_nodes.get(scope2_id);
    assert_eq!(scope2.flags_b6 & 8, 8);
}

#[test]
fn build_predefined_nodes_decl1_has_sec_flags_8_and_type_info_5() {
    let mut ctx = CompilerContext::new();
    let parent = make_parent_decl(&mut ctx);
    ctx.build_predefined_nodes(None, parent);
    // decl1 = first DeclNode created (index 1 — K09 parent is index 0).
    use vb6_core::frontend::ast::NodeId;
    let decl1 = NodeId(1);
    let d = ctx.decls.get(decl1);
    assert_eq!(d.sec_flags & 8, 8);
    assert_eq!(d.type_info, 5);
}

#[test]
fn build_predefined_nodes_sets_first_predefined_child() {
    let mut ctx = CompilerContext::new();
    let parent = make_parent_decl(&mut ctx);
    ctx.build_predefined_nodes(None, parent);
    assert!(ctx.first_predefined_child.is_some());
    // first_predefined_child should be the ModuleList child of scope1.
    let scope1 = ctx.scope_nodes.get(ctx.first_predefined_scope.unwrap());
    assert_eq!(ctx.first_predefined_child, scope1.child);
}

#[test]
fn build_predefined_nodes_sets_second_predefined_child() {
    let mut ctx = CompilerContext::new();
    let parent = make_parent_decl(&mut ctx);
    ctx.build_predefined_nodes(None, parent);
    assert!(ctx.second_predefined_child.is_some());
}

#[test]
fn build_predefined_nodes_interns_second_scope_name() {
    // VB6 names the second predefined scope by interning keyword_string(0xba)
    // = "Unknown". After building, that symbol must exist and head the chain
    // containing the second scope's decl (decl2 = NodeId(2) in the DeclArena).
    use vb6_core::frontend::ast::NodeId;
    let mut ctx = CompilerContext::new();
    let parent = make_parent_decl(&mut ctx);
    ctx.build_predefined_nodes(None, parent);

    let sym = ctx.interned_symbol("Unknown").expect("`Unknown` must be interned");
    assert_eq!(ctx.scopes.get(sym).head, Some(NodeId(2)),
        "the interned symbol must head decl2's declaration chain");
    // Case-insensitive lookup.
    assert_eq!(ctx.interned_symbol("UNKNOWN"), Some(sym));
    // The predefined-scope symbol is interned with kind 4, original-case name.
    assert_eq!(ctx.interned_kind("Unknown"), Some(4));
    assert_eq!(ctx.interned_name(sym), Some("Unknown"));
}

#[test]
fn intern_string_is_deduped_and_case_insensitive() {
    let mut ctx = CompilerContext::new();
    let a = ctx.intern_string("Foo", 4);
    let b = ctx.intern_string("foo", 4);
    let c = ctx.intern_string("Bar", 2);
    assert_eq!(a, b, "same name (any case) must return the same symbol");
    assert_ne!(a, c, "different names must return different symbols");
}

#[test]
fn build_predefined_nodes_parent_ext_flag_propagates_to_decls() {
    // When parent.ext_flags bit 0 is set, both DeclNodes get flags_10 |= 8.
    let mut ctx = CompilerContext::new();
    let parent = make_parent_decl(&mut ctx);
    ctx.decls.get_mut(parent).ext_flags = 1;
    ctx.build_predefined_nodes(None, parent);
    use vb6_core::frontend::ast::NodeId;
    let decl1 = ctx.decls.get(NodeId(1));
    let decl2 = ctx.decls.get(NodeId(2));
    assert_eq!(decl1.flags_10 & 8, 8);
    assert_eq!(decl2.flags_10 & 8, 8);
}
