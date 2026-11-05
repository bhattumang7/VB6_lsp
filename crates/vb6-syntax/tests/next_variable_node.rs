//! The variable written after `Next` is kept on the loop it closes.

mod common;
use common::parse;
use vb6_syntax::frontend::ast::{ExprNode, NodeId};

fn name_of(p: &common::Parsed, id: NodeId) -> String {
    let ExprNode::NameRef { sym, .. } = p.node(id) else { panic!("not a name") };
    p.name(*sym).to_string()
}

#[test]
fn a_single_next_variable_belongs_to_its_loop() {
    let p = parse("Sub S()\nFor i = 1 To 3\nNext i\nEnd Sub\n");
    assert_eq!(p.diags, 0);
    let b = p.proc_body(0);
    let ExprNode::For { var, next_var: Some(nv), .. } = p.node(b[0]) else { panic!("not For") };
    assert_eq!(name_of(&p, *var), "i");
    assert_eq!(name_of(&p, *nv), "i");
    assert_ne!(var, nv);
}

#[test]
fn next_without_a_variable_leaves_none() {
    let p = parse("Sub S()\nFor i = 1 To 3\nNext\nFor Each v In c\nNext\nEnd Sub\n");
    assert_eq!(p.diags, 0);
    let b = p.proc_body(0);
    assert!(matches!(p.node(b[0]), ExprNode::For { next_var: None, .. }));
    assert!(matches!(p.node(b[1]), ExprNode::ForEach { next_var: None, .. }));
}

#[test]
fn for_each_keeps_its_next_variable() {
    let p = parse("Sub S()\nFor Each v In c\nNext v\nEnd Sub\n");
    assert_eq!(p.diags, 0);
    let b = p.proc_body(0);
    let ExprNode::ForEach { next_var: Some(nv), .. } = p.node(b[0]) else { panic!("not ForEach") };
    assert_eq!(name_of(&p, *nv), "v");
}

#[test]
fn a_shared_next_gives_each_loop_its_own_variable() {
    let p = parse("Sub S()\nFor j = 1 To 2\nFor i = 1 To 3\nNext i, j\nEnd Sub\n");
    assert_eq!(p.diags, 0);
    let b = p.proc_body(0);
    let ExprNode::For { body, next_var: Some(outer), .. } = p.node(b[0]) else { panic!("not For") };
    assert_eq!(name_of(&p, *outer), "j");
    let ExprNode::Block { stmts } = p.node(*body) else { panic!("not a block") };
    let ExprNode::For { next_var: Some(inner), .. } = p.node(stmts[0]) else { panic!("not inner For") };
    assert_eq!(name_of(&p, *inner), "i");
}
