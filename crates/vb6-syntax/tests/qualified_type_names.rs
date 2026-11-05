//! A type name may carry any number of qualifiers; each further name nests
//! as the child of the one before it.

mod common;
use common::{parse, Parsed};
use vb6_syntax::frontend::ast::{ExprNode, NodeId};

fn declared_type(src: &str) -> (Parsed, NodeId) {
    let p = parse(src);
    assert_eq!(p.found, vec![], "{src}");
    let ExprNode::Block { stmts } = p.node(p.top[0]) else { panic!("block") };
    let ExprNode::DimItem { type_node: Some(t), .. } = p.node(stmts[0]) else { panic!("dim item") };
    let t = *t;
    (p, t)
}

#[test]
fn two_part_name_keeps_its_single_child() {
    let (p, t) = declared_type("Dim a As Foo.Bar\n");
    let ExprNode::UserType { name, child: Some(c) } = p.node(t) else { panic!("user type") };
    assert_eq!(p.name(*name), "Foo");
    let ExprNode::NameRef { sym, .. } = p.node(*c) else { panic!("name") };
    assert_eq!(p.name(*sym), "Bar");
}

#[test]
fn a_third_part_nests_under_the_second() {
    let (p, t) = declared_type("Dim a As Foo.Bar.Baz\n");
    let ExprNode::UserType { name, child: Some(c) } = p.node(t) else { panic!("user type") };
    assert_eq!(p.name(*name), "Foo");
    let ExprNode::UserType { name: second, child: Some(c2) } = p.node(*c) else { panic!("nested type") };
    assert_eq!(p.name(*second), "Bar");
    let ExprNode::NameRef { sym, .. } = p.node(*c2) else { panic!("name") };
    assert_eq!(p.name(*sym), "Baz");
}

#[test]
fn a_fourth_part_nests_once_more() {
    let (p, t) = declared_type("Dim a As A.B.C.D\n");
    let ExprNode::UserType { child: Some(c), .. } = p.node(t) else { panic!("user type") };
    let ExprNode::UserType { child: Some(c2), .. } = p.node(*c) else { panic!("second") };
    let ExprNode::UserType { child: Some(c3), .. } = p.node(*c2) else { panic!("third") };
    assert!(matches!(p.node(*c3), ExprNode::NameRef { .. }));
}
