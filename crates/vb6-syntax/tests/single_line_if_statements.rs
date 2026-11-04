//! A single-line `If` owns every statement on its line: all the statements of
//! the `Then` part, and all of those after `Else`.

mod common;
use common::{offset_of, parse};
use vb6_syntax::frontend::ast::ExprNode;

fn assigned_name<'a>(p: &'a common::Parsed, id: vb6_syntax::frontend::ast::NodeId) -> &'a str {
    let ExprNode::Assign { target, .. } = p.node(id) else { panic!("expected assignment") };
    let ExprNode::NameRef { sym, .. } = p.node(*target) else { panic!("expected name") };
    p.name(*sym)
}

#[test]
fn then_and_else_parts_keep_all_their_statements() {
    let p = parse("Sub S()\nIf cond Then one = 1: two = 2 Else three = 3: four = 4\nfive = 5\nEnd Sub\n");
    assert_eq!(p.found, vec![]);
    let body = p.proc_body(0);
    assert_eq!(body.len(), 2);
    let ExprNode::If { then_body, else_body, .. } = p.node(body[0]) else { panic!("expected If") };
    let then: Vec<_> = p.block(*then_body).iter().map(|n| assigned_name(&p, *n).to_string()).collect();
    assert_eq!(then, ["one", "two"]);
    let els: Vec<_> = p.block(else_body.unwrap()).iter().map(|n| assigned_name(&p, *n).to_string()).collect();
    assert_eq!(els, ["three", "four"]);
    assert_eq!(assigned_name(&p, body[1]), "five");
}

#[test]
fn statements_after_the_then_part_stay_inside_without_else() {
    let p = parse("Sub S()\nIf cond Then one = 1: two = 2\nthree = 3\nEnd Sub\n");
    assert_eq!(p.found, vec![]);
    let body = p.proc_body(0);
    assert_eq!(body.len(), 2);
    let ExprNode::If { then_body, else_body, .. } = p.node(body[0]) else { panic!("expected If") };
    assert_eq!(p.block(*then_body).len(), 2);
    assert!(else_body.is_none());
    assert_eq!(assigned_name(&p, body[1]), "three");
}

#[test]
fn empty_then_part_and_empty_statements() {
    let p = parse("Sub S()\nIf a Then Else b = 1\nIf a Then: c = 2\nEnd Sub\n");
    assert_eq!(p.found, vec![]);
    let body = p.proc_body(0);
    assert_eq!(body.len(), 2);
    let ExprNode::If { then_body, else_body, .. } = p.node(body[0]) else { panic!("expected If") };
    assert!(p.block(*then_body).is_empty());
    assert_eq!(p.block(else_body.unwrap()).len(), 1);
    let ExprNode::If { then_body, else_body, .. } = p.node(body[1]) else { panic!("expected If") };
    assert_eq!(p.block(*then_body).len(), 1);
    assert!(else_body.is_none());
}

#[test]
fn an_else_on_the_next_line_does_not_continue_the_if() {
    let src = "Sub S()\nIf a Then b = 1\nElse\nEnd Sub\n";
    let p = parse(src);
    assert_eq!(p.found, vec![(0x9c96, offset_of(src, "Else", 0), 4)]);
    let body = p.proc_body(0);
    assert_eq!(body.len(), 1);
    let ExprNode::If { else_body, .. } = p.node(body[0]) else { panic!("expected If") };
    assert!(else_body.is_none());
}

#[test]
fn nested_single_line_if_takes_the_else() {
    let p = parse("Sub S()\nIf a Then If b Then c = 1 Else d = 2\nEnd Sub\n");
    assert_eq!(p.found, vec![]);
    let body = p.proc_body(0);
    assert_eq!(body.len(), 1);
    let ExprNode::If { then_body, else_body, .. } = p.node(body[0]) else { panic!("expected If") };
    assert!(else_body.is_none());
    let inner = p.block(*then_body);
    assert_eq!(inner.len(), 1);
    let ExprNode::If { then_body, else_body, .. } = p.node(inner[0]) else { panic!("expected inner If") };
    assert_eq!(p.block(*then_body).len(), 1);
    assert_eq!(p.block(else_body.unwrap()).len(), 1);
}
