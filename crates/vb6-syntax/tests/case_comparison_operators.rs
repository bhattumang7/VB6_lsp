//! A comparison operator that starts a `Case` item stands for `Is` before it.

mod common;
use common::{parse, Parsed};
use vb6_syntax::frontend::ast::{AstLit, BinOpKind, ExprNode, NodeId};

fn first_case_items(p: &Parsed) -> Vec<NodeId> {
    let body = p.proc_body(0);
    let ExprNode::SelectCase { cases, .. } = p.node(body[0]) else { panic!("expected Select Case") };
    let ExprNode::CaseBlock { test, .. } = p.node(cases[0]) else { panic!("expected Case block") };
    let ExprNode::ArgList { args } = p.node(*test) else { panic!("expected items") };
    args.clone()
}

fn case_is(p: &Parsed, id: NodeId) -> (BinOpKind, i32) {
    let ExprNode::CaseIs { op, rhs } = p.node(id) else { panic!("expected CaseIs") };
    let ExprNode::Literal { lit: AstLit::Int(n) } = p.node(*rhs) else { panic!("expected an integer") };
    (*op, *n)
}

fn parse_case(items: &str) -> Parsed {
    parse(&format!("Sub S()\nSelect Case x\nCase {items}\ny = 1\nEnd Select\nEnd Sub\n"))
}

#[test]
fn every_comparison_operator_without_is() {
    for (text, op) in [
        ("=", BinOpKind::Eq), ("<>", BinOpKind::Ne), ("<", BinOpKind::Lt),
        (">", BinOpKind::Gt), ("<=", BinOpKind::Le), (">=", BinOpKind::Ge),
    ] {
        let p = parse_case(&format!("{text} 5"));
        assert_eq!(p.found, vec![], "{text}");
        let items = first_case_items(&p);
        assert_eq!(items.len(), 1);
        assert_eq!(case_is(&p, items[0]), (op, 5), "{text}");
    }
}

#[test]
fn explicit_is_gives_the_same_node() {
    let a = parse_case("Is >= 5");
    let b = parse_case(">= 5");
    assert_eq!(case_is(&a, first_case_items(&a)[0]), case_is(&b, first_case_items(&b)[0]));
}

#[test]
fn mixed_items_keep_their_order_and_shapes() {
    let p = parse_case("1, 2 To 3, > 4");
    assert_eq!(p.found, vec![]);
    let items = first_case_items(&p);
    assert_eq!(items.len(), 3);
    assert!(matches!(p.node(items[0]), ExprNode::Literal { lit: AstLit::Int(1) }));
    assert!(matches!(p.node(items[1]), ExprNode::RangeTo { .. }));
    assert_eq!(case_is(&p, items[2]), (BinOpKind::Gt, 4));
}
