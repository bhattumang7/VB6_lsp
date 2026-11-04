//! `^` groups left to right like every other binary operator: `2 ^ 3 ^ 2`
//! is `(2 ^ 3) ^ 2`, and unary minus binds looser than `^`.

mod common;
use common::parse;
use vb6_syntax::frontend::ast::{AstLit, BinOpKind, ExprNode, NodeId, UnOpKind};

fn rhs_of_first_assignment(src: &str) -> (common::Parsed, NodeId) {
    let p = parse(&format!("Sub S()\n x = {src}\nEnd Sub\n"));
    assert_eq!(p.diags, 0);
    let body = p.proc_body(0);
    assert_eq!(body.len(), 1);
    let ExprNode::Assign { value, .. } = p.node(body[0]) else { panic!("not an assignment") };
    let value = *value;
    (p, value)
}

fn binop(p: &common::Parsed, id: NodeId, want: BinOpKind) -> (NodeId, NodeId) {
    match p.node(id) {
        ExprNode::BinOp { op, lhs, rhs } if *op == want => (*lhs, *rhs),
        _ => panic!("expected {want:?}"),
    }
}

fn int(p: &common::Parsed, id: NodeId) -> i32 {
    match p.node(id) {
        ExprNode::Literal { lit: AstLit::Int(n) } => *n,
        _ => panic!("expected an integer literal"),
    }
}

#[test]
fn chained_exponent_is_left_associative() {
    let (p, v) = rhs_of_first_assignment("2 ^ 3 ^ 2");
    let (lhs, rhs) = binop(&p, v, BinOpKind::Pow);
    assert_eq!(int(&p, rhs), 2);
    let (a, b) = binop(&p, lhs, BinOpKind::Pow);
    assert_eq!((int(&p, a), int(&p, b)), (2, 3));
}

#[test]
fn negation_applies_to_the_whole_chain() {
    let (p, v) = rhs_of_first_assignment("-2 ^ 3 ^ 2");
    let ExprNode::UnOp { op: UnOpKind::Neg, operand } = p.node(v) else { panic!("expected Neg") };
    let (lhs, rhs) = binop(&p, *operand, BinOpKind::Pow);
    assert_eq!(int(&p, rhs), 2);
    let (a, b) = binop(&p, lhs, BinOpKind::Pow);
    assert_eq!((int(&p, a), int(&p, b)), (2, 3));
}

#[test]
fn exponent_still_binds_tighter_than_multiplication() {
    let (p, v) = rhs_of_first_assignment("2 * 3 ^ 2 ^ 2");
    let (m, rhs) = binop(&p, v, BinOpKind::Mul);
    assert_eq!(int(&p, m), 2);
    let (lhs, last) = binop(&p, rhs, BinOpKind::Pow);
    assert_eq!(int(&p, last), 2);
    let (a, b) = binop(&p, lhs, BinOpKind::Pow);
    assert_eq!((int(&p, a), int(&p, b)), (3, 2));
}

#[test]
fn reversed_and_spaced_comparison_operators_parse() {
    for (text, want) in [
        ("a =< b", BinOpKind::Le),
        ("a => b", BinOpKind::Ge),
        ("a >< b", BinOpKind::Ne),
        ("a < > b", BinOpKind::Ne),
        ("a > = b", BinOpKind::Ge),
    ] {
        let (p, v) = rhs_of_first_assignment(text);
        let (lhs, rhs) = binop(&p, v, want);
        assert!(matches!(p.node(lhs), ExprNode::NameRef { .. }), "{text}");
        assert!(matches!(p.node(rhs), ExprNode::NameRef { .. }), "{text}");
    }
}
