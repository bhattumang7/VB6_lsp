//! Unary operators and the trees they build: `Not` takes everything that binds
//! tighter than `And`, unary minus takes only `^`, and a unary plus builds no
//! node and takes everything that binds tighter than `+` and `-`.

mod common;
use common::{parse, Parsed};
use vb6_syntax::frontend::ast::{BinOpKind, ExprNode, NodeId, UnOpKind};

fn rhs(src: &str) -> (Parsed, NodeId) {
    let p = parse(&format!("Sub S()\n    x = {src}\nEnd Sub\n"));
    assert_eq!(p.found, vec![], "{src}");
    let body = p.proc_body(0);
    let ExprNode::Assign { value, .. } = p.node(body[0]) else { panic!("expected an assignment") };
    let value = *value;
    (p, value)
}

fn name(p: &Parsed, id: NodeId) -> String {
    let ExprNode::NameRef { sym, .. } = p.node(id) else { panic!("expected a name, got another node") };
    p.name(*sym).to_string()
}

fn binop(p: &Parsed, id: NodeId, want: BinOpKind) -> (NodeId, NodeId) {
    let ExprNode::BinOp { op, lhs, rhs } = p.node(id) else { panic!("expected a binary operation") };
    assert_eq!(*op, want);
    (*lhs, *rhs)
}

#[test]
fn unary_plus_builds_no_node() {
    let (p, v) = rhs("+a");
    assert_eq!(name(&p, v), "a");
}

#[test]
fn unary_plus_operand_takes_multiplicative_operators() {
    let (p, v) = rhs("+a * bb");
    let (l, r) = binop(&p, v, BinOpKind::Mul);
    assert_eq!((name(&p, l), name(&p, r)), ("a".to_string(), "bb".to_string()));
}

#[test]
fn unary_plus_operand_takes_mod_integer_division_and_power() {
    for (text, op) in [("+a Mod bb", BinOpKind::Mod), ("+a \\ bb", BinOpKind::IDiv), ("+a ^ bb", BinOpKind::Pow), ("+a / bb", BinOpKind::Div)] {
        let (p, v) = rhs(text);
        let (l, r) = binop(&p, v, op);
        assert_eq!((name(&p, l), name(&p, r)), ("a".to_string(), "bb".to_string()), "{text}");
    }
}

#[test]
fn unary_plus_operand_stops_at_addition() {
    let (p, v) = rhs("+a + bb");
    let (l, r) = binop(&p, v, BinOpKind::Add);
    assert_eq!((name(&p, l), name(&p, r)), ("a".to_string(), "bb".to_string()));
}

#[test]
fn unary_plus_before_minus_keeps_the_minus() {
    let (p, v) = rhs("+-a");
    let ExprNode::UnOp { op: UnOpKind::Neg, operand } = p.node(v) else { panic!("expected Neg") };
    assert_eq!(name(&p, *operand), "a");
}

#[test]
fn minus_takes_only_power() {
    let (p, v) = rhs("-a * bb");
    let (l, r) = binop(&p, v, BinOpKind::Mul);
    let ExprNode::UnOp { op: UnOpKind::Neg, operand } = p.node(l) else { panic!("expected Neg on the left") };
    assert_eq!((name(&p, *operand), name(&p, r)), ("a".to_string(), "bb".to_string()));
    let (p, v) = rhs("-a ^ bb");
    let ExprNode::UnOp { op: UnOpKind::Neg, operand } = p.node(v) else { panic!("expected Neg") };
    binop(&p, *operand, BinOpKind::Pow);
}

#[test]
fn not_takes_comparisons_but_not_and() {
    let (p, v) = rhs("Not a = bb");
    let ExprNode::UnOp { op: UnOpKind::Not, operand } = p.node(v) else { panic!("expected Not") };
    binop(&p, *operand, BinOpKind::Eq);
    let (p, v) = rhs("Not a And bb");
    let (l, _) = binop(&p, v, BinOpKind::And);
    assert!(matches!(p.node(l), ExprNode::UnOp { op: UnOpKind::Not, .. }));
}
