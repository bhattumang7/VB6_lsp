//! `Not` between two operands is an infix operator at the level of `And`; a
//! call-site `ByVal` needs an operand; after a named argument every argument is
//! named; a named argument is not a statement.

mod common;
use common::{offset_of, parse, Parsed};
use vb6_syntax::frontend::ast::{BinOpKind, ExprNode, NodeId, UnOpKind};

fn rhs(src: &str) -> (Parsed, NodeId) {
    let p = parse(&format!("Sub S()\n    x = {src}\nEnd Sub\n"));
    assert_eq!(p.found, vec![], "{src}");
    let ExprNode::Assign { value, .. } = p.node(p.proc_body(0)[0]) else { panic!("assignment") };
    let value = *value;
    (p, value)
}

fn name(p: &Parsed, id: NodeId) -> String {
    let ExprNode::NameRef { sym, .. } = p.node(id) else { panic!("name") };
    p.name(*sym).to_string()
}

#[test]
fn not_between_two_operands_is_an_infix_node() {
    let (p, v) = rhs("aa Not bb");
    let ExprNode::BinOp { op: BinOpKind::Not, lhs, rhs } = p.node(v) else { panic!("expected infix Not") };
    assert_eq!((name(&p, *lhs), name(&p, *rhs)), ("aa".to_string(), "bb".to_string()));
}

#[test]
fn infix_not_binds_like_and_and_is_left_associative() {
    let (p, v) = rhs("aa Not bb And cc");
    let ExprNode::BinOp { op: BinOpKind::And, lhs, .. } = p.node(v) else { panic!("And at the top") };
    assert!(matches!(p.node(*lhs), ExprNode::BinOp { op: BinOpKind::Not, .. }));
    let (p, v) = rhs("aa Not bb + cc");
    let ExprNode::BinOp { op: BinOpKind::Not, rhs, .. } = p.node(v) else { panic!("Not at the top") };
    assert!(matches!(p.node(*rhs), ExprNode::BinOp { op: BinOpKind::Add, .. }));
}

#[test]
fn prefix_not_is_unchanged() {
    let (p, v) = rhs("Not aa");
    assert!(matches!(p.node(v), ExprNode::UnOp { op: UnOpKind::Not, .. }));
    let (p, v) = rhs("aa And Not bb");
    let ExprNode::BinOp { op: BinOpKind::And, rhs, .. } = p.node(v) else { panic!("And") };
    assert!(matches!(p.node(*rhs), ExprNode::UnOp { op: UnOpKind::Not, .. }));
}

#[test]
fn byval_without_an_operand_is_expected_expression() {
    for text in ["f a, ByVal", "x = f(a, ByVal)", "f ByVal"] {
        let src = format!("Sub S()\n    {text}\nEnd Sub\n");
        let p = parse(&src);
        assert!(!p.found.is_empty(), "{text}");
        assert!(p.found.iter().any(|d| d.1 as usize >= src.find("ByVal").unwrap()), "{text}: {:?}", p.found);
    }
    assert_eq!(parse("Sub S()\n    f a, ByVal b\nEnd Sub\n").found, vec![]);
}

#[test]
fn a_positional_argument_after_a_named_one_is_an_error_at_that_argument() {
    for (text, at) in [("x = f(aa:=1, bb)", "bb"), ("f aa:=1, bb", "bb")] {
        let src = format!("Sub S()\n    {text}\nEnd Sub\n");
        let p = parse(&src);
        assert!(p.found.iter().any(|d| d.1 == offset_of(&src, at, 0)), "{text}: {:?}", p.found);
    }
    for ok in ["x = f(aa:=1, bb:=2)", "x = f(1, 2, bb:=3)", "f 1, bb:=2"] {
        assert_eq!(parse(&format!("Sub S()\n    {ok}\nEnd Sub\n")).found, vec![], "{ok}");
    }
}

#[test]
fn a_named_argument_is_not_a_statement() {
    let src = "Sub S()\n    f := a\nEnd Sub\n";
    let p = parse(src);
    assert_eq!(p.found[0], (0x9c70, offset_of(src, ":=", 0), 2));
}
