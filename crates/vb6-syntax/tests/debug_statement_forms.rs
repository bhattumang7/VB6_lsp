//! `Debug.Print [#channel,] items` and `Debug.Assert condition`.

mod common;
use common::{offset_of, parse, Parsed};
use vb6_syntax::frontend::ast::{AstLit, BinOpKind, ExprNode, NodeId};

fn only_statement(p: &Parsed) -> NodeId {
    let b = p.proc_body(0);
    assert_eq!(b.len(), 1);
    b[0]
}

fn is_name(p: &Parsed, id: NodeId, want: &str) -> bool {
    matches!(p.node(id), ExprNode::NameRef { sym, .. } if p.name(*sym) == want)
}

#[test]
fn print_items_without_a_channel() {
    let p = parse("Sub S()\n Debug.Print \"a\", item\nEnd Sub\n");
    assert_eq!(p.diags, 0);
    let ExprNode::DebugPrint { channel: None, args } = p.node(only_statement(&p)) else { panic!("not Debug.Print") };
    assert_eq!(args.len(), 2);
    assert!(matches!(p.node(args[0]), ExprNode::Literal { lit: AstLit::Str(s) } if &**s == "a"));
    assert!(is_name(&p, args[1], "item"));
}

#[test]
fn print_with_a_channel() {
    let p = parse("Sub S()\n Debug.Print #1, first; second\nEnd Sub\n");
    assert_eq!(p.diags, 0);
    let ExprNode::DebugPrint { channel: Some(ch), args } = p.node(only_statement(&p)) else { panic!("not Debug.Print #") };
    assert!(matches!(p.node(*ch), ExprNode::Literal { lit: AstLit::Int(1) }));
    assert_eq!(args.len(), 2);
    assert!(is_name(&p, args[0], "first"));
    assert!(is_name(&p, args[1], "second"));
}

#[test]
fn print_with_a_channel_and_no_items() {
    let p = parse("Sub S()\n Debug.Print #1,\nEnd Sub\n");
    assert_eq!(p.diags, 0);
    let ExprNode::DebugPrint { channel: Some(_), args } = p.node(only_statement(&p)) else { panic!("not Debug.Print #") };
    assert!(args.is_empty());
}

#[test]
fn a_channel_needs_its_comma_and_comes_first() {
    // (text after `Debug.Print `, offset within it of the first rejected token)
    let cases: [(&str, u32); 5] = [("#1;", 2), ("#1 item", 3), ("#", 1), ("item, #1", 6), ("#1, #2, item", 4)];
    for (text, rel) in cases {
        let src = format!("Sub S()\n Debug.Print {text}\nEnd Sub\n");
        let p = parse(&src);
        assert_eq!(p.found[0].1, offset_of(&src, "Print ", 0) + 6 + rel, "{text}: {:?}", p.found);
    }
}

#[test]
fn assert_takes_one_expression() {
    let p = parse("Sub S()\n Debug.Assert item > 1\nEnd Sub\n");
    assert_eq!(p.diags, 0);
    let ExprNode::DebugAssert { cond } = p.node(only_statement(&p)) else { panic!("not Debug.Assert") };
    let ExprNode::BinOp { op: BinOpKind::Gt, lhs, rhs } = p.node(*cond) else { panic!("not a comparison") };
    assert!(is_name(&p, *lhs, "item"));
    assert!(matches!(p.node(*rhs), ExprNode::Literal { lit: AstLit::Int(1) }));
}

#[test]
fn assert_rejects_anything_but_one_expression() {
    // (text after `Debug.Assert`, offset within it of the first rejected token)
    let cases: [(&str, u32); 7] = [("", 0), (" a, b", 2), (" a;", 2), (" a,", 2), (" a b", 3), (" Spc(1)", 1), (" Tab", 1)];
    for (text, rel) in cases {
        let src = format!("Sub S()\n Debug.Assert{text}\nEnd Sub\n");
        let p = parse(&src);
        assert_eq!(p.found[0].1, offset_of(&src, "Assert", 0) + 6 + rel, "{text:?}: {:?}", p.found);
    }
}
