//! A single-line `If` cannot open a block `If`; a loop, `With` or `Select Case` opened in
//! one must be closed on the same line; and a statement followed by `Else` is only valid
//! inside a single-line `If`.

mod common;
use common::{offset_of, parse};
use vb6_syntax::frontend::ast::ExprNode;

const SYNTAX: u32 = 0x9c70;
const END_IF_WITHOUT_BLOCK_IF: u32 = 0x9c97;

#[test]
fn openers_closed_on_their_own_line_are_accepted() {
    for body in [
        "If x Then For i = 1 To 2: y = 1: Next",
        "If x Then Do: y = 1: Loop",
        "If x Then While y: y = 0: Wend",
        "If x Then y = 1 Else For i = 1 To 2: Next",
    ] {
        let p = parse(&format!("Sub S()\n{body}\nEnd Sub\n"));
        assert_eq!(p.found, vec![], "{body}");
    }
}

#[test]
fn openers_left_open_at_the_line_end_report_end_if_without_block_if() {
    for (body, opener) in [
        ("If x Then For i = 1 To 2\nNext", "For"),
        ("If x Then Do\nLoop", "Do"),
        ("If x Then While y\nWend", "While"),
        ("If x Then With o\nEnd With", "With"),
        ("If x Then Select Case x\nEnd Select", "Select"),
        ("If x Then y = 1 Else For i = 1 To 2\nNext", "For"),
        ("If x Then For i = 1 To 2: y = 1\nNext", "For"),
    ] {
        let src = format!("Sub S()\n{body}\nEnd Sub\n");
        let p = parse(&src);
        assert_eq!(p.found[0].0, END_IF_WITHOUT_BLOCK_IF, "{body}");
        assert_eq!(p.found[0].1, offset_of(&src, opener, 0), "{body}");
    }
}

#[test]
fn a_block_if_cannot_open_inside_a_single_line_if() {
    for body in ["If x Then If y Then\nz = 1\nEnd If", "If x Then y = 1 Else If z Then\nz = 1\nEnd If"] {
        let src = format!("Sub S()\n{body}\nEnd Sub\n");
        let p = parse(&src);
        assert_eq!(p.found[0].0, SYNTAX, "{body}");
        assert_eq!(p.found[0].1, offset_of(&src, "Then\n", 0) + 4, "{body}");
    }
}

#[test]
fn a_nested_single_line_if_still_parses() {
    let p = parse("Sub S()\nIf x Then If y Then z = 1 Else z = 2\nEnd Sub\n");
    assert_eq!(p.found, vec![]);
    let b = p.proc_body(0);
    let ExprNode::If { then_body, else_body: None, .. } = p.node(b[0]) else { panic!("not an If") };
    let ExprNode::Block { stmts } = p.node(*then_body) else { panic!("not a block") };
    assert!(matches!(p.node(stmts[0]), ExprNode::If { else_body: Some(_), .. }));
}

#[test]
fn a_statement_followed_by_else_in_a_block_if_is_an_error() {
    for body in [
        "If x Then\ny = 1 Else y = 2\nEnd If",
        "If x Then\nElseIf z Then y = 1 Else y = 2\nEnd If",
    ] {
        let src = format!("Sub S()\n{body}\nEnd Sub\n");
        let p = parse(&src);
        assert_eq!(p.found[0], (0x9c6f, offset_of(&src, "Else y", 0), 4), "{body}");
    }
}
