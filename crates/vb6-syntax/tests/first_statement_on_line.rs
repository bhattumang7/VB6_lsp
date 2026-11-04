//! A block `If`, `ElseIf` and `Else` must be the first statement on their
//! line; the single-line forms may follow other statements.

mod common;
use common::{offset_of, parse};
use vb6_syntax::frontend::ast::ExprNode;

const MUST_BE_FIRST_ON_LINE: u32 = 0x9c71;

#[test]
fn block_if_after_a_colon_is_reported_at_the_if() {
    let src = "Sub S()\nx = 1: If a Then\ny = 2\nEnd If\nEnd Sub\n";
    let p = parse(src);
    assert_eq!(p.found, vec![(MUST_BE_FIRST_ON_LINE, offset_of(src, "If", 0), 2)]);
    // The block `If` is still built.
    let body = p.proc_body(0);
    assert_eq!(body.len(), 2);
    let ExprNode::If { then_body, .. } = p.node(body[1]) else { panic!("expected If") };
    assert_eq!(p.block(*then_body).len(), 1);
}

#[test]
fn else_and_elseif_after_a_colon_are_reported() {
    let src = "Sub S()\nIf a Then\nx = 1: Else\ny = 2\nEnd If\nEnd Sub\n";
    assert_eq!(parse(src).found, vec![(MUST_BE_FIRST_ON_LINE, offset_of(src, "Else", 0), 4)]);
    let src = "Sub S()\nIf a Then\nx = 1: ElseIf b Then\ny = 2\nEnd If\nEnd Sub\n";
    assert_eq!(parse(src).found, vec![(MUST_BE_FIRST_ON_LINE, offset_of(src, "ElseIf", 0), 6)]);
}

#[test]
fn single_line_if_may_follow_other_statements() {
    for ok in [
        "x = 1: If a Then y = 2",
        "x = 1: If a Then y = 2 Else y = 3",
        "x = 1: If a Then: y = 2",
        "If a Then x = 1 Else y = 2",
    ] {
        let p = parse(&format!("Sub S()\n{ok}\nEnd Sub\n"));
        assert_eq!(p.found, vec![], "for {ok:?}");
    }
}

#[test]
fn a_label_or_line_number_does_not_count_as_a_statement() {
    for ok in [
        "Done: If a Then\nEnd If",
        "10 If a Then\nEnd If",
        "20: If a Then\nEnd If",
    ] {
        let p = parse(&format!("Sub S()\n{ok}\nEnd Sub\n"));
        assert_eq!(p.found, vec![], "for {ok:?}");
    }
}

#[test]
fn the_next_line_starts_fresh() {
    let p = parse("Sub S()\nx = 1:\nIf a Then\nEnd If\nEnd Sub\n");
    assert_eq!(p.found, vec![]);
}
