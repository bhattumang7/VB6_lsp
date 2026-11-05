//! After a statement-separating colon a `Rem` comment ends the statement; before any
//! colon, after a label's or line number's colon, and after the operand lists of
//! `Close`, `Input`, `Call` and `Erase` it is an error.

mod common;
use common::{offset_of, parse};
use vb6_syntax::frontend::ast::ExprNode;

const EXPECTED: u32 = 0x9c6f;

fn sub(body: &str) -> String {
    format!("Sub S()\n{body}\nEnd Sub\n")
}

#[test]
fn colon_then_assignment_then_rem_is_two_statements() {
    let p = parse(&sub(" x = 1 : y = 2 Rem c"));
    assert_eq!(p.diags, 0);
    let b = p.proc_body(0);
    assert_eq!(b.len(), 2);
    assert!(matches!(p.node(b[0]), ExprNode::Assign { .. }));
    assert!(matches!(p.node(b[1]), ExprNode::Assign { .. }));
}

#[test]
fn colon_then_other_statements_then_rem_are_accepted() {
    for stmt in ["Print 1", "Open \"a\" For Input As #1", "Dim a As Long", "GoTo L"] {
        let p = parse(&format!("Sub S()\nL:\n x = 1 : {stmt} Rem c\nEnd Sub\n"));
        assert_eq!(p.diags, 0, "{stmt}");
    }
}

#[test]
fn rem_before_any_colon_is_rejected() {
    let src = sub(" x = 1 Rem c");
    let p = parse(&src);
    assert_eq!(p.found, vec![(EXPECTED, offset_of(&src, "Rem", 0), 5)]);
}

#[test]
fn rem_after_the_operand_list_of_four_statements_is_rejected() {
    for stmt in ["Close 1", "Input #1, a", "Call S", "Erase a"] {
        let src = sub(&format!(" x = 1 : {stmt} Rem c"));
        let p = parse(&src);
        assert_eq!(p.found[0], (EXPECTED, offset_of(&src, "Rem", 0), 5), "{stmt}");
    }
}

#[test]
fn a_label_or_line_number_colon_does_not_allow_rem() {
    for lead in ["L:", "10:"] {
        let src = sub(&format!(" {lead} x = 1 Rem c"));
        let p = parse(&src);
        assert_eq!(p.found, vec![(EXPECTED, offset_of(&src, "Rem", 0), 5)], "{lead}");
    }
}

#[test]
fn block_else_without_a_colon_does_not_allow_rem() {
    let src = sub(" If x Then\n y = 1\n Else y = 2 Rem c\n End If");
    let p = parse(&src);
    assert_eq!(p.found[0], (EXPECTED, offset_of(&src, "Rem", 0), 5));
}

#[test]
fn case_and_do_colons_allow_rem() {
    let p = parse(&sub(" Select Case 1\n Case 1: x = 1 Rem c\n End Select"));
    assert_eq!(p.diags, 0);
    let p = parse(&sub(" Do: x = 1 Rem c\n Loop"));
    assert_eq!(p.diags, 0);
}
