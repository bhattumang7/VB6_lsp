//! Block statements are matched on a stack, as VB6's compiler does: a
//! terminator is valid only for the innermost open block, a block still open
//! at the end of its procedure is reported once (the innermost), and a
//! procedure's own `End` is checked.

mod common;
use common::{offset_of, parse};
use vb6_syntax::frontend::ast::ExprNode;

const ELSE_WITHOUT_IF: u32 = 0x9c96;
const END_IF_WITHOUT_BLOCK_IF: u32 = 0x9c97;
const NEXT_WITHOUT_FOR: u32 = 0x9c98;
const LOOP_WITHOUT_DO: u32 = 0x9c99;
const WEND_WITHOUT_WHILE: u32 = 0x9c9a;
const CASE_ELSE_OUTSIDE_SELECT: u32 = 0x9ca9;
const CASE_WITHOUT_SELECT: u32 = 0x9caa;
const END_SELECT_WITHOUT_SELECT: u32 = 0x9cab;
const END_WITH_WITHOUT_WITH: u32 = 0x9cc6;

/// A stray terminator is reported at its keyword, skipped, and builds no node.
fn assert_stray(line: &str, code: u32, len: u32) {
    let src = format!("Sub S()\n{line}\nx = 1\nEnd Sub\n");
    let p = parse(&src);
    assert_eq!(p.found, vec![(code, offset_of(&src, line, 0), len)], "for `{line}`");
    let body = p.proc_body(0);
    assert_eq!(body.len(), 1, "only `x = 1` remains for `{line}`");
    assert!(matches!(p.node(body[0]), ExprNode::Assign { .. }));
}

#[test]
fn stray_terminators_are_reported_and_skipped() {
    assert_stray("End If", END_IF_WITHOUT_BLOCK_IF, 3);
    assert_stray("EndIf", END_IF_WITHOUT_BLOCK_IF, 5);
    assert_stray("End With", END_WITH_WITHOUT_WITH, 3);
    assert_stray("End Select", END_SELECT_WITHOUT_SELECT, 3);
    assert_stray("Next", NEXT_WITHOUT_FOR, 4);
    assert_stray("Loop", LOOP_WITHOUT_DO, 4);
    assert_stray("Wend", WEND_WITHOUT_WHILE, 4);
    assert_stray("Else", ELSE_WITHOUT_IF, 4);
    assert_stray("ElseIf a Then", ELSE_WITHOUT_IF, 6);
    assert_stray("Case 1", CASE_WITHOUT_SELECT, 4);
    assert_stray("Case Else", CASE_ELSE_OUTSIDE_SELECT, 4);
}

#[test]
fn terminator_must_match_the_innermost_block() {
    // `Next` inside an `If` inside a `For` closes neither.
    let src = "Sub S()\nFor i = 1 To 2\nIf a Then\nNext\nEnd If\nNext\nEnd Sub\n";
    let p = parse(src);
    assert_eq!(p.found, vec![(NEXT_WITHOUT_FOR, offset_of(src, "Next", 0), 4)]);
    let body = p.proc_body(0);
    assert_eq!(body.len(), 1);
    let ExprNode::For { body: for_body, .. } = p.node(body[0]) else { panic!("expected For") };
    let inner = p.block(*for_body);
    assert_eq!(inner.len(), 1);
    let ExprNode::If { then_body, else_body, .. } = p.node(inner[0]) else { panic!("expected If") };
    assert!(p.block(*then_body).is_empty());
    assert!(else_body.is_none());
}

#[test]
fn second_else_and_elseif_after_else_are_stray() {
    let src = "Sub S()\nIf a Then\nElse\nElse\nElseIf b Then\nEnd If\nEnd Sub\n";
    let p = parse(src);
    assert_eq!(
        p.found,
        vec![
            (ELSE_WITHOUT_IF, offset_of(src, "Else", 1), 4),
            (ELSE_WITHOUT_IF, offset_of(src, "ElseIf", 0), 6),
        ]
    );
    let body = p.proc_body(0);
    assert_eq!(body.len(), 1);
    let ExprNode::If { then_body, else_body, .. } = p.node(body[0]) else { panic!("expected If") };
    assert!(p.block(*then_body).is_empty());
    assert!(p.block(else_body.unwrap()).is_empty());
}

#[test]
fn case_inside_an_if_inside_a_select_is_stray() {
    let src = "Sub S()\nSelect Case x\nCase 1\nIf a Then\nCase 2\nEnd If\nEnd Select\nEnd Sub\n";
    let p = parse(src);
    assert_eq!(p.found, vec![(CASE_WITHOUT_SELECT, offset_of(src, "Case 2", 0), 4)]);
}

#[test]
fn wend_does_not_close_a_do() {
    let src = "Sub S()\nDo\nWend\nLoop\nEnd Sub\n";
    let p = parse(src);
    assert_eq!(p.found, vec![(WEND_WITHOUT_WHILE, offset_of(src, "Wend", 0), 4)]);
    let body = p.proc_body(0);
    let ExprNode::Do { body: do_body, .. } = p.node(body[0]) else { panic!("expected Do") };
    assert!(p.block(*do_body).is_empty());
}

fn assert_unclosed(block: &str, code: u32) {
    let src = format!("Sub S()\n{block}\nx = 1\nEnd Sub\n");
    let p = parse(&src);
    assert_eq!(p.found, vec![(code, offset_of(&src, "End Sub", 0), 3)], "for `{block}`");
    // The unclosed block takes every following statement of the procedure.
    let body = p.proc_body(0);
    assert_eq!(body.len(), 1, "for `{block}`");
}

#[test]
fn unclosed_blocks_are_reported_at_the_end_of_the_procedure() {
    assert_unclosed("If a Then", 0x9c92);
    assert_unclosed("While a", 0x9c93);
    assert_unclosed("Do", 0x9c94);
    assert_unclosed("For i = 1 To 2", 0x9c95);
    assert_unclosed("For Each v In c", 0x9c95);
    assert_unclosed("Select Case a", 0x9cac);
    assert_unclosed("With o", 0x9cc7);
}

#[test]
fn only_the_innermost_unclosed_block_is_reported() {
    let src = "Sub S()\nIf a Then\nFor i = 1 To 2\nEnd Sub\n";
    let p = parse(src);
    assert_eq!(p.found, vec![(0x9c95, offset_of(src, "End Sub", 0), 3)]);
}

#[test]
fn unclosed_block_does_not_leak_into_the_next_procedure() {
    let src = "Sub A()\nIf a Then\nEnd Sub\nSub B()\nx = 1\nEnd Sub\n";
    let p = parse(src);
    assert_eq!(p.found, vec![(0x9c92, offset_of(src, "End Sub", 0), 3)]);
    let second = p.proc_body(1);
    assert_eq!(second.len(), 1);
    assert!(matches!(p.node(second[0]), ExprNode::Assign { .. }));
}

#[test]
fn end_of_file_reports_the_missing_procedure_end_only() {
    let eof = |src: &str| src.len() as u32;
    let src = "Sub S()\nx = 1\n";
    assert_eq!(parse(src).found, vec![(0x9c8f, eof(src), 0)]);
    let src = "Function F()\nx = 1\n";
    assert_eq!(parse(src).found, vec![(0x9c90, eof(src), 0)]);
    let src = "Property Get P()\nx = 1\n";
    assert_eq!(parse(src).found, vec![(0x9cbe, eof(src), 0)]);
    let src = "Property Let P(v)\nx = 1\n";
    assert_eq!(parse(src).found, vec![(0x9cbe, eof(src), 0)]);
    // A block still open at the end of the file is not reported separately.
    let src = "Sub S()\nIf a Then\n";
    assert_eq!(parse(src).found, vec![(0x9c8f, eof(src), 0)]);
}

#[test]
fn end_of_the_wrong_kind_is_reported_at_the_end_statement() {
    let src = "Function F()\nEnd Sub\n";
    assert_eq!(parse(src).found, vec![(0x9c90, offset_of(src, "End Sub", 0), 3)]);
    let src = "Sub S()\nEnd Function\n";
    assert_eq!(parse(src).found, vec![(0x9c8f, offset_of(src, "End Function", 0), 3)]);
    let src = "Property Get P()\nEnd Sub\n";
    assert_eq!(parse(src).found, vec![(0x9cbe, offset_of(src, "End Sub", 0), 3)]);
    let src = "Sub S()\nEnd Property\n";
    assert_eq!(parse(src).found, vec![(0x9c8f, offset_of(src, "End Property", 0), 3)]);
}

#[test]
fn only_a_comment_may_follow_the_procedure_end() {
    let src = "Sub S()\nEnd Sub x = 1\n";
    let p = parse(src);
    assert_eq!(p.found, vec![(0x9c91, offset_of(src, "x", 0), 1)]);
    assert_eq!(p.top.len(), 1);

    let src = "Sub S()\nEnd Sub: Dim a\n";
    assert_eq!(parse(src).found, vec![(0x9c91, offset_of(src, ":", 0), 1)]);

    for ok in ["Sub S()\nEnd Sub ' done\n", "Sub S()\nEnd Sub\n", "Sub S()\nEnd Sub", "Sub S()\nEnd Sub Rem done\n"] {
        assert_eq!(parse(ok).found, vec![], "for {ok:?}");
    }
}

#[test]
fn bare_end_and_balanced_blocks_stay_clean() {
    let p = parse(
        "Sub S()\nIf a Then\nFor i = 1 To 2\nDo\nWhile b\nWith o\nSelect Case i\nCase 1\nEnd\nEnd Select\nEnd With\nWend\nLoop\nNext\nElse\nEnd\nEnd If\nEnd Sub\n",
    );
    assert_eq!(p.found, vec![]);
}
