//! `Select Case`: nothing may stand between the header and the first `Case`,
//! and nothing may follow `Case Else`. `If`: the legacy form without `Then`
//! takes `GoTo` only, and a block `If` cannot start in the middle of a line.

mod common;
use common::{offset_of, parse};

const STATEMENT_BEFORE_FIRST_CASE: u32 = 0x9d95; // 40341
const CASE_WITHOUT_SELECT: u32 = 0x9caa; // 40106
const MUST_BE_FIRST_ON_LINE: u32 = 0x9c71; // 40049

#[test]
fn a_statement_between_select_case_and_the_first_case_is_reported_at_the_statement() {
    let src = "Sub S()\nSelect Case x\ny = 1\nCase 1\nEnd Select\nEnd Sub\n";
    let p = parse(src);
    assert_eq!(p.found, vec![(STATEMENT_BEFORE_FIRST_CASE, offset_of(src, "y = 1", 0), 1)]);
}

#[test]
fn a_label_between_select_case_and_the_first_case_is_reported_too() {
    let src = "Sub S()\nSelect Case x\nL:\nCase 1\nEnd Select\nEnd Sub\n";
    let p = parse(src);
    assert!(p.found.iter().any(|d| d.0 == STATEMENT_BEFORE_FIRST_CASE && d.1 == offset_of(src, "L:", 0)), "{:?}", p.found);
}

#[test]
fn blank_lines_and_comments_before_the_first_case_are_fine() {
    let p = parse("Sub S()\nSelect Case x\n\n' c\nCase 1\ny = 1\nEnd Select\nEnd Sub\n");
    assert_eq!(p.found, vec![]);
}

#[test]
fn a_case_after_case_else_is_case_without_select_case_at_that_case() {
    let src = "Sub S()\nSelect Case x\nCase 1\nCase Else\nCase 2\nEnd Select\nEnd Sub\n";
    let p = parse(src);
    assert_eq!(p.found, vec![(CASE_WITHOUT_SELECT, offset_of(src, "Case 2", 0), 4)]);
}

#[test]
fn a_second_case_else_is_case_without_select_case() {
    let src = "Sub S()\nSelect Case x\nCase Else\nCase Else\nEnd Select\nEnd Sub\n";
    let p = parse(src);
    assert_eq!(p.found, vec![(CASE_WITHOUT_SELECT, offset_of(src, "Case Else", 1), 4)]);
}

#[test]
fn case_else_last_is_fine() {
    let p = parse("Sub S()\nSelect Case x\nCase 1\nCase 2\nCase Else\nEnd Select\nEnd Sub\n");
    assert_eq!(p.found, vec![]);
}

#[test]
fn legacy_if_without_then_takes_goto_only() {
    let ok = parse("Sub S()\nIf x GoTo L\nL:\nEnd Sub\n");
    assert_eq!(ok.found, vec![]);
    let src = "Sub S()\nIf x GoSub L\nL:\nEnd Sub\n";
    let p = parse(src);
    assert!(!p.found.is_empty());
    assert!(p.found.iter().any(|d| d.1 >= offset_of(src, "GoSub", 0) - 1 && d.1 <= offset_of(src, "GoSub", 0) + 5), "{:?}", p.found);
}

#[test]
fn a_block_if_after_else_on_the_same_line_must_be_first_on_the_line() {
    let src = "Sub S()\nIf x Then\nElse If y Then\nEnd If\nEnd If\nEnd Sub\n";
    let p = parse(src);
    assert!(
        p.found.iter().any(|d| d.0 == MUST_BE_FIRST_ON_LINE && d.1 == offset_of(src, "If y", 0)),
        "{:?}",
        p.found
    );
}

#[test]
fn else_followed_by_a_statement_or_a_colon_is_fine() {
    for body in ["If x Then\nElse y = 1\nEnd If\n", "If x Then\nElse: y = 1\nEnd If\n"] {
        let p = parse(&format!("Sub S()\n{body}End Sub\n"));
        assert_eq!(p.found, vec![], "{body:?}");
    }
}
