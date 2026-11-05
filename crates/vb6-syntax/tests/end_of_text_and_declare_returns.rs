//! A last logical line that ends in a continuation mark is not read at all;
//! a Declare return type may be an array but never a new object.

mod common;
use common::{offset_of, parse};

#[test]
fn a_last_line_continued_into_the_end_of_the_text_is_dropped() {
    for text in [
        "Dim x As _\n",
        "Const c = _\n",
        "x = _\n",
        "Declare Function f _\n",
        "Option Explicit _",
        "Dim x As Long, _  \n",
        "' comment _\n",
        "Dim x As _\n    _\n",
    ] {
        assert_eq!(parse(text).found, vec![], "{text:?}");
    }
}

#[test]
fn a_continued_line_followed_by_an_empty_line_is_joined_to_it() {
    let src = "Dim x As Long, _\n\n";
    assert!(!parse(src).found.is_empty());
    let src = "Sub S()\nx = 1\nEnd Sub _\n\n";
    assert_eq!(parse(src).found, vec![]);
}

#[test]
fn only_the_last_logical_line_is_dropped() {
    let src = "Dim x As Long\nDim y As _\n";
    let p = parse(src);
    assert_eq!(p.found, vec![]);
    assert_eq!(p.top.len(), 1);
    // A continued line followed by more text is joined to it as usual.
    let src = "Dim x As _\n    Long\n";
    let p = parse(src);
    assert_eq!(p.found, vec![]);
    assert_eq!(p.top.len(), 1);
}

#[test]
fn a_dropped_closing_line_leaves_the_procedure_open() {
    let src = "Sub S()\nEnd Sub _\n";
    let found = parse(src).found;
    assert_eq!(found.len(), 1);
    assert_eq!(found[0].0, 0x9c8f);
}

#[test]
fn a_mark_without_a_blank_before_it_is_not_a_continuation() {
    let src = "x = 1+_\n";
    assert!(!parse(src).found.is_empty());
}

#[test]
fn a_declare_function_may_return_an_array() {
    assert_eq!(parse("Declare Function f Lib \"k\" () As Long()\n").found, vec![]);
}

#[test]
fn a_declare_function_never_returns_a_new_object() {
    let src = "Declare Function f Lib \"k\" () As New Class1\n";
    let found = parse(src).found;
    assert_eq!(found[0], (0x9c6f, offset_of(src, "New", 0), 3));
}
