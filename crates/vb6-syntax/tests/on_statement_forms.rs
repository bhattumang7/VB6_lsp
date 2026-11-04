//! `On`: `Local` may only be followed by `Error`; the forms around it keep
//! the shapes they had.

mod common;
use common::{offset_of, parse};
use vb6_syntax::frontend::ast::{ExprNode, OnErrorKind};

#[test]
fn local_before_anything_but_error_is_a_syntax_error_at_that_word() {
    let src = "Sub S()\n    On Local x GoTo L\nL:\nEnd Sub\n";
    let p = parse(src);
    assert_eq!(p.found, vec![(0x9c70, offset_of(src, "x", 0), 1)]);
}

#[test]
fn local_error_forms_are_unchanged() {
    let p = parse("Sub S()\n    On Local Error Resume Next\nEnd Sub\n");
    assert_eq!(p.found, vec![]);
    let body = p.proc_body(0);
    assert!(matches!(p.node(body[0]), ExprNode::OnError { kind: OnErrorKind::ResumeNext }));
}

#[test]
fn a_second_local_is_a_syntax_error() {
    let src = "Sub S()\n    On Local Local Error Resume Next\nEnd Sub\n";
    let p = parse(src);
    assert_eq!(p.found[0], (0x9c70, src.rfind("Local").unwrap() as u32, 5));
}
