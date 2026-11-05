//! After `On Error` only `Resume Next` and `GoTo` follow.

mod common;
use common::{offset_of, parse};
use vb6_syntax::frontend::ast::{ExprNode, LabelRef, OnErrorKind};

#[test]
fn the_three_handler_forms() {
    let p = parse("Sub S()\n On Error Resume Next\n On Error GoTo L\n On Error GoTo 0\nL:\nEnd Sub\n");
    assert_eq!(p.diags, 0);
    let b = p.proc_body(0);
    assert!(matches!(p.node(b[0]), ExprNode::OnError { kind: OnErrorKind::ResumeNext }));
    let ExprNode::OnError { kind: OnErrorKind::Goto(LabelRef::Name(s)) } = p.node(b[1]) else { panic!("not On Error GoTo") };
    assert_eq!(p.name(*s), "L");
    assert!(matches!(p.node(b[2]), ExprNode::OnError { kind: OnErrorKind::Disable }));
}

#[test]
fn an_expression_after_error_is_rejected() {
    for tail in ["x GoTo L", "1 GoTo L", "-1 GoTo L", "Len(c) GoTo L", "x GoSub L", "(x) GoTo L", "\"a\" GoTo L", "Not x GoTo L"] {
        let src = format!("Sub S()\n On Error {tail}\nL:\nEnd Sub\n");
        let p = parse(&src);
        assert_eq!(p.found[0].1, offset_of(&src, "Error ", 0) + 6, "{tail}: {:?}", p.found);
    }
    let src = "Sub S()\n On Local Error x GoTo L\nL:\nEnd Sub\n";
    let p = parse(src);
    assert_eq!(p.found[0].1, offset_of(src, "x GoTo", 0));
}

#[test]
fn on_expression_goto_is_unchanged() {
    let p = parse("Sub S()\n On x GoTo A, B\nA:\nB:\nEnd Sub\n");
    assert_eq!(p.diags, 0);
    let b = p.proc_body(0);
    let ExprNode::OnGo { is_gosub: false, labels, .. } = p.node(b[0]) else { panic!("not On GoTo") };
    assert_eq!(labels.len(), 2);
}
