//! Every label a jump names must be defined in the same procedure.

mod common;
use common::{offset_of, parse};
use vb6_syntax::frontend::ast::{ExprNode, LabelRef, OnErrorKind};

const LABEL_NOT_DEFINED: u32 = 0x9ca4;

#[test]
fn every_jump_form_is_checked() {
    for (stmt, target, len) in [
        ("GoTo Nope", "Nope", 4),
        ("GoSub Nope", "Nope", 4),
        ("On Error GoTo Nope", "Nope", 4),
        ("Resume Nope", "Nope", 4),
        ("On x GoTo Nope", "Nope", 4),
        ("On x GoSub 10, Nope", "Nope", 4),
        ("GoTo 99", "99", 2),
        ("On Error GoTo 99", "99", 2),
        ("Resume 99", "99", 2),
    ] {
        let src = format!("Sub S()\n{stmt}\n10 y = 1\nEnd Sub\n");
        let p = parse(&src);
        let mut found = p.found.clone();
        found.retain(|d| d.0 == LABEL_NOT_DEFINED);
        assert_eq!(
            found,
            vec![(LABEL_NOT_DEFINED, offset_of(&src, target, 0), len)],
            "for {stmt:?}"
        );
    }
}

#[test]
fn defined_labels_satisfy_every_form() {
    let src = "Sub S()\n\
        On Error GoTo Handler\n\
        GoTo Done\n\
        GoSub Sub1\n\
        Resume Done\n\
        On x GoTo Done, Handler\n\
        GoTo 20\n\
        On Error GoTo 20\n\
        Resume 20\n\
        Exit Sub\n\
        Handler:\n\
        Done:\n\
        Sub1:\n\
        20 y = 1\n\
        End Sub\n";
    let p = parse(src);
    assert_eq!(p.found, vec![]);
}

#[test]
fn labels_are_case_insensitive_and_scoped_to_the_procedure() {
    assert_eq!(parse("Sub S()\nGoTo done\nDONE:\nEnd Sub\n").found, vec![]);
    let src = "Sub A()\nLbl:\nEnd Sub\nSub B()\nGoTo Lbl\nEnd Sub\n";
    let p = parse(src);
    assert_eq!(p.found, vec![(LABEL_NOT_DEFINED, offset_of(src, "Lbl", 1), 3)]);
}

#[test]
fn a_label_defined_after_the_jump_counts() {
    assert_eq!(parse("Sub S()\nGoTo Later\nx = 1\nLater:\nEnd Sub\n").found, vec![]);
}

#[test]
fn on_error_goto_zero_and_resume_zero_name_no_label() {
    let p = parse("Sub S()\nOn Error GoTo 0\nResume 0\nResume Next\nResume\nOn Error Resume Next\nEnd Sub\n");
    assert_eq!(p.found, vec![]);
}

#[test]
fn a_label_may_be_named_error() {
    let src = "Sub S()\nOn Error GoTo Error\nExit Sub\nError:\nMsgBox x\nEnd Sub\n";
    let p = parse(src);
    assert_eq!(p.found, vec![]);
    let body = p.proc_body(0);
    assert_eq!(body.len(), 4);
    let ExprNode::OnError { kind: OnErrorKind::Goto(LabelRef::Name(on_error)) } = p.node(body[0]) else {
        panic!("expected On Error GoTo");
    };
    let ExprNode::Label { target: LabelRef::Name(defined) } = p.node(body[2]) else {
        panic!("expected a label");
    };
    assert_eq!(on_error, defined);
    assert_eq!(p.name(*defined).to_ascii_lowercase(), "error");
    assert!(matches!(p.node(body[3]), ExprNode::CallStmt { .. }));
}

#[test]
fn error_statement_with_an_operand_is_unchanged() {
    let p = parse("Sub S()\nError 5\nEnd Sub\n");
    assert_eq!(p.found, vec![]);
    let body = p.proc_body(0);
    assert_eq!(body.len(), 1);
    assert!(matches!(p.node(body[0]), ExprNode::ErrorStmt { .. }));
}
