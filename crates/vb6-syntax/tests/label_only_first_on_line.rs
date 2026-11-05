//! A `Name:` is a line label only as the first thing on its logical line (a line
//! number written without a colon does not count); anywhere else the name is an
//! ordinary statement followed by a separator. `Randomize:` is never a label.

mod common;
use common::{offset_of, parse, Parsed};
use vb6_syntax::frontend::ast::{ExprNode, LabelRef, NodeId};
use vb6_syntax::frontend::token::TypeSuffix;

const LABEL_NOT_DEFINED: u32 = 0x9ca4;

fn is_label_named(p: &Parsed, id: NodeId, want: &str) -> bool {
    matches!(p.node(id), ExprNode::Label { target: LabelRef::Name(s) } if p.name(*s) == want)
}

fn is_call_of(p: &Parsed, id: NodeId, want: &str) -> bool {
    matches!(p.node(id), ExprNode::NameRef { sym, suffix: TypeSuffix::None } if p.name(*sym) == want)
}

#[test]
fn first_on_line_is_a_label() {
    let p = parse("Sub S()\n Foo:\nEnd Sub\n");
    assert_eq!(p.diags, 0);
    let b = p.proc_body(0);
    assert_eq!(b.len(), 1);
    assert!(is_label_named(&p, b[0], "Foo"));
}

#[test]
fn name_after_a_colon_is_a_call() {
    let src = "Sub S()\n x = 1 : Foo:\nEnd Sub\nSub Foo()\nEnd Sub\n";
    let p = parse(src);
    assert_eq!(p.diags, 0);
    let b = p.proc_body(0);
    assert_eq!(b.len(), 2);
    assert!(matches!(p.node(b[0]), ExprNode::Assign { .. }));
    assert!(is_call_of(&p, b[1], "Foo"));
}

#[test]
fn a_goto_does_not_find_a_name_that_follows_a_colon() {
    let src = "Sub S()\n GoTo Foo\n x = 1 : Foo:\nEnd Sub\n";
    let p = parse(src);
    assert_eq!(p.found, vec![(LABEL_NOT_DEFINED, offset_of(src, "Foo", 0), 3)]);
}

#[test]
fn name_after_a_label_is_a_call() {
    let p = parse("Sub S()\n Foo: Bar:\nEnd Sub\nSub Bar()\nEnd Sub\n");
    assert_eq!(p.diags, 0);
    let b = p.proc_body(0);
    assert_eq!(b.len(), 2);
    assert!(is_label_named(&p, b[0], "Foo"));
    assert!(is_call_of(&p, b[1], "Bar"));
}

#[test]
fn name_after_a_line_number_with_a_colon_is_a_call() {
    let p = parse("Sub S()\n 10: Foo:\nEnd Sub\nSub Foo()\nEnd Sub\n");
    assert_eq!(p.diags, 0);
    let b = p.proc_body(0);
    assert_eq!(b.len(), 2);
    assert!(matches!(p.node(b[0]), ExprNode::Label { target: LabelRef::Line(10) }));
    assert!(is_call_of(&p, b[1], "Foo"));
}

#[test]
fn name_after_a_line_number_without_a_colon_is_a_label() {
    let p = parse("Sub S()\n 10 Foo:\nEnd Sub\n");
    assert_eq!(p.diags, 0);
    let b = p.proc_body(0);
    assert_eq!(b.len(), 2);
    assert!(matches!(p.node(b[0]), ExprNode::Label { target: LabelRef::Line(10) }));
    assert!(is_label_named(&p, b[1], "Foo"));
}

#[test]
fn name_after_then_or_a_case_colon_does_not_define_a_label() {
    let src = "Sub S()
 GoTo Foo
 If x Then Foo:
End Sub
";
    let p = parse(src);
    assert_eq!(p.found, vec![(LABEL_NOT_DEFINED, offset_of(src, "Foo", 0), 3)]);
    let src = "Sub S()
 GoTo Foo
 Select Case x
 Case 1: Foo:
 End Select
End Sub
";
    let p = parse(src);
    assert_eq!(p.found, vec![(LABEL_NOT_DEFINED, offset_of(src, "Foo", 0), 3)]);
}

#[test]
fn randomize_before_a_colon_is_never_a_label() {
    let src = "Sub S()\n GoTo Randomize\n Randomize:\nEnd Sub\n";
    let p = parse(src);
    assert_eq!(p.found, vec![(LABEL_NOT_DEFINED, offset_of(src, "Randomize", 0), 9)]);
    let b = p.proc_body(0);
    assert_eq!(b.len(), 2);
    assert!(is_call_of(&p, b[1], "Randomize"));
}
