//! `Declare` library and alias operands, and `As New` targets.

mod common;
use common::{offset_of, parse};
use vb6_syntax::frontend::ast::{AstLit, ExprNode};

const EXPECTED: u32 = 0x9c6f;

#[test]
fn lib_and_alias_are_string_literals() {
    let p = parse("Declare Function F Lib \"kernel32\" Alias \"FooA\" (ByVal x As Long) As Long\n");
    assert_eq!(p.found, vec![]);
    let ExprNode::DeclareDecl { lib, alias, .. } = p.node(p.top[0]) else { panic!("not a Declare") };
    assert!(matches!(p.node(*lib), ExprNode::Literal { lit: AstLit::Str(s) } if &**s == "kernel32"));
    let alias = alias.expect("alias");
    assert!(matches!(p.node(alias), ExprNode::Literal { lit: AstLit::Str(s) } if &**s == "FooA"));
}

#[test]
fn calling_convention_word_may_precede_lib() {
    let p = parse("Declare Function F CDecl Lib \"k\" () As Long\n");
    assert_eq!(p.found, vec![]);
    assert!(matches!(p.node(p.top[0]), ExprNode::DeclareDecl { .. }));
}

#[test]
fn non_literal_lib_or_alias_is_reported_at_the_operand() {
    for (src, bad) in [
        ("Declare Function F Lib kernel32 ()\n", "kernel32"),
        ("Declare Function F Lib \"k\" Alias foo ()\n", "foo"),
        ("Declare Function F Lib \"k\" Alias 5 ()\n", "5"),
    ] {
        let p = parse(src);
        assert_eq!(p.found, vec![(EXPECTED, offset_of(src, bad, 0), bad.len() as u32)], "{src:?}");
    }
}

#[test]
fn object_is_not_creatable() {
    let src = "Sub S()\n Dim x As New Object\nEnd Sub\n";
    let p = parse(src);
    assert_eq!(p.found, vec![(EXPECTED, offset_of(src, "Object", 0), 6)]);
    let p = parse("Sub S()\n Dim x As New Collection\nEnd Sub\n");
    assert_eq!(p.found, vec![]);
}
