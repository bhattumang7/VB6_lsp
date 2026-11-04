//! `Attribute` lines in a module or after a procedure: the compiler reads a
//! name (with optional `.name` parts joined without blanks) and the `=`, and
//! ignores what follows the `=`; a line that does not begin with a name is not
//! read at all. The fixture has the compiler's answer for each shape.

mod common;
use common::answers::{assert_none_wrong, disagreements};
use vb6_syntax::frontend::ast::ExprNode;

const TEMPLATES: &str = include_str!("fixtures/attribute_lines/templates.txt");
const ANSWERS: &str = include_str!("fixtures/attribute_lines/oracle.tsv");

#[test]
fn attribute_lines_match_the_compiler() {
    let (checked, wrong) = disagreements(TEMPLATES, ANSWERS, |_, _, _, _| false);
    assert!(checked > 100, "fixture is unexpectedly small: {checked}");
    assert_none_wrong(checked, &wrong);
}

fn parse(src: &str) -> common::Parsed {
    common::parse(src)
}

#[test]
fn a_well_formed_attribute_keeps_its_target_name_and_values() {
    let p = parse("Attribute Foo.VB_Description = \"x\", 2\n");
    assert_eq!(p.found, vec![]);
    let ExprNode::Attribute { target, name, values } = p.node(p.top[0]) else { panic!("expected an Attribute") };
    assert_eq!(p.name(target.expect("a target")), "Foo");
    assert_eq!(p.name(*name), "VB_Description");
    assert_eq!(values.len(), 2);
}

#[test]
fn an_attribute_without_a_target() {
    let p = parse("Attribute VB_Name = \"Module1\"\n");
    assert_eq!(p.found, vec![]);
    let ExprNode::Attribute { target, name, values } = p.node(p.top[0]) else { panic!("expected an Attribute") };
    assert_eq!(*target, None);
    assert_eq!(p.name(*name), "VB_Name");
    assert_eq!(values.len(), 1);
}

#[test]
fn a_line_that_does_not_begin_with_a_name_is_ignored() {
    for src in ["Attribute\n", "Attribute 1 2\n", "Attribute = 1\n", "Attribute Call\n", "Attribute .VB_Description = \"x\"\n"] {
        let p = parse(src);
        assert_eq!(p.found, vec![], "{src}");
        assert!(matches!(p.node(p.top[0]), ExprNode::Block { stmts } if stmts.is_empty()), "{src}");
    }
}

#[test]
fn what_follows_the_equals_sign_is_never_an_error() {
    let p = parse("Attribute Foo.VB_Description = \"x\" \"y\" +\nAttribute Foo.VB_Description = = .\n");
    assert_eq!(p.found, vec![]);
    assert_eq!(p.top.len(), 2);
}

#[test]
fn a_blank_before_the_dot_is_a_syntax_error() {
    let src = "Attribute Foo .VB_Description = \"x\"\n";
    let p = parse(src);
    assert_eq!(p.found, vec![(0x9c70, src.find(".VB").unwrap() as u32, 1)]);
}

#[test]
fn a_name_chain_without_an_equals_sign_is_a_syntax_error() {
    for src in ["Attribute foo\n", "Attribute VB_Name\n", "Attribute Foo.VB_Description \"x\"\n"] {
        let p = parse(src);
        assert_eq!(p.found.len(), 1, "{src}");
        assert_eq!(p.found[0].0, 0x9c70, "{src}");
    }
}
