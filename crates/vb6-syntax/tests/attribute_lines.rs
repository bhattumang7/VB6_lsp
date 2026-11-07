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
    assert_eq!(values.len(), 1);
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
    assert_eq!(p.found, vec![(0x9c70, src.len() as u32, 0)]);
}

#[test]
fn a_name_chain_without_an_equals_sign_is_a_syntax_error() {
    for src in ["Attribute foo\n", "Attribute VB_Name\n", "Attribute Foo.VB_Description \"x\"\n"] {
        let p = parse(src);
        assert_eq!(p.found.len(), 1, "{src}");
        assert_eq!(p.found[0].0, 0x9c70, "{src}");
    }
}

#[test]
fn the_rest_of_an_ignored_line_includes_colons() {
    for src in [
        "Attribute Attribute VB_Name = \"m\"\n",
        "Attribute : VB_Name = \"m\"\n",
        "Attribute VB_Name = : \"m\"\n",
        "Attribute VB_Name = \"m\": Dim y\n",
    ] {
        let p = parse(src);
        assert_eq!(p.found, vec![], "{src:?}");
        assert_eq!(p.top.len(), 1, "{src:?}: the whole line is one attribute line");
    }
}

#[test]
fn a_plain_name_before_the_equals_sign_is_a_syntax_error() {
    for src in ["Attribute x = \"m\"\n", "Attribute Step = \"m\"\n"] {
        let p = parse(src);
        assert!(!p.found.is_empty(), "{src:?}");
        assert!(p.found.iter().all(|d| d.0 == 0x9c70 || d.0 == 0x9c6f), "{src:?}: {:?}", p.found);
    }
    // A dotted name is fine; an attribute word not in the table is an error at the line that follows.
    assert_eq!(parse("Attribute VB_Anything = 1\n").found, vec![(0x9c70, 26, 0)]);
    {
        let src = "Attribute Foo.x = \"m\"\n";
        assert_eq!(parse(src).found, vec![], "{src:?}");
    }
}

#[test]
fn as_after_attribute_is_expected_type_name() {
    for src in ["Attribute As = \"m\"\n", "Attribute As VB_Name = \"m\"\n"] {
        let p = parse(src);
        assert_eq!(p.found.len(), 1, "{src:?}: {:?}", p.found);
        assert_eq!(p.found[0].0, 0x9c6f, "{src:?}");
        assert_eq!(p.found[0].1, src.find("As").unwrap() as u32, "{src:?}");
    }
}

#[test]
fn else_or_a_line_number_before_an_attribute_line_still_makes_an_attribute_line() {
    for src in [
        "Else Attribute VB_Name = \"m\"\n",
        "1 Attribute VB_Name = \"m\"\n",
        "Sub S()\nElse Attribute VB_Name = \"m\"\nEnd Sub\n",
        "Sub S()\n1 Attribute VB_Name = \"m\"\nEnd Sub\n",
    ] {
        assert_eq!(parse(src).found, vec![], "{src:?}");
    }
}

#[test]
fn a_prefixed_attribute_line_follows_the_same_name_rules() {
    for src in ["Else Attribute x = \"m\"\n", "1 Attribute x = \"m\"\n"] {
        assert!(!parse(src).found.is_empty(), "{src:?}");
    }
}

#[test]
fn a_bare_attribute_word_after_else_or_a_line_number_is_not_an_attribute_line() {
    // `Else` without an `If` is still reported, and a line number plus a bare word at module level is
    // still a statement outside a procedure.
    let body = "Sub S()\nElse Attribute\nEnd Sub\n";
    assert!(parse(body).found.iter().any(|d| d.0 == 0x9c96), "{:?}", parse(body).found);
    let module = "1 Attribute\n";
    assert!(parse(module).found.iter().any(|d| d.0 == 0x9c73), "{:?}", parse(module).found);
}

#[test]
fn a_prefix_and_a_colon_before_an_attribute_line_still_make_an_attribute_line() {
    for src in [
        "Else: Attribute VB_Name = \"m\"\n",
        "10: Attribute VB_Name = \"m\"\n",
        "Sub S()\nElse: Attribute VB_Name = \"m\"\nEnd Sub\n",
    ] {
        assert_eq!(parse(src).found, vec![], "{src:?}");
    }
}

#[test]
fn a_statement_before_a_trailing_attribute_line_belongs_to_that_line_at_module_level() {
    assert_eq!(parse("x = 1: Attribute VB_Name = \"m\"\n").found, vec![]);
    // Without the trailing attribute line it is still a statement outside a procedure.
    let p = parse("x = 1: y = 2\n");
    assert_eq!(p.found.len(), 1);
    assert_eq!(p.found[0].0, 0x9c73);
}

#[test]
fn a_second_line_number_on_one_line_is_a_syntax_error() {
    let src = "Sub S()\n10 20 x = 1\nEnd Sub\n";
    let p = parse(src);
    assert_eq!(p.found[0], (0x9c70, src.find("20").unwrap() as u32, 2));
}
