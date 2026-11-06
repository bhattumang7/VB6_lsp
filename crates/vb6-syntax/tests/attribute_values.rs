//! The value of an `Attribute` line as the line reader builds it: one literal (a string with its
//! escapes applied), a negated number, or the two strings of an extended key; the lines that are
//! wrong are reported at the next line that is not an attribute line.

mod common;
use common::{offset_of, parse, Parsed};
use vb6_syntax::frontend::ast::{AstLit, ExprNode, UnOpKind};

/// The target name, the attribute name and the value nodes of the `index`-th top-level node.
fn attribute(p: &Parsed, index: usize) -> (Option<String>, String, Vec<AstLit>) {
    let ExprNode::Attribute { target, name, values } = p.node(p.top[index]) else {
        panic!("expected an Attribute")
    };
    (
        target.map(|t| p.name(t).to_string()),
        p.name(*name).to_string(),
        values
            .iter()
            .map(|v| match p.node(*v) {
                ExprNode::Literal { lit } => lit.clone(),
                _ => panic!("expected a literal value"),
            })
            .collect(),
    )
}

fn lit(l: AstLit) -> AstLit {
    l
}

fn text(s: &str) -> AstLit {
    AstLit::Str(s.into())
}

#[test]
fn a_string_value_has_its_escapes_applied() {
    for (written, shown) in [
        ("x", "x"),
        ("a\\nb", "a\nb"),
        ("a\\rb", "a\rb"),
        ("a\\\\b", "a\\b"),
        ("a\\qb", "aqb"),
        ("\\", ""),
        ("", ""),
    ] {
        let p = parse(&format!("Attribute VB_Description = \"{written}\"\n"));
        assert_eq!(p.found, vec![], "{written}");
        assert_eq!(attribute(&p, 0), (None, "VB_Description".to_string(), vec![text(shown)]), "{written}");
    }
}

#[test]
fn a_number_value_is_one_literal_and_a_negated_one_is_a_negation() {
    let p = parse("Attribute VB_HelpID = 7\n");
    assert_eq!(p.found, vec![]);
    let (t, n, v) = attribute(&p, 0);
    assert_eq!((t, n.as_str()), (None, "VB_HelpID"));
    assert_eq!(v.len(), 1);
    assert!(matches!(v[0], AstLit::Int(7) | AstLit::Long(7)), "{:?}", v[0]);

    let p = parse("Attribute VB_HelpID = -7\n");
    assert_eq!(p.found, vec![]);
    let ExprNode::Attribute { values, .. } = p.node(p.top[0]) else { panic!() };
    assert_eq!(values.len(), 1);
    let ExprNode::UnOp { op, operand } = p.node(values[0]) else { panic!("expected a negation") };
    assert_eq!(*op, UnOpKind::Neg);
    assert!(matches!(p.node(*operand), ExprNode::Literal { lit: AstLit::Int(7) | AstLit::Long(7) }));
}

#[test]
fn a_boolean_value_is_one_literal() {
    for (written, value) in [("True", true), ("False", false)] {
        let p = parse(&format!("Attribute VB_Exposed = {written}\n"));
        assert_eq!(p.found, vec![], "{written}");
        assert_eq!(attribute(&p, 0), (None, "VB_Exposed".to_string(), vec![lit(AstLit::Bool(value))]), "{written}");
    }
}

#[test]
fn a_member_attribute_keeps_its_target() {
    let p = parse("Attribute Foo.VB_Description = \"d\"\n");
    assert_eq!(p.found, vec![]);
    assert_eq!(attribute(&p, 0), (Some("Foo".to_string()), "VB_Description".to_string(), vec![text("d")]));
}

#[test]
fn an_extended_key_has_a_pair_of_strings_that_keep_their_backslashes() {
    let p = parse("Attribute Foo.VB_Ext_KEY = \"a\\\\b\", \"c\"\n");
    assert_eq!(p.found, vec![]);
    assert_eq!(
        attribute(&p, 0),
        (Some("Foo".to_string()), "VB_Ext_KEY".to_string(), vec![text("a\\\\b"), text("c")])
    );
}

#[test]
fn an_extended_key_without_its_second_string_is_wrong() {
    let src = "Attribute Foo.VB_Ext_KEY = \"a\"\nPublic Sub S()\nEnd Sub\n";
    let p = parse(src);
    assert_eq!(p.found, vec![(0x9c70, offset_of(src, "Public", 0), 0)]);
}

#[test]
fn a_wrong_line_is_reported_at_the_next_line_that_is_not_an_attribute_line() {
    let src = "Attribute VB_Name = \"m\"\nAttribute foo\nAttribute VB_Description = \"x\"\nOption Explicit\n";
    let p = parse(src);
    assert_eq!(p.found, vec![(0x9c70, offset_of(src, "Option", 0), 0)]);
}

#[test]
fn a_wrong_line_at_the_end_of_the_text_is_reported_at_the_end() {
    let src = "Attribute foo\n";
    assert_eq!(parse(src).found, vec![(0x9c70, src.len() as u32, 0)]);
    let src = "Attribute foo";
    assert_eq!(parse(src).found, vec![(0x9c70, src.len() as u32, 0)]);
}

#[test]
fn an_indented_attribute_line_is_reported_at_the_end_of_its_own_line() {
    for lead in ["  ", "\t", " \t "] {
        let src = format!("{lead}Attribute foo\nOption Explicit\n");
        assert_eq!(parse(&src).found, vec![(0x9c70, offset_of(&src, "
", 0), 0)], "{lead:?}");
    }
}

#[test]
fn a_run_of_wrong_lines_is_reported_once() {
    let src = "Attribute foo\nAttribute bar\nOption Explicit\n";
    let at = offset_of(src, "Option", 0);
    assert_eq!(parse(src).found, vec![(0x9c70, at, 0)]);
}

#[test]
fn a_prefixed_attribute_line_is_read_and_one_after_end_sub_is_still_text_after_the_end() {
    let p = parse("Public Sub S()\nEnd Sub\n10: Attribute S.VB_Description = \"d\"\n");
    assert_eq!(p.found, vec![]);
    let src = "Public Sub S()\nEnd Sub: Attribute S.VB_Description = \"d\"\n";
    assert_eq!(parse(src).found, vec![(0x9c91, offset_of(src, ":", 0), 1)]);
}

#[test]
fn an_attribute_line_between_enum_members_is_not_a_member() {
    let p = parse("Public Enum E\nA\nAttribute VB_Description = \"x\"\nB\nEnd Enum\n");
    assert_eq!(p.found, vec![]);
    let ExprNode::EnumDecl { members, .. } = p.node(p.top[0]) else { panic!("expected an Enum") };
    assert_eq!(members.len(), 2);
}

#[test]
fn an_attribute_line_between_type_members_is_not_a_member() {
    let p = parse("Public Type T\nA As Long\nAttribute VB_Description = \"x\"\nB As Long\nEnd Type\n");
    assert_eq!(p.found, vec![]);
    let ExprNode::TypeDecl { members, .. } = p.node(p.top[0]) else { panic!("expected a Type") };
    assert_eq!(members.len(), 2);
}

#[test]
fn an_attribute_line_between_select_and_the_first_case_is_accepted() {
    let p = parse("Public Sub Q()\nSelect Case 1\nAttribute VB_Description = \"x\"\nCase 1\nEnd Select\nEnd Sub\n");
    assert_eq!(p.found, vec![]);
}

#[test]
fn a_statement_between_select_and_the_first_case_is_still_reported() {
    let p = parse("Public Sub Q()\nSelect Case 1\nDim a\nCase 1\nEnd Select\nEnd Sub\n");
    assert_eq!(p.found.len(), 1);
}
