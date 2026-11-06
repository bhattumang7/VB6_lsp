//! A type word or a built-in function word written in brackets: which of them keep the
//! meaning of the keyword they spell, and which are ordinary names.

mod common;
use common::{offset_of, parse};
use vb6_syntax::frontend::ast::{AstLit, ExprNode, NodeId};

const EXPECTED: u32 = 0x9c6f;
const SYNTAX_ERROR: u32 = 0x9c70;

// ── Type words ─────────────────────────────────────────────────────────────

/// The declaration item of the only variable declared in `src`, and the module parsed.
fn only_variable(src: &str) -> (common::Parsed, NodeId) {
    let p = parse(src);
    assert_eq!(p.found, vec![], "for {src:?}");
    let ExprNode::Block { stmts } = p.node(p.top[0]) else { panic!("expected a declaration block") };
    assert_eq!(stmts.len(), 1);
    let item = stmts[0];
    (p, item)
}

#[test]
fn a_bracketed_type_word_is_that_type() {
    for (word, kind) in [
        ("Integer", 2),
        ("Long", 3),
        ("Single", 4),
        ("Double", 5),
        ("Currency", 6),
        ("Date", 7),
        ("Byte", 17),
        ("Boolean", 11),
        ("Variant", 12),
        ("Object", 9),
    ] {
        let (p, item) = only_variable(&format!("Private o As [{word}]\n"));
        let ExprNode::DimItem { type_node: Some(t), is_new: false, .. } = p.node(item) else {
            panic!("expected a typed item")
        };
        assert!(matches!(p.node(*t), ExprNode::BuiltinType { kind: k } if *k == kind), "for {word}");
    }
}

#[test]
fn a_bracketed_string_takes_a_length() {
    let (p, item) = only_variable("Private o As [String] * 3\n");
    let ExprNode::DimItem { type_node: Some(t), .. } = p.node(item) else { panic!("expected a typed item") };
    let ExprNode::StringType { fixed_len: Some(len) } = p.node(*t) else { panic!("expected a fixed string") };
    assert!(matches!(p.node(*len), ExprNode::Literal { lit: AstLit::Int(3) }));

    let (p, item) = only_variable("Private o As [string]\n");
    let ExprNode::DimItem { type_node: Some(t), .. } = p.node(item) else { panic!("expected a typed item") };
    assert!(matches!(p.node(*t), ExprNode::StringType { fixed_len: None }));
}

#[test]
fn another_bracketed_word_is_a_type_name_and_takes_no_length() {
    let (p, item) = only_variable("Private o As [Foo]\n");
    let ExprNode::DimItem { type_node: Some(t), .. } = p.node(item) else { panic!("expected a typed item") };
    let ExprNode::UserType { name, child: None } = p.node(*t) else { panic!("expected a type name") };
    assert_eq!(p.name(*name), "Foo");

    // A bracketed word that is not a type word is no string either: `* 3` is left over.
    let src = "Private o As [Foo] * 3\n";
    let p = parse(src);
    assert_eq!(p.found[0], (EXPECTED, offset_of(src, "*", 0), 1));
}

#[test]
fn new_does_not_take_a_bracketed_object() {
    let src = "Private o As New [Object]\n";
    let p = parse(src);
    assert_eq!(p.messages[0], "Expected: identifier");
    assert_eq!(p.found[0], (EXPECTED, offset_of(src, "[Object]", 0), 8));

    let src = "Sub S()\nDim o As Object\nSet o = New [Object]\nEnd Sub\n";
    let p = parse(src);
    assert_eq!(p.messages[0], "Syntax error");
    assert_eq!(p.found[0].1, offset_of(src, "[Object]", 0));

    // Any other bracketed word names a class, the built-in type words included.
    let (p, item) = only_variable("Private o As New [Integer]\n");
    let ExprNode::DimItem { is_new: true, type_node: Some(t), .. } = p.node(item) else { panic!("expected New item") };
    let ExprNode::UserType { name, child: None } = p.node(*t) else { panic!("expected a class name") };
    assert_eq!(p.name(*name), "Integer");
}

// ── Built-in functions spelled in brackets ───────────────────────────────────

fn assigned_value(src: &str) -> (common::Parsed, NodeId) {
    let p = parse(src);
    let body = p.proc_body(0);
    let ExprNode::Assign { value, .. } = p.node(body[0]) else { panic!("expected an assignment") };
    let value = *value;
    (p, value)
}

#[test]
fn a_bracketed_word_that_spells_a_function_with_a_fixed_shape_keeps_the_shape() {
    for call in [
        "[CByte]",
        "[CByte](1, 2)",
        "[CByte]()",
        "[InStr](\"a\")",
        "[InStr](1, 2, 3, 4, 5)",
        "[InStrB](\"a\")",
        "[StrComp](\"a\")",
    ] {
        let src = format!("Sub S()\nv = {call}\nEnd Sub\n");
        let p = parse(&src);
        assert!(!p.found.is_empty(), "for {call}");
        assert_eq!(p.found[0].0, SYNTAX_ERROR, "for {call}");
    }
}

#[test]
fn a_bracketed_function_word_with_the_right_shape_is_the_call() {
    for (call, count) in [
        ("[CByte](1)", 1),
        ("[InStr](\"a\", \"b\")", 2),
        ("[InStr](1, \"a\", \"b\")", 3),
        ("[InStr](1, \"a\", \"b\", 1)", 4),
        ("[InStrB](\"a\", \"b\")", 2),
        ("[StrComp](\"a\", \"b\")", 2),
        ("[StrComp](\"a\", \"b\", 1)", 3),
    ] {
        let (p, value) = assigned_value(&format!("Sub S()\nv = {call}\nEnd Sub\n"));
        assert_eq!(p.found, vec![], "for {call}");
        let ExprNode::Call { func, args } = p.node(value) else { panic!("expected a call for {call}") };
        assert!(matches!(p.node(*func), ExprNode::NameRef { .. }), "for {call}");
        let ExprNode::ArgList { args: items } = p.node(*args) else { panic!("expected arguments") };
        assert_eq!(items.len(), count, "for {call}");
    }
}

#[test]
fn every_other_bracketed_function_word_is_an_ordinary_name() {
    for word in ["Abs", "CInt", "Left", "String", "Error$", "Foo"] {
        let (p, value) = assigned_value(&format!("Sub S()\nv = [{word}]\nEnd Sub\n"));
        assert_eq!(p.found, vec![], "for {word}");
        let ExprNode::NameRef { sym, .. } = p.node(value) else { panic!("expected a name for {word}") };
        assert_eq!(p.name(*sym).to_ascii_lowercase(), word.to_ascii_lowercase());
    }
    // Called with any number of arguments.
    let (p, value) = assigned_value("Sub S()\nv = [Abs](1, 2, 3)\nEnd Sub\n");
    assert_eq!(p.found, vec![]);
    assert!(matches!(p.node(value), ExprNode::Call { .. }));
}
