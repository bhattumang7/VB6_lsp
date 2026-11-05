//! Declaration items: New takes a class name, WithEvents needs a type, ReDim
//! may name a bracketed array, and the bracketed and bounded forms of a Type
//! member.

mod common;
use common::{offset_of, parse};

const EXPECTED: u32 = 0x9c6f;
const SYNTAX_ERROR: u32 = 0x9c70;
const STATEMENT_IN_TYPE: u32 = 0x9c76;

fn in_sub(stmt: &str) -> String {
    format!("Sub S()\n    {stmt}\nEnd Sub\n")
}

#[test]
fn new_before_a_built_in_type_is_an_error_at_the_type() {
    for ty in ["Long", "Integer", "Variant", "Date", "Byte", "Boolean", "Object"] {
        let src = format!("Dim a As New {ty}\n");
        assert_eq!(parse(&src).found, vec![(EXPECTED, offset_of(&src, ty, 0), ty.len() as u32)], "{ty}");
        let src = in_sub(&format!("Dim a As New {ty}"));
        assert_eq!(parse(&src).found, vec![(SYNTAX_ERROR, offset_of(&src, ty, 0), ty.len() as u32)], "{ty}");
    }
    assert_eq!(parse("Dim a As New Foo\n").found, vec![]);
}

#[test]
fn with_events_needs_a_type() {
    let src = "WithEvents a\n";
    let found = parse(src).found;
    assert!(found.iter().any(|d| d.0 == EXPECTED), "{found:?}");
}

#[test]
fn redim_names_a_bracketed_array() {
    for text in ["ReDim [a b](3) As Long", "ReDim [a b](1 To 3, 2 To 5)", "ReDim Preserve [a b](3)"] {
        assert_eq!(parse(&in_sub(text)).found, vec![], "{text}");
    }
    // A declaration still cannot take a bracketed name.
    assert!(!parse(&in_sub("Dim [a b](3) As Long")).found.is_empty());
}

#[test]
fn a_type_member_with_bounds_needs_no_as() {
    for text in [
        "Type T\n    a(1 To 3, 2 To 5)\nEnd Type\n",
        "Type T\n    a(1 To 3)\nEnd Type\n",
        "Type T\n    a(3, 1 To 4)\nEnd Type\n",
        "Type T\n    a(-1 To 3)\nEnd Type\n",
        "Type T\n    Date(1 To 3, 2 To 5)\nEnd Type\n",
    ] {
        assert_eq!(parse(text).found, vec![], "{text}");
    }
    for text in [
        "Type T\n    a\nEnd Type\n",
        "Type T\n    a()\nEnd Type\n",
        "Type T\n    a(3)\nEnd Type\n",
        "Type T\n    a(3, 4)\nEnd Type\n",
        "Type T\n    a(x)\nEnd Type\n",
    ] {
        assert!(parse(text).found.iter().any(|d| d.0 == STATEMENT_IN_TYPE), "{text}");
    }
}

#[test]
fn a_bracketed_type_member_needs_bounds() {
    for text in ["Type T\n    [a b](3) As Long\nEnd Type\n", "Type T\n    [a b]() As Long\nEnd Type\n"] {
        assert_eq!(parse(text).found, vec![], "{text}");
    }
    let src = "Type T\n    [a b]\nEnd Type\n";
    assert_eq!(parse(src).found[0], (STATEMENT_IN_TYPE, offset_of(src, "[a b]", 0), 5));
    let src = "Type T\n    [a b] As Long\nEnd Type\n";
    assert_eq!(parse(src).found[0].0, EXPECTED);
    assert_eq!(parse(src).found[0].1, offset_of(src, "[a b]", 0));
}

#[test]
fn a_constant_takes_a_built_in_type_only() {
    for (text, name) in [("Const a As Foo = 1
", "Foo"), ("Const a As New Foo = 1
", "Foo"), ("Const a As Class1 = 1
", "Class1")] {
        let found = parse(text).found;
        assert_eq!(found[0], (EXPECTED, offset_of(text, name, 0), name.len() as u32), "{text}");
    }
    let src = in_sub("Const a As Foo = 1");
    assert_eq!(parse(&src).found[0].0, 0x9c70);
    for text in ["Const a As Long = 1\n", "Const a As String = \"x\"\n", "Const a As Variant = 1\n", "Const a = 1\n"] {
        assert_eq!(parse(text).found, vec![], "{text}");
    }
}

#[test]
fn a_function_or_getter_returns_neither_a_new_object_nor_a_fixed_string() {
    for (text, at) in [
        ("Function F() As New Foo\nEnd Function\n", "New"),
        ("Property Get P() As New Foo\nEnd Property\n", "New"),
        ("Declare Function f Lib \"k\" () As New Foo\n", "New"),
    ] {
        let found = parse(text).found;
        assert_eq!(found[0], (EXPECTED, offset_of(text, at, 0), 3), "{text}");
    }
    for text in ["Function F() As String * 3\nEnd Function\n", "Property Get P() As String * 3\nEnd Property\n"] {
        let found = parse(text).found;
        assert_eq!(found[0].0, EXPECTED, "{text}");
        assert_eq!(found[0].1, offset_of(text, "String", 0), "{text}");
    }
    for text in ["Function F() As String\nEnd Function\n", "Function F() As Long()\nEnd Function\n", "Function F() As Foo\nEnd Function\n"] {
        assert_eq!(parse(text).found, vec![], "{text}");
    }
}

#[test]
fn static_is_given_once_on_a_procedure() {
    for text in [
        "Static Sub S() Static\nEnd Sub\n",
        "Public Static Sub S() Static\nEnd Sub\n",
        "Static Private Function F() As Long Static\nEnd Function\n",
        "Static Property Get P() As Long Static\nEnd Property\n",
    ] {
        let found = parse(text).found;
        assert_eq!(found.len(), 1, "{text}");
        assert_eq!(found[0].0, SYNTAX_ERROR, "{text}");
        assert_eq!(found[0].1, text.rfind("Static").unwrap() as u32, "{text}");
    }
    for text in ["Static Sub S()\nEnd Sub\n", "Sub S() Static\nEnd Sub\n", "Private Static Sub S()\nEnd Sub\n"] {
        assert_eq!(parse(text).found, vec![], "{text}");
    }
}
