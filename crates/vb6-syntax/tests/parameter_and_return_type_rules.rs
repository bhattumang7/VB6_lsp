//! Parameter lists and return types: a comma needs a parameter after it,
//! `New` takes a class name and no built-in type, and the array marker on a
//! type needs both brackets.

mod common;
use common::{offset_of, parse};

const ERR_EXPECTED: u32 = 0x9c6f;

#[test]
fn trailing_comma_in_a_parameter_list_is_expected_identifier_at_the_bracket() {
    for (text, owner) in [
        ("Sub S(a As Long, )\nEnd Sub\n", "Sub"),
        ("Function F(a As Long, ) As Long\nEnd Function\n", "Function"),
        ("Property Let P(v As Long, )\nEnd Property\n", "Property"),
        ("Declare Function f Lib \"k\" (a As Long, ) As Long\n", "Declare"),
        ("Event E(a As Long, )\n", "Event"),
    ] {
        let p = parse(text);
        let at = offset_of(text, ")", 0);
        assert!(p.found.iter().any(|d| d.0 == ERR_EXPECTED && d.1 == at), "{owner}: {:?}", p.found);
    }
}

#[test]
fn parameter_list_without_trailing_comma_is_clean() {
    let p = parse("Sub S(a As Long, ByVal b, ParamArray c())\nEnd Sub\n");
    assert_eq!(p.found, vec![]);
}

#[test]
fn new_before_a_built_in_type_is_expected_type_name() {
    for (text, word) in [
        ("Sub S(a As New Long)\nEnd Sub\n", "Long"),
        ("Function F() As New String\nEnd Function\n", "String"),
        ("Property Get P() As New Long\nEnd Property\n", "Long"),
        ("Declare Function f Lib \"k\" (a As New Long) As Long\n", "Long"),
        ("Declare Function f Lib \"k\" () As New Long\n", "Long"),
    ] {
        let p = parse(text);
        let at = offset_of(text, word, 0);
        assert!(p.found.iter().any(|d| d.0 == ERR_EXPECTED && d.1 == at), "{text}: {:?}", p.found);
    }
}

#[test]
fn new_before_a_class_name_is_clean_in_a_return_type() {
    let p = parse("Function F() As New Foo
End Function
");
    assert_eq!(p.found, vec![]);
}

#[test]
fn new_in_a_parameter_is_expected_type_name_at_new() {
    for text in [
        "Sub S(a As New Foo)
End Sub
",
        "Function F(ByVal a As New Foo) As Long
End Function
",
        "Declare Sub f Lib \"k\" (a As New Foo)
",
    ] {
        let p = parse(text);
        let at = offset_of(text, "New", 0);
        assert_eq!(p.found[0], (ERR_EXPECTED, at, 3), "{text}");
    }
}

#[test]
fn array_marker_on_a_type_needs_its_closing_bracket() {
    for text in ["Function F() As String(\nEnd Function\n", "Property Get P() As Long(\nEnd Property\n"] {
        let p = parse(text);
        assert!(!p.found.is_empty(), "{text}");
        assert!(p.found.iter().any(|d| d.0 == 0x9c6f), "{text}: {:?}", p.found);
    }
    assert_eq!(parse("Function F() As String()\nEnd Function\n").found, vec![]);
}
