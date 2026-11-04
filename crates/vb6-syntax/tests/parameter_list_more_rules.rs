//! Further parameter-list rules: the object value of a property assignment,
//! repeated names, fixed-length strings, ParamArray after Optional, and the
//! procedure rules for Declare.

mod common;
use common::{offset_of, parse};

const PROPERTY_PROCEDURES_INCONSISTENT: u32 = 0x9c43;
const DUPLICATE_DECLARATION: u32 = 0x9c9f;
const EXPECTED_ARRAY: u32 = 0x9c8a;

#[test]
fn property_set_value_is_an_object_parameter() {
    for ty in ["Long", "String", "Date", "Boolean", "Integer", "Double"] {
        let src = format!("Property Set Prp(v As {ty})\nEnd Property\n");
        assert_eq!(
            parse(&src).found,
            vec![(PROPERTY_PROCEDURES_INCONSISTENT, offset_of(&src, "v", 0), 1)],
            "{ty}"
        );
    }
    for (text, len) in [("Property Set Prp(v())\nEnd Property\n", 1), ("Property Set Prp(v%)\nEnd Property\n", 2)] {
        assert_eq!(
            parse(text).found,
            vec![(PROPERTY_PROCEDURES_INCONSISTENT, offset_of(text, "v", 0), len)],
            "{text}"
        );
    }
    for ty in ["", " As Variant", " As Object", " As Foo"] {
        let src = format!("Property Set Prp(v{ty})\nEnd Property\n");
        assert_eq!(parse(&src).found, vec![], "{ty}");
    }
    // Only the last parameter is the value.
    assert_eq!(parse("Property Set Prp(i As Long, v As Object)\nEnd Property\n").found, vec![]);
}

#[test]
fn a_parameter_name_cannot_repeat_except_in_declare_and_event() {
    for text in [
        "Sub S(a, a)\nEnd Sub\n",
        "Function F(a As Long, a As Long) As Long\nEnd Function\n",
        "Property Let Prp(a, a)\nEnd Property\n",
    ] {
        let found = parse(text).found;
        assert_eq!(found.len(), 1, "{text}");
        assert_eq!(found[0].0, DUPLICATE_DECLARATION, "{text}");
        assert_eq!(found[0].1, offset_of(text, "a", 1), "{text}");
    }
    assert_eq!(parse("Declare Sub f Lib \"k\" (a, a)\n").found, vec![]);
}

#[test]
fn a_fixed_length_string_parameter_is_expected_array() {
    for text in ["Sub S(a As String * 3)\nEnd Sub\n", "Declare Sub f Lib \"k\" (ByVal a As String * 3)\n"] {
        let found = parse(text).found;
        assert_eq!(found[0], (EXPECTED_ARRAY, offset_of(text, "String", 0), 6), "{text}");
    }
    let found = parse("Sub S(a() As String * 3)\nEnd Sub\n").found;
    assert!(found.iter().all(|d| d.0 != EXPECTED_ARRAY), "{found:?}");
}

#[test]
fn param_array_cannot_follow_optional() {
    let src = "Sub S(Optional ParamArray a())\nEnd Sub\n";
    let found = parse(src).found;
    assert_eq!(found[0].1, offset_of(src, "ParamArray", 0));
}

#[test]
fn declare_follows_the_procedure_parameter_rules() {
    // An array cannot be ByVal.
    let src = "Declare Sub f Lib \"k\" (ByVal a())\n";
    assert_eq!(parse(src).found.len(), 1);
    // Optional ones come last.
    let src = "Declare Sub f Lib \"k\" (Optional a, b)\n";
    let found = parse(src).found;
    assert_eq!(found[0].1, offset_of(src, "b)", 0));
    // ParamArray comes last and is an array of Variant.
    let src = "Declare Sub f Lib \"k\" (ParamArray a(), b)\n";
    let found = parse(src).found;
    assert_eq!(found[0].1, offset_of(src, "a()", 0));
    let src = "Declare Sub f Lib \"k\" (ParamArray a() As Long)\n";
    assert_eq!(parse(src).found.len(), 1);
    assert_eq!(parse("Declare Sub f Lib \"k\" (ParamArray a() As Variant)\n").found, vec![]);
}

#[test]
fn optional_in_a_property_assignment_ends_the_parameter_clause() {
    // The Optional word is itself the first error here; the array marker after the type
    // is not read.
    for text in [
        "Property Let Prp(Optional a As Long())\nEnd Property\n",
        "Property Set Prp(Optional a As Long())\nEnd Property\n",
    ] {
        assert_eq!(parse(text).found, vec![], "{text}");
    }
    // The same text in a procedure is a plain syntax error at the bracket.
    let src = "Sub S(Optional a As Long())\nEnd Sub\n";
    assert_eq!(parse(src).found[0].1, offset_of(src, "()", 0));
}
