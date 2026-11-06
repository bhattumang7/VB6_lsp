//! Errors raised while working out constants, bounds, fixed string lengths and Optional
//! defaults: the code, where it is reported and its text.

mod common;
use common::*;

/// `(code, offset, length, message)` of every report of `src`.
fn reports(src: &str) -> Vec<(u32, u32, u32, String)> {
    let p = parse(src);
    p.found.iter().zip(&p.messages).map(|((c, o, l), m)| (*c, *o, *l, m.clone())).collect()
}

fn one(code: u32, at: u32, len: u32, message: &str) -> Vec<(u32, u32, u32, String)> {
    vec![(code, at, len, message.to_string())]
}

#[test]
fn a_value_out_of_the_declared_type_overflows_at_the_name() {
    assert_eq!(reports("Const c As Byte = 300\n"), one(6, 6, 1, "Overflow"));
    assert_eq!(reports("Const c As Integer = 40000\n"), one(6, 6, 1, "Overflow"));
}

#[test]
fn division_by_zero_is_reported_at_the_name() {
    assert_eq!(reports("Const c = 1 / 0\n"), one(11, 6, 1, "Division by zero"));
    assert_eq!(reports("Enum E\na = 1 / 0\nEnd Enum\n"), one(11, 7, 1, "Division by zero"));
}

#[test]
fn an_empty_array_range_is_reported_at_the_array_name() {
    assert_eq!(reports("Dim a(1 To 0) As Long\n"), one(40109, 4, 1, "Range has no values"));
}

#[test]
fn a_fixed_string_length_outside_one_to_65526_is_reported_at_the_length() {
    assert_eq!(reports("Dim s As String * 0\n"), one(40112, 18, 1, "Invalid length for fixed-length string"));
}

#[test]
fn an_optional_default_out_of_the_parameter_type_overflows_at_the_parameter_name() {
    assert_eq!(reports("Sub S(Optional x As Byte = 300)\nEnd Sub\n"), one(6, 4, 1, "Overflow"));
}

#[test]
fn option_compare_text_decides_how_constant_strings_compare() {
    let src = "Option Compare Text\nConst c = \"A\" = \"a\"\nConst d = 1 \\ (c + 1)\n";
    assert_eq!(reports(src), one(11, 46, 1, "Division by zero"));
    let src = "Option Compare Binary\nConst c = \"A\" = \"a\"\nConst d = 1 \\ (c + 1)\n";
    assert_eq!(reports(src), vec![]);
}

#[test]
fn a_constant_that_is_fine_leaves_the_tree_as_read() {
    let p = parse("Const c As Long = 2 + 3\n");
    assert_eq!(p.found, vec![]);
    assert_eq!(p.top.iter().map(|n| p.show(*n)).collect::<Vec<_>>(), vec!["{(item c const bounds=- type=type3 init=(+ 2% 3%))}".to_string()]);
}
