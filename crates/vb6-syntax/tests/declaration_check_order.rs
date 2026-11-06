//! The order the declaration checks report in: a repeated name before any value is worked out,
//! then the module's declarations as written, then the defaults in procedure headers, then the
//! bodies; a declaration after the end of a procedure is refused where it is read and declares
//! nothing.

mod common;
use common::*;

const AMBIGUOUS: u32 = 0x802c;
const OVERFLOW: u32 = 6;
const AFTER_END: u32 = 0x9c91;
const RANGE_HAS_NO_VALUES: u32 = 0x9cad;
const INVALID_LENGTH: u32 = 0x9cb0;
const SYNTAX_ERROR: u32 = 0x9c70;

fn codes(src: &str) -> Vec<(u32, u32, u32)> {
    parse(src).found
}

#[test]
fn a_repeated_name_is_reported_before_any_value_error_wherever_either_stands() {
    // Value error first, repeated name after.
    let src = "Const c As Byte = 300\nDim v\nDim V\n";
    assert_eq!(codes(src), vec![(AMBIGUOUS, 32, 1), (OVERFLOW, 6, 1)]);
    // Repeated name first, value error after.
    let src = "Dim v\nDim V\nConst c As Byte = 300\n";
    assert_eq!(codes(src), vec![(AMBIGUOUS, 10, 1), (OVERFLOW, 18, 1)]);
    // An empty range and an invalid length behind a repeated procedure name.
    let src = "Dim a(1 To 0)\nDim s As String * 0\nSub P()\nEnd Sub\nSub p()\nEnd Sub\n";
    assert_eq!(
        codes(src),
        vec![(AMBIGUOUS, 54, 1), (RANGE_HAS_NO_VALUES, 4, 1), (INVALID_LENGTH, 32, 1)]
    );
}

#[test]
fn module_declarations_are_reported_in_the_order_written() {
    let src = "Const m As Byte = 300\nDim a(1 To 0)\n";
    assert_eq!(codes(src), vec![(OVERFLOW, 6, 1), (RANGE_HAS_NO_VALUES, 26, 1)]);
    let src = "Dim a(1 To 0)\nConst m As Byte = 300\n";
    assert_eq!(codes(src), vec![(RANGE_HAS_NO_VALUES, 4, 1), (OVERFLOW, 20, 1)]);
}

#[test]
fn the_defaults_in_every_header_are_reported_before_any_body() {
    let src = "Sub P()\nConst l As Byte = 300\nEnd Sub\nSub Q(Optional o As Byte = 300)\nEnd Sub\n";
    assert_eq!(codes(src), vec![(OVERFLOW, 42, 1), (OVERFLOW, 14, 1)]);
}

#[test]
fn a_declaration_after_the_end_of_a_procedure_declares_nothing() {
    // The second `V` is refused where it is read, so the first is not repeated.
    let src = "Dim v\nSub Q()\nEnd Sub\nDim V\n";
    assert_eq!(codes(src), vec![(AFTER_END, 22, 3)]);
    // The value in a declaration after a procedure is not worked out either.
    let src = "Sub Q()\nEnd Sub\nConst c As Byte = 300\n";
    assert_eq!(codes(src), vec![(AFTER_END, 16, 5)]);
}

#[test]
fn a_declaration_after_a_procedure_is_refused_after_the_module_declarations_and_headers() {
    // A module-level value error before the procedure.
    let src = "Const m As Byte = 300\nSub Q()\nEnd Sub\nDim late As Long\n";
    assert_eq!(codes(src), vec![(OVERFLOW, 6, 1), (AFTER_END, 38, 3)]);
    // The default of a later procedure's header.
    let src = "Sub Q()\nEnd Sub\nDim late As Long\nSub P(Optional o As Byte = 300)\nEnd Sub\n";
    assert_eq!(codes(src), vec![(OVERFLOW, 37, 1), (AFTER_END, 16, 3)]);
}

#[test]
fn a_body_is_reported_among_the_refused_declarations_by_where_it_stands() {
    // The body before the refused declaration is reported first.
    let src = "Sub P()\nConst l As Byte = 300\nEnd Sub\nDim late As Long\n";
    assert_eq!(codes(src), vec![(OVERFLOW, 14, 1), (AFTER_END, 38, 3)]);
    // The body after it is reported second.
    let src = "Sub Q()\nEnd Sub\nDim late As Long\nSub P()\nConst l As Byte = 300\nEnd Sub\n";
    assert_eq!(codes(src), vec![(AFTER_END, 16, 3), (OVERFLOW, 47, 1)]);
}

#[test]
fn a_literal_length_too_large_for_16_bits_is_refused_where_it_is_read() {
    // Outside a procedure it is an invalid length, ahead of a repeated name.
    let src = "Dim s As String * 65536\nSub P()\nEnd Sub\nSub p()\nEnd Sub\n";
    assert_eq!(codes(src), vec![(INVALID_LENGTH, 18, 5), (AMBIGUOUS, 44, 1)]);
    // Inside a procedure it is a syntax error at the literal.
    for text in ["65536", "70000", "&H10000", "65536&", "00065536", "&O200000", "2147483647"] {
        let src = format!("Sub P()\nDim s As String * {text}\nEnd Sub\n");
        let at = offset_of(&src, text, 0);
        assert_eq!(codes(&src), vec![(SYNTAX_ERROR, at, text.len() as u32)], "{text}");
    }
}

#[test]
fn a_literal_length_from_65527_to_65535_or_zero_is_judged_with_the_declarations() {
    for text in ["0", "65527", "65535", "&HFFFF", "0&"] {
        let src = format!("Dim s As String * {text}\nSub P()\nEnd Sub\nSub p()\nEnd Sub\n");
        let at = offset_of(&src, text, 0);
        let second_p = offset_of(&src, "p()", 0);
        assert_eq!(
            codes(&src),
            vec![(AMBIGUOUS, second_p, 1), (INVALID_LENGTH, at, text.len() as u32)],
            "{text}"
        );
        let src = format!("Sub P()\nDim s As String * {text}\nEnd Sub\n");
        let at = offset_of(&src, text, 0);
        assert_eq!(codes(&src), vec![(INVALID_LENGTH, at, text.len() as u32)], "{text}");
    }
}
