//! Semantic checks: Erase/ReDim operands, reversed Case ranges, repeated named
//! arguments, member references outside a With.

use vb6_core::session::Session;

const EXPECTED_ARRAY: u32 = 40074;
const RANGE_HAS_NO_VALUES: u32 = 40109;
const NAMED_ARGUMENT_REPEATED: u32 = 40128;
const UNQUALIFIED_REFERENCE: u32 = 40136;
const INVALID_REDIM: u32 = 40149;

fn session(src: &str) -> Session {
    Session::from_sources(vec![("m.bas".to_string(), src.as_bytes().to_vec())])
}

fn codes(src: &str) -> Vec<(u32, u32, u32)> {
    session(src).diagnostics(0).iter().map(|d| (d.code, d.span.start, d.span.len)).collect()
}

fn at(src: &str, ctx: &str, member: &str) -> u32 {
    (src.find(ctx).unwrap() + ctx.find(member).unwrap()) as u32
}

#[test]
fn erase_of_a_scalar_is_reported() {
    let src = "Sub Run()\n    Dim n As Long\n    Erase n\nEnd Sub\n";
    assert_eq!(codes(src), vec![(EXPECTED_ARRAY, at(src, "Erase n", "n"), 1)]);
}

#[test]
fn erase_of_arrays_and_variants_is_fine() {
    let src = "Sub Run()\n    Dim a(3) As Long\n    Dim d() As String\n    Dim v As Variant\n    Erase a, d, v\nEnd Sub\n";
    assert_eq!(codes(src), vec![]);
}

#[test]
fn redim_of_a_scalar_is_reported() {
    let src = "Sub Run()\n    Dim n As Long\n    ReDim n(3)\nEnd Sub\n";
    assert_eq!(codes(src), vec![(EXPECTED_ARRAY, at(src, "n(3)", "n"), 1)]);
}

#[test]
fn redim_of_a_constant_is_invalid() {
    let src = "Const Size = 5\nSub Run()\n    ReDim Size(3)\nEnd Sub\n";
    let found = codes(src);
    assert_eq!(found.len(), 1);
    assert_eq!(found[0].0, INVALID_REDIM);
}

#[test]
fn redim_of_dynamic_arrays_and_variants_is_fine() {
    let src = "Sub Run()\n    Dim d() As Long\n    Dim v As Variant\n    ReDim d(3)\n    ReDim Preserve d(5)\n    ReDim v(2)\nEnd Sub\n";
    assert_eq!(codes(src), vec![]);
}

#[test]
fn reversed_case_range_has_no_values() {
    let src = "Sub Run(x As Long)\n    Select Case x\n        Case 5 To 1\n        Case 1 To 5\n        Case 3 To 3\n    End Select\nEnd Sub\n";
    assert_eq!(codes(src), vec![(RANGE_HAS_NO_VALUES, at(src, "Case 5 To 1", "5 To 1"), 6)]);
}

#[test]
fn a_named_argument_may_be_given_once() {
    let src = "Sub F(Optional a As Long, Optional b As Long)\nEnd Sub\nSub Run()\n    F a:=1, b:=2\n    F a:=1, a:=2\nEnd Sub\n";
    let found = codes(src);
    assert_eq!(found.len(), 1);
    assert_eq!(found[0].0, NAMED_ARGUMENT_REPEATED);
    assert!(found[0].1 as usize > src.find("F a:=1, a:=2").unwrap());
}

#[test]
fn a_leading_dot_needs_an_enclosing_with() {
    let src = "Sub Run()\n    .x = 1\nEnd Sub\n";
    let found = codes(src);
    assert_eq!(found.len(), 1);
    assert_eq!(found[0].0, UNQUALIFIED_REFERENCE);
    assert_eq!(found[0].1, at(src, ".x = 1", "."));
    let ok = "Sub Run(o As Object)\n    With o\n        .x = 1\n        With .y\n            .z = 2\n        End With\n    End With\nEnd Sub\n";
    assert_eq!(codes(ok), vec![]);
}

#[test]
fn messages_are_the_compiler_texts() {
    let src = "Sub Run()\n    .x = 1\nEnd Sub\n";
    let s = session(src);
    assert_eq!(s.diagnostics(0)[0].message().as_deref(), Some("Invalid or unqualified reference"));
}

const EVENT_NOT_FOUND: u32 = 57150;

fn class_codes(src: &str) -> Vec<(u32, u32, u32)> {
    let s = Session::from_sources(vec![("c.cls".to_string(), src.as_bytes().to_vec())]);
    s.diagnostics(0).iter().map(|d| (d.code, d.span.start, d.span.len)).collect()
}

#[test]
fn raising_an_undeclared_event_is_reported() {
    let src = "Public Event Done(ByVal n As Long)\nSub Run()\n    RaiseEvent Done(1)\n    RaiseEvent Nope\nEnd Sub\n";
    assert_eq!(class_codes(src), vec![(EVENT_NOT_FOUND, at(src, "RaiseEvent Nope", "Nope"), 4)]);
}

#[test]
fn an_event_declared_after_its_use_is_found() {
    let src = "Sub Run()\n    RaiseEvent done\nEnd Sub\nEvent Done()\n";
    assert_eq!(class_codes(src), vec![]);
}

const UDT_BYVAL: u32 = 40140;

#[test]
fn a_user_defined_type_cannot_be_passed_byval() {
    let src = "Type Pt\n    X As Long\nEnd Type\nSub A(ByVal p As Pt)\nEnd Sub\nSub B(ByRef p As Pt)\nEnd Sub\nSub C(p As Pt)\nEnd Sub\nSub D(ByVal n As Long)\nEnd Sub\n";
    assert_eq!(codes(src), vec![(UDT_BYVAL, at(src, "ByVal p As Pt", "p"), 1)]);
}

const READ_ONLY_PROPERTY: u32 = 57094;

#[test]
fn a_get_only_property_cannot_be_assigned() {
    let src = "Private m As Long\nProperty Get Ro() As Long\n    Ro = m\nEnd Property\nProperty Get Rw() As Long\n    Rw = m\nEnd Property\nProperty Let Rw(v As Long)\n    m = v\nEnd Property\nSub Run()\n    Ro = 3\n    Rw = 3\nEnd Sub\n";
    assert_eq!(class_codes(src), vec![(READ_ONLY_PROPERTY, at(src, "Ro = 3", "Ro"), 2)]);
}

#[test]
fn a_property_with_a_set_accessor_can_be_assigned() {
    let src = "Property Get Item()\nEnd Property\nProperty Set Item(o As Object)\nEnd Property\nSub Run(o As Object)\n    Item = o\nEnd Sub\n";
    assert_eq!(class_codes(src), vec![]);
}
