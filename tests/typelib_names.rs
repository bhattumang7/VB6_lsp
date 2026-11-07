//! How automation types are spelled for the engine.

use vb6_lsp::typelib::names::*;

#[test]
fn plain_types_have_their_vb6_spelling() {
    let spelled: Vec<(u16, Option<&str>)> = [
        VT_I2, VT_I4, VT_R4, VT_R8, VT_CY, VT_DATE, VT_BSTR, VT_BOOL, VT_VARIANT, VT_DISPATCH, VT_UNKNOWN,
        VT_DECIMAL, VT_UI1, VT_INT, VT_UI4, VT_UINT, VT_ERROR, VT_LPSTR, VT_LPWSTR, VT_I8, VT_UI8,
    ]
    .into_iter()
    .map(|vt| (vt, plain_type_name(vt)))
    .collect();
    assert_eq!(
        spelled,
        vec![
            (VT_I2, Some("Integer")),
            (VT_I4, Some("Long")),
            (VT_R4, Some("Single")),
            (VT_R8, Some("Double")),
            (VT_CY, Some("Currency")),
            (VT_DATE, Some("Date")),
            (VT_BSTR, Some("String")),
            (VT_BOOL, Some("Boolean")),
            (VT_VARIANT, Some("Variant")),
            (VT_DISPATCH, Some("Object")),
            (VT_UNKNOWN, Some("IUnknown")),
            (VT_DECIMAL, Some("Decimal")),
            (VT_UI1, Some("Byte")),
            (VT_INT, Some("Long")),
            (VT_UI4, Some("Long")),
            (VT_UINT, Some("Long")),
            (VT_ERROR, Some("Long")),
            (VT_LPSTR, Some("String")),
            (VT_LPWSTR, Some("String")),
            (VT_I8, Some("LongLong")),
            (VT_UI8, Some("LongLong")),
        ]
    );
}

#[test]
fn types_without_a_vb6_spelling_have_none() {
    let none: Vec<u16> = [
        VT_EMPTY, VT_NULL, VT_I1, VT_UI2, VT_VOID, VT_HRESULT, VT_PTR, VT_SAFEARRAY, VT_CARRAY,
        VT_USERDEFINED,
    ]
    .into_iter()
    .filter(|&vt| plain_type_name(vt).is_none())
    .collect();
    assert_eq!(
        none,
        vec![VT_EMPTY, VT_NULL, VT_I1, VT_UI2, VT_VOID, VT_HRESULT, VT_PTR, VT_SAFEARRAY, VT_CARRAY, VT_USERDEFINED]
    );
}

#[test]
fn an_array_is_spelled_with_empty_parentheses() {
    assert_eq!(array_of("Long"), "Long()");
    assert_eq!(array_of("ADODB.Field"), "ADODB.Field()");
}

#[test]
fn only_void_hresult_and_empty_have_no_value() {
    let valueless: Vec<u16> = [VT_VOID, VT_HRESULT, VT_EMPTY, VT_NULL, VT_I4, VT_BSTR, VT_VARIANT, VT_PTR]
        .into_iter()
        .filter(|&vt| is_valueless(vt))
        .collect();
    assert_eq!(valueless, vec![VT_VOID, VT_HRESULT, VT_EMPTY, VT_NULL]);
}
