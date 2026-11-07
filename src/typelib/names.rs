//! How an automation type is spelled in a [`vb6_engine::Library`], and the flag
//! bits the type-library format assigns to types, members and parameters.

// `VARENUM` values.
pub const VT_EMPTY: u16 = 0;
pub const VT_NULL: u16 = 1;
pub const VT_I2: u16 = 2;
pub const VT_I4: u16 = 3;
pub const VT_R4: u16 = 4;
pub const VT_R8: u16 = 5;
pub const VT_CY: u16 = 6;
pub const VT_DATE: u16 = 7;
pub const VT_BSTR: u16 = 8;
pub const VT_DISPATCH: u16 = 9;
pub const VT_ERROR: u16 = 10;
pub const VT_BOOL: u16 = 11;
pub const VT_VARIANT: u16 = 12;
pub const VT_UNKNOWN: u16 = 13;
pub const VT_DECIMAL: u16 = 14;
pub const VT_I1: u16 = 16;
pub const VT_UI1: u16 = 17;
pub const VT_UI2: u16 = 18;
pub const VT_UI4: u16 = 19;
pub const VT_I8: u16 = 20;
pub const VT_UI8: u16 = 21;
pub const VT_INT: u16 = 22;
pub const VT_UINT: u16 = 23;
pub const VT_VOID: u16 = 24;
pub const VT_HRESULT: u16 = 25;
pub const VT_PTR: u16 = 26;
pub const VT_SAFEARRAY: u16 = 27;
pub const VT_CARRAY: u16 = 28;
pub const VT_USERDEFINED: u16 = 29;
pub const VT_LPSTR: u16 = 30;
pub const VT_LPWSTR: u16 = 31;

// Type flags (`TYPEFLAGS`).
pub const TYPEFLAG_FAPPOBJECT: u16 = 0x0001;
pub const TYPEFLAG_FCONTROL: u16 = 0x0020;

// Implemented-type flags (`IMPLTYPEFLAGS`).
pub const IMPLTYPEFLAG_FDEFAULT: i32 = 0x1;
pub const IMPLTYPEFLAG_FSOURCE: i32 = 0x2;

// Parameter flags (`PARAMFLAGS`).
pub const PARAMFLAG_FLCID: u16 = 0x0004;
pub const PARAMFLAG_FRETVAL: u16 = 0x0008;
pub const PARAMFLAG_FOPT: u16 = 0x0010;
pub const PARAMFLAG_FHASDEFAULT: u16 = 0x0020;

/// The VB6 spelling of a plain automation type, or `None` for a type that has
/// no VB6 spelling on its own (a pointer, an array, a user-defined type), no
/// value (`void`, `HRESULT`), or that VB6 cannot use (`char`, unsigned 16-bit integers).
pub fn plain_type_name(vt: u16) -> Option<&'static str> {
    Some(match vt {
        VT_I2 => "Integer",
        VT_I4 | VT_INT | VT_UI4 | VT_UINT | VT_ERROR => "Long",
        VT_R4 => "Single",
        VT_R8 => "Double",
        VT_CY => "Currency",
        VT_DATE => "Date",
        VT_BSTR | VT_LPSTR | VT_LPWSTR => "String",
        VT_BOOL => "Boolean",
        VT_VARIANT => "Variant",
        VT_DISPATCH => "Object",
        VT_UNKNOWN => "IUnknown",
        VT_DECIMAL => "Decimal",
        VT_UI1 => "Byte",
        VT_I8 | VT_UI8 => "LongLong",
        _ => return None,
    })
}

/// The spelling of an array of `element`.
pub fn array_of(element: &str) -> String {
    format!("{element}()")
}

/// Whether `vt` has no value to spell: a `Sub`'s result.
pub fn is_valueless(vt: u16) -> bool {
    matches!(vt, VT_VOID | VT_HRESULT | VT_EMPTY | VT_NULL)
}
