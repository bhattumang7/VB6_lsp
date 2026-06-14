/// 16-byte VARIANT as laid out in 32-bit COM (VARIANTARG / VARIANT).
///
/// Matches the x86 COM ABI exactly (`#[repr(C)]`). Used by all
/// VARIANT-based `rtc*` builtins in the engine.
#[repr(C)]
#[derive(Copy, Clone, Debug, Default, PartialEq)]
pub struct Variant32 {
    /// VARTYPE discriminant (VT_* constants below).
    pub vt: u16,
    pub res1: u16,
    pub res2: u16,
    pub res3: u16,
    /// Offset +8: low 32 bits of payload (e.g. i16/i32/f32/BSTR ptr low half).
    pub data_lo: u32,
    /// Offset +12: high 32 bits (used by VT_R8, VT_CY, VT_I8, etc.).
    pub data_hi: u32,
}

impl Variant32 {
    pub const VT_EMPTY:    u16 = 0;
    pub const VT_NULL:     u16 = 1;
    pub const VT_I2:       u16 = 2;
    pub const VT_I4:       u16 = 3;
    pub const VT_R4:       u16 = 4;
    pub const VT_R8:       u16 = 5;
    pub const VT_CY:       u16 = 6;
    pub const VT_DATE:     u16 = 7;
    pub const VT_BSTR:     u16 = 8;
    pub const VT_DISPATCH: u16 = 9;
    pub const VT_ERROR:    u16 = 10;
    pub const VT_BOOL:     u16 = 11;
    pub const VT_DECIMAL:  u16 = 14;
    pub const VT_UI1:      u16 = 17;
    pub const VT_BYREF:    u16 = 0x4000;

    /// VT_EMPTY (uninitialized / numeric 0).
    pub fn empty() -> Self { Self::default() }

    /// VT_NULL (propagates through numeric expressions).
    pub fn null() -> Self {
        Self { vt: Self::VT_NULL, ..Default::default() }
    }

    /// VT_BSTR wrapping a raw BSTR pointer.
    pub fn bstr(ptr: *mut u16) -> Self {
        Self { vt: Self::VT_BSTR, data_lo: ptr as u32, ..Default::default() }
    }

    /// VT_I2 from a signed 16-bit integer.
    pub fn i2(v: i16) -> Self {
        Self { vt: Self::VT_I2, data_lo: v as u16 as u32, ..Default::default() }
    }

    /// VT_I4 from a signed 32-bit integer.
    pub fn i4(v: i32) -> Self {
        Self { vt: Self::VT_I4, data_lo: v as u32, ..Default::default() }
    }

    /// VT_R4 from a 32-bit float.
    pub fn r4(v: f32) -> Self {
        Self { vt: Self::VT_R4, data_lo: v.to_bits(), ..Default::default() }
    }

    /// VT_R8 from a 64-bit float.
    pub fn r8(v: f64) -> Self {
        let bits = v.to_bits();
        Self {
            vt: Self::VT_R8,
            data_lo: bits as u32,
            data_hi: (bits >> 32) as u32,
            ..Default::default()
        }
    }

    /// VT_BOOL: True = 0xFFFF (−1 as i16), False = 0.
    pub fn bool_vba(v: bool) -> Self {
        Self { vt: Self::VT_BOOL, data_lo: if v { 0xFFFF } else { 0 }, ..Default::default() }
    }

    /// VT_CY (Currency) from raw 64-bit fixed-point value (scaled by 10 000).
    pub fn cy(lo: u32, hi: u32) -> Self {
        Self { vt: Self::VT_CY, data_lo: lo, data_hi: hi, ..Default::default() }
    }

    /// VT_ERROR sentinel for "missing optional argument" (DISP_E_PARAMNOTFOUND).
    pub fn missing() -> Self {
        Self { vt: Self::VT_ERROR, data_lo: 0x8002_0004, ..Default::default() }
    }

    // ── Payload accessors ──────────────────────────────────────────────────

    pub fn as_i2(&self) -> i16 { self.data_lo as u16 as i16 }
    pub fn as_i4(&self) -> i32 { self.data_lo as i32 }
    pub fn as_r4(&self) -> f32 { f32::from_bits(self.data_lo) }
    pub fn as_r8(&self) -> f64 {
        f64::from_bits(self.data_lo as u64 | (self.data_hi as u64) << 32)
    }
    pub fn as_bool(&self) -> bool { self.data_lo as u16 != 0 }

    /// Read result of a VB6 comparison call: `None` = VT_NULL, `Some(i)` = VT_I2 comparison.
    pub fn compare_result(&self) -> Option<i32> {
        match self.vt {
            Self::VT_NULL => None,
            _ => Some(self.data_lo as i16 as i32),
        }
    }
}
