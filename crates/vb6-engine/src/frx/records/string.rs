use crate::frx::{FrxError, FrxReader};

/// A `u32`-length-prefixed byte blob (the `$"file.frx":N` long-string form).
///
/// On-disk layout:
/// ```text
/// byteLen: u32 LE          — number of bytes of string data
/// data:    [u8; byteLen]   — ANSI/MBCS bytes (no NUL terminator)
/// ```
///
/// Used for Caption / Text / long string properties referenced as `$"...":N`
/// in the `.frm` file.
#[derive(Debug)]
pub struct LenPrefixedBytes<'a> {
    pub data: &'a [u8],
}

impl<'a> LenPrefixedBytes<'a> {
    pub fn read(r: &mut FrxReader<'a>) -> Result<Self, FrxError> {
        let data = r.read_len_prefixed_bytes()?;
        Ok(LenPrefixedBytes { data })
    }

    pub fn write(data: &[u8], out: &mut Vec<u8>) {
        let len = data.len() as u32;
        out.extend_from_slice(&len.to_le_bytes());
        out.extend_from_slice(data);
    }
}

/// A `u8`-length-prefixed byte blob (the `"file.frx":N` short-string form, no `$`).
///
/// On-disk layout:
/// ```text
/// byteLen: u8              — number of bytes of string data (max 255)
/// data:    [u8; byteLen]   — ANSI bytes (no NUL terminator)
/// ```
///
/// Used for short Caption / Tag properties referenced without `$` in the `.frm`
/// file.  Strings longer than 255 bytes use `LenPrefixedBytes` (the `$` form).
#[derive(Debug)]
pub struct StringShortBytes<'a> {
    pub data: &'a [u8],
}

impl<'a> StringShortBytes<'a> {
    pub fn read(r: &mut FrxReader<'a>) -> Result<Self, FrxError> {
        let at = r.pos();
        let len = r.read_u8()? as usize;
        if len > r.remaining() {
            return Err(FrxError::LengthOverflow {
                pos: at,
                declared: len as u32,
                remaining: r.remaining(),
            });
        }
        let data = r.read_bytes(len)?;
        Ok(StringShortBytes { data })
    }

    pub fn write(data: &[u8], out: &mut Vec<u8>) {
        let len = data.len().min(255);
        out.push(len as u8);
        out.extend_from_slice(&data[..len]);
    }
}
