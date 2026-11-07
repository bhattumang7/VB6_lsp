use crate::frx::{FrxError, FrxReader};

/// StdFont binary persistence record.
///
/// Written by `OLEPRO32.DLL!StdFont::IPersistStream::Save` and read back by
/// `StdFont::IPersistStream::Load`. VB6 stores this in FRX for Font-typed
/// properties (e.g. `Form.Font`, `Label.Font`).
///
/// On-disk layout:
/// ```text
/// version:  u8             — always 1
/// charset:  u16 LE         — ANSI_CHARSET=0, SYMBOL_CHARSET=2, …
/// flags:    u8             — bit 0=italic, bit 1=underline, bit 2=strikethrough
/// weight:   u16 LE         — FW_NORMAL=400, FW_BOLD=700
/// size:     u32 LE         — point size × 10000 (e.g. 8.25pt → 82500)
/// name_len: u8             — byte length of font name
/// name:     [u8; name_len] — ANSI font face name (no NUL terminator)
/// ```
///
/// Bold is encoded via `weight >= 700`, not via `flags`.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct StdFontRecord {
    pub charset: u16,
    /// bit 0 = italic, bit 1 = underline, bit 2 = strikethrough
    pub flags: u8,
    /// Font weight: 400 = normal, 700 = bold.
    pub weight: u16,
    /// Point size × 10000 (so 8.25pt is stored as 82500).
    pub size_times_10k: u32,
    /// ANSI font face name.
    pub name: Vec<u8>,
}

impl StdFontRecord {
    pub const VERSION: u8 = 1;

    pub fn is_italic(&self) -> bool { self.flags & 0x01 != 0 }
    pub fn is_underline(&self) -> bool { self.flags & 0x02 != 0 }
    pub fn is_strikethrough(&self) -> bool { self.flags & 0x04 != 0 }
    pub fn is_bold(&self) -> bool { self.weight >= 700 }

    /// Parse a StdFont record from the current stream position.
    pub fn read(r: &mut FrxReader<'_>) -> Result<Self, FrxError> {
        let version = r.read_u8()?;
        if version != Self::VERSION {
            // treat as a bad-magic-style error — version byte acts as a magic
            return Err(FrxError::BadMagic {
                pos: r.pos() - 1,
                expected: Self::VERSION as u16,
                got: version as u16,
            });
        }
        let charset          = r.read_u16_le()?;
        let flags            = r.read_u8()?;
        let weight           = r.read_u16_le()?;
        let size_times_10k   = r.read_u32_le()?;
        let name_len         = r.read_u8()? as usize;
        let name_bytes       = r.read_bytes(name_len)?;
        Ok(StdFontRecord {
            charset,
            flags,
            weight,
            size_times_10k,
            name: name_bytes.to_vec(),
        })
    }

    /// Serialize back to the on-disk format.
    pub fn write(&self, out: &mut Vec<u8>) {
        out.push(Self::VERSION);
        out.extend_from_slice(&self.charset.to_le_bytes());
        out.push(self.flags);
        out.extend_from_slice(&self.weight.to_le_bytes());
        out.extend_from_slice(&self.size_times_10k.to_le_bytes());
        let name_len = self.name.len().min(255) as u8;
        out.push(name_len);
        out.extend_from_slice(&self.name[..name_len as usize]);
    }
}
