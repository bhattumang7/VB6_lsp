use crate::frx::{FrxError, FrxReader};

/// A `PropertyPages` record listing the property-page GUIDs / names for a control.
///
/// On-disk layout:
/// ```text
/// count:  u32 LE
/// for each page:
///   name_len: u16 LE             — byte length INCLUDING the trailing NUL
///   name:     [u8; name_len]     — ANSI page name with NUL terminator
/// ```
#[derive(Debug)]
pub struct PropertyPagesRecord {
    /// Page name bytes (ANSI, without trailing NUL terminator).
    pub pages: Vec<Vec<u8>>,
}

impl PropertyPagesRecord {
    pub fn read(r: &mut FrxReader<'_>) -> Result<Self, FrxError> {
        let count = r.read_u32_le()? as usize;
        let mut pages = Vec::with_capacity(count);
        for _ in 0..count {
            let len_pos = r.pos();
            let len = r.read_u16_le()? as usize;
            if len > r.remaining() {
                return Err(FrxError::LengthOverflow {
                    pos: len_pos,
                    declared: len as u32,
                    remaining: r.remaining(),
                });
            }
            let raw = r.read_bytes(len)?;
            // len includes the trailing NUL; strip it
            let name = raw.split(|&c| c == 0).next().unwrap_or(raw);
            pages.push(name.to_vec());
        }
        Ok(PropertyPagesRecord { pages })
    }

    pub fn write(&self, out: &mut Vec<u8>) {
        out.extend_from_slice(&(self.pages.len() as u32).to_le_bytes());
        for page in &self.pages {
            // stored length includes the trailing NUL
            out.extend_from_slice(&((page.len() + 1) as u16).to_le_bytes());
            out.extend_from_slice(page);
            out.push(0); // NUL terminator
        }
    }
}
