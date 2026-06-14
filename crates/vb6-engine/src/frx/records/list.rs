use crate::frx::{FrxError, FrxReader};

/// A ListBox / ComboBox `List` record.
///
/// On-disk layout:
/// ```text
/// count:  u16 LE                      — number of items (0 = empty list, stops here)
/// sig:    u16 LE                      — type signature (present only when count > 0)
/// for each item:
///   item_len: u16 LE                  — byte length of item text
///   text:     [u8; item_len]          — item text (ANSI, no NUL)
/// ```
///
/// `ItemData` values are stored in a *separate* FRX record (at a different byte
/// offset, referenced by a distinct property in the `.frm` file) — they are NOT
/// a trailing section of this record.
#[derive(Debug)]
pub struct ListRecord {
    /// The u16 signature field written by VB6 when count > 0 (retained for
    /// byte-exact round-trips). Zero when the list is empty.
    pub sig: u16,
    /// Item text bytes (ANSI, no NUL).
    pub items: Vec<Vec<u8>>,
}

impl ListRecord {
    pub fn read(r: &mut FrxReader<'_>) -> Result<Self, FrxError> {
        let count = r.read_u16_le()? as usize;
        if count == 0 {
            return Ok(ListRecord { sig: 0, items: Vec::new() });
        }
        let sig = r.read_u16_le()?;
        let mut items = Vec::with_capacity(count);
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
            let bytes = r.read_bytes(len)?;
            items.push(bytes.to_vec());
        }
        Ok(ListRecord { sig, items })
    }

    pub fn write(&self, out: &mut Vec<u8>) {
        out.extend_from_slice(&(self.items.len() as u16).to_le_bytes());
        if !self.items.is_empty() {
            out.extend_from_slice(&self.sig.to_le_bytes());
            for item in &self.items {
                out.extend_from_slice(&(item.len() as u16).to_le_bytes());
                out.extend_from_slice(item);
            }
        }
    }
}
