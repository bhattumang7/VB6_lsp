use crate::frx::{FrxError, FrxReader};

/// A ListBox / ComboBox `ItemData` record.
///
/// Shares the same on-disk framing as `ListRecord` but carries the parallel
/// array of programmer-defined Long values associated with each list item.
/// It is stored as a **separate** FRX record at its own byte offset, referenced
/// by the `ItemData` property in the `.frm` file.
///
/// On-disk layout:
/// ```text
/// count:  u16 LE                      — number of items (0 = empty, stops here)
/// sig:    u16 LE                      — type signature (present only when count > 0)
/// for each item:
///   item_len: u16 LE                  — byte length of the raw value (1–4 bytes typical)
///   value:    [u8; item_len]          — ItemData Long value, little-endian
/// ```
///
/// Use [`ItemDataRecord::item_value`] to decode a raw item as a signed 32-bit integer.
#[derive(Debug)]
pub struct ItemDataRecord {
    /// Signature field (retained for byte-exact round-trips). Zero when empty.
    pub sig: u16,
    /// Raw item bytes (variable-length LE i32 representation per item).
    pub items: Vec<Vec<u8>>,
}

impl ItemDataRecord {
    pub fn read(r: &mut FrxReader<'_>) -> Result<Self, FrxError> {
        let count = r.read_u16_le()? as usize;
        if count == 0 {
            return Ok(ItemDataRecord { sig: 0, items: Vec::new() });
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
        Ok(ItemDataRecord { sig, items })
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

    /// Interpret a raw item (little-endian, 1–4 bytes) as a signed 32-bit integer.
    pub fn item_value(raw: &[u8]) -> i32 {
        let mut b = [0u8; 4];
        for (i, &x) in raw.iter().take(4).enumerate() {
            b[i] = x;
        }
        i32::from_le_bytes(b)
    }
}
