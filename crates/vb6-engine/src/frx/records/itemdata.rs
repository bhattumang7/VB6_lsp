use super::counted_items::{read_counted_items, write_counted_items};
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
        let (sig, items) = read_counted_items(r)?;
        Ok(ItemDataRecord { sig, items })
    }

    pub fn write(&self, out: &mut Vec<u8>) {
        write_counted_items(self.sig, &self.items, out);
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
