use super::counted_items::{read_counted_items, write_counted_items};
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
        let (sig, items) = read_counted_items(r)?;
        Ok(ListRecord { sig, items })
    }

    pub fn write(&self, out: &mut Vec<u8>) {
        write_counted_items(self.sig, &self.items, out);
    }
}
