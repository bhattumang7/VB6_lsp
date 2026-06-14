use crate::frx::{FrxError, FrxReader};

/// A single named property in a VBPropertyBag record.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct BagEntry {
    /// Property name (ANSI bytes, no NUL).
    pub name: Vec<u8>,
    /// Property value bytes (ANSI text representation, e.g. `b"True"`, `b"42"`).
    pub value: Vec<u8>,
}

/// VBPropertyBag (named text property bag) record.
///
/// VB6's `IPersistPropertyBag` path stores control properties as name=value
/// text pairs. This is used for controls that implement `IPersistPropertyBag`
/// rather than `IPersistStream`.  It is also used for the VBA `PropertyBag`
/// object (`UserControl` `InitProperties` / `ReadProperties` / `WriteProperties`
/// methods).
///
/// On-disk layout:
/// ```text
/// entry_count: u32 LE
/// for each entry:
///   name_len:  u32 LE            — byte length of property name
///   name:      [u8; name_len]    — ANSI property name
///   value_len: u32 LE            — byte length of value text
///   value:     [u8; value_len]   — ANSI text value (e.g. b"True", b"42")
/// ```
#[derive(Debug)]
pub struct PropertyBagRecord {
    pub entries: Vec<BagEntry>,
}

impl PropertyBagRecord {
    /// Parse a PropertyBag record from the current stream position.
    pub fn read(r: &mut FrxReader<'_>) -> Result<Self, FrxError> {
        let count = r.read_u32_le()? as usize;
        let mut entries = Vec::with_capacity(count);
        for _ in 0..count {
            let name  = r.read_len_prefixed_bytes()?.to_vec();
            let value = r.read_len_prefixed_bytes()?.to_vec();
            entries.push(BagEntry { name, value });
        }
        Ok(PropertyBagRecord { entries })
    }

    /// Serialize back to the on-disk format.
    pub fn write(&self, out: &mut Vec<u8>) {
        let count = self.entries.len() as u32;
        out.extend_from_slice(&count.to_le_bytes());
        for e in &self.entries {
            let name_len = e.name.len() as u32;
            out.extend_from_slice(&name_len.to_le_bytes());
            out.extend_from_slice(&e.name);
            let val_len = e.value.len() as u32;
            out.extend_from_slice(&val_len.to_le_bytes());
            out.extend_from_slice(&e.value);
        }
    }
}
