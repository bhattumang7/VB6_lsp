use crate::frx::{FrxError, FrxReader};

/// Read the framing shared by the list-style records.
///
/// ```text
/// count:  u16 LE                      — number of items (0 = empty, stops here)
/// sig:    u16 LE                      — type signature (present only when count > 0)
/// for each item:
///   item_len: u16 LE                  — byte length of the item
///   bytes:    [u8; item_len]          — the item
/// ```
///
/// Returns the signature (zero when empty) and the raw items.
pub(crate) fn read_counted_items(r: &mut FrxReader<'_>) -> Result<(u16, Vec<Vec<u8>>), FrxError> {
    let count = r.read_u16_le()? as usize;
    if count == 0 {
        return Ok((0, Vec::new()));
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
    Ok((sig, items))
}

/// Write the framing read by [`read_counted_items`].
pub(crate) fn write_counted_items(sig: u16, items: &[Vec<u8>], out: &mut Vec<u8>) {
    out.extend_from_slice(&(items.len() as u16).to_le_bytes());
    if !items.is_empty() {
        out.extend_from_slice(&sig.to_le_bytes());
        for item in items {
            out.extend_from_slice(&(item.len() as u16).to_le_bytes());
            out.extend_from_slice(item);
        }
    }
}
