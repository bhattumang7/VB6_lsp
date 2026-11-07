use crate::frx::{FrxError, FrxReader};

/// An opaque vendor control property bag (Tier-3 proprietary state).
///
/// On-disk layout:
/// ```text
/// outer:  u32 LE              — byte length of the body that follows
/// body:   [u8; outer]         — vendor-defined bytes; may start with a 16-byte CLSID
/// ```
///
/// When the body starts with 16 bytes that look like a GUID (mostly non-printable),
/// the CLSID is extracted and stored in [`OcxBagRecord::clsid`] for re-encoding.
/// The full body (including the CLSID bytes if present) is in [`OcxBagRecord::data`].
#[derive(Debug)]
pub struct OcxBagRecord<'a> {
    /// CLSID extracted from the first 16 body bytes when they look like a GUID.
    pub clsid: Option<[u8; 16]>,
    /// Full body bytes (zero-copy slice). Starts with the CLSID bytes when `clsid` is `Some`.
    pub data: &'a [u8],
}

impl<'a> OcxBagRecord<'a> {
    pub fn read(r: &mut FrxReader<'a>) -> Result<Self, FrxError> {
        let at = r.pos();
        let outer = r.read_u32_le()? as usize;
        if outer > r.remaining() {
            return Err(FrxError::LengthOverflow {
                pos: at,
                declared: outer as u32,
                remaining: r.remaining(),
            });
        }
        let body = r.read_bytes(outer)?;
        let clsid = if body.len() >= 16 && looks_like_guid(&body[..16]) {
            let mut g = [0u8; 16];
            g.copy_from_slice(&body[..16]);
            Some(g)
        } else {
            None
        };
        Ok(OcxBagRecord { clsid, data: body })
    }

    pub fn write(&self, out: &mut Vec<u8>) {
        out.extend_from_slice(&(self.data.len() as u32).to_le_bytes());
        out.extend_from_slice(self.data);
    }
}

/// A 16-byte sequence that is mostly non-printable is likely a binary GUID.
fn looks_like_guid(b: &[u8]) -> bool {
    let printable = b.iter().filter(|&&c| (0x20..0x7f).contains(&c)).count();
    printable < 12
}
