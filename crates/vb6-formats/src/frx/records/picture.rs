use crate::frx::{FrxError, FrxReader};

/// `"lt\0\0"` as a little-endian u32 (bytes: 6C 74 00 00).
const LT_MAGIC: u32 = 0x0000_746C;

/// An OLE picture blob as stored in the FRX file.
///
/// On-disk layout (standard form):
/// ```text
/// outer:    u32 LE              — total framing size = 8 + data_len
/// magic:    [u8; 4] = "lt\0\0"  — OleLoadPicture stream header
/// data_len: u32 LE              — byte length of the raw image data
/// data:     [u8; data_len]      — raw image bytes (BMP, ICO, WMF, EMF, …)
/// ```
///
/// ImageList / collection-bag variant prefixes a 16-byte CLSID between `outer`
/// and `magic`. In that case `outer = 24 + data_len` and `clsid` is `Some(_)`.
///
/// An empty picture slot (removed icon) has `data_len == 0`. When there is no
/// CLSID, the slot encodes as `[u32=8]["lt\0\0"][u32=0]`.
///
/// The image format (BMP, ICO, WMF, …) is inferred from the data magic bytes
/// at the call site — there is no discriminant field in the FRX record itself.
#[derive(Debug)]
pub struct PictureRecord<'a> {
    /// CLSID present for ImageList / collection-bag framing; `None` for
    /// standard Form/Control picture properties.
    pub clsid: Option<[u8; 16]>,
    /// Raw image bytes. Empty slice for an empty picture slot.
    pub data: &'a [u8],
}

impl<'a> PictureRecord<'a> {
    pub fn read(r: &mut FrxReader<'a>) -> Result<Self, FrxError> {
        let _outer = r.read_u32_le()?;

        // A 16-byte CLSID may appear between `outer` and `"lt\0\0"`.
        // Peek to decide which layout we have.
        let clsid = if r.peek_u32_le()? == LT_MAGIC {
            None
        } else {
            if r.remaining() < 16 {
                return Err(FrxError::UnexpectedEof {
                    pos: r.pos(),
                    needed: 16,
                    available: r.remaining(),
                });
            }
            let raw = r.read_bytes(16)?;
            let mut g = [0u8; 16];
            g.copy_from_slice(raw);
            Some(g)
        };

        // Consume and verify "lt\0\0" magic.
        let magic_pos = r.pos();
        let magic = r.read_u32_le()?;
        if magic != LT_MAGIC {
            return Err(FrxError::BadMagic {
                pos: magic_pos,
                expected: 0x746C,
                got: (magic & 0xFFFF) as u16,
            });
        }

        let data_len_pos = r.pos();
        let data_len = r.read_u32_le()? as usize;
        if data_len > r.remaining() {
            return Err(FrxError::LengthOverflow {
                pos: data_len_pos,
                declared: data_len as u32,
                remaining: r.remaining(),
            });
        }
        let data = r.read_bytes(data_len)?;
        Ok(PictureRecord { clsid, data })
    }

    pub fn write(&self, out: &mut Vec<u8>) {
        let data_len = self.data.len() as u32;
        if let Some(c) = &self.clsid {
            // CLSID-framed: outer = CLSID(16) + magic(4) + data_len(4) + data
            out.extend_from_slice(&(data_len + 24).to_le_bytes());
            out.extend_from_slice(c);
        } else {
            // Standard: outer = magic(4) + data_len(4) + data
            out.extend_from_slice(&(data_len + 8).to_le_bytes());
        }
        out.extend_from_slice(b"lt\0\0");
        out.extend_from_slice(&data_len.to_le_bytes());
        out.extend_from_slice(self.data);
    }
}
