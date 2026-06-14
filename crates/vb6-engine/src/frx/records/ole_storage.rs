use crate::frx::{FrxError, FrxReader};
use super::clsid::Clsid;

/// Persistence mechanism used to store a control's state blob.
///
/// Maps to the OLE VT type codes and to the interface probed at load time:
///
/// | Variant   | Wire byte | OLE VARTYPE                  | Interface                  |
/// |-----------|-----------|------------------------------|----------------------------|
/// | `Stream`  |  0x01     | `VT_STREAMED_OBJECT` (0x44)  | `IPersistStreamInit` / `IPersistStream` |
/// | `Storage` |  0x02     | `VT_STORED_OBJECT`  (0x45)   | `IPersistStorage`          |
///
/// The capability probe uses bit 0x1 = StreamInit, bit 0x8 = Stream
/// (both collapse to `Stream`), bit 0x4 = Storage (→ `Storage`).
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum PersistMechanism {
    /// `IPersistStreamInit` (preferred) or `IPersistStream` fallback.
    ///
    /// Capability bits 0x1 / 0x8.  `Load` is at vtbl+0x14 for both.
    Stream,
    /// `IPersistStorage` — nested OLE compound document (`IStorage`).
    ///
    /// Capability bit 0x4.  Written via `WriteClassStg` + `IPersistStorage::Save`.
    Storage,
}

/// An opaque control-state blob.
///
/// On-disk layout (STREAMED_OBJECT path, type tag 0x44 in the OLE property type
/// system):
/// ```text
/// mechanism:  u8    — 0x01 = Stream, 0x02 = Storage
/// clsid:      [u8; 16]  — the control's CLSID (LE layout, see Clsid)
/// size:       u32 LE    — byte length of the opaque state blob
/// data:       [u8; size] — the raw IPersistStream / IPersistStorage bytes
/// ```
///
/// The runtime calls `WriteClassStg(pStorage, &clsid)` then
/// `pStorage->IPersistStorage::Save` (or the Stream equivalents), so the
/// `data` bytes are exactly what the control's own `IPersistStream::Save`
/// writes.  This parser does NOT interpret them further — they are forwarded
/// to the control on load via `IPersistStreamInit::Load` (or the fallback
/// chain).
#[derive(Debug)]
pub struct OleStorageRecord<'a> {
    pub mechanism: PersistMechanism,
    pub clsid: Clsid,
    /// Raw IPersistStream/IPersistStorage bytes.  Zero-copy slice.
    pub data: &'a [u8],
}

impl<'a> OleStorageRecord<'a> {
    pub fn read(r: &mut FrxReader<'a>) -> Result<Self, FrxError> {
        let at = r.pos();
        let mech_byte = r.read_u8()?;
        let mechanism = match mech_byte {
            0x01 => PersistMechanism::Stream,
            0x02 => PersistMechanism::Storage,
            other => return Err(FrxError::BadMagic { pos: at, expected: 0x01, got: other as u16 }),
        };
        let clsid = Clsid::read(r)?;
        let size = r.read_u32_le()? as usize;
        if size > r.remaining() {
            return Err(FrxError::LengthOverflow {
                pos: r.pos() - 4,
                declared: size as u32,
                remaining: r.remaining(),
            });
        }
        let data = r.read_bytes(size)?;
        Ok(OleStorageRecord { mechanism, clsid, data })
    }

    /// Serialize back to the on-disk format.
    pub fn write(&self, out: &mut Vec<u8>) {
        out.push(match self.mechanism {
            PersistMechanism::Stream  => 0x01,
            PersistMechanism::Storage => 0x02,
        });
        self.clsid.write(out);
        out.extend_from_slice(&(self.data.len() as u32).to_le_bytes());
        out.extend_from_slice(self.data);
    }
}
