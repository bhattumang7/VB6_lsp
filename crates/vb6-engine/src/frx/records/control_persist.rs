use crate::frx::{FrxError, FrxReader};
use super::clsid::Clsid;
use super::ole_storage::OleStorageRecord;
use super::property_bag::PropertyBagRecord;

/// The persistence path chosen for a custom control.
///
/// The VB6 runtime probes COM interfaces in order:
/// 1. `IPersistStreamInit` (IID `{7FD52380}`) — `Save`/`Load` + `InitNew`
/// 2. `IPersistStorage`   (IID `{0000010A}`) — nested compound-document store
/// 3. `IPersistStream`    (IID `{00000109}`) — older stream-only interface
///
/// It additionally falls back to `IPersistPropertyBag`.
///
/// **Why `Stream` covers both `IPersistStreamInit` and `IPersistStream`:**
/// Both interfaces expose `Load` at vtbl offset +0x14 (they share the same
/// vtable layout up to `GetSizeMax`; `InitNew` at +0x20 exists only in
/// `IPersistStreamInit`).  The loader always probes `IPersistStreamInit`
/// first, falls back to `IPersistStream`, then calls `Load` at vtbl+0x14 on
/// whichever one answered.  Because the wire format written by `Save` is
/// identical for both interfaces, a single tag (`0x01 = Stream`) is
/// sufficient; the loader re-probes at runtime.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
#[repr(u8)]
pub enum ControlPersistPath {
    /// `IPersistStreamInit` (preferred) or `IPersistStream` fallback.
    ///
    /// Both interfaces use the same on-disk format (`Save` at vtbl+0x18,
    /// `Load` at vtbl+0x14).  The runtime re-probes which QI to issue.
    Stream        = 0x01,
    /// `IPersistStorage` — nested OLE compound document (IStorage).
    Storage       = 0x02,
    /// `IPersistPropertyBag` — named text properties.
    PropertyBag   = 0x03,
}

/// A custom ActiveX control's persisted state.
///
/// On-disk layout:
/// ```text
/// path:    u8         — ControlPersistPath discriminant
/// clsid:   [u8; 16]  — control's CLSID
/// one of (selected by `path`):
///   Stream:      OleStorageRecord (mechanism=Stream, clsid omitted — already above)
///   Storage:     OleStorageRecord (mechanism=Storage)
///   PropertyBag: PropertyBagRecord
/// ```
///
/// Control creation and the initial load call `IClassFactory2::CreateInstanceLic`
/// for licensed controls.
#[derive(Debug)]
pub enum ControlPersistRecord<'a> {
    Stream(Clsid, OleStorageRecord<'a>),
    Storage(Clsid, OleStorageRecord<'a>),
    PropertyBag(Clsid, PropertyBagRecord),
}

impl<'a> ControlPersistRecord<'a> {
    pub fn read(r: &mut FrxReader<'a>) -> Result<Self, FrxError> {
        let at = r.pos();
        let path_byte = r.read_u8()?;
        let clsid = Clsid::read(r)?;
        match path_byte {
            0x01 => {
                let blob = OleStorageRecord::read(r)?;
                Ok(ControlPersistRecord::Stream(clsid, blob))
            }
            0x02 => {
                let blob = OleStorageRecord::read(r)?;
                Ok(ControlPersistRecord::Storage(clsid, blob))
            }
            0x03 => {
                let bag = PropertyBagRecord::read(r)?;
                Ok(ControlPersistRecord::PropertyBag(clsid, bag))
            }
            other => Err(FrxError::BadMagic { pos: at, expected: 0x01, got: other as u16 }),
        }
    }

    pub fn clsid(&self) -> &Clsid {
        match self {
            ControlPersistRecord::Stream(c, _)      => c,
            ControlPersistRecord::Storage(c, _)     => c,
            ControlPersistRecord::PropertyBag(c, _) => c,
        }
    }

    pub fn write(&self, out: &mut Vec<u8>) {
        match self {
            ControlPersistRecord::Stream(clsid, blob) => {
                out.push(0x01);
                clsid.write(out);
                blob.write(out);
            }
            ControlPersistRecord::Storage(clsid, blob) => {
                out.push(0x02);
                clsid.write(out);
                blob.write(out);
            }
            ControlPersistRecord::PropertyBag(clsid, bag) => {
                out.push(0x03);
                clsid.write(out);
                bag.write(out);
            }
        }
    }
}
