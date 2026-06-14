pub mod clsid;
pub mod control_create;
pub mod control_persist;
pub mod font;
pub mod itemdata;
pub mod list;
pub mod ocx_bag;
pub mod ole_storage;
pub mod picture;
pub mod property_bag;
pub mod property_pages;
pub mod string;

pub use clsid::Clsid;
pub use control_create::ControlCreateRecord;
pub use control_persist::ControlPersistRecord;
pub use font::StdFontRecord;
pub use itemdata::ItemDataRecord;
pub use list::ListRecord;
pub use ocx_bag::OcxBagRecord;
pub use ole_storage::OleStorageRecord;
pub use picture::PictureRecord;
pub use property_bag::PropertyBagRecord;
pub use property_pages::PropertyPagesRecord;
pub use string::{LenPrefixedBytes, StringShortBytes};

use crate::frx::{FrxError, FrxReader};

/// Discriminant for an FRX record, matching the VB6 property type system.
///
/// The type of each record in the FRX file is determined by the `.frm`
/// property metadata (not by a magic in the FRX bytes themselves).  This enum
/// encodes the full set of property types that VB6 persists to FRX, including
/// every type-tag case handled by the control-load dispatcher.
#[derive(Debug)]
pub enum FrxRecord<'a> {
    // -----------------------------------------------------------------------
    // Scalar / simple types — stored inline in .frm; NOT in .frx.
    // -----------------------------------------------------------------------
    Empty,
    Integer(i16),
    Long(i32),
    Single(f32),
    Double(f64),
    Currency(i64),
    Date(f64),
    Bool(bool),

    // -----------------------------------------------------------------------
    // Heap types — stored in .frx, referenced by `"form.frx":NNNNNNNN`
    // -----------------------------------------------------------------------

    /// Short string: 1-byte length prefix (no `$` in `.frm`).
    StringShort(StringShortBytes<'a>),

    /// Long string: 4-byte length prefix (`$"..."` in `.frm`).
    BinaryString(LenPrefixedBytes<'a>),

    /// StdFont binary serialization (OLEPRO32.DLL).
    Font(StdFontRecord),

    /// OLE picture blob (`[u32 outer]["lt\0\0"][u32 dataLen][image bytes]`).
    /// Covers Picture, Icon, MouseIcon, DragIcon, DisabledPicture, DownPicture,
    /// MaskPicture, TabPicture.
    Picture(PictureRecord<'a>),

    /// ListBox / ComboBox `List` items.
    ListItems(ListRecord),

    /// ListBox / ComboBox `ItemData` values (separate FRX record from `ListItems`).
    ItemData(ItemDataRecord),

    /// PropertyPages page-name list.
    PropertyPages(PropertyPagesRecord),

    /// Opaque vendor control property bag.
    OcxBag(OcxBagRecord<'a>),

    // -----------------------------------------------------------------------
    // Custom ActiveX control state
    // -----------------------------------------------------------------------

    /// Control creation metadata — CLSID + ProgID + optional license key.
    ControlCreate(ControlCreateRecord),

    /// Custom control's persisted state blob.
    ControlState(ControlPersistRecord<'a>),

    /// Opaque OLE compound-storage blob.
    OleStorage(OleStorageRecord<'a>),

    /// VBPropertyBag named-property record.
    PropertyBag(PropertyBagRecord),

    // -----------------------------------------------------------------------
    // Control-array and advanced properties
    // -----------------------------------------------------------------------

    ControlArray {
        index: u32,
        visible: bool,
        tab_stop: bool,
    },

    Menu {
        caption: Vec<u8>,
        negotiate_position: u8,
    },

    DataBinding {
        prop_name: Vec<u8>,
        data_field: Vec<u8>,
    },

    /// AsyncRead picture — same format as `Picture`.
    AsyncPicture(PictureRecord<'a>),
}

impl<'a> FrxRecord<'a> {
    pub fn read(kind: RecordKind, r: &mut FrxReader<'a>) -> Result<Self, FrxError> {
        match kind {
            RecordKind::StringShort   => Ok(FrxRecord::StringShort(StringShortBytes::read(r)?)),
            RecordKind::BinaryString  => Ok(FrxRecord::BinaryString(LenPrefixedBytes::read(r)?)),
            RecordKind::Font          => Ok(FrxRecord::Font(StdFontRecord::read(r)?)),
            RecordKind::Picture       => Ok(FrxRecord::Picture(PictureRecord::read(r)?)),
            RecordKind::AsyncPicture  => Ok(FrxRecord::AsyncPicture(PictureRecord::read(r)?)),
            RecordKind::ListItems     => Ok(FrxRecord::ListItems(ListRecord::read(r)?)),
            RecordKind::ItemData      => Ok(FrxRecord::ItemData(ItemDataRecord::read(r)?)),
            RecordKind::PropertyPages => Ok(FrxRecord::PropertyPages(PropertyPagesRecord::read(r)?)),
            RecordKind::OcxBag        => Ok(FrxRecord::OcxBag(OcxBagRecord::read(r)?)),
            RecordKind::ControlCreate => Ok(FrxRecord::ControlCreate(ControlCreateRecord::read(r)?)),
            RecordKind::ControlState  => Ok(FrxRecord::ControlState(ControlPersistRecord::read(r)?)),
            RecordKind::OleStorage    => Ok(FrxRecord::OleStorage(OleStorageRecord::read(r)?)),
            RecordKind::PropertyBag   => Ok(FrxRecord::PropertyBag(PropertyBagRecord::read(r)?)),
            RecordKind::ControlArray  => {
                let index   = r.read_u32_le()?;
                let flags   = r.read_u8()?;
                let visible  = flags & 0x01 != 0;
                let tab_stop = flags & 0x02 != 0;
                Ok(FrxRecord::ControlArray { index, visible, tab_stop })
            }
            RecordKind::Menu => {
                let caption           = r.read_len_prefixed_bytes()?.to_vec();
                let negotiate_position = r.read_u8()?;
                Ok(FrxRecord::Menu { caption, negotiate_position })
            }
            RecordKind::DataBinding => {
                let prop_name  = r.read_len_prefixed_bytes()?.to_vec();
                let data_field = r.read_len_prefixed_bytes()?.to_vec();
                Ok(FrxRecord::DataBinding { prop_name, data_field })
            }
        }
    }

    pub fn write(&self, out: &mut Vec<u8>) {
        match self {
            FrxRecord::Empty | FrxRecord::Integer(_) | FrxRecord::Long(_)
            | FrxRecord::Single(_) | FrxRecord::Double(_) | FrxRecord::Currency(_)
            | FrxRecord::Date(_) | FrxRecord::Bool(_) => {}
            FrxRecord::StringShort(s)   => StringShortBytes::write(s.data, out),
            FrxRecord::BinaryString(s)  => LenPrefixedBytes::write(s.data, out),
            FrxRecord::Font(f)          => f.write(out),
            FrxRecord::Picture(p) | FrxRecord::AsyncPicture(p) => p.write(out),
            FrxRecord::ListItems(l)     => l.write(out),
            FrxRecord::ItemData(d)      => d.write(out),
            FrxRecord::PropertyPages(p) => p.write(out),
            FrxRecord::OcxBag(b)        => b.write(out),
            FrxRecord::ControlCreate(c) => c.write(out),
            FrxRecord::ControlState(c)  => c.write(out),
            FrxRecord::OleStorage(o)    => o.write(out),
            FrxRecord::PropertyBag(b)   => b.write(out),
            FrxRecord::ControlArray { index, visible, tab_stop } => {
                out.extend_from_slice(&index.to_le_bytes());
                let flags = (*visible as u8) | ((*tab_stop as u8) << 1);
                out.push(flags);
            }
            FrxRecord::Menu { caption, negotiate_position } => {
                let len = caption.len() as u32;
                out.extend_from_slice(&len.to_le_bytes());
                out.extend_from_slice(caption);
                out.push(*negotiate_position);
            }
            FrxRecord::DataBinding { prop_name, data_field } => {
                let n = prop_name.len() as u32;
                out.extend_from_slice(&n.to_le_bytes());
                out.extend_from_slice(prop_name);
                let d = data_field.len() as u32;
                out.extend_from_slice(&d.to_le_bytes());
                out.extend_from_slice(data_field);
            }
        }
    }
}

/// Tag used to select which `FrxRecord` variant to parse.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum RecordKind {
    StringShort,
    BinaryString,
    Font,
    Picture,
    AsyncPicture,
    ListItems,
    ItemData,
    PropertyPages,
    OcxBag,
    ControlCreate,
    ControlState,
    OleStorage,
    PropertyBag,
    ControlArray,
    Menu,
    DataBinding,
}
