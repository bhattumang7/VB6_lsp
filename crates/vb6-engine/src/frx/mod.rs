/// FRX binary stream reader — a cursor over a byte slice.
///
/// VB6 `.frx` files are a flat byte store referenced by offset from the
/// accompanying `.frm` text file.  Every record type (picture, font, string,
/// list items, …) lives at a specific byte offset in the file; this reader
/// is seeked to that offset and then the caller reads the record.
///
/// All multi-byte values in FRX records are **little-endian**.
pub mod reader;
pub mod records;
pub mod reference;

pub use reader::{FrxError, FrxReader};
pub use records::{FrxRecord, RecordKind};
pub use reference::{FrxRef, PropKind, kind_for_property, parse_frx_reference};
