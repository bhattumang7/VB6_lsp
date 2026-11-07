//! Decoded `.frx`/`.ctx` blobs.
//!
//! [`decode`] reads the blob at an offset as a given [`PropKind`] and returns an owned
//! [`FrxValue`]; [`encode`] writes a value back to its on-disk bytes. Parsing and writing of the
//! blob layouts is done by the [`records`](super::records); this layer adds the interpretation
//! (image format, text decoding, font style) and the offset/length bookkeeping.

use encoding_rs::WINDOWS_1252;

use super::records::{
    ItemDataRecord, LenPrefixedBytes, ListRecord, OcxBagRecord, PictureRecord,
    PropertyPagesRecord, StdFontRecord, StringShortBytes,
};
use super::{FrxError, FrxReader, PropKind};

/// The container format of an image blob, detected from its leading bytes.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum ImageFormat {
    Bmp,
    Ico,
    Cur,
    Gif,
    Jpeg,
    Png,
    Wmf,
    Emf,
    Unknown,
}

impl ImageFormat {
    /// The file extension for an image of this format.
    pub fn ext(self) -> &'static str {
        match self {
            ImageFormat::Bmp => "bmp",
            ImageFormat::Ico => "ico",
            ImageFormat::Cur => "cur",
            ImageFormat::Gif => "gif",
            ImageFormat::Jpeg => "jpg",
            ImageFormat::Png => "png",
            ImageFormat::Wmf => "wmf",
            ImageFormat::Emf => "emf",
            ImageFormat::Unknown => "bin",
        }
    }
}

/// Font description recovered from a `StdFont` blob.
#[derive(Debug, Clone, PartialEq)]
pub struct FontInfo {
    pub name: String,
    pub size_pt: f64,
    /// Raw weight (400 = normal, 700 = bold).
    pub weight: u16,
    pub charset: u16,
    /// The raw style flags byte as stored, retained for byte-exact re-encode.
    pub raw_flags: u8,
    pub bold: bool,
    pub italic: bool,
    pub underline: bool,
    pub strikethrough: bool,
}

/// A decoded blob value.
#[derive(Debug, Clone, PartialEq)]
pub enum FrxValue {
    /// Picture / Icon / MouseIcon / ToolboxBitmap: an image with detected format.
    /// `clsid` holds the 16-byte class id when the on-disk framing carried one (the
    /// form used inside ImageList/collection bags), retained for byte-exact re-encode.
    Picture { format: ImageFormat, data: Vec<u8>, clsid: Option<[u8; 16]> },
    /// A `StdFont`.
    Font(FontInfo),
    /// Caption / Text / long string value.
    Text(String),
    /// `List` items for ListBox/ComboBox. `sig` is the 2-byte type signature that
    /// follows the count, retained so the blob re-encodes byte-for-byte.
    List { items: Vec<String>, sig: u16 },
    /// `ItemData` values paralleling a list. Each item is the raw little-endian
    /// bytes of the Long exactly as stored (1-4 bytes); use [`itemdata_value`] to
    /// read it as an `i32`. `sig` + raw bytes are retained for byte-exact re-encode.
    ItemData { items: Vec<Vec<u8>>, sig: u16 },
    /// `PropertyPages` page-name list.
    PropertyPages(Vec<String>),
    /// A proprietary control property bag, surfaced opaque so nothing is lost and
    /// nothing is silently mis-read. `clsid` is filled when a leading class id is
    /// detectable in the blob.
    OcxBag { clsid: Option<[u8; 16]>, data: Vec<u8> },
    /// A proprietary control bag decoded *live* via COM into typed properties
    /// `(name, value-as-text)`. Produced only by a
    /// [`crate::controls::resources::ComDecoder`].
    DecodedBag { clsid: Option<[u8; 16]>, properties: Vec<(String, String)> },
    /// An explicitly empty resource (e.g. a removed icon).
    Empty,
}

/// Decode the blob at `offset` interpreting it as `kind`.
pub fn decode(buf: &[u8], offset: usize, kind: PropKind) -> Result<FrxValue, FrxError> {
    decode_span(buf, offset, kind).map(|(v, _)| v)
}

/// Like [`decode`], but also returns the blob's byte span (how many bytes it
/// occupied). Used by coverage accounting to prove every byte is attributed.
pub fn decode_span(
    buf: &[u8],
    offset: usize,
    kind: PropKind,
) -> Result<(FrxValue, usize), FrxError> {
    let mut r = FrxReader::new(buf);
    r.seek(offset)?;
    let value = match kind {
        PropKind::Picture => picture_value(PictureRecord::read(&mut r)?),
        PropKind::Font => FrxValue::Font(font_info(&StdFontRecord::read(&mut r)?)),
        PropKind::StringShort => FrxValue::Text(decode_text(StringShortBytes::read(&mut r)?.data)),
        PropKind::StringLong => FrxValue::Text(decode_text(LenPrefixedBytes::read(&mut r)?.data)),
        PropKind::List => {
            let rec = ListRecord::read(&mut r)?;
            FrxValue::List {
                items: rec.items.iter().map(|b| decode_ansi(b)).collect(),
                sig: rec.sig,
            }
        }
        PropKind::ItemData => {
            let rec = ItemDataRecord::read(&mut r)?;
            FrxValue::ItemData { items: rec.items, sig: rec.sig }
        }
        PropKind::PropertyPages => {
            let rec = PropertyPagesRecord::read(&mut r)?;
            FrxValue::PropertyPages(rec.pages.iter().map(|b| decode_ansi(b)).collect())
        }
        PropKind::OcxBag => return Ok(ocx_bag_value(buf, offset)),
    };
    Ok((value, r.pos() - offset))
}

fn picture_value(rec: PictureRecord<'_>) -> FrxValue {
    match (rec.data.is_empty(), rec.clsid) {
        // A bare empty slot is a removed icon. A class-id-framed empty slot (an unset
        // TabPicture) keeps its framing so it re-encodes to the same bytes.
        (true, None) => FrxValue::Empty,
        (_, clsid) => FrxValue::Picture {
            format: detect_image_format(rec.data),
            data: rec.data.to_vec(),
            clsid,
        },
    }
}

fn font_info(rec: &StdFontRecord) -> FontInfo {
    FontInfo {
        name: decode_ansi(&rec.name),
        size_pt: rec.size_times_10k as f64 / 10000.0,
        weight: rec.weight,
        charset: rec.charset,
        raw_flags: rec.flags,
        bold: rec.is_bold(),
        italic: rec.is_italic(),
        underline: rec.is_underline(),
        strikethrough: rec.is_strikethrough(),
    }
}

/// A proprietary control bag. A leading `[u32 len]` frames the body, and a class id at the start
/// of the body is extracted. A length that runs past the end of the file keeps what is there; a
/// blob with no usable length claims everything to the end of the file (coverage reports that as
/// an opaque tail, since the bag's true length is unknown without the control).
pub fn decode_ocx_bag(buf: &[u8], offset: usize) -> FrxValue {
    ocx_bag_value(buf, offset).0
}

fn ocx_bag_value(buf: &[u8], offset: usize) -> (FrxValue, usize) {
    let mut r = FrxReader::new(buf);
    if r.seek(offset).is_ok() {
        if let Ok(outer) = r.read_u32_le() {
            let start = offset + 4;
            let end = start.saturating_add(outer as usize).min(buf.len());
            if end > start {
                let rec = OcxBagRecord::from_body(&buf[start..end]);
                return (FrxValue::OcxBag { clsid: rec.clsid, data: rec.data.to_vec() }, end - offset);
            }
        }
    }
    let data = buf.get(offset..).map(|s| s.to_vec()).unwrap_or_default();
    let span = data.len();
    (FrxValue::OcxBag { clsid: None, data }, span)
}

/// Interpret a raw `ItemData` item (little-endian, 1-4 bytes) as a 32-bit Long.
pub fn itemdata_value(raw: &[u8]) -> i32 {
    ItemDataRecord::item_value(raw)
}

/// Re-encode a decoded value back to its on-disk byte form.
///
/// Byte-exact inverse of [`decode_span`] for `Picture`, `Empty`, `Font`, `List`, `ItemData`,
/// `PropertyPages`, and `OcxBag`. `Text` is best-effort only: its decode applies a character-set
/// heuristic (a blob that held UTF-16 comes back as one-byte text).
pub fn encode(value: &FrxValue, kind: PropKind) -> Vec<u8> {
    let mut out = Vec::new();
    match value {
        FrxValue::Empty => PictureRecord { clsid: None, data: &[] }.write(&mut out),
        FrxValue::Picture { data, clsid, .. } => {
            PictureRecord { clsid: *clsid, data }.write(&mut out)
        }
        FrxValue::Font(f) => StdFontRecord {
            charset: f.charset,
            flags: f.raw_flags,
            weight: f.weight,
            size_times_10k: (f.size_pt * 10000.0).round() as u32,
            name: encode_ansi(&f.name),
        }
        .write(&mut out),
        FrxValue::Text(s) => {
            let bytes = encode_ansi(s);
            match kind {
                PropKind::StringShort => StringShortBytes::write(&bytes, &mut out),
                _ => LenPrefixedBytes::write(&bytes, &mut out),
            }
        }
        FrxValue::List { items, sig } => {
            ListRecord { sig: *sig, items: items.iter().map(|s| encode_ansi(s)).collect() }
                .write(&mut out)
        }
        FrxValue::ItemData { items, sig } => {
            ItemDataRecord { sig: *sig, items: items.clone() }.write(&mut out)
        }
        FrxValue::PropertyPages(pages) => {
            PropertyPagesRecord { pages: pages.iter().map(|s| encode_ansi(s)).collect() }
                .write(&mut out)
        }
        FrxValue::OcxBag { data, .. } => OcxBagRecord { clsid: None, data }.write(&mut out),
        // A live-decoded bag is a semantic view, not a byte form; it is not round-tripped.
        FrxValue::DecodedBag { .. } => {}
    }
    out
}

/// Detect an image container from its leading bytes.
pub fn detect_image_format(data: &[u8]) -> ImageFormat {
    if data.len() < 4 {
        return ImageFormat::Unknown;
    }
    if data.starts_with(&[0x42, 0x4D]) {
        ImageFormat::Bmp
    } else if data.starts_with(&[0x00, 0x00, 0x01, 0x00]) {
        ImageFormat::Ico
    } else if data.starts_with(&[0x00, 0x00, 0x02, 0x00]) {
        ImageFormat::Cur
    } else if data.starts_with(&[0x47, 0x49, 0x46, 0x38]) {
        ImageFormat::Gif
    } else if data.starts_with(&[0xFF, 0xD8, 0xFF]) {
        ImageFormat::Jpeg
    } else if data.starts_with(&[0x89, 0x50, 0x4E, 0x47]) {
        ImageFormat::Png
    } else if data.starts_with(&[0xD7, 0xCD, 0xC6, 0x9A])
        || data.starts_with(&[0x01, 0x00, 0x09, 0x00])
    {
        ImageFormat::Wmf
    } else if data.len() >= 44 && &data[40..44] == b" EMF" {
        ImageFormat::Emf
    } else {
        ImageFormat::Unknown
    }
}

/// Decode bytes that may be ANSI (Windows-1252) or UTF-16LE.
fn decode_text(bytes: &[u8]) -> String {
    if is_utf16le(bytes) {
        let units: Vec<u16> =
            bytes.chunks_exact(2).map(|c| u16::from_le_bytes([c[0], c[1]])).collect();
        String::from_utf16_lossy(&units)
    } else {
        decode_ansi(bytes)
    }
}

/// Heuristic: even length and a high proportion of zero high-bytes => UTF-16LE.
fn is_utf16le(bytes: &[u8]) -> bool {
    if bytes.len() < 4 || !bytes.len().is_multiple_of(2) {
        return false;
    }
    let pairs = bytes.len() / 2;
    let zero_hi = bytes.chunks_exact(2).filter(|c| c[1] == 0).count();
    zero_hi * 5 >= pairs * 4 // >= 80% of code units in U+0000..U+00FF
}

/// Decode Windows-1252 bytes to a `String`.
fn decode_ansi(bytes: &[u8]) -> String {
    WINDOWS_1252.decode_without_bom_handling(bytes).0.into_owned()
}

/// Inverse of [`decode_ansi`]; a character with no Windows-1252 byte becomes `?`.
fn encode_ansi(s: &str) -> Vec<u8> {
    s.chars()
        .map(|c| {
            let mut buf = [0u8; 4];
            let (bytes, _, unmappable) = WINDOWS_1252.encode(c.encode_utf8(&mut buf));
            if unmappable {
                b'?'
            } else {
                bytes[0]
            }
        })
        .collect()
}
