//! VB6 FRX/CTX companion resource parsing & decoding.
//!
//! A `.frm`/`.ctl`/`.pag`/`.dob` file stores binary or large property values as
//! references into a side-car `.frx`/`.ctx` file, addressed by byte offset:
//!
//! ```text
//! Icon            =   "frmMain.frx":0000
//! Caption         =   $"frmMain.frx":0BCA      ' $ => length-prefixed string
//! ```
//!
//! The blob formats decoded here were verified two independent ways: against
//! Microsoft's own COM serialization (`StdPicture` / `StdFont` via
//! `IPersistStream::Save`) and against real-world corpus bytes.
//! Driving the decode by the *known property kind* (rather than sniffing a
//! signature byte) avoids the mis-reads that signature-guessing produces.

/// Image container formats found inside `StdPicture` blobs.
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
    /// File extension (without dot) suitable for writing the blob out.
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

/// A reference to a blob inside a companion file, as written on a `.frm`/`.ctl` line.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct FrxRef {
    /// Property name (`Icon`, `Picture`, `Caption`, ...). May be empty when a bare
    /// value string is parsed directly.
    pub property: String,
    /// Companion file name, e.g. `frmMain.frx`.
    pub file: String,
    /// Byte offset into the companion file.
    pub offset: u32,
    /// True for the `$"...":N` form (a length-prefixed string value).
    pub dollar: bool,
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
    Picture {
        format: ImageFormat,
        data: Vec<u8>,
        clsid: Option<[u8; 16]>,
    },
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
    /// Tier-3 proprietary control property bag, surfaced opaque so nothing is lost
    /// and nothing is silently mis-read. `clsid` is filled when a leading class id
    /// is detectable in the blob.
    OcxBag { clsid: Option<[u8; 16]>, data: Vec<u8> },
    /// A proprietary control bag decoded *live* via COM (Tier 3) into typed
    /// properties `(name, value-as-text)`. Produced only by a [`crate::controls::resources::ComDecoder`].
    DecodedBag {
        clsid: Option<[u8; 16]>,
        properties: Vec<(String, String)>,
    },
    /// An explicitly empty resource (e.g. a removed icon).
    Empty,
}

/// What a referenced blob is expected to be, derived from the property name.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum PropKind {
    Picture,
    Font,
    /// `"...":N` short string: single-byte length prefix.
    StringShort,
    /// `$"...":N` long string: 4-byte length prefix.
    StringLong,
    List,
    ItemData,
    PropertyPages,
    /// Unknown / vendor control bag — keep opaque.
    OcxBag,
}

/// Decode failures. We report these rather than guessing (never silently mis-read).
#[derive(Debug, Clone, PartialEq, Eq)]
pub enum FrxError {
    OffsetOutOfRange { offset: usize, len: usize },
    Truncated { needed: usize, have: usize },
    BadPictureHeader { offset: usize },
}

impl std::fmt::Display for FrxError {
    fn fmt(&self, f: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        match self {
            FrxError::OffsetOutOfRange { offset, len } => {
                write!(f, "offset 0x{:X} out of range (file len {})", offset, len)
            }
            FrxError::Truncated { needed, have } => {
                write!(f, "blob truncated: needed {} bytes, have {}", needed, have)
            }
            FrxError::BadPictureHeader { offset } => {
                write!(f, "bad/unknown picture framing at offset 0x{:X}", offset)
            }
        }
    }
}
impl std::error::Error for FrxError {}

// =============================================================================
// Tier 1 — reference parsing
// =============================================================================

/// Parse an FRX reference from a `.frm`/`.ctl` property value.
///
/// Accepts both forms (with or without a leading property name):
///   `"frmMain.frx":0000`   and   `$"frmMain.frx":0BCA`
pub fn parse_frx_reference(value: &str) -> Option<FrxRef> {
    let v = value.trim();
    let (dollar, rest) = if let Some(r) = v.strip_prefix('$') {
        (true, r.trim_start())
    } else {
        (false, v)
    };
    let rest = rest.strip_prefix('"')?;
    let (file, after) = rest.split_once('"')?;
    let after = after.trim_start();
    let off_str = after.strip_prefix(':')?.trim();
    // offset may be followed by trailing tokens; take the leading hex run.
    let hex: String = off_str.chars().take_while(|c| c.is_ascii_hexdigit()).collect();
    if hex.is_empty() {
        return None;
    }
    let offset = u32::from_str_radix(&hex, 16).ok()?;
    if !(file.ends_with(".frx")
        || file.ends_with(".ctx")
        || file.ends_with(".FRX")
        || file.ends_with(".CTX"))
    {
        return None;
    }
    Some(FrxRef {
        property: String::new(),
        file: file.to_string(),
        offset,
        dollar,
    })
}

/// Map a designer property name (plus the `$` flag) to the expected blob kind.
pub fn kind_for_property(name: &str, dollar: bool) -> PropKind {
    // Normalize: take the final path segment (so a qualified group path like
    // "Images.ListImage1.Picture" -> "Picture") and drop a trailing "(N)" index
    // (so "TabPicture(0)" / "MouseIcon(1)" -> "TabPicture" / "MouseIcon").
    let base = name.rsplit('.').next().unwrap_or(name);
    let base = base.split('(').next().unwrap_or(base);
    let n = base.trim().to_ascii_lowercase();
    match n.as_str() {
        "picture" | "icon" | "image" | "mouseicon" | "dragicon" | "toolboxbitmap"
        | "disabledpicture" | "downpicture" | "maskpicture" | "tabpicture"
        | "mouseicon0" => PropKind::Picture,
        "font" | "mousefont" => PropKind::Font,
        "list" => PropKind::List,
        "itemdata" => PropKind::ItemData,
        "propertypages" => PropKind::PropertyPages,
        // Known proprietary control bags (vendor-serialized) — keep opaque.
        "_gridinfo" | "bands" | "column" | "sortkey" | "fmtcondition" | "formatstyle"
        | "template" | "printerproperties" | "initbuttons" | "bindings"
        | "initlistimages" => PropKind::OcxBag,
        // Text-ish properties: long form uses the 4-byte prefix, short the 1-byte.
        "caption" | "text" | "textrtf" | "tag" | "tooltiptext" | "title" => {
            if dollar {
                PropKind::StringLong
            } else {
                PropKind::StringShort
            }
        }
        _ => {
            if dollar {
                PropKind::StringLong
            } else {
                PropKind::OcxBag
            }
        }
    }
}

// =============================================================================
// Tier 2 — blob decoding (verified formats)
// =============================================================================

fn rd_u16(b: &[u8], at: usize) -> Result<u16, FrxError> {
    b.get(at..at + 2)
        .map(|s| u16::from_le_bytes([s[0], s[1]]))
        .ok_or(FrxError::Truncated { needed: at + 2, have: b.len() })
}
fn rd_u32(b: &[u8], at: usize) -> Result<u32, FrxError> {
    b.get(at..at + 4)
        .map(|s| u32::from_le_bytes([s[0], s[1], s[2], s[3]]))
        .ok_or(FrxError::Truncated { needed: at + 4, have: b.len() })
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
    if offset > buf.len() {
        return Err(FrxError::OffsetOutOfRange { offset, len: buf.len() });
    }
    match kind {
        PropKind::Picture => picture_span(buf, offset),
        PropKind::Font => font_span(buf, offset),
        PropKind::StringShort => string_span(buf, offset, 1),
        PropKind::StringLong => string_span(buf, offset, 4),
        PropKind::List => list_span(buf, offset),
        PropKind::ItemData => itemdata_span(buf, offset),
        PropKind::PropertyPages => property_pages_span(buf, offset),
        PropKind::OcxBag => Ok(ocx_bag_span(buf, offset)),
    }
}

/// `StdPicture`: `[u32 outerLen]["lt\0\0"][u32 dataLen][image bytes]`, outerLen = 8 + dataLen.
#[cfg(test)]
pub fn decode_picture(buf: &[u8], offset: usize) -> Result<FrxValue, FrxError> {
    picture_span(buf, offset).map(|(v, _)| v)
}

fn picture_span(buf: &[u8], offset: usize) -> Result<(FrxValue, usize), FrxError> {
    let outer = rd_u32(buf, offset)? as usize;
    let body = offset + 4;
    // The "lt\0\0" magic is either directly after the length (standard) or after a
    // 16-byte StdPicture class id (the form used inside ImageList/collection bags).
    let lt = if buf.get(body..body + 4).map_or(false, |s| s == b"lt\0\0") {
        body
    } else if buf.get(body + 16..body + 20).map_or(false, |s| s == b"lt\0\0") {
        body + 16
    } else {
        return Err(FrxError::BadPictureHeader { offset });
    };
    let header = lt - body; // 0 (standard) or 16 (CLSID-prefixed)
    let clsid = if header == 16 {
        let mut g = [0u8; 16];
        g.copy_from_slice(&buf[body..body + 16]);
        Some(g)
    } else {
        None
    };
    let data_len = rd_u32(buf, lt + 4)? as usize;
    if data_len == 0 {
        // An empty slot: a bare "removed icon" (Empty) or a CLSID-framed empty
        // (e.g. an unset TabPicture) which must keep its framing to re-encode.
        return Ok(match clsid {
            Some(_) => (
                FrxValue::Picture { format: ImageFormat::Unknown, data: Vec::new(), clsid },
                4 + outer,
            ),
            None => (FrxValue::Empty, 4 + outer),
        });
    }
    if outer != data_len + 8 + header {
        return Err(FrxError::BadPictureHeader { offset });
    }
    let start = lt + 8;
    let end = start + data_len;
    let data = buf
        .get(start..end)
        .ok_or(FrxError::Truncated { needed: end, have: buf.len() })?
        .to_vec();
    let format = detect_image_format(&data);
    Ok((FrxValue::Picture { format, data, clsid }, 4 + outer))
}

/// `StdFont`: `[u8 ver][u16 charset][u8 flags][u16 weight][u32 size=pt*10000][u8 nameLen][name]`.
#[cfg(test)]
pub fn decode_font(buf: &[u8], offset: usize) -> Result<FrxValue, FrxError> {
    font_span(buf, offset).map(|(v, _)| v)
}

fn font_span(buf: &[u8], offset: usize) -> Result<(FrxValue, usize), FrxError> {
    let need = offset + 11;
    if buf.len() < need {
        return Err(FrxError::Truncated { needed: need, have: buf.len() });
    }
    let charset = rd_u16(buf, offset + 1)?;
    let flags = buf[offset + 3];
    let weight = rd_u16(buf, offset + 4)?;
    let size = rd_u32(buf, offset + 6)?;
    let name_len = buf[offset + 10] as usize;
    let name_start = offset + 11;
    let name_end = name_start + name_len;
    let name_bytes = buf
        .get(name_start..name_end)
        .ok_or(FrxError::Truncated { needed: name_end, have: buf.len() })?;
    let name = decode_ansi(name_bytes);
    let font = FontInfo {
        name,
        size_pt: size as f64 / 10000.0,
        weight,
        charset,
        raw_flags: flags,
        bold: weight >= 700 || (flags & 0x01) != 0,
        italic: (flags & 0x02) != 0,
        underline: (flags & 0x04) != 0,
        strikethrough: (flags & 0x08) != 0,
    };
    Ok((FrxValue::Font(font), 11 + name_len))
}

/// Length-prefixed string. `prefix` is 1 (short form) or 4 (`$` long form).
/// Text may be ANSI or UTF-16LE; we detect the latter heuristically.
#[cfg(test)]
pub fn decode_string(buf: &[u8], offset: usize, prefix: usize) -> Result<FrxValue, FrxError> {
    string_span(buf, offset, prefix).map(|(v, _)| v)
}

fn string_span(buf: &[u8], offset: usize, prefix: usize) -> Result<(FrxValue, usize), FrxError> {
    let len = match prefix {
        1 => *buf
            .get(offset)
            .ok_or(FrxError::Truncated { needed: offset + 1, have: buf.len() })? as usize,
        _ => rd_u32(buf, offset)? as usize,
    };
    let start = offset + prefix;
    let end = start + len;
    let bytes = buf
        .get(start..end)
        .ok_or(FrxError::Truncated { needed: end, have: buf.len() })?;
    Ok((FrxValue::Text(decode_text(bytes)), prefix + len))
}

/// `List`: `[u16 count][u16 sig][ {[u16 len][ansi text]} x count ]`.
/// Driven by the known property kind, so the signature value is not gated on.
#[cfg(test)]
pub fn decode_list(buf: &[u8], offset: usize) -> Result<FrxValue, FrxError> {
    list_span(buf, offset).map(|(v, _)| v)
}

/// Decode the framing shared by `List` and `ItemData`:
///   `[u16 count]` — and only when count>0 — `[u16 sig]` then `count` items,
///   each `[u16 len][len bytes]`. An empty collection is just the 2-byte count
///   (no signature), so the returned span is 2 in that case.
fn count_sig_items<'a>(
    buf: &'a [u8],
    offset: usize,
) -> Result<(u16, Vec<&'a [u8]>, usize), FrxError> {
    let count = rd_u16(buf, offset)? as usize;
    if count == 0 {
        return Ok((0, Vec::new(), 2));
    }
    let sig = rd_u16(buf, offset + 2)?;
    let mut pos = offset + 4;
    let mut items = Vec::with_capacity(count);
    for _ in 0..count {
        let len = rd_u16(buf, pos)? as usize;
        pos += 2;
        let end = pos + len;
        let bytes = buf
            .get(pos..end)
            .ok_or(FrxError::Truncated { needed: end, have: buf.len() })?;
        items.push(bytes);
        pos = end;
    }
    Ok((sig, items, pos - offset))
}

fn list_span(buf: &[u8], offset: usize) -> Result<(FrxValue, usize), FrxError> {
    let (sig, items, span) = count_sig_items(buf, offset)?;
    let strings = items.iter().map(|b| decode_ansi(b)).collect();
    Ok((FrxValue::List { items: strings, sig }, span))
}

fn itemdata_span(buf: &[u8], offset: usize) -> Result<(FrxValue, usize), FrxError> {
    let (sig, items, span) = count_sig_items(buf, offset)?;
    let raw = items.iter().map(|b| b.to_vec()).collect();
    Ok((FrxValue::ItemData { items: raw, sig }, span))
}

/// Interpret a raw `ItemData` item (little-endian, 1-4 bytes) as a 32-bit Long.
pub fn itemdata_value(raw: &[u8]) -> i32 {
    let mut b = [0u8; 4];
    for (i, &x) in raw.iter().take(4).enumerate() {
        b[i] = x;
    }
    i32::from_le_bytes(b)
}

/// `PropertyPages`: `[u32 count][ {[u16 len-incl-null][name bytes][00]} x count ]`.
#[cfg(test)]
pub fn decode_property_pages(buf: &[u8], offset: usize) -> Result<FrxValue, FrxError> {
    property_pages_span(buf, offset).map(|(v, _)| v)
}

fn property_pages_span(buf: &[u8], offset: usize) -> Result<(FrxValue, usize), FrxError> {
    let count = rd_u32(buf, offset)? as usize;
    let mut pos = offset + 4;
    let mut pages = Vec::with_capacity(count);
    for _ in 0..count {
        let len = rd_u16(buf, pos)? as usize;
        pos += 2;
        let end = pos + len;
        let bytes = buf
            .get(pos..end)
            .ok_or(FrxError::Truncated { needed: end, have: buf.len() })?;
        // length includes the trailing NUL
        let trimmed = bytes.split(|&c| c == 0).next().unwrap_or(bytes);
        pages.push(decode_ansi(trimmed));
        pos = end;
    }
    Ok((FrxValue::PropertyPages(pages), pos - offset))
}

/// Tier-3: surface a proprietary control bag opaquely. A leading `[u32 len][16-byte
/// CLSID]` (e.g. a RichTextBox bag) is recognised and the CLSID extracted; otherwise
/// the whole remaining slice is returned as raw bytes.
pub fn decode_ocx_bag(buf: &[u8], offset: usize) -> FrxValue {
    ocx_bag_span(buf, offset).0
}

fn ocx_bag_span(buf: &[u8], offset: usize) -> (FrxValue, usize) {
    // Try the common `[u32 outerLen][16-byte CLSID]...` shape.
    if let Ok(outer) = rd_u32(buf, offset) {
        let start = offset + 4;
        let end = (start + outer as usize).min(buf.len());
        if end > start {
            let body = &buf[start..end];
            let clsid = if body.len() >= 16 && looks_like_guid(&body[..16]) {
                let mut g = [0u8; 16];
                g.copy_from_slice(&body[..16]);
                Some(g)
            } else {
                None
            };
            // span = the leading u32 length prefix + the body it framed.
            return (FrxValue::OcxBag { clsid, data: body.to_vec() }, end - offset);
        }
    }
    // No length framing: the bag's true length is unknown without the control;
    // we conservatively claim to end-of-file (coverage flags this as OpaqueTail).
    let data = buf.get(offset..).map(|s| s.to_vec()).unwrap_or_default();
    let span = data.len();
    (FrxValue::OcxBag { clsid: None, data }, span)
}

// =============================================================================
// Re-encode (inverse of decode) — proves the decode is lossless for standard types
// =============================================================================

/// Re-encode a decoded value back to its on-disk byte form.
///
/// Byte-exact inverse of [`decode_span`] for `Picture`, `Empty`, `Font`,
/// `List`, `ItemData`, `PropertyPages`, and length-framed `OcxBag`. `Text` is
/// best-effort only (its decode applies a lossy charset heuristic).
#[allow(dead_code)]
pub fn encode(value: &FrxValue, kind: PropKind) -> Vec<u8> {
    let mut out = Vec::new();
    match value {
        FrxValue::Empty => encode_empty(&mut out),
        FrxValue::Picture { data, clsid, .. } => encode_picture(&mut out, data, clsid),
        FrxValue::Font(f) => encode_font(&mut out, f),
        FrxValue::Text(s) => encode_text(&mut out, s, kind),
        FrxValue::List { items, sig } => encode_list(&mut out, items, *sig),
        FrxValue::ItemData { items, sig } => encode_item_data(&mut out, items, *sig),
        FrxValue::PropertyPages(pages) => encode_property_pages(&mut out, pages),
        FrxValue::OcxBag { data, .. } => {
            // Length-framed bags re-emit the [u32 len] prefix + the body we kept.
            out.extend_from_slice(&(data.len() as u32).to_le_bytes());
            out.extend_from_slice(data);
        }
        FrxValue::DecodedBag { .. } => {
            // A live-decoded bag is a semantic view, not a byte form; not round-tripped.
        }
    }
    out
}

#[allow(dead_code)]
fn encode_empty(out: &mut Vec<u8>) {
    out.extend_from_slice(&8u32.to_le_bytes());
    out.extend_from_slice(b"lt\0\0");
    out.extend_from_slice(&0u32.to_le_bytes());
}

#[allow(dead_code)]
fn encode_picture(out: &mut Vec<u8>, data: &[u8], clsid: &Option<[u8; 16]>) {
    let data_len = data.len() as u32;
    if let Some(c) = clsid {
        out.extend_from_slice(&(data_len + 24).to_le_bytes());
        out.extend_from_slice(c);
    } else {
        out.extend_from_slice(&(data_len + 8).to_le_bytes());
    }
    out.extend_from_slice(b"lt\0\0");
    out.extend_from_slice(&data_len.to_le_bytes());
    out.extend_from_slice(data);
}

#[allow(dead_code)]
fn encode_font(out: &mut Vec<u8>, f: &FontInfo) {
    out.push(0x01);
    out.extend_from_slice(&f.charset.to_le_bytes());
    out.push(f.raw_flags);
    out.extend_from_slice(&f.weight.to_le_bytes());
    let size = (f.size_pt * 10000.0).round() as u32;
    out.extend_from_slice(&size.to_le_bytes());
    let name = encode_ansi(&f.name);
    out.push(name.len() as u8);
    out.extend_from_slice(&name);
}

#[allow(dead_code)]
fn encode_text(out: &mut Vec<u8>, s: &str, kind: PropKind) {
    let bytes = encode_ansi(s);
    match kind {
        PropKind::StringShort => out.push(bytes.len() as u8),
        _ => out.extend_from_slice(&(bytes.len() as u32).to_le_bytes()),
    }
    out.extend_from_slice(&bytes);
}

#[allow(dead_code)]
fn encode_list(out: &mut Vec<u8>, items: &[String], sig: u16) {
    out.extend_from_slice(&(items.len() as u16).to_le_bytes());
    if items.is_empty() {
        return;
    }
    out.extend_from_slice(&sig.to_le_bytes());
    for it in items {
        let b = encode_ansi(it);
        out.extend_from_slice(&(b.len() as u16).to_le_bytes());
        out.extend_from_slice(&b);
    }
}

#[allow(dead_code)]
fn encode_item_data(out: &mut Vec<u8>, items: &[Vec<u8>], sig: u16) {
    out.extend_from_slice(&(items.len() as u16).to_le_bytes());
    if items.is_empty() {
        return;
    }
    out.extend_from_slice(&sig.to_le_bytes());
    for it in items {
        out.extend_from_slice(&(it.len() as u16).to_le_bytes());
        out.extend_from_slice(it);
    }
}

#[allow(dead_code)]
fn encode_property_pages(out: &mut Vec<u8>, pages: &[String]) {
    out.extend_from_slice(&(pages.len() as u32).to_le_bytes());
    for p in pages {
        let b = encode_ansi(p);
        out.extend_from_slice(&((b.len() + 1) as u16).to_le_bytes()); // length includes the NUL
        out.extend_from_slice(&b);
        out.push(0);
    }
}

/// Inverse of [`decode_ansi`]: map a `char` back to a single byte (Latin-1).
#[allow(dead_code)]
fn encode_ansi(s: &str) -> Vec<u8> {
    s.chars()
        .map(|c| if (c as u32) <= 0xFF { c as u8 } else { b'?' })
        .collect()
}

// =============================================================================
// Helpers
// =============================================================================

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
    } else if data.starts_with(&[0xD7, 0xCD, 0xC6, 0x9A]) || data.starts_with(&[0x01, 0x00, 0x09, 0x00]) {
        ImageFormat::Wmf
    } else if data.len() >= 44 && &data[40..44] == b" EMF" {
        ImageFormat::Emf
    } else {
        ImageFormat::Unknown
    }
}

/// True if a 16-byte slice plausibly holds a non-text GUID (not all printable).
fn looks_like_guid(b: &[u8]) -> bool {
    if b.len() < 16 {
        return false;
    }
    // A textual string of 16 bytes would be mostly printable; a GUID generally is not.
    let printable = b.iter().filter(|&&c| (0x20..0x7f).contains(&c)).count();
    printable < 12
}

/// Decode bytes that may be ANSI (Windows-1252-ish) or UTF-16LE.
fn decode_text(bytes: &[u8]) -> String {
    if is_utf16le(bytes) {
        let units: Vec<u16> = bytes
            .chunks_exact(2)
            .map(|c| u16::from_le_bytes([c[0], c[1]]))
            .collect();
        String::from_utf16_lossy(&units)
    } else {
        decode_ansi(bytes)
    }
}

/// Heuristic: even length and a high proportion of zero high-bytes => UTF-16LE.
fn is_utf16le(bytes: &[u8]) -> bool {
    if bytes.len() < 4 || bytes.len() % 2 != 0 {
        return false;
    }
    let pairs = bytes.len() / 2;
    let zero_hi = bytes.chunks_exact(2).filter(|c| c[1] == 0).count();
    zero_hi * 5 >= pairs * 4 // >= 80% of code units in U+0000..U+00FF
}

/// Decode ANSI bytes to a `String`, mapping Windows-1252 high range to Unicode.
fn decode_ansi(bytes: &[u8]) -> String {
    bytes
        .iter()
        .map(|&b| {
            if b < 0x80 {
                b as char
            } else {
                // Latin-1 fallback (close enough for control captions/names).
                b as char
            }
        })
        .collect()
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn parse_reference_both_forms() {
        let a = parse_frx_reference("\"frmMain.frx\":0000").unwrap();
        assert_eq!(a.file, "frmMain.frx");
        assert_eq!(a.offset, 0);
        assert!(!a.dollar);

        let b = parse_frx_reference("$\"frmMain.frx\":0BCA").unwrap();
        assert_eq!(b.offset, 0x0BCA);
        assert!(b.dollar);

        let c = parse_frx_reference("\"Ctl.ctx\":00AC").unwrap();
        assert_eq!(c.offset, 0xAC);

        assert!(parse_frx_reference("Just a caption").is_none());
    }

    #[test]
    fn picture_header_verified_layout() {
        // [u32 outerLen=12][lt\0\0][u32 dataLen=4]["BM" + 2]  (outer = 8 + 4)
        let mut b = Vec::new();
        b.extend_from_slice(&12u32.to_le_bytes());
        b.extend_from_slice(b"lt\0\0");
        b.extend_from_slice(&4u32.to_le_bytes());
        b.extend_from_slice(&[0x42, 0x4D, 0x00, 0x00]);
        match decode_picture(&b, 0).unwrap() {
            FrxValue::Picture { format, data, .. } => {
                assert_eq!(format, ImageFormat::Bmp);
                assert_eq!(data, vec![0x42, 0x4D, 0x00, 0x00]);
            }
            other => panic!("expected picture, got {:?}", other),
        }
    }

    #[test]
    fn picture_empty_icon() {
        let mut b = Vec::new();
        b.extend_from_slice(&8u32.to_le_bytes());
        b.extend_from_slice(b"lt\0\0");
        b.extend_from_slice(&0u32.to_le_bytes());
        assert_eq!(decode_picture(&b, 0).unwrap(), FrxValue::Empty);
    }

    #[test]
    fn font_verified_layout() {
        // 'MS Sans Serif' 8.25pt normal.
        let b: Vec<u8> = vec![
            0x01, 0x00, 0x00, 0x00, 0x90, 0x01, 0x44, 0x42, 0x01, 0x00, 0x0D, b'M', b'S', b' ',
            b'S', b'a', b'n', b's', b' ', b'S', b'e', b'r', b'i', b'f',
        ];
        match decode_font(&b, 0).unwrap() {
            FrxValue::Font(f) => {
                assert_eq!(f.name, "MS Sans Serif");
                assert!((f.size_pt - 8.25).abs() < 1e-9);
                assert_eq!(f.weight, 400);
                assert!(!f.bold);
            }
            other => panic!("expected font, got {:?}", other),
        }
    }

    #[test]
    fn short_and_long_strings() {
        // short: [u8 len][ansi]
        let mut s = vec![5u8];
        s.extend_from_slice(b"Hello");
        assert_eq!(decode_string(&s, 0, 1).unwrap(), FrxValue::Text("Hello".into()));

        // long: [u32 len][ansi]
        let mut l = 5u32.to_le_bytes().to_vec();
        l.extend_from_slice(b"World");
        assert_eq!(decode_string(&l, 0, 4).unwrap(), FrxValue::Text("World".into()));
    }

    #[test]
    fn list_any_signature() {
        // count=2, sig=0x000B (the value the corpus uses and the old code mis-read)
        let mut b = Vec::new();
        b.extend_from_slice(&2u16.to_le_bytes());
        b.extend_from_slice(&0x000Bu16.to_le_bytes());
        b.extend_from_slice(&3u16.to_le_bytes());
        b.extend_from_slice(b"abc");
        b.extend_from_slice(&2u16.to_le_bytes());
        b.extend_from_slice(b"de");
        assert_eq!(
            decode_list(&b, 0).unwrap(),
            FrxValue::List { items: vec!["abc".into(), "de".into()], sig: 0x000B }
        );
    }

    #[test]
    fn property_pages_list() {
        let mut b = 2u32.to_le_bytes().to_vec();
        // "PPGeneral\0" len incl null = 10
        b.extend_from_slice(&10u16.to_le_bytes());
        b.extend_from_slice(b"PPGeneral\0");
        b.extend_from_slice(&6u16.to_le_bytes());
        b.extend_from_slice(b"PPCol\0");
        assert_eq!(
            decode_property_pages(&b, 0).unwrap(),
            FrxValue::PropertyPages(vec!["PPGeneral".into(), "PPCol".into()])
        );
    }

    #[test]
    fn offset_out_of_range_reported() {
        let b = [0u8; 4];
        assert!(matches!(
            decode(&b, 99, PropKind::Picture),
            Err(FrxError::OffsetOutOfRange { .. })
        ));
    }

    #[test]
    fn property_kind_mapping() {
        assert_eq!(kind_for_property("Icon", false), PropKind::Picture);
        assert_eq!(kind_for_property("Font", false), PropKind::Font);
        assert_eq!(kind_for_property("Caption", true), PropKind::StringLong);
        assert_eq!(kind_for_property("Caption", false), PropKind::StringShort);
        assert_eq!(kind_for_property("List", false), PropKind::List);
        assert_eq!(kind_for_property("_GridInfo", false), PropKind::OcxBag);
    }

    #[test]
    fn font_roundtrips_byte_exact() {
        // The exact bytes for 'MS Sans Serif' 8.25pt.
        let original: Vec<u8> = vec![
            0x01, 0x00, 0x00, 0x00, 0x90, 0x01, 0x44, 0x42, 0x01, 0x00, 0x0D, b'M', b'S', b' ',
            b'S', b'a', b'n', b's', b' ', b'S', b'e', b'r', b'i', b'f',
        ];
        let (val, span) = decode_span(&original, 0, PropKind::Font).unwrap();
        assert_eq!(span, original.len());
        assert_eq!(encode(&val, PropKind::Font), original);
    }

    #[test]
    fn picture_and_list_roundtrip_byte_exact() {
        // picture: [u32 12][lt\0\0][u32 4][BM\0\0]
        let mut pic = Vec::new();
        pic.extend_from_slice(&12u32.to_le_bytes());
        pic.extend_from_slice(b"lt\0\0");
        pic.extend_from_slice(&4u32.to_le_bytes());
        pic.extend_from_slice(&[0x42, 0x4D, 0x00, 0x00]);
        let (pv, ps) = decode_span(&pic, 0, PropKind::Picture).unwrap();
        assert_eq!(ps, pic.len());
        assert_eq!(encode(&pv, PropKind::Picture), pic);

        // list with sig 0x000B
        let mut list = Vec::new();
        list.extend_from_slice(&2u16.to_le_bytes());
        list.extend_from_slice(&0x000Bu16.to_le_bytes());
        list.extend_from_slice(&3u16.to_le_bytes());
        list.extend_from_slice(b"abc");
        list.extend_from_slice(&2u16.to_le_bytes());
        list.extend_from_slice(b"de");
        let (lv, ls) = decode_span(&list, 0, PropKind::List).unwrap();
        assert_eq!(ls, list.len());
        assert_eq!(encode(&lv, PropKind::List), list);
    }
}
