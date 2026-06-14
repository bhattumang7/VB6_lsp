//! Designer-file (`.frm`/`.ctl`/…) support for the session: form-control
//! symbols and FRX-reference hover decoding.
//!
//! VB6 designer files are two documents in one: a `Begin…End` designer section
//! (parsed by [`crate::frm`]) followed by VB6 code. This module surfaces the
//! designer side to the LSP layer:
//!   * [`form_controls`] flattens the control tree into a symbol list, and
//!   * [`describe_frx_reference`] decodes a `"form.frx":NNNN` property value into
//!     a human-readable hover string using the [`crate::frx`] deserializers.

use crate::frm::parse_frm;
use crate::frm::parser::BeginBlock;
use crate::frontend::ast::Span;
use crate::frx::records::{FrxRecord, RecordKind};
use crate::frx::{parse_frx_reference, kind_for_property, FrxReader, PropKind};

/// A control declared in a designer file.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct FormControl {
    /// Instance name, e.g. `cmdOK`.
    pub name: String,
    /// Control type, e.g. `VB.CommandButton`.
    pub control_type: String,
    /// Nesting depth (0 = the form/usercontrol itself).
    pub depth: u32,
    /// Byte span of the control name in the source (best-effort).
    pub span: Span,
}

/// Parse a designer file's source and return its controls (the root form and
/// every nested control), depth-first.
///
/// Returns an empty list for `.cls`-style files with no designer block or on a
/// designer parse error (the code side is still handled by the normal binder).
pub fn form_controls(src: &str) -> Vec<FormControl> {
    let Ok(frm) = parse_frm(src) else {
        return Vec::new();
    };
    let mut out = Vec::new();
    if let Some(root) = frm.root.as_ref() {
        walk(root, 0, src, &mut out);
    }
    out
}

fn walk(block: &BeginBlock, depth: u32, src: &str, out: &mut Vec<FormControl>) {
    out.push(FormControl {
        name: block.name.clone(),
        control_type: block.control_type.clone(),
        depth,
        span: name_span(src, &block.control_type, &block.name),
    });
    for child in &block.children {
        walk(child, depth + 1, src, out);
    }
}

/// Best-effort byte span of a control's name on its `Begin <type> <name>` line.
fn name_span(src: &str, control_type: &str, name: &str) -> Span {
    // Prefer the `<type> <name>` pairing to disambiguate; fall back to a bare
    // search for the name.
    let pat = format!("{control_type} {name}");
    let start = if let Some(p) = src.find(&pat) {
        p + control_type.len() + 1
    } else if let Some(p) = src.find(name) {
        p
    } else {
        return Span::DUMMY;
    };
    Span { start: start as u32, len: name.len() as u32 }
}

/// Decode a designer FRX reference (`prop = "form.frx":NNNN`) into hover text,
/// given the property name, the property value string, and the companion file's
/// bytes. Returns `None` if the value is not an FRX reference or cannot be read.
pub fn describe_frx_reference(prop_name: &str, value: &str, frx_bytes: &[u8]) -> Option<String> {
    let r = parse_frx_reference(value)?;
    let kind = kind_for_property(prop_name, r.dollar);
    let record_kind = record_kind_for(kind)?;
    let mut reader = FrxReader::new(frx_bytes);
    reader.seek(r.offset as usize).ok()?;
    let record = FrxRecord::read(record_kind, &mut reader).ok()?;
    Some(describe_record(&record))
}

/// Map the designer property kind to the FRX record deserializer to run.
fn record_kind_for(kind: PropKind) -> Option<RecordKind> {
    Some(match kind {
        PropKind::Picture => RecordKind::Picture,
        PropKind::Font => RecordKind::Font,
        PropKind::StringShort => RecordKind::StringShort,
        PropKind::StringLong => RecordKind::BinaryString,
        PropKind::List => RecordKind::ListItems,
        PropKind::ItemData => RecordKind::ItemData,
        PropKind::PropertyPages => RecordKind::PropertyPages,
        PropKind::OcxBag => RecordKind::OcxBag,
    })
}

fn describe_record(rec: &FrxRecord) -> String {
    match rec {
        FrxRecord::Picture(p) | FrxRecord::AsyncPicture(p) => {
            if p.data.is_empty() {
                "Picture: (empty)".to_string()
            } else {
                format!("Picture: {} ({} bytes)", image_format(p.data), p.data.len())
            }
        }
        FrxRecord::Font(f) => {
            let name = String::from_utf8_lossy(&f.name);
            let size = f.size_times_10k as f64 / 10_000.0;
            let mut styles = Vec::new();
            if f.is_bold() {
                styles.push("bold");
            }
            if f.is_italic() {
                styles.push("italic");
            }
            if f.is_underline() {
                styles.push("underline");
            }
            if f.is_strikethrough() {
                styles.push("strikethrough");
            }
            let suffix = if styles.is_empty() {
                String::new()
            } else {
                format!(", {}", styles.join(" "))
            };
            format!("Font: {name} {size}pt{suffix}")
        }
        FrxRecord::StringShort(s) => format!("Text: {:?}", lossy_preview(s.data)),
        FrxRecord::BinaryString(s) => format!("Text: {:?}", lossy_preview(s.data)),
        FrxRecord::OcxBag(_) => "Control property bag (opaque)".to_string(),
        FrxRecord::PropertyPages(_) => "Property pages".to_string(),
        FrxRecord::ListItems(_) => "List items".to_string(),
        FrxRecord::ItemData(_) => "Item data".to_string(),
        _ => "FRX resource".to_string(),
    }
}

/// Identify an image format from its leading magic bytes.
fn image_format(data: &[u8]) -> &'static str {
    match data {
        [0x42, 0x4D, ..] => "BMP",
        [0x00, 0x00, 0x01, 0x00, ..] => "ICO",
        [0x00, 0x00, 0x02, 0x00, ..] => "CUR",
        [0x47, 0x49, 0x46, 0x38, ..] => "GIF",
        [0xFF, 0xD8, 0xFF, ..] => "JPEG",
        [0x89, 0x50, 0x4E, 0x47, ..] => "PNG",
        [0xD7, 0xCD, 0xC6, 0x9A, ..] => "WMF",
        [0x01, 0x00, 0x00, 0x00, ..] => "EMF",
        _ => "image",
    }
}

/// A short, printable preview of CP-1252 text bytes.
fn lossy_preview(data: &[u8]) -> String {
    const MAX: usize = 64;
    let slice = &data[..data.len().min(MAX)];
    let mut s: String = slice.iter().map(|&b| b as char).collect();
    if data.len() > MAX {
        s.push('…');
    }
    s
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::frx::records::PictureRecord;
    use crate::frx::records::StdFontRecord;

    #[test]
    fn lists_form_controls() {
        let src = "VERSION 5.00\n\
                   Begin VB.Form Form1\n\
                   Caption = \"Hi\"\n\
                   Begin VB.CommandButton cmdOK\n\
                   End\n\
                   End\n\
                   Attribute VB_Name = \"Form1\"\n";
        let controls = form_controls(src);
        assert_eq!(controls.len(), 2);
        assert_eq!(controls[0].name, "Form1");
        assert_eq!(controls[0].control_type, "VB.Form");
        assert_eq!(controls[0].depth, 0);
        assert_eq!(controls[1].name, "cmdOK");
        assert_eq!(controls[1].depth, 1);
        // Span points at the control name in source.
        let s = controls[1].span;
        assert_eq!(&src[s.start as usize..(s.start + s.len) as usize], "cmdOK");
    }

    /// Build a minimal `.frx` containing a single record at offset 0.
    fn frx_with(record: &FrxRecord) -> Vec<u8> {
        let mut out = Vec::new();
        record.write(&mut out);
        out
    }

    #[test]
    fn describes_picture_reference() {
        // A 4-byte BMP-magic "image".
        let pic = FrxRecord::Picture(PictureRecord { clsid: None, data: &[0x42, 0x4D, 0x01, 0x02] });
        let frx = frx_with(&pic);
        let desc = describe_frx_reference("Icon", "\"Form1.frx\":0000", &frx).unwrap();
        assert_eq!(desc, "Picture: BMP (4 bytes)");
    }

    #[test]
    fn describes_font_reference() {
        let font = FrxRecord::Font(StdFontRecord {
            charset: 0,
            flags: 0x01, // italic
            weight: 700, // bold
            size_times_10k: 82_500,
            name: b"MS Sans Serif".to_vec(),
        });
        let frx = frx_with(&font);
        let desc = describe_frx_reference("Font", "\"Form1.frx\":0000", &frx).unwrap();
        assert_eq!(desc, "Font: MS Sans Serif 8.25pt, bold italic");
    }

    #[test]
    fn non_frx_value_is_none() {
        assert!(describe_frx_reference("Caption", "\"plain text\"", &[]).is_none());
    }
}
