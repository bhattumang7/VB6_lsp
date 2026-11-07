//! Structured export / IR for a VB6 form's design.
//!
//! Emits the full control tree + every property + every resolved companion
//! resource (images as base64, fonts/strings/lists decoded, opaque bags labelled
//! with their CLSID) plus the byte-accounting coverage report. This is the
//! consumable artifact a re-implementation in a new technology reads.
//!
//! View-structs derive `Serialize` here so the core decode types
//! ([`super::frx::FrxValue`], [`super::form_designer::DesignerControl`]) stay
//! serde-free and the JSON wire shape is controlled in one place.

use std::path::Path;

use base64::Engine as _;
use serde::Serialize;

use super::coverage::{self, CoverageReport};
use super::form_designer::{self, DesignerControl};
use super::frx::FrxValue;
use super::resources::{self, OcxBagPolicy, ResolveError, ResolvedResource};

/// Top-level form export.
#[derive(Debug, Serialize)]
pub struct FormExport {
    pub form_file: String,
    pub objects: Vec<ObjectEntryView>,
    pub root: Option<ControlView>,
    pub resources: Vec<ResourceView>,
    pub coverage: Vec<CoverageView>,
}

#[derive(Debug, Serialize)]
pub struct ObjectEntryView {
    pub clsid: String,
    pub version: String,
    pub file: String,
}

#[derive(Debug, Serialize)]
pub struct ControlView {
    pub type_name: String,
    pub name: String,
    pub library: Option<String>,
    pub is_intrinsic: bool,
    pub properties: Vec<PropView>,
    pub children: Vec<ControlView>,
}

#[derive(Debug, Serialize)]
pub struct PropView {
    pub name: String,
    pub value: String,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub frx: Option<FrxRefView>,
}

#[derive(Debug, Serialize)]
pub struct FrxRefView {
    pub file: String,
    pub offset: u32,
    pub dollar: bool,
}

#[derive(Debug, Serialize)]
pub struct ResourceView {
    pub control_path: String,
    pub control_type: String,
    pub property: String,
    pub kind: String,
    pub value: ResourceValueView,
}

/// The decoded value of a resource, or the error explaining why it couldn't decode.
#[derive(Debug, Serialize)]
#[serde(tag = "type")]
pub enum ResourceValueView {
    Picture { format: String, base64: String },
    Font {
        name: String,
        size_pt: f64,
        weight: u16,
        charset: u16,
        bold: bool,
        italic: bool,
        underline: bool,
        strikethrough: bool,
    },
    Text { value: String },
    List { items: Vec<String> },
    ItemData { items: Vec<i32> },
    PropertyPages { pages: Vec<String> },
    OcxBag { clsid: Option<String>, base64: String },
    DecodedBag { clsid: Option<String>, properties: Vec<(String, String)> },
    Empty,
    Error { kind: String, message: String },
}

#[derive(Debug, Serialize)]
pub struct CoverageView {
    pub file: String,
    pub total: usize,
    pub covered: usize,
    pub percent: f64,
    pub complete: bool,
    pub unexplained_bytes: usize,
    pub overlaps: usize,
    pub opaque_spans: usize,
    pub errors: usize,
}

/// Read a `.frm`/`.ctl`, parse + resolve + account, and assemble the export IR
/// (Tier-3 bags left opaque).
pub fn build_form_export(frm_path: &Path) -> std::io::Result<FormExport> {
    build_form_export_with(frm_path, OcxBagPolicy::Opaque, None)
}

/// Like [`build_form_export`] but with an explicit Tier-3 policy and optional COM
/// decoder (for live proprietary-bag decoding where a licensed control exists).
pub fn build_form_export_with(
    frm_path: &Path,
    policy: OcxBagPolicy,
    com: Option<&dyn resources::ComDecoder>,
) -> std::io::Result<FormExport> {
    let raw = std::fs::read(frm_path)?;
    let source = String::from_utf8_lossy(&raw);
    let designer = form_designer::parse_designer(&source);
    let resolved = resources::resolve_form_resources(frm_path, policy, com)?;
    let cov = coverage::coverage_for_form(frm_path)?;

    Ok(FormExport {
        form_file: frm_path
            .file_name()
            .map(|s| s.to_string_lossy().to_string())
            .unwrap_or_default(),
        objects: designer
            .objects
            .iter()
            .map(|o| ObjectEntryView {
                clsid: o.clsid.clone(),
                version: o.version.clone(),
                file: o.file.clone(),
            })
            .collect(),
        root: designer.root.as_ref().map(control_view),
        resources: resolved.iter().map(resource_view).collect(),
        coverage: cov.iter().map(coverage_view).collect(),
    })
}

fn control_view(c: &DesignerControl) -> ControlView {
    ControlView {
        type_name: c.type_name.clone(),
        name: c.name.clone(),
        library: c.library().map(|s| s.to_string()),
        is_intrinsic: c.is_intrinsic(),
        properties: c
            .properties
            .iter()
            .map(|p| PropView {
                name: p.name.clone(),
                value: p.value.clone(),
                frx: p.frx.as_ref().map(|f| FrxRefView {
                    file: f.file.clone(),
                    offset: f.offset,
                    dollar: f.dollar,
                }),
            })
            .collect(),
        children: c.children.iter().map(control_view).collect(),
    }
}

fn resource_view(r: &ResolvedResource) -> ResourceView {
    ResourceView {
        control_path: r.reference.control_path.clone(),
        control_type: r.reference.control_type.clone(),
        property: r.reference.property.clone(),
        kind: format!("{:?}", r.kind),
        value: match &r.value {
            Ok(v) => value_view(v),
            Err(e) => ResourceValueView::Error {
                kind: error_kind(e),
                message: e.to_string(),
            },
        },
    }
}

fn value_view(v: &FrxValue) -> ResourceValueView {
    match v {
        FrxValue::Picture { format, data, .. } => ResourceValueView::Picture {
            format: format!("{:?}", format),
            base64: base64::engine::general_purpose::STANDARD.encode(data),
        },
        FrxValue::Font(f) => ResourceValueView::Font {
            name: f.name.clone(),
            size_pt: f.size_pt,
            weight: f.weight,
            charset: f.charset,
            bold: f.bold,
            italic: f.italic,
            underline: f.underline,
            strikethrough: f.strikethrough,
        },
        FrxValue::Text(s) => ResourceValueView::Text { value: s.clone() },
        FrxValue::List { items, .. } => ResourceValueView::List {
            items: items.clone(),
        },
        FrxValue::ItemData { items, .. } => ResourceValueView::ItemData {
            items: items.iter().map(|b| super::frx::itemdata_value(b)).collect(),
        },
        FrxValue::PropertyPages(p) => ResourceValueView::PropertyPages { pages: p.clone() },
        FrxValue::OcxBag { clsid, data } => ResourceValueView::OcxBag {
            clsid: clsid.as_ref().map(|g| format_guid(*g)),
            base64: base64::engine::general_purpose::STANDARD.encode(data),
        },
        FrxValue::DecodedBag { clsid, properties } => ResourceValueView::DecodedBag {
            clsid: clsid.as_ref().map(|g| format_guid(*g)),
            properties: properties.clone(),
        },
        FrxValue::Empty => ResourceValueView::Empty,
    }
}

fn coverage_view(r: &CoverageReport) -> CoverageView {
    CoverageView {
        file: r.file.clone(),
        total: r.total_len,
        covered: r.covered,
        percent: r.coverage_pct(),
        complete: r.is_complete(),
        unexplained_bytes: r.unexplained_bytes(),
        overlaps: r.overlaps.len(),
        opaque_spans: r.opaque_spans,
        errors: r.errors.len(),
    }
}

fn error_kind(e: &ResolveError) -> String {
    match e {
        ResolveError::MissingCompanion { .. } => "MissingCompanion",
        ResolveError::Read(_) => "Read",
        ResolveError::Decode(_) => "Decode",
        ResolveError::LicensedBagUnavailable { .. } => "LicensedBagUnavailable",
        ResolveError::NoComDecoder => "NoComDecoder",
        ResolveError::ComDecodeFailed { .. } => "ComDecodeFailed",
    }
    .to_string()
}

/// Format a 16-byte COM GUID (Data1/2/3 little-endian, Data4 big-endian).
pub fn format_guid(g: [u8; 16]) -> String {
    let d1 = u32::from_le_bytes([g[0], g[1], g[2], g[3]]);
    let d2 = u16::from_le_bytes([g[4], g[5]]);
    let d3 = u16::from_le_bytes([g[6], g[7]]);
    format!(
        "{{{:08X}-{:04X}-{:04X}-{:02X}{:02X}-{:02X}{:02X}{:02X}{:02X}{:02X}{:02X}}}",
        d1, d2, d3, g[8], g[9], g[10], g[11], g[12], g[13], g[14], g[15]
    )
}

#[cfg(test)]
mod tests {
    use super::*;
    use std::io::Write;

    fn write_temp(name: &str, bytes: &[u8]) -> std::path::PathBuf {
        let mut p = std::env::temp_dir();
        p.push(format!("frxexp_{}_{}", std::process::id(), name));
        std::fs::File::create(&p).unwrap().write_all(bytes).unwrap();
        p
    }

    #[test]
    fn builds_and_serializes_full_export() {
        // companion: one [u32 12][lt\0\0][u32 4][BM..] picture, fully covered.
        let mut frx = Vec::new();
        frx.extend_from_slice(&12u32.to_le_bytes());
        frx.extend_from_slice(b"lt\0\0");
        frx.extend_from_slice(&4u32.to_le_bytes());
        frx.extend_from_slice(&[0x42, 0x4D, 0x00, 0x00]);
        let frx_path = write_temp("exp.frx", &frx);
        let frx_name = frx_path.file_name().unwrap().to_string_lossy().to_string();
        let frm = format!("Begin VB.Form frmX\n   Icon = \"{}\":0000\nEnd\n", frx_name);
        let frm_path = write_temp("exp.frm", frm.as_bytes());

        let exp = build_form_export(&frm_path).unwrap();
        let root = exp.root.as_ref().expect("root control");
        assert_eq!(root.type_name, "VB.Form");
        assert!(root.is_intrinsic);
        // exactly one resolved resource, decoded as a BMP picture
        assert_eq!(exp.resources.len(), 1);
        assert_eq!(exp.resources[0].property, "Icon");
        assert!(matches!(
            exp.resources[0].value,
            ResourceValueView::Picture { .. }
        ));
        // coverage proves the companion is fully accounted for
        assert_eq!(exp.coverage.len(), 1);
        assert!(exp.coverage[0].complete);
        assert_eq!(exp.coverage[0].covered, exp.coverage[0].total);

        // serialises to JSON with the tagged resource value
        let json = serde_json::to_string(&exp).unwrap();
        assert!(json.contains("\"type\":\"Picture\""));
        assert!(json.contains("\"format\":\"Bmp\""));

        let _ = std::fs::remove_file(frx_path);
        let _ = std::fs::remove_file(frm_path);
    }
}
