//! Byte-accounting / coverage for VB6 companion files (`.frx`/`.ctx`).
//!
//! A companion file has no internal directory — blobs are addressed only by byte
//! offsets from the `.frm`/`.ctl`. This module attributes every byte of each
//! companion to a decoded reference and reports any unattributed range.
//!
//! Zero unexplained bytes (modulo trailing DWORD padding) **proves** we read the
//! whole file; an overlap (two references claiming the same bytes) **proves** a bug.

use std::collections::BTreeMap;
use std::path::Path;

use super::form_designer::{self, ResourceRef};
use super::frx::{self, PropKind};

/// A decoded byte range within one companion file.
#[derive(Debug, Clone)]
pub struct ByteSpan {
    pub start: usize,
    /// Exclusive end.
    pub end: usize,
    pub property: String,
    pub control_path: String,
    pub kind: PropKind,
    /// True when the length could not be proven (proprietary bag without framing).
    pub opaque: bool,
}

/// An unattributed range between or after decoded spans.
#[derive(Debug, Clone, PartialEq, Eq)]
pub enum Gap {
    /// Bytes nothing decoded — a missed reference or a wrong length (bug signal).
    Unexplained { start: usize, end: usize },
    /// A short all-zero run padding to a DWORD boundary (benign).
    TrailingPadding { start: usize, end: usize },
}

/// Coverage for a single companion file.
#[derive(Debug, Clone)]
pub struct CoverageReport {
    pub file: String,
    pub total_len: usize,
    pub covered: usize,
    pub spans: Vec<ByteSpan>,
    pub gaps: Vec<Gap>,
    /// Pairs of (sorted) span indices whose ranges overlap — always a bug.
    pub overlaps: Vec<(usize, usize)>,
    /// Count of spans whose length is opaque/unproven (Tier-3 bags).
    pub opaque_spans: usize,
    /// References that failed to decode: `(property@0xOFFSET, message)`.
    pub errors: Vec<(String, String)>,
}

impl CoverageReport {
    /// True when every non-padding byte is attributed, nothing overlaps, and
    /// nothing failed to decode. Opaque (Tier-3) spans count as covered but are
    /// surfaced separately via [`opaque_spans`](Self::opaque_spans).
    pub fn is_complete(&self) -> bool {
        self.overlaps.is_empty()
            && self.errors.is_empty()
            && !self.gaps.iter().any(|g| matches!(g, Gap::Unexplained { .. }))
    }

    pub fn coverage_pct(&self) -> f64 {
        if self.total_len == 0 {
            100.0
        } else {
            self.covered as f64 * 100.0 / self.total_len as f64
        }
    }

    pub fn unexplained_bytes(&self) -> usize {
        self.gaps
            .iter()
            .map(|g| match g {
                Gap::Unexplained { start, end } => end - start,
                Gap::TrailingPadding { .. } => 0,
            })
            .sum()
    }
}

/// Compute coverage for every companion referenced by a `.frm`/`.ctl`.
pub fn coverage_for_form(frm_path: &Path) -> std::io::Result<Vec<CoverageReport>> {
    let raw = std::fs::read(frm_path)?;
    let source = String::from_utf8_lossy(&raw);
    let designer = form_designer::parse_designer(&source);
    let dir = frm_path.parent().unwrap_or_else(|| Path::new("."));

    let mut by_file: BTreeMap<String, Vec<ResourceRef>> = BTreeMap::new();
    for r in designer.resource_refs() {
        by_file.entry(r.frx.file.clone()).or_default().push(r);
    }

    let mut reports = Vec::new();
    for (file, refs) in by_file {
        let companion = dir.join(&file);
        match std::fs::read(&companion) {
            Ok(bytes) => reports.push(build_report(file, &bytes, &refs)),
            Err(_) => {
                let errors = refs
                    .iter()
                    .map(|r| {
                        (
                            format!("{}@0x{:X}", r.property, r.frx.offset),
                            "missing companion".to_string(),
                        )
                    })
                    .collect();
                reports.push(CoverageReport {
                    file,
                    total_len: 0,
                    covered: 0,
                    spans: vec![],
                    gaps: vec![],
                    overlaps: vec![],
                    opaque_spans: 0,
                    errors,
                });
            }
        }
    }
    Ok(reports)
}

fn build_report(file: String, bytes: &[u8], refs: &[ResourceRef]) -> CoverageReport {
    let total_len = bytes.len();
    let mut spans: Vec<ByteSpan> = Vec::new();
    let mut errors: Vec<(String, String)> = Vec::new();

    for r in refs {
        let kind = frx::kind_for_property(&r.property, r.frx.dollar);
        let offset = r.frx.offset as usize;
        match frx::decode_span(bytes, offset, kind) {
            Ok((_, span)) => spans.push(ByteSpan {
                start: offset,
                end: (offset + span).min(total_len),
                property: r.property.clone(),
                control_path: r.control_path.clone(),
                kind,
                opaque: kind == PropKind::OcxBag,
            }),
            Err(e) => errors.push((format!("{}@0x{:X}", r.property, offset), e.to_string())),
        }
    }

    spans.sort_by_key(|s| s.start);

    // Opaque (Tier-3) bags have an unprovable length. Blobs are packed sequentially,
    // so clamp an opaque span to the next blob's start (or EOF) — it then neither
    // over-claims bytes nor produces a false overlap with the following reference.
    for i in 0..spans.len() {
        if spans[i].opaque {
            let next_start = spans.get(i + 1).map(|s| s.start).unwrap_or(total_len);
            if next_start > spans[i].start && next_start < spans[i].end {
                spans[i].end = next_start;
            }
            if spans[i].end > total_len {
                spans[i].end = total_len;
            }
        }
    }

    // Overlaps: adjacent spans where the later starts before the earlier ends.
    let mut overlaps = Vec::new();
    for i in 1..spans.len() {
        if spans[i].start < spans[i - 1].end {
            overlaps.push((i - 1, i));
        }
    }

    // Sweep for covered union length + unattributed gaps.
    let mut covered = 0usize;
    let mut gaps = Vec::new();
    let mut cursor = 0usize;
    for s in &spans {
        if s.start > cursor {
            classify_gap(bytes, cursor, s.start, &mut gaps);
            cursor = s.start;
        }
        if s.end > cursor {
            covered += s.end - cursor;
            cursor = s.end;
        }
    }
    if cursor < total_len {
        classify_gap(bytes, cursor, total_len, &mut gaps);
    }

    let opaque_spans = spans.iter().filter(|s| s.opaque).count();

    CoverageReport {
        file,
        total_len,
        covered,
        spans,
        gaps,
        overlaps,
        opaque_spans,
        errors,
    }
}

/// Classify an unattributed range as DWORD padding (short all-zero run) or a
/// genuinely unexplained gap.
fn classify_gap(bytes: &[u8], start: usize, end: usize, gaps: &mut Vec<Gap>) {
    let slice = &bytes[start..end.min(bytes.len())];
    if slice.iter().all(|&b| b == 0) && (end - start) < 4 {
        gaps.push(Gap::TrailingPadding { start, end });
    } else {
        gaps.push(Gap::Unexplained { start, end });
    }
}

#[cfg(test)]
mod tests {
    use super::*;
    use std::io::Write;

    fn write_temp(name: &str, bytes: &[u8]) -> std::path::PathBuf {
        let mut p = std::env::temp_dir();
        p.push(format!("frxcov_{}_{}", std::process::id(), name));
        std::fs::File::create(&p).unwrap().write_all(bytes).unwrap();
        p
    }

    #[test]
    fn full_coverage_of_a_single_picture() {
        // companion: [u32 12][lt\0\0][u32 4][BM\0\0] = 16 bytes, fully covered by one ref
        let mut frx = Vec::new();
        frx.extend_from_slice(&12u32.to_le_bytes());
        frx.extend_from_slice(b"lt\0\0");
        frx.extend_from_slice(&4u32.to_le_bytes());
        frx.extend_from_slice(&[0x42, 0x4D, 0x00, 0x00]);
        let frx_path = write_temp("cov.frx", &frx);
        let name = frx_path.file_name().unwrap().to_string_lossy().to_string();
        let frm = format!("Begin VB.Form f\n   Icon = \"{}\":0000\nEnd\n", name);
        let frm_path = write_temp("cov.frm", frm.as_bytes());

        let reports = coverage_for_form(&frm_path).unwrap();
        assert_eq!(reports.len(), 1);
        let r = &reports[0];
        assert_eq!(r.total_len, 16);
        assert_eq!(r.covered, 16);
        assert!(r.is_complete(), "expected complete, got {:?}", r.gaps);
        assert!(r.overlaps.is_empty());
        let _ = std::fs::remove_file(frx_path);
        let _ = std::fs::remove_file(frm_path);
    }

    #[test]
    fn unexplained_gap_is_flagged() {
        // 16-byte picture + 8 trailing non-zero bytes that no ref points at.
        let mut frx = Vec::new();
        frx.extend_from_slice(&12u32.to_le_bytes());
        frx.extend_from_slice(b"lt\0\0");
        frx.extend_from_slice(&4u32.to_le_bytes());
        frx.extend_from_slice(&[0x42, 0x4D, 0x00, 0x00]);
        frx.extend_from_slice(&[1, 2, 3, 4, 5, 6, 7, 8]); // unreferenced
        let frx_path = write_temp("gap.frx", &frx);
        let name = frx_path.file_name().unwrap().to_string_lossy().to_string();
        let frm = format!("Begin VB.Form f\n   Icon = \"{}\":0000\nEnd\n", name);
        let frm_path = write_temp("gap.frm", frm.as_bytes());

        let reports = coverage_for_form(&frm_path).unwrap();
        let r = &reports[0];
        assert!(!r.is_complete());
        assert_eq!(r.unexplained_bytes(), 8);
        let _ = std::fs::remove_file(frx_path);
        let _ = std::fs::remove_file(frm_path);
    }
}
