//! Byte-exact round-trip for VB6 companion blobs over the real corpus.
//!
//! For every standard-type reference: read the companion slice, `decode_span` it,
//! `encode` the value back, and assert the bytes are identical. A mismatch means
//! our model of that blob is lossy or wrong — a concrete bug. `Text` (lossy charset
//! heuristic) and unframed `OcxBag` are measured but excluded from the strict set.
//! Skips automatically when the corpus is absent.

use std::collections::HashMap;
use std::path::{Path, PathBuf};

use vb6_lsp::controls::form_designer;
use vb6_lsp::controls::frx::{self, PropKind};

fn collect_forms(dir: &Path, out: &mut Vec<PathBuf>) {
    if let Ok(entries) = std::fs::read_dir(dir) {
        for e in entries.flatten() {
            let p = e.path();
            if p.is_dir() {
                collect_forms(&p, out);
            } else if let Some(ext) = p.extension() {
                let ext = ext.to_string_lossy().to_ascii_lowercase();
                if ext == "frm" || ext == "ctl" {
                    out.push(p);
                }
            }
        }
    }
}

fn strict(kind: PropKind) -> bool {
    matches!(
        kind,
        PropKind::Picture
            | PropKind::Font
            | PropKind::List
            | PropKind::ItemData
            | PropKind::PropertyPages
    )
}

#[test]
fn corpus_roundtrip_byte_exact() {
    let corpus = Path::new(env!("CARGO_MANIFEST_DIR"))
        .join("re_lab")
        .join("corpus");
    if !corpus.exists() {
        eprintln!("SKIP: corpus not present at {}", corpus.display());
        return;
    }
    let mut forms = Vec::new();
    collect_forms(&corpus, &mut forms);

    // per-kind (pass, total)
    let mut tally: HashMap<String, (usize, usize)> = HashMap::new();
    let mut failures: Vec<String> = Vec::new();

    for form in &forms {
        let raw = match std::fs::read(form) {
            Ok(b) => b,
            Err(_) => continue,
        };
        let src = String::from_utf8_lossy(&raw);
        let designer = form_designer::parse_designer(&src);
        let dir = form.parent().unwrap_or_else(|| Path::new("."));
        let mut cache: HashMap<String, Option<Vec<u8>>> = HashMap::new();

        for r in designer.resource_refs() {
            let kind = frx::kind_for_property(&r.property, r.frx.dollar);
            let bytes = cache
                .entry(r.frx.file.clone())
                .or_insert_with(|| std::fs::read(dir.join(&r.frx.file)).ok());
            let bytes = match bytes {
                Some(b) => b,
                None => continue,
            };
            let off = r.frx.offset as usize;
            let (val, span) = match frx::decode_span(bytes, off, kind) {
                Ok(x) => x,
                Err(_) => continue, // decode failures are a D1/D3 concern, not round-trip
            };
            let end = (off + span).min(bytes.len());
            let original = &bytes[off..end];
            let re = frx::encode(&val, kind);
            let ok = re.as_slice() == original;

            let label = format!("{:?}", kind);
            let entry = tally.entry(label.clone()).or_insert((0, 0));
            entry.1 += 1;
            if ok {
                entry.0 += 1;
            } else if strict(kind) && failures.len() < 20 {
                failures.push(format!(
                    "{:?} {}@0x{:X} in {} ({} -> {} bytes)",
                    kind,
                    r.property,
                    off,
                    r.frx.file,
                    original.len(),
                    re.len()
                ));
            }
        }
    }

    eprintln!("=== FRX/CTX BYTE-EXACT ROUND-TRIP (re_lab/corpus) ===");
    let mut kinds: Vec<_> = tally.iter().collect();
    kinds.sort_by(|a, b| b.1 .1.cmp(&a.1 .1));
    for (k, (pass, total)) in kinds {
        eprintln!("  {:<14} {}/{} byte-exact", k, pass, total);
    }
    if !failures.is_empty() {
        eprintln!("--- strict-set mismatches ---");
        for f in &failures {
            eprintln!("  {}", f);
        }
    }

    // Strict-set blobs must re-encode byte-for-byte.
    let strict_fail: usize = failures.len();
    assert_eq!(
        strict_fail, 0,
        "{} strict-set blob(s) did not round-trip byte-exact",
        strict_fail
    );
}
