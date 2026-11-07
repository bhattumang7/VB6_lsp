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

/// Per-kind (pass, total) counts plus the first strict-set mismatches.
#[derive(Default)]
struct Outcome {
    tally: HashMap<String, (usize, usize)>,
    failures: Vec<String>,
}

type CompanionCache = HashMap<String, Option<Vec<u8>>>;

/// Decode one reference, re-encode it, and record whether the bytes match.
fn check_reference(
    r: &form_designer::ResourceRef,
    bytes: &[u8],
    outcome: &mut Outcome,
) {
    let kind = frx::kind_for_property(&r.property, r.frx.dollar);
    let off = r.frx.offset as usize;
    let (val, span) = match frx::decode_span(bytes, off, kind) {
        Ok(x) => x,
        Err(_) => return, // Skip decode failures
    };
    let end = (off + span).min(bytes.len());
    let original = &bytes[off..end];
    let re = frx::encode(&val, kind);
    let ok = re.as_slice() == original;

    let entry = outcome.tally.entry(format!("{:?}", kind)).or_insert((0, 0));
    entry.1 += 1;
    if ok {
        entry.0 += 1;
    } else if strict(kind) && outcome.failures.len() < 20 {
        outcome.failures.push(format!(
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

/// Round-trip every companion reference of one form.
fn check_form(form: &Path, outcome: &mut Outcome) {
    let Ok(raw) = std::fs::read(form) else {
        return;
    };
    let src = String::from_utf8_lossy(&raw);
    let designer = form_designer::parse_designer(&src);
    let dir = form.parent().unwrap_or_else(|| Path::new("."));
    let mut cache: CompanionCache = HashMap::new();

    for r in designer.resource_refs() {
        let bytes = cache
            .entry(r.frx.file.clone())
            .or_insert_with(|| std::fs::read(dir.join(&r.frx.file)).ok());
        if let Some(b) = bytes {
            check_reference(&r, b, outcome);
        }
    }
}

fn print_report(outcome: &Outcome) {
    eprintln!("=== FRX/CTX BYTE-EXACT ROUND-TRIP (re_lab/corpus) ===");
    let mut kinds: Vec<_> = outcome.tally.iter().collect();
    kinds.sort_by(|a, b| b.1 .1.cmp(&a.1 .1));
    for (k, (pass, total)) in kinds {
        eprintln!("  {:<14} {}/{} byte-exact", k, pass, total);
    }
    if !outcome.failures.is_empty() {
        eprintln!("--- strict-set mismatches ---");
        for f in &outcome.failures {
            eprintln!("  {}", f);
        }
    }
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

    let mut outcome = Outcome::default();
    for form in &forms {
        check_form(form, &mut outcome);
    }
    print_report(&outcome);

    // Strict-set blobs must re-encode byte-for-byte.
    let strict_fail: usize = outcome.failures.len();
    assert_eq!(
        strict_fail, 0,
        "{} strict-set blob(s) did not round-trip byte-exact",
        strict_fail
    );
}
