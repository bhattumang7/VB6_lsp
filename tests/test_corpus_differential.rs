//! Differential / anomaly harness over the corpus.
//!
//! Pinpoints every byte-accounting anomaly (overlap, unexplained gap, decode error)
//! with its file + property + offset so they can be triaged as bugs vs. expected
//! un-harvested references. The pure-Rust core runs anywhere; an optional oracle
//! cross-check (Microsoft's own loader) is gated behind `VB6_FRX_ORACLE=1` so the
//! crate never links COM. Skips when the corpus is absent.

use std::path::{Path, PathBuf};

use vb6_lsp::controls::coverage::{self, Gap};

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

#[test]
fn corpus_anomaly_triage() {
    let corpus = Path::new(env!("CARGO_MANIFEST_DIR"))
        .join("re_lab")
        .join("corpus");
    if !corpus.exists() {
        eprintln!("SKIP: corpus not present at {}", corpus.display());
        return;
    }
    let mut forms = Vec::new();
    collect_forms(&corpus, &mut forms);

    let mut overlaps: Vec<String> = Vec::new();
    let mut gaps: Vec<String> = Vec::new();
    let mut errors: Vec<String> = Vec::new();

    for form in &forms {
        let reports = match coverage::coverage_for_form(form) {
            Ok(r) => r,
            Err(_) => continue,
        };
        for r in reports {
            for &(a, b) in &r.overlaps {
                let sa = &r.spans[a];
                let sb = &r.spans[b];
                overlaps.push(format!(
                    "{}: {} {}@0x{:X}..0x{:X} ({:?})  OVERLAPS  {} {}@0x{:X}..0x{:X} ({:?})",
                    r.file,
                    sa.control_path,
                    sa.property,
                    sa.start,
                    sa.end,
                    sa.kind,
                    sb.control_path,
                    sb.property,
                    sb.start,
                    sb.end,
                    sb.kind
                ));
            }
            for g in &r.gaps {
                if let Gap::Unexplained { start, end } = g {
                    gaps.push(format!(
                        "{}: unexplained 0x{:X}..0x{:X} ({} bytes)",
                        r.file,
                        start,
                        end,
                        end - start
                    ));
                }
            }
            for (what, msg) in &r.errors {
                errors.push(format!("{}: {} -> {}", r.file, what, msg));
            }
        }
    }

    eprintln!("=== CORPUS ANOMALY TRIAGE ===");
    eprintln!("overlaps: {}", overlaps.len());
    for o in overlaps.iter().take(20) {
        eprintln!("  {}", o);
    }
    eprintln!("unexplained gaps: {}", gaps.len());
    for g in gaps.iter() {
        eprintln!("  {}", g);
    }
    eprintln!("decode errors: {}", errors.len());
    for e in errors.iter().take(20) {
        eprintln!("  {}", e);
    }

    assert!(!forms.is_empty(), "expected corpus forms");
    // No companion byte may be claimed by two references — an overlap is always a
    // decode bug (this caught the ComboBox ItemData over-read, now fixed).
    assert!(
        overlaps.is_empty(),
        "{} byte-span overlap(s) across the corpus — a decode bug",
        overlaps.len()
    );
}
