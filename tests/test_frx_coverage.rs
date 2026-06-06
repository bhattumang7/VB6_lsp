//! Real-corpus byte-accounting for VB6 companion files (.frx/.ctx).
//!
//! Measures, across the `re_lab/corpus` sample, how many bytes of every companion
//! we attribute to a decoded reference — the empirical answer to "do we read the
//! whole file". Skips automatically when the corpus is absent (e.g. in CI).

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
fn corpus_byte_accounting_report() {
    let corpus = Path::new(env!("CARGO_MANIFEST_DIR"))
        .join("re_lab")
        .join("corpus");
    if !corpus.exists() {
        eprintln!("SKIP: corpus not present at {}", corpus.display());
        return;
    }

    let mut forms = Vec::new();
    collect_forms(&corpus, &mut forms);

    let mut companions = 0usize;
    let mut total_bytes = 0usize;
    let mut covered_bytes = 0usize;
    let mut unexplained_ranges = 0usize;
    let mut unexplained_bytes = 0usize;
    let mut overlaps = 0usize;
    let mut errors = 0usize;
    let mut opaque = 0usize;
    let mut complete_files = 0usize;
    let mut worst: Vec<(String, usize)> = Vec::new();

    for form in &forms {
        let reports = match coverage::coverage_for_form(form) {
            Ok(r) => r,
            Err(_) => continue,
        };
        for r in reports {
            companions += 1;
            total_bytes += r.total_len;
            covered_bytes += r.covered;
            overlaps += r.overlaps.len();
            errors += r.errors.len();
            opaque += r.opaque_spans;
            if r.is_complete() {
                complete_files += 1;
            }
            let ub = r.unexplained_bytes();
            if ub > 0 {
                unexplained_bytes += ub;
                unexplained_ranges += r
                    .gaps
                    .iter()
                    .filter(|g| matches!(g, Gap::Unexplained { .. }))
                    .count();
                worst.push((format!("{} ({} spans)", r.file, r.spans.len()), ub));
            }
        }
    }

    let pct = if total_bytes > 0 {
        covered_bytes as f64 * 100.0 / total_bytes as f64
    } else {
        100.0
    };
    worst.sort_by(|a, b| b.1.cmp(&a.1));

    eprintln!("=== FRX/CTX BYTE-ACCOUNTING (re_lab/corpus) ===");
    eprintln!("forms/ctls scanned  : {}", forms.len());
    eprintln!("companions analysed : {}", companions);
    eprintln!(
        "complete files      : {}/{} (no unexplained bytes / overlaps / errors)",
        complete_files, companions
    );
    eprintln!("total bytes         : {}", total_bytes);
    eprintln!("covered bytes       : {} ({:.2}%)", covered_bytes, pct);
    eprintln!(
        "unexplained         : {} ranges, {} bytes",
        unexplained_ranges, unexplained_bytes
    );
    eprintln!("overlaps (bugs)     : {}", overlaps);
    eprintln!("decode errors       : {}", errors);
    eprintln!("opaque Tier-3 spans : {}", opaque);
    eprintln!("--- top files by unexplained bytes ---");
    for (f, b) in worst.iter().take(15) {
        eprintln!("  {:>8} B  {}", b, f);
    }

    // First-pass goal is measurement, not a hard 100%: just assert the harness ran.
    assert!(companions > 0, "expected to analyse some companions");
}
