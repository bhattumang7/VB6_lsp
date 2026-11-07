//! Shared by the tests that run over the third-party VB6 corpus.

#![allow(dead_code)]

use std::path::{Path, PathBuf};

/// The corpus directory: `VB6_CORPUS_DIR`, else `re_lab/corpus` beside the repository.
pub fn corpus_dir() -> PathBuf {
    match std::env::var_os("VB6_CORPUS_DIR") {
        Some(dir) => dir.into(),
        None => Path::new(env!("CARGO_MANIFEST_DIR"))
            .join("..")
            .join("..")
            .join("..")
            .join("re_lab")
            .join("corpus"),
    }
}

/// Every `.frm` and `.ctl` file under `dir`, recursively.
pub fn collect_forms(dir: &Path, out: &mut Vec<PathBuf>) {
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
