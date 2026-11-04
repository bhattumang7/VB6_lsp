//! Command line for the differential harness.
//!
//! ```text
//! oracle-diff controls
//! oracle-diff run [--taxonomy FILE] [--report FILE] <file-or-dir>...
//! ```
//!
//! Environment: `VB6_EXE` (compiler path), `ORACLE_WORK` (scratch directory),
//! `ORACLE_CACHE` (verdict cache directory).

use oracle_diff::compare::{classify, Outcome, Taxonomy};
use oracle_diff::oracle::{default_vb6_exe, Case, ModuleKind, Oracle, Verdict};
use oracle_diff::ours;
use std::collections::BTreeMap;
use std::path::{Path, PathBuf};
use std::process::ExitCode;

fn env_path(name: &str, default: impl FnOnce() -> PathBuf) -> PathBuf {
    std::env::var_os(name).map(PathBuf::from).unwrap_or_else(default)
}

fn build_oracle() -> std::io::Result<Oracle> {
    let work = env_path("ORACLE_WORK", || std::env::temp_dir().join("vb6-oracle-work"));
    let cache = env_path("ORACLE_CACHE", || {
        env_path("LOCALAPPDATA", std::env::temp_dir).join("vb6-oracle-diff").join("cache")
    });
    Oracle::new(env_path("VB6_EXE", default_vb6_exe), work, cache)
}

fn collect(path: &Path, out: &mut Vec<PathBuf>) {
    if path.is_dir() {
        if let Ok(entries) = std::fs::read_dir(path) {
            let mut entries: Vec<_> = entries.flatten().map(|e| e.path()).collect();
            entries.sort();
            for entry in entries {
                collect(&entry, out);
            }
        }
    } else if let Some(ext) = path.extension() {
        let ext = ext.to_string_lossy().to_ascii_lowercase();
        if ext == "bas" || ext == "cls" {
            out.push(path.to_path_buf());
        }
    }
}

/// Controls run first on every batch: a known-good and a known-bad module.
fn controls(oracle: &Oracle) -> bool {
    let good = Case::from_text("control-accept", ModuleKind::Bas, "Sub Probe()\n    x = 1 + 2\nEnd Sub\n");
    let bad = Case::from_text("control-reject", ModuleKind::Bas, "Sub Probe()\n    x = 1 +* 2\nEnd Sub\n");
    let verdicts = oracle.run(&[good, bad]);
    let ok = verdicts[0] == Verdict::Accepted
        && matches!(&verdicts[1], Verdict::Rejected { message, .. } if message == "Syntax error");
    println!("control accept: {:?}", verdicts[0]);
    println!("control reject: {:?}", verdicts[1]);
    println!("controls {}", if ok { "PASS" } else { "FAIL" });
    ok
}

fn run(args: &[String]) -> Result<ExitCode, String> {
    let mut taxonomy_path = PathBuf::from("tools/oracle-diff/taxonomy.tsv");
    let mut report_path = None;
    let mut inputs = Vec::new();
    let mut it = args.iter();
    while let Some(a) = it.next() {
        match a.as_str() {
            "--taxonomy" => taxonomy_path = it.next().ok_or("--taxonomy needs a file")?.into(),
            "--report" => report_path = Some(PathBuf::from(it.next().ok_or("--report needs a file")?)),
            other => collect(Path::new(other), &mut inputs),
        }
    }
    if inputs.is_empty() {
        return Err("no .bas/.cls inputs given".into());
    }
    let taxonomy = Taxonomy::parse(&std::fs::read_to_string(&taxonomy_path).unwrap_or_default())?;
    let oracle = build_oracle().map_err(|e| format!("oracle setup: {e}"))?;
    if !controls(&oracle) {
        return Err("controls failed; the oracle is not trustworthy".into());
    }
    let cases: Vec<Case> = inputs
        .iter()
        .filter_map(|p| {
            let kind = if p.extension().is_some_and(|e| e.eq_ignore_ascii_case("cls")) {
                ModuleKind::Cls
            } else {
                ModuleKind::Bas
            };
            std::fs::read(p).ok().map(|source| Case { id: p.display().to_string(), kind, source })
        })
        .collect();
    let verdicts = oracle.run(&cases);

    let mut counts: BTreeMap<String, usize> = BTreeMap::new();
    let mut unclassified: BTreeMap<String, usize> = BTreeMap::new();
    let mut report = String::from("case\toutcome\toracle\toracle_line\tours\tours_line\n");
    for (case, verdict) in cases.iter().zip(&verdicts) {
        let diags = ours::diagnostics(&case.source, case.kind);
        let outcome = classify(verdict, &diags, &taxonomy);
        *counts.entry(format!("{outcome:?}")).or_default() += 1;
        if let (Outcome::Unclassified, Verdict::Rejected { message, .. }) = (&outcome, verdict) {
            *unclassified.entry(message.clone()).or_default() += 1;
        }
        let (omsg, oline) = match verdict {
            Verdict::Accepted => ("accepted".to_string(), String::new()),
            Verdict::Rejected { message, line, .. } => {
                (message.clone(), line.map(|l| (l + 1).to_string()).unwrap_or_default())
            }
            Verdict::ProjectError { message } => (format!("PROJECT ERROR: {message}"), String::new()),
            Verdict::TimedOut => ("TIMED OUT".to_string(), String::new()),
            Verdict::Unrecognised { exit_code, log } => {
                let flat = log.replace(['\r', '\n'], " ");
                (format!("UNRECOGNISED exit={exit_code:?} log={}", flat.trim()), String::new())
            }
        };
        let (umsg, uline) = diags
            .first()
            .map(|d| (d.message.clone(), d.line.to_string()))
            .unwrap_or_default();
        report.push_str(&format!("{}\t{outcome:?}\t{omsg}\t{oline}\t{umsg}\t{uline}\n", case.id));
    }
    if let Some(path) = report_path {
        std::fs::write(&path, &report).map_err(|e| format!("report: {e}"))?;
        println!("report written to {}", path.display());
    }
    println!("{} cases", cases.len());
    for (k, v) in &counts {
        println!("  {k}: {v}");
    }
    if !unclassified.is_empty() {
        println!("messages needing classification:");
        for (m, n) in &unclassified {
            println!("  {n:>5}  {m}");
        }
    }
    let bad = counts
        .iter()
        .any(|(k, _)| matches!(k.as_str(), "FalseReject" | "FalseAccept" | "Unclassified" | "NoVerdict"));
    Ok(if bad { ExitCode::from(2) } else { ExitCode::SUCCESS })
}

fn main() -> ExitCode {
    let args: Vec<String> = std::env::args().skip(1).collect();
    let result = match args.first().map(String::as_str) {
        Some("controls") => build_oracle()
            .map_err(|e| format!("oracle setup: {e}"))
            .map(|o| if controls(&o) { ExitCode::SUCCESS } else { ExitCode::FAILURE }),
        Some("run") => run(&args[1..]),
        Some("ours") => {
            let mut inputs = Vec::new();
            args[1..].iter().for_each(|a| collect(Path::new(a), &mut inputs));
            for path in &inputs {
                let kind = if path.extension().is_some_and(|e| e.eq_ignore_ascii_case("cls")) {
                    ModuleKind::Cls
                } else {
                    ModuleKind::Bas
                };
                let source = std::fs::read(path).unwrap_or_default();
                let diags = ours::diagnostics(&source, kind);
                println!("{}: {} diagnostics", path.display(), diags.len());
                for d in diags {
                    println!("  line {}: {} ({:#x})", d.line, d.message, d.code);
                }
            }
            Ok(ExitCode::SUCCESS)
        }
        _ => Err("usage: oracle-diff controls | ours <path>... | run [--taxonomy F] [--report F] <path>...".into()),
    };
    result.unwrap_or_else(|e| {
        eprintln!("{e}");
        ExitCode::FAILURE
    })
}
