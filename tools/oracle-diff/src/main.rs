//! Command line for the differential harness.
//!
//! ```text
//! oracle-diff controls
//! oracle-diff ours <path>...
//! oracle-diff run [--taxonomy FILE] [--report FILE] <file-or-dir>...
//! oracle-diff mutate [--seed N] [--count N] [--out DIR] [--taxonomy FILE] [--report FILE] <file-or-dir>...
//! ```
//!
//! Environment: `VB6_EXE` (compiler path), `ORACLE_WORK` (scratch directory),
//! `ORACLE_CACHE` (verdict cache directory).

use oracle_diff::compare::{classify, Outcome, Taxonomy};
use oracle_diff::mutate::{header_lines, mutants};
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

/// Our side of a comparison. By default every diagnostic the language server shows (parse errors
/// and name-resolution errors); `ORACLE_LAYERS=parser` limits it to the parser.
fn our_diagnostics(source: &[u8], kind: ModuleKind) -> Vec<ours::OurDiagnostic> {
    if std::env::var("ORACLE_LAYERS").is_ok_and(|v| v == "parser") {
        ours::diagnostics(source, kind)
    } else {
        ours::full_diagnostics(source, kind)
    }
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

fn kind_of(path: &Path) -> ModuleKind {
    if path.extension().is_some_and(|e| e.eq_ignore_ascii_case("cls")) {
        ModuleKind::Cls
    } else {
        ModuleKind::Bas
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

/// Options shared by the `run` and `mutate` commands.
#[derive(Default)]
struct Options {
    taxonomy: Option<PathBuf>,
    report: Option<PathBuf>,
    out: Option<PathBuf>,
    seed: u64,
    count: usize,
    inputs: Vec<PathBuf>,
}

fn parse_options(args: &[String]) -> Result<Options, String> {
    let mut o = Options { seed: 1, count: 200, ..Options::default() };
    let mut it = args.iter();
    while let Some(a) = it.next() {
        let mut value = |what: &str| it.next().cloned().ok_or(format!("{what} needs a value"));
        match a.as_str() {
            "--taxonomy" => o.taxonomy = Some(value("--taxonomy")?.into()),
            "--report" => o.report = Some(value("--report")?.into()),
            "--out" => o.out = Some(value("--out")?.into()),
            "--seed" => o.seed = value("--seed")?.parse().map_err(|_| "--seed needs a number")?,
            "--count" => o.count = value("--count")?.parse().map_err(|_| "--count needs a number")?,
            other => collect(Path::new(other), &mut o.inputs),
        }
    }
    if o.inputs.is_empty() {
        return Err("no .bas/.cls inputs given".into());
    }
    Ok(o)
}

fn load_taxonomy(o: &Options) -> Result<Taxonomy, String> {
    let path = o.taxonomy.clone().unwrap_or_else(|| PathBuf::from("tools/oracle-diff/taxonomy.tsv"));
    Taxonomy::parse(&std::fs::read_to_string(&path).unwrap_or_default())
}

/// The verdict's message and its line in the file (the compiler counts from 0 and
/// leaves the module header out).
fn describe(verdict: &Verdict, header: u32) -> (String, String) {
    match verdict {
        Verdict::Accepted => ("accepted".to_string(), String::new()),
        Verdict::Rejected { message, line, .. } => {
            (message.clone(), line.map(|l| (l + 1 + header).to_string()).unwrap_or_default())
        }
        Verdict::ProjectError { message } => (format!("PROJECT ERROR: {message}"), String::new()),
        Verdict::TimedOut => ("TIMED OUT".to_string(), String::new()),
        Verdict::Unrecognised { exit_code, log } => {
            let flat = log.replace(['\r', '\n'], " ");
            (format!("UNRECOGNISED exit={exit_code:?} log={}", flat.trim()), String::new())
        }
    }
}

/// Compares every case, prints the summary and writes the report. Returns the
/// outcome of each case, in order, and whether anything needs attention.
fn compare_all(
    oracle: &Oracle,
    taxonomy: &Taxonomy,
    cases: &[Case],
    report_path: Option<&Path>,
) -> Result<(Vec<Outcome>, bool), String> {
    let verdicts = oracle.run(cases);
    let mut outcomes = Vec::new();
    let mut counts: BTreeMap<String, usize> = BTreeMap::new();
    let mut unclassified: BTreeMap<String, usize> = BTreeMap::new();
    let mut report = String::from("case\toutcome\toracle\toracle_line\tours\tours_line\n");
    for (case, verdict) in cases.iter().zip(&verdicts) {
        let diags = our_diagnostics(&case.source, case.kind);
        let outcome = classify(verdict, &diags, taxonomy);
        *counts.entry(format!("{outcome:?}")).or_default() += 1;
        if let (Outcome::Unclassified, Verdict::Rejected { message, .. }) = (&outcome, verdict) {
            *unclassified.entry(message.clone()).or_default() += 1;
        }
        let (omsg, oline) = describe(verdict, header_lines(&case.source));
        let (umsg, uline) = diags.first().map(|d| (d.message.clone(), d.line.map(|l| l.to_string()).unwrap_or_default())).unwrap_or_default();
        report.push_str(&format!("{}\t{outcome:?}\t{omsg}\t{oline}\t{umsg}\t{uline}\n", case.id));
        outcomes.push(outcome);
    }
    if let Some(path) = report_path {
        std::fs::write(path, &report).map_err(|e| format!("report: {e}"))?;
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
    let attention = outcomes.iter().any(|o| {
        matches!(
            o,
            Outcome::FalseReject
                | Outcome::FalseAccept
                | Outcome::Unverifiable
                | Outcome::Unclassified
                | Outcome::NoVerdict
        )
    });
    Ok((outcomes, attention))
}

fn trusted_oracle() -> Result<Oracle, String> {
    let oracle = build_oracle().map_err(|e| format!("oracle setup: {e}"))?;
    if !controls(&oracle) {
        return Err("controls failed; the oracle is not trustworthy".into());
    }
    Ok(oracle)
}

fn run(args: &[String]) -> Result<ExitCode, String> {
    let o = parse_options(args)?;
    let taxonomy = load_taxonomy(&o)?;
    let oracle = trusted_oracle()?;
    let cases: Vec<Case> = o
        .inputs
        .iter()
        .filter_map(|p| {
            std::fs::read(p).ok().map(|source| Case { id: p.display().to_string(), kind: kind_of(p), source })
        })
        .collect();
    let (_, attention) = compare_all(&oracle, &taxonomy, &cases, o.report.as_deref())?;
    Ok(if attention { ExitCode::from(2) } else { ExitCode::SUCCESS })
}

/// Parents are the modules the compiler accepts and the parser passes, so a
/// mutant's first error comes from its one change.
fn accepted_parents(oracle: &Oracle, o: &Options) -> Vec<Case> {
    let parents: Vec<Case> = o
        .inputs
        .iter()
        .filter_map(|p| {
            std::fs::read(p).ok().map(|source| Case { id: p.display().to_string(), kind: kind_of(p), source })
        })
        .collect();
    let parent_verdicts = oracle.run(&parents);
    parents
        .into_iter()
        .zip(parent_verdicts)
        .filter(|(c, v)| *v == Verdict::Accepted && our_diagnostics(&c.source, c.kind).is_empty())
        .map(|(c, _)| c)
        .collect()
}

fn mutant_cases(parents: &[Case], o: &Options) -> Vec<Case> {
    let per_parent = o.count.div_ceil(parents.len());
    let mut cases = Vec::new();
    for (n, parent) in parents.iter().enumerate() {
        for (m, mutant) in mutants(&parent.source, o.seed.wrapping_add(n as u64), per_parent).into_iter().enumerate() {
            cases.push(Case {
                id: format!("{}#{m} line {}: {}", parent.id, mutant.line, mutant.description),
                kind: parent.kind,
                source: mutant.source,
            });
        }
    }
    cases
}

/// Writes every case whose outcome needs a look into `dir`.
fn write_interesting(dir: &Path, cases: &[Case], outcomes: &[Outcome]) -> Result<(), String> {
    std::fs::create_dir_all(dir).map_err(|e| format!("out dir: {e}"))?;
    for (n, (case, outcome)) in cases.iter().zip(outcomes).enumerate() {
        if matches!(outcome, Outcome::AgreeAccept | Outcome::AgreeRejectSameMessage | Outcome::AgreeRejectDifferentMessage | Outcome::SemanticOnly) {
            continue;
        }
        let ext = if case.kind == ModuleKind::Cls { "cls" } else { "bas" };
        std::fs::write(dir.join(format!("{n:05}-{outcome:?}.{ext}")), &case.source)
            .map_err(|e| format!("out file: {e}"))?;
    }
    Ok(())
}

fn mutate(args: &[String]) -> Result<ExitCode, String> {
    let o = parse_options(args)?;
    let taxonomy = load_taxonomy(&o)?;
    let oracle = trusted_oracle()?;

    let parents = accepted_parents(&oracle, &o);
    if parents.is_empty() {
        return Err("no parent module is accepted by both sides".into());
    }
    println!("{} parent modules", parents.len());

    let cases = mutant_cases(&parents, &o);
    let (outcomes, attention) = compare_all(&oracle, &taxonomy, &cases, o.report.as_deref())?;
    if let Some(dir) = &o.out {
        write_interesting(dir, &cases, &outcomes)?;
    }
    Ok(if attention { ExitCode::from(2) } else { ExitCode::SUCCESS })
}

/// `matrix <templates> <words> <output.tsv>`: asks the compiler about every
/// template with every word and writes `template<TAB>word<TAB>class<TAB>message`.
fn matrix_command(args: &[String]) -> Result<ExitCode, String> {
    let [templates, words, output] = args else {
        return Err("usage: oracle-diff matrix <templates> <words> <output.tsv>".into());
    };
    let templates = oracle_diff::matrix::parse_templates(
        &std::fs::read_to_string(templates).map_err(|e| format!("templates: {e}"))?,
    )?;
    let words = oracle_diff::matrix::parse_words(&std::fs::read_to_string(words).map_err(|e| format!("words: {e}"))?);
    let taxonomy = load_taxonomy(&Options::default())?;
    let oracle = trusted_oracle()?;
    let rows = oracle_diff::matrix::run(&oracle, &taxonomy, &templates, &words)?;
    let mut text = String::from("# template\tword\tclass\tmessage\n");
    for (template, word, fate) in &rows {
        text.push_str(&format!("{template}\t{word}\t{}\n", fate.column()));
    }
    std::fs::write(output, text).map_err(|e| format!("output: {e}"))?;
    println!("{} rows written to {output}", rows.len());
    Ok(ExitCode::SUCCESS)
}

fn main() -> ExitCode {
    let args: Vec<String> = std::env::args().skip(1).collect();
    let result = match args.first().map(String::as_str) {
        Some("controls") => build_oracle()
            .map_err(|e| format!("oracle setup: {e}"))
            .map(|o| if controls(&o) { ExitCode::SUCCESS } else { ExitCode::FAILURE }),
        Some("run") => run(&args[1..]),
        Some("mutate") => mutate(&args[1..]),
        Some("matrix") => matrix_command(&args[1..]),
        Some("ours") => {
            let mut inputs = Vec::new();
            args[1..].iter().for_each(|a| collect(Path::new(a), &mut inputs));
            for path in &inputs {
                let source = std::fs::read(path).unwrap_or_default();
                let diags = our_diagnostics(&source, kind_of(path));
                println!("{}: {} diagnostics", path.display(), diags.len());
                for d in diags {
                    println!("  line {}: {} ({:#x})", d.line.map(|l| l.to_string()).unwrap_or_else(|| "-".into()), d.message, d.code);
                }
            }
            Ok(ExitCode::SUCCESS)
        }
        _ => Err("usage: oracle-diff controls | ours <path>... | run [opts] <path>... | mutate [opts] <path>...".into()),
    };
    result.unwrap_or_else(|e| {
        eprintln!("{e}");
        ExitCode::FAILURE
    })
}
