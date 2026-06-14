//! Grammar regression tests derived from the old tree-sitter corpus.
//!
//! Each corpus file (`tests/fixtures/corpus/*.txt`) captures a VB6 syntax
//! edge case that was previously verified against the tree-sitter grammar.
//! This harness re-runs every snippet through the vb6-syntax Parser and
//! asserts that it parses without any syntax errors.
//!
//! Semantic diagnostics ("variable not defined", "sub not defined") are
//! excluded: snippets are minimal stubs where undefined names are expected.

use std::path::PathBuf;

use vb6_engine::frontend::ast::ExprArena;
use vb6_engine::frontend::parser::Parser;
use vb6_engine::frontend::scanner::ScannerContext;

// ── Corpus parsing ────────────────────────────────────────────────────────────

struct CorpusEntry {
    name: String,
    source: String,
}

/// Parse one corpus `.txt` file into a list of (name, VB6 source) pairs.
fn parse_corpus(text: &str) -> Vec<CorpusEntry> {
    let mut entries = Vec::new();

    // Split on lines that consist entirely of `=` signs (≥ 10 of them).
    // The format is:  ...preamble...
    //   ==================
    //   Test Name
    //   ==================
    //   <VB6 source>
    //   ---
    //   (expected tree)
    //   [next preamble / separator / entry]
    let sep_re = |line: &str| -> bool {
        let t = line.trim();
        t.len() >= 10 && t.chars().all(|c| c == '=')
    };

    let lines: Vec<&str> = text.lines().collect();
    let mut i = 0;
    while i < lines.len() {
        if sep_re(lines[i]) {
            // Next non-empty line(s) are the test name, until the closing separator.
            let name_start = i + 1;
            let mut close = name_start;
            while close < lines.len() && !sep_re(lines[close]) {
                close += 1;
            }
            if close >= lines.len() {
                break;
            }
            let name = lines[name_start..close]
                .iter()
                .map(|l| l.trim())
                .filter(|l| !l.is_empty())
                .collect::<Vec<_>>()
                .join(" ");

            // Body starts after the closing separator, ends at `---` (alone on a line)
            // or the next opening separator.
            let body_start = close + 1;
            let mut body_end = body_start;
            while body_end < lines.len() {
                let t = lines[body_end].trim();
                if t == "---" || sep_re(lines[body_end]) {
                    break;
                }
                body_end += 1;
            }

            let source = lines[body_start..body_end]
                .iter()
                .map(|l| *l)
                .collect::<Vec<_>>()
                .join("\n")
                .trim()
                .to_string();

            if !name.is_empty() && !source.is_empty() {
                entries.push(CorpusEntry { name, source });
            }

            // Skip past the `---` separator and the expected-tree body so we
            // land on the next preamble / separator.
            i = body_end;
            if i < lines.len() && lines[i].trim() == "---" {
                i += 1;
                // Skip the expected-tree block until a separator or EOF
                while i < lines.len() && !sep_re(lines[i]) {
                    i += 1;
                }
            }
        } else {
            i += 1;
        }
    }

    entries
}

// ── Parse helpers ─────────────────────────────────────────────────────────────

const PARSE_ERROR_CODE: u32 = 0x9c6f; // "Expected: <various>"
const CLASS_ONLY_CODE: u32 = 0xdee1; // "Only valid in object module"

/// Keywords that, at the start of a source snippet, mean the snippet is already
/// a module-level construct and should be fed to `parse_module` directly.
fn is_module_level(src: &str) -> bool {
    let first = src
        .lines()
        .map(str::trim)
        .find(|l| !l.is_empty() && !l.starts_with(';'))
        .unwrap_or("");
    first.starts_with("Sub ")
        || first.starts_with("Function ")
        || first.starts_with("Public ")
        || first.starts_with("Private ")
        || first.starts_with("Friend ")
        || first.starts_with("Option ")
        || first.starts_with("Type ")
        || first.starts_with("Enum ")
        || first.starts_with('#')
        || first.starts_with("Attribute ")
}

/// Run `src` through the VB6 parser and return only syntax-error diagnostics
/// (excluding semantic codes for undefined names / missing subs).
fn parse_errors(src: &str) -> Vec<u32> {
    let mut ctx = ScannerContext::new(1, 1, 0x0409);
    ctx.intern_keywords();
    let mut arena = ExprArena::new();
    let mut parser = Parser::new(&mut ctx, src.as_bytes());
    parser.parse_module(&mut arena);
    parser
        .diagnostics
        .items()
        .iter()
        .map(|d| d.code)
        .filter(|&c| c == PARSE_ERROR_CODE || c == CLASS_ONLY_CODE)
        .collect()
}

/// Parse `src` as a full module; if it looks like bare statements, wrap it in
/// a `Sub S()…End Sub` first.
fn check_source(src: &str) -> Vec<u32> {
    if is_module_level(src) {
        parse_errors(src)
    } else {
        let wrapped = format!("Sub S()\n{src}\nEnd Sub");
        parse_errors(&wrapped)
    }
}

// ── Known parser gaps ─────────────────────────────────────────────────────────
//
// Entries listed here are skipped with a printed note rather than failed.
// Each entry should reference a clear reason so it can be found and un-skipped
// once the underlying parser gap is closed.

/// `(file, entry_name_prefix, reason)`.
/// The prefix match is a substring of the entry name (case-sensitive).
const SKIP: &[(&str, &str, &str)] = &[];

fn should_skip(filename: &str, name: &str) -> Option<&'static str> {
    SKIP.iter()
        .find(|(f, prefix, _)| *f == filename && name.starts_with(prefix))
        .map(|(_, _, reason)| *reason)
}

// ── Test runner ───────────────────────────────────────────────────────────────

fn corpus_dir() -> PathBuf {
    let manifest = env!("CARGO_MANIFEST_DIR");
    PathBuf::from(manifest).join("tests/fixtures/corpus")
}

fn run_corpus_file(filename: &str) -> Vec<String> {
    let path = corpus_dir().join(filename);
    let text = std::fs::read_to_string(&path)
        .unwrap_or_else(|e| panic!("cannot read {}: {e}", path.display()));

    let entries = parse_corpus(&text);
    assert!(!entries.is_empty(), "{filename}: no entries parsed");

    let mut failures = Vec::new();
    for entry in &entries {
        if let Some(reason) = should_skip(filename, &entry.name) {
            eprintln!("  SKIP [{filename}] \"{}\": {reason}", entry.name);
            continue;
        }
        let errors = check_source(&entry.source);
        if !errors.is_empty() {
            failures.push(format!(
                "  [{filename}] \"{}\": syntax error code(s) {:?}\n    source: {:?}",
                entry.name,
                errors,
                &entry.source[..entry.source.len().min(120)],
            ));
        }
    }
    failures
}

macro_rules! corpus_test {
    ($name:ident, $file:literal) => {
        #[test]
        fn $name() {
            let failures = run_corpus_file($file);
            if !failures.is_empty() {
                panic!(
                    "{} snippet(s) failed to parse:\n{}",
                    failures.len(),
                    failures.join("\n")
                );
            }
        }
    };
}

corpus_test!(corpus_bugs, "bugs.txt");
corpus_test!(corpus_bugs_round2, "bugs_round2.txt");
corpus_test!(corpus_bugs_round3, "bugs_round3.txt");
corpus_test!(corpus_bugs_round4, "bugs_round4.txt");
corpus_test!(corpus_bugs_round5, "bugs_round5.txt");
corpus_test!(corpus_bugs_round6, "bugs_round6.txt");
corpus_test!(corpus_bugs_round7, "bugs_round7.txt");
