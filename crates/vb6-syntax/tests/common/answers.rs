//! Replays a fixture of compiler answers against the parser: every row names a
//! template and a word, and says whether the compiler found the form wrong
//! (`syntax`), found nothing wrong with it (`accepted`, or `semantic` when it
//! objected only to what the names mean), or gave no answer (`unanswered`).
#![allow(dead_code)]

use std::collections::BTreeMap;
use vb6_syntax::frontend::ast::ExprArena;
use vb6_syntax::frontend::parser::{ModuleKind, Parser};
use vb6_syntax::frontend::scanner::ScannerContext;

pub struct Template {
    pub kind: ModuleKind,
    pub text: String,
}

/// Parses a templates file: `=== <name> <bas|cls>` starts a block.
pub fn templates(text: &str) -> BTreeMap<String, Template> {
    let mut out: BTreeMap<String, Template> = BTreeMap::new();
    let mut current: Option<String> = None;
    for line in text.lines() {
        if let Some(header) = line.strip_prefix("=== ") {
            let mut parts = header.split_whitespace();
            let name = parts.next().expect("a template name").to_string();
            let kind = match parts.next() {
                Some("bas") => ModuleKind::Standard,
                Some("cls") => ModuleKind::Class,
                other => panic!("unknown module kind {other:?}"),
            };
            out.insert(name.clone(), Template { kind, text: String::new() });
            current = Some(name);
        } else if let Some(name) = &current {
            let t = out.get_mut(name).unwrap();
            t.text.push_str(line);
            t.text.push('\n');
        }
    }
    out
}

/// A word as the fixture writes it: `\n` and `\t` stand for a line break and a tab.
pub fn unescape(word: &str) -> String {
    word.replace("\\n", "\n").replace("\\t", "\t").replace("\\s", " ")
}

/// The diagnostics that report a block left open at the end of the source. When the
/// source holds a semantic error as well, the compiler reports that one first and
/// never gets to the open block, so such a row says nothing about the parse.
const MISSING_END_CODES: [u32; 14] = [
    0x9c8f, 0x9c90, 0x9c92, 0x9c93, 0x9c94, 0x9c95, 0x9cac, 0x9cbe, 0x9cc7, 0x9c76, 0xdeff, 0x9c91,
    // A jump to a label that is never defined is found after the semantic checks as well.
    0x9ca4,
    // The scope check on declared names is the binder's; the parser also reports a repeated
    // parameter or label, which the compiler finds after the semantic checks as well.
    0x9c9f,
];

pub fn diagnostics(source: &str, kind: ModuleKind) -> usize {
    diagnostic_codes(source, kind).len()
}

pub fn diagnostic_codes(source: &str, kind: ModuleKind) -> Vec<u32> {
    let mut ctx = ScannerContext::new(1, 1, 0x0409);
    ctx.intern_keywords();
    let mut arena = ExprArena::new();
    let mut parser = Parser::with_module_kind(&mut ctx, source.as_bytes(), kind);
    parser.parse_module(&mut arena);
    parser.diagnostics.items().iter().map(|d| d.code).collect()
}

/// Every row on which the parser and the compiler disagree, as readable lines.
/// `skip(template, word, class, message)` leaves out rows whose answer says nothing about the parse.
pub fn disagreements(
    templates_text: &str,
    answers: &str,
    skip: impl Fn(&str, &str, &str, &str) -> bool,
) -> (usize, Vec<String>) {
    let templates = templates(templates_text);
    let mut wrong = Vec::new();
    let mut checked = 0;
    for line in answers.lines().filter(|l| !l.starts_with('#') && !l.trim().is_empty()) {
        let mut cols = line.split('\t');
        let (template, word, class) = (cols.next().unwrap(), cols.next().unwrap(), cols.next().unwrap());
        let message = cols.next().unwrap_or("");
        let expect_parses = match class {
            "accepted" | "semantic" => true,
            "syntax" => false,
            "unanswered" => continue,
            other => panic!("unknown class `{other}`"),
        };
        if skip(template, word, class, message) {
            continue;
        }
        let t = &templates[template];
        let source = t.text.replace("{k}", &unescape(word));
        let codes = diagnostic_codes(&source, t.kind);
        let found = codes.len();
        // A semantic answer hides a block that is left open.
        let masked = class == "semantic" && !codes.is_empty() && codes.iter().all(|c| MISSING_END_CODES.contains(c));
        checked += 1;
        if !masked && (found == 0) != expect_parses {
            let verdict = if expect_parses { "accepts" } else { "rejects" };
            wrong.push(format!("{template} {word}: the compiler {verdict} ({class} {message}), parser reported {found}"));
        }
    }
    (checked, wrong)
}

/// Fails with a summary by template and a few examples of each.
pub fn assert_none_wrong(checked: usize, wrong: &[String]) {
    let mut per_template: BTreeMap<&str, usize> = BTreeMap::new();
    for w in wrong {
        *per_template.entry(w.split(' ').next().unwrap()).or_default() += 1;
    }
    let mut examples: Vec<String> = Vec::new();
    for template in per_template.keys() {
        examples.extend(wrong.iter().filter(|w| w.starts_with(&format!("{template} "))).take(std::env::var("EXAMPLES").ok().and_then(|v| v.parse().ok()).unwrap_or(12)).cloned());
    }
    assert!(
        wrong.is_empty(),
        "{} of {checked} rows disagree {:?}, examples:\n{}",
        wrong.len(),
        per_template,
        examples.join("\n")
    );
}
