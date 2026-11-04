//! Where a reserved word may be used as a name. The fixture holds the
//! compiler's answer for every keyword in every declaration position of
//! `templates.txt`; the parser must accept exactly the forms the compiler
//! accepts (or rejects only for what the names mean) and reject the others.

use std::collections::BTreeMap;
use vb6_syntax::frontend::ast::ExprArena;
use vb6_syntax::frontend::parser::{ModuleKind, Parser};
use vb6_syntax::frontend::scanner::ScannerContext;

const TEMPLATES: &str = include_str!("fixtures/keyword_names/templates.txt");
const ANSWERS: &str = include_str!("fixtures/keyword_names/oracle.tsv");

struct Template {
    kind: ModuleKind,
    text: String,
}

fn templates() -> BTreeMap<String, Template> {
    let mut out: BTreeMap<String, Template> = BTreeMap::new();
    let mut current: Option<String> = None;
    for line in TEMPLATES.lines() {
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

fn diagnostics(source: &str, kind: ModuleKind) -> usize {
    let mut ctx = ScannerContext::new(1, 1, 0x0409);
    ctx.intern_keywords();
    let mut arena = ExprArena::new();
    let mut parser = Parser::with_module_kind(&mut ctx, source.as_bytes(), kind);
    parser.parse_module(&mut arena);
    parser.diagnostics.items().len()
}

#[test]
fn keywords_as_names_match_the_compiler_in_every_position() {
    let templates = templates();
    let mut wrong: Vec<String> = Vec::new();
    let mut checked = 0;
    for line in ANSWERS.lines().filter(|l| !l.starts_with('#') && !l.trim().is_empty()) {
        let mut cols = line.split('\t');
        let (template, word, class) = (cols.next().unwrap(), cols.next().unwrap(), cols.next().unwrap());
        let message = cols.next().unwrap_or("");
        let expect_parses = match class {
            "accepted" | "semantic" => true,
            "syntax" => false,
            "unanswered" => continue,
            other => panic!("unknown class `{other}`"),
        };
        // A block opener whose templates do not close the block: the compiler's semantic
        // error comes first and hides the missing closer, so its answer says nothing
        // about the parse.
        if class == "semantic" && ["With", "While", "Case"].contains(&word) && template.starts_with("start") {
            continue;
        }
        let t = &templates[template];
        let source = t.text.replace("{k}", word);
        let found = diagnostics(&source, t.kind);
        checked += 1;
        if (found == 0) != expect_parses {
            let verdict = if expect_parses { "accepts" } else { "rejects" };
            wrong.push(format!("{template} {word}: the compiler {verdict} ({class} {message}), parser reported {found}"));
        }
    }
    assert!(checked > 5000, "fixture is unexpectedly small: {checked}");
    let mut per_template: BTreeMap<&str, usize> = BTreeMap::new();
    for w in &wrong {
        *per_template.entry(w.split(' ').next().unwrap()).or_default() += 1;
    }
    let mut examples: Vec<String> = Vec::new();
    for template in per_template.keys() {
        examples.extend(wrong.iter().filter(|w| w.starts_with(&format!("{template} "))).take(10).cloned());
    }
    assert!(
        wrong.is_empty(),
        "{} disagreements {:?}, examples:\n{}",
        wrong.len(),
        per_template,
        examples.join("\n")
    );
}
