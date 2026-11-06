//! Constant values, array bounds, fixed string lengths, Optional defaults, enum members and
//! repeated names, each against the compiler's answer: its message and the line it reports, or
//! no report at all.

use vb6_syntax::frontend::ast::ExprArena;
use vb6_syntax::frontend::parser::{ModuleKind, Parser};
use vb6_syntax::frontend::scanner::ScannerContext;

const CASES: &str = include_str!("fixtures/constant_declarations/cases.tsv");

fn unescape(s: &str) -> String {
    let mut out = String::new();
    let mut chars = s.chars();
    while let Some(c) = chars.next() {
        if c != '\\' {
            out.push(c);
            continue;
        }
        match chars.next() {
            Some('r') => out.push('\r'),
            Some('n') => out.push('\n'),
            Some('t') => out.push('\t'),
            Some('\\') => out.push('\\'),
            other => panic!("bad escape {other:?}"),
        }
    }
    out
}

/// The first thing the parser reports, as its message and the line it falls on.
fn first_report(source: &str, kind: ModuleKind) -> Option<(String, u32)> {
    let mut ctx = ScannerContext::new(1, 1, 0x0409);
    ctx.intern_keywords();
    let mut arena = ExprArena::new();
    let mut parser = Parser::with_module_kind(&mut ctx, source.as_bytes(), kind);
    parser.parse_module(&mut arena);
    parser.diagnostics.items().first().map(|d| {
        let start = (d.span.start as usize).min(source.len());
        let line = 1 + source.as_bytes()[..start].iter().filter(|b| **b == b'\n').count() as u32;
        (d.message().unwrap_or_else(|| format!("(code {:#x})", d.code)), line)
    })
}

#[test]
fn constants_bounds_lengths_defaults_and_names_match_the_compiler() {
    let mut checked = 0;
    let mut wrong = Vec::new();
    for line in CASES.lines().filter(|l| !l.starts_with('#') && !l.is_empty()) {
        let mut cols = line.split('\t');
        let kind = match cols.next().unwrap() {
            "bas" => ModuleKind::Standard,
            "cls" => ModuleKind::Class,
            other => panic!("unknown kind {other}"),
        };
        let message = cols.next().unwrap();
        let at: u32 = cols.next().unwrap().parse().unwrap_or(0);
        let source = unescape(cols.next().unwrap());
        let want = if message.is_empty() { None } else { Some((message.to_string(), at)) };
        let got = first_report(&source, kind);
        checked += 1;
        if got != want {
            wrong.push(format!("{source:?}: the compiler says {want:?}, the parser {got:?}"));
        }
    }
    assert!(checked > 400, "fixture is unexpectedly small: {checked}");
    assert!(
        wrong.is_empty(),
        "{} of {checked} differ:\n{}",
        wrong.len(),
        wrong.iter().take(40).cloned().collect::<Vec<_>>().join("\n")
    );
}
