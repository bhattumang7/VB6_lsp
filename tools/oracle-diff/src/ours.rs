//! Runs the `vb6-syntax` parser on a case and reports its diagnostics in the
//! same shape the compiler's verdicts use.

use vb6_engine::session::Session;
use vb6_syntax::frontend::ast::ExprArena;
use vb6_syntax::frontend::diagnostics::Diagnostic;
use vb6_syntax::frontend::diagnostics::message_for_code;
use crate::oracle::ModuleKind;
use vb6_syntax::frontend::parser::{ModuleKind as ParserModuleKind, Parser};
use vb6_syntax::frontend::line_reader::LineTable;
use vb6_syntax::frontend::scanner::ScannerContext;

/// One parser diagnostic.
#[derive(Clone, Debug, PartialEq, Eq)]
pub struct OurDiagnostic {
    pub code: u32,
    /// The line the compiler would report: counted as it counts (every line feed, carriage return
    /// or carriage return with line feed is one break, and every piece of a long line is a line),
    /// and for an error inside a continued statement the first line of that statement. `None`
    /// when the compiler gives the error no line.
    pub line: Option<u32>,
    pub message: String,
}

/// Parses `source` as a module of the given kind and returns every diagnostic in report order.
pub fn diagnostics(source: &[u8], kind: ModuleKind) -> Vec<OurDiagnostic> {
    let mut ctx = ScannerContext::new(1, 1, 0x0409);
    ctx.intern_keywords();
    let mut arena = ExprArena::new();
    let parser_kind = match kind {
        ModuleKind::Bas => ParserModuleKind::Standard,
        ModuleKind::Cls => ParserModuleKind::Class,
    };
    let mut parser = Parser::with_module_kind(&mut ctx, source, parser_kind);
    parser.parse_module(&mut arena);
    let table = LineTable::new(source);
    // A limit broken while the text is read comes first and has no line.
    let unnumbered = usize::from(parser.read_problem().is_some());
    parser
        .diagnostics
        .items()
        .iter()
        .enumerate()
        .map(|(n, d)| {
            let line = (n >= unnumbered).then(|| table.logical_line_of(d.span.start));
            let base = message_for_code(d.code)
                .map(str::to_string)
                .unwrap_or_else(|| format!("(code {:#x})", d.code));
            let message = match d.label {
                Some(label) if d.code == 0x9c6f => format!("Expected: {label}"),
                _ => base,
            };
            OurDiagnostic { code: d.code, line, message }
        })
        .collect()
}

/// Every diagnostic the language server would show for `source` as the only module of a
/// standard project (bundled runtime libraries): parse errors and name-resolution errors.
pub fn full_diagnostics(source: &[u8], kind: ModuleKind) -> Vec<OurDiagnostic> {
    let path = match kind {
        ModuleKind::Bas => "m.bas",
        ModuleKind::Cls => "m.cls",
    };
    let session = Session::from_sources(vec![(path.to_string(), source.to_vec())]);
    session.diagnostics(0).iter().map(|d| describe(source, d)).collect()
}

fn describe(source: &[u8], d: &Diagnostic) -> OurDiagnostic {
    let start = (d.span.start as usize).min(source.len());
    let line = Some(LineTable::new(source).logical_line_of(start as u32));
    let base = message_for_code(d.code)
        .map(str::to_string)
        .unwrap_or_else(|| format!("(code {:#x})", d.code));
    let message = match d.label {
        Some(label) if d.code == 0x9c6f => format!("Expected: {label}"),
        _ => base,
    };
    OurDiagnostic { code: d.code, line, message }
}
