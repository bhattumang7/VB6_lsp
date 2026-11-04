//! Runs the `vb6-syntax` parser on a case and reports its diagnostics in the
//! same shape the compiler's verdicts use.

use vb6_syntax::frontend::ast::ExprArena;
use vb6_syntax::frontend::diagnostics::message_for_code;
use crate::oracle::ModuleKind;
use vb6_syntax::frontend::parser::{ModuleKind as ParserModuleKind, Parser};
use vb6_syntax::frontend::scanner::ScannerContext;

/// One parser diagnostic.
#[derive(Clone, Debug, PartialEq, Eq)]
pub struct OurDiagnostic {
    pub code: u32,
    /// 1-based line of the start of the diagnostic's span.
    pub line: u32,
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
    parser
        .diagnostics
        .items()
        .iter()
        .map(|d| {
            let start = (d.span.start as usize).min(source.len());
            let line = 1 + source[..start].iter().filter(|b| **b == b'\n').count() as u32;
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
