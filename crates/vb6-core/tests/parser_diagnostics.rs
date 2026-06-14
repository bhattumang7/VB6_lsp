//! Tests for parser diagnostics and error reporting.

use vb6_core::frontend::ast::ExprArena;
use vb6_core::frontend::parser::Parser;
use vb6_core::frontend::scanner::ScannerContext;

fn parse_errors(src: &[u8]) -> Vec<u32> {
    let mut c = ScannerContext::new(1, 1, 0x0409);
    c.intern_keywords();
    let mut arena = ExprArena::new();
    let mut parser = Parser::new(&mut c, src);
    parser.parse_module(&mut arena);
    parser.diagnostics.items().iter().map(|d| d.code).collect()
}

// VB6 "Expected: <token>" code (all structural errors share this code).
const ERR_EXPECTED: u32 = 0x9c6f;
// VB6 "Only valid in object module" code.
const ERR_OBJECT_ONLY: u32 = 0xdee1;

#[test]
fn missing_end_if() {
    let codes = parse_errors(b"Sub T()\nIf True Then\nEnd Sub");
    assert!(codes.contains(&ERR_EXPECTED));
}

#[test]
fn invalid_as_clause() {
    let codes = parse_errors(b"Dim x As 123");
    assert!(codes.contains(&ERR_EXPECTED));
}

#[test]
fn expected_then() {
    let codes = parse_errors(b"Sub T()\nIf True\nEnd If\nEnd Sub");
    assert!(codes.contains(&ERR_EXPECTED));
}

#[test]
fn class_only_keywords_in_standard_module() {
    // Event and Implements are not allowed in standard modules.
    let codes = parse_errors(b"Event Click()\nImplements IInterface");
    assert!(codes.contains(&ERR_OBJECT_ONLY));
}
