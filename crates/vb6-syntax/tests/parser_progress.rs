//! Regression tests: the module-level parser loop must always make progress.
//!
//! A comment (`'...`) or `Rem` is a statement-terminator token, so
//! `skip_to_stmt_end` stops *at* it without consuming. A `.frm` form-designer
//! header line such as `StartUpPosition = 3  'Windows Default` once left the
//! top-level loop spinning on that comment forever, allocating an unbounded
//! number of "Expected" diagnostics (gigabytes of RAM in seconds). These tests
//! pin that down: parsing must terminate and must not produce a diagnostic
//! storm.

use vb6_syntax::frontend::ast::ExprArena;
use vb6_syntax::frontend::parser::Parser;
use vb6_syntax::frontend::scanner::ScannerContext;

/// Parse `src` as a module and return how many diagnostics were produced.
/// If parsing ever spins, this never returns and the test fails by timeout.
fn diagnostic_count(src: &[u8]) -> usize {
    let mut ctx = ScannerContext::new(1, 1, 0x0409);
    ctx.intern_keywords();
    let mut arena = ExprArena::new();
    let mut parser = Parser::new(&mut ctx, src);
    let _ = parser.parse_module(&mut arena);
    parser.diagnostics.items().len()
}

/// An inline comment after an unrecognised module-level line must not spin.
#[test]
fn module_level_inline_comment_terminates() {
    let src = b"StartUpPosition =   3  'Windows Default\r\n";
    assert!(diagnostic_count(src) < 1000);
}

/// A bare comment-only line at module level must be skipped, not treated as an
/// unrecognised statement.
#[test]
fn module_level_comment_only_line() {
    let src = b"'just a comment\r\nRem another\r\n";
    assert_eq!(diagnostic_count(src), 0);
}

/// The whole form-designer header of a `.frm` (which is not VB code) must parse
/// to completion without spinning.
#[test]
fn frm_header_terminates() {
    let src = b"VERSION 5.00\r\n\
Begin VB.Form Form1 \r\n\
   Caption         =   \"Form1\"\r\n\
   StartUpPosition =   3  'Windows Default\r\n\
End\r\n\
Attribute VB_Name = \"Form1\"\r\n";
    assert!(diagnostic_count(src) < 1000);
}
