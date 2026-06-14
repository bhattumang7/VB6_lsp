//! Structural tests for VB6 semantic analysis.

use vb6_core::frontend::ast::{ExprArena, NodeId};
use vb6_core::frontend::parser::{Parser, ModuleKind};
use vb6_core::frontend::scanner::ScannerContext;
use vb6_core::context::CompilerContext;

fn analyze(src: &[u8]) -> CompilerContext {
    let mut ctx = CompilerContext::new();
    let mut sc = ScannerContext::new(1, 1, 0x0409);
    sc.intern_keywords();
    let mut arena = ExprArena::new();
    let mut parser = Parser::new(&mut sc, src);
    let _nodes = parser.parse_module(&mut arena);
    // Binder/Sema would run here in a full compiler.
    // These tests currently focus on the infrastructure.
    ctx
}

#[test]
fn compiler_context_initialization() {
    let ctx = CompilerContext::new();
    assert!(ctx.decls.is_empty());
    assert!(ctx.scopes.is_empty());
}

#[test]
fn basic_analysis_smoke_test() {
    let _ctx = analyze(b"Dim x As Integer");
}
