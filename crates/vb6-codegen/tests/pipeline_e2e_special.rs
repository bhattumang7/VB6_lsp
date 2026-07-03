//! The two pipeline_e2e cases that don't fit the byte-exact fixture harness
//! (`fixture_harness.rs`): one asserts a gated *lowering error*, the other
//! asserts a partial byte-range / substring property rather than an exact
//! full-stream comparison. Everything else migrated to
//! `tests/fixtures/<case>/`.

use vb6_codegen::lower_proc;
use vb6_sema::frontend::ast::ExprArena;
use vb6_sema::frontend::parser::Parser;
use vb6_sema::frontend::scanner::ScannerContext;
use vb6_sema::sema::bind;

fn try_compile(decl: &str, stmt: &str) -> Result<Vec<u8>, vb6_codegen::LowerError> {
    let src =
        format!("Attribute VB_Name = \"Module1\"\r\nSub Main()\r\n{decl}\r\n{stmt}\r\nEnd Sub\r\n");
    let mut ctx = ScannerContext::new(1, 1, 0x0409);
    ctx.intern_keywords();
    let mut arena = ExprArena::new();
    let mut parser = Parser::new(&mut ctx, src.as_bytes());
    let top = parser.parse_module(&mut arena);
    let spans = std::mem::take(&mut parser.node_spans);
    let vis = std::mem::take(&mut parser.decl_public);
    drop(parser);
    let module = bind(&ctx, &arena, &top, &spans, &vis);
    lower_proc(&module, 0, &arena, 0x0008)
}

fn compile_module(src: &str, module_desc: u16) -> Vec<Vec<u8>> {
    let mut ctx = ScannerContext::new(1, 1, 0x0409);
    ctx.intern_keywords();
    let mut arena = ExprArena::new();
    let mut parser = Parser::new(&mut ctx, src.as_bytes());
    let top = parser.parse_module(&mut arena);
    let spans = std::mem::take(&mut parser.node_spans);
    let vis = std::mem::take(&mut parser.decl_public);
    drop(parser);
    let module = bind(&ctx, &arena, &top, &spans, &vis);
    vb6_codegen::lower_module(&module, &arena, module_desc)
        .unwrap_or_else(|e| panic!("lower_module failed: {e:?}"))
}

#[test]
fn e2e_instr_four_args_gated() {
    // An explicit compare-mode argument (4-arg form) needs Option Compare handling
    // → gated, not mis-emitted.
    assert!(try_compile(
        "Dim a As String, b As String, r As Long",
        "r = InStr(1, a, b, 1)"
    )
    .is_err());
}

#[test]
fn e2e_module_global_string_pool_dedup() {
    // A repeated string value is interned once across the module: proc 1's "aaa"
    // reuses index 0; its new "ccc" gets index 1.
    let procs = compile_module(
        "Attribute VB_Name = \"Module1\"\r\n\
         Sub Main()\r\n\
         Dim s As String\r\n\
         s = \"aaa\"\r\n\
         End Sub\r\n\
         Sub Two()\r\n\
         Dim t As String, u As String\r\n\
         t = \"aaa\"\r\n\
         u = \"ccc\"\r\n\
         End Sub\r\n",
        0x0008,
    );
    assert_eq!(procs.len(), 2);
    // proc 1 first stores "aaa" (reused index 0) then "ccc" (new index 1).
    assert_eq!(&procs[1][0..3], &[0x1b, 0x00, 0x00]);
    assert!(procs[1].windows(3).any(|w| w == [0x1b, 0x01, 0x00]));
}
