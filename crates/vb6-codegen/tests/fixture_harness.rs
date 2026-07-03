//! Data-driven fixture harness: every directory under `tests/fixtures/<case>/`
//! holds `input.bas` (a full VB6 module) and `expected.pcode` (raw bytes).
//! `build.rs` discovers each fixture directory at compile time and generates
//! one `#[test]` per case (included below), so adding coverage means adding
//! a fixture directory, not writing test code.
//!
//! Each fixture is run through the real pipeline: ScannerContext -> Parser ->
//! bind -> lower_module, and the emitted bytes for the target procedure
//! (`proc_index`, default 0) must equal `expected.pcode` byte-for-byte.

use std::fs;
use std::path::PathBuf;

use vb6_codegen::lower_module;
use vb6_sema::frontend::ast::ExprArena;
use vb6_sema::frontend::parser::Parser;
use vb6_sema::frontend::scanner::ScannerContext;
use vb6_sema::sema::bind;

const MODULE_DESC: u16 = 0x0008;

fn fixture_dir(case_name: &str) -> PathBuf {
    PathBuf::from(env!("CARGO_MANIFEST_DIR"))
        .join("tests")
        .join("fixtures")
        .join(case_name)
}

fn compile_module_bytes(src: &str) -> Vec<Vec<u8>> {
    let mut ctx = ScannerContext::new(1, 1, 0x0409);
    ctx.intern_keywords();
    let mut arena = ExprArena::new();
    let mut parser = Parser::new(&mut ctx, src.as_bytes());
    let top = parser.parse_module(&mut arena);
    let spans = std::mem::take(&mut parser.node_spans);
    let vis = std::mem::take(&mut parser.decl_public);
    drop(parser);
    let module = bind(&ctx, &arena, &top, &spans, &vis);
    lower_module(&module, &arena, MODULE_DESC)
        .unwrap_or_else(|e| panic!("lower_module failed: {e:?}"))
}

fn hex_window(bytes: &[u8], center: usize, radius: usize) -> String {
    let start = center.saturating_sub(radius);
    let end = (center + radius).min(bytes.len());
    bytes[start..end]
        .iter()
        .map(|b| format!("{b:02x}"))
        .collect::<Vec<_>>()
        .join(" ")
}

fn run_fixture(case_name: &str) {
    let dir = fixture_dir(case_name);
    let src = fs::read_to_string(dir.join("input.bas"))
        .unwrap_or_else(|e| panic!("{case_name}: cannot read input.bas: {e}"));
    let expected = fs::read(dir.join("expected.pcode"))
        .unwrap_or_else(|e| panic!("{case_name}: cannot read expected.pcode: {e}"));
    let proc_index: usize = match fs::read_to_string(dir.join("proc_index")) {
        Ok(s) => s
            .trim()
            .parse()
            .unwrap_or_else(|e| panic!("{case_name}: bad proc_index: {e}")),
        Err(_) => 0,
    };

    let procs = compile_module_bytes(&src);
    let actual = procs.get(proc_index).unwrap_or_else(|| {
        panic!(
            "{case_name}: proc index {proc_index} out of range (module lowered {} procs)",
            procs.len()
        )
    });

    if actual.as_slice() != expected.as_slice() {
        let mismatch_at = actual
            .iter()
            .zip(expected.iter())
            .position(|(a, b)| a != b)
            .unwrap_or_else(|| actual.len().min(expected.len()));
        panic!(
            "{case_name}: byte mismatch at offset {mismatch_at} (expected {} bytes, actual {} bytes)\n  expected: .. {} ..\n  actual:   .. {} ..",
            expected.len(),
            actual.len(),
            hex_window(&expected, mismatch_at, 8),
            hex_window(actual, mismatch_at, 8),
        );
    }
}

include!(concat!(env!("OUT_DIR"), "/fixture_tests.rs"));
