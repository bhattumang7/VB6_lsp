//! Data-driven fixture harness: every directory under `tests/fixtures/<case>/`
//! holds `input.bas` (a full VB6 module) and `expected.pcode` (raw bytes), and
//! optionally `input.cls` (a Class module `input.bas` references via `Dim o
//! As New ClassName` — the class's name comes from its own `Attribute
//! VB_Name = "..."` line). `build.rs` discovers each fixture directory at
//! compile time and generates one `#[test]` per case (included below), so
//! adding coverage means adding a fixture directory, not writing test code.
//!
//! Each fixture is run through the real pipeline: ScannerContext -> Parser ->
//! bind (-> bind_with_classes when a class is present) -> lower_module, and
//! the emitted bytes for the target procedure (`proc_index`, default 0) must
//! equal `expected.pcode` byte-for-byte.

use std::collections::HashMap;
use std::fs;
use std::path::PathBuf;

use vb6_codegen::lower_module_with_classes;
use vb6_sema::frontend::ast::{ExprArena, ProcKind};
use vb6_sema::frontend::parser::Parser;
use vb6_sema::frontend::scanner::ScannerContext;
use vb6_sema::sema::{bind, bind_with_classes, ExternalClass, ExternalProperty};

const MODULE_DESC: u16 = 0x0008;

fn fixture_dir(case_name: &str) -> PathBuf {
    PathBuf::from(env!("CARGO_MANIFEST_DIR"))
        .join("tests")
        .join("fixtures")
        .join(case_name)
}

/// The class name declared by a `.cls` file's `Attribute VB_Name = "..."`
/// line (VB6's own convention for a module's name — not parsed by this
/// front end's grammar, which has no "class module" concept; read directly
/// off the source text here since a fixture's `.cls` name is exactly this).
fn class_name_from_source(src: &str) -> String {
    src.lines()
        .find_map(|line| {
            let line = line.trim();
            let rest = line.strip_prefix("Attribute VB_Name")?;
            let rest = rest.trim_start().strip_prefix('=')?;
            let rest = rest.trim();
            rest.strip_prefix('"')?.strip_suffix('"').map(str::to_string)
        })
        .unwrap_or_else(|| panic!("class source has no `Attribute VB_Name = \"...\"` line"))
}

/// Bind a `.cls` source (an ordinary module, front-end-wise — Public fields
/// bind exactly like a Standard module's) into an `ExternalClass` field list,
/// keyed by the class's declared name.
fn bind_class(src: &str) -> (String, ExternalClass) {
    let mut ctx = ScannerContext::new(1, 1, 0x0409);
    ctx.intern_keywords();
    let mut arena = ExprArena::new();
    let mut parser = Parser::new(&mut ctx, src.as_bytes());
    let top = parser.parse_module(&mut arena);
    let spans = std::mem::take(&mut parser.node_spans);
    let vis = std::mem::take(&mut parser.decl_public);
    drop(parser);
    let module = bind(&ctx, &arena, &top, &spans, &vis);
    let fields = module
        .module_vars
        .iter()
        .filter(|v| v.is_public)
        .map(|v| (ctx.symbol(v.sym_id as usize).name.clone(), v.vba_type.clone()))
        .collect();
    // Group Property Get/Let procs by name (declaration order, first-seen),
    // recording which accessors each named property actually has — codegen
    // numbers vtable slots over accessors present, so the Get/Let split (not
    // just the type) matters.
    let mut properties: Vec<ExternalProperty> = Vec::new();
    for proc in &module.procs {
        if !proc.is_public {
            continue;
        }
        let (is_get, is_let) = match proc.kind {
            ProcKind::PropGet => (true, false),
            ProcKind::PropLet => (false, true),
            _ => continue,
        };
        let name = ctx.symbol(proc.sym_id as usize).name.clone();
        let ty = if is_get {
            proc.ret_type.clone()
        } else {
            proc.params.first().map(|p| p.vba_type.clone()).unwrap_or_default()
        };
        if let Some(existing) = properties.iter_mut().find(|p| p.name.eq_ignore_ascii_case(&name)) {
            existing.has_get |= is_get;
            existing.has_let |= is_let;
        } else {
            properties.push(ExternalProperty { name, vba_type: ty, has_get: is_get, has_let: is_let });
        }
    }
    (class_name_from_source(src), ExternalClass { fields, properties })
}

fn compile_module_bytes(src: &str, class_src: Option<&str>) -> Vec<Vec<u8>> {
    let known_classes: HashMap<String, ExternalClass> = match class_src {
        Some(class_src) => {
            let (name, class) = bind_class(class_src);
            HashMap::from([(name.to_ascii_lowercase(), class)])
        }
        None => HashMap::new(),
    };

    let mut ctx = ScannerContext::new(1, 1, 0x0409);
    ctx.intern_keywords();
    let mut arena = ExprArena::new();
    let mut parser = Parser::new(&mut ctx, src.as_bytes());
    let top = parser.parse_module(&mut arena);
    let spans = std::mem::take(&mut parser.node_spans);
    let vis = std::mem::take(&mut parser.decl_public);
    drop(parser);
    let module = bind_with_classes(&ctx, &arena, &top, &spans, &vis, &known_classes);
    lower_module_with_classes(&module, &arena, MODULE_DESC, &known_classes)
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
    let class_src = fs::read_to_string(dir.join("input.cls")).ok();
    let expected = fs::read(dir.join("expected.pcode"))
        .unwrap_or_else(|e| panic!("{case_name}: cannot read expected.pcode: {e}"));
    let proc_index: usize = match fs::read_to_string(dir.join("proc_index")) {
        Ok(s) => s
            .trim()
            .parse()
            .unwrap_or_else(|e| panic!("{case_name}: bad proc_index: {e}")),
        Err(_) => 0,
    };

    let procs = compile_module_bytes(&src, class_src.as_deref());
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
