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
use vb6_sema::sema::{bind, bind_with_classes, AccessorKind, ClassMemberSlot, ExternalClass, VbaType};

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

/// True for an object/reference-assignable type (or Variant) — these
/// synthesize a `Set` accessor alongside `Get`/`Let` when used as a field.
fn is_object_type(ty: &VbaType) -> bool {
    matches!(ty, VbaType::Object | VbaType::Variant | VbaType::UserDefined(_))
}

/// Bind a `.cls` source (an ordinary module, front-end-wise — Public fields
/// bind exactly like a Standard module's) into an `ExternalClass`'s ordered
/// member list (see `ExternalClass::members`), keyed by the class's declared
/// name. Fields and procedures are merged into ONE list sorted by source
/// position (`name_span.start`) — the vtable slot counter runs across the
/// whole class's declaration sequence, not per-kind, so losing cross-kind
/// order here would silently mis-layout slots for any class that interleaves
/// fields and procedures.
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

    let mut events: Vec<(u32, ClassMemberSlot)> = Vec::new();
    for v in &module.module_vars {
        if !v.is_public || v.is_const {
            continue;
        }
        events.push((
            v.name_span.start,
            ClassMemberSlot::Field {
                name: ctx.symbol(v.sym_id as usize).name.clone(),
                vba_type: v.vba_type.clone(),
                is_object: is_object_type(&v.vba_type),
            },
        ));
    }
    for proc in &module.procs {
        if !proc.is_public {
            continue;
        }
        let name = ctx.symbol(proc.sym_id as usize).name.clone();
        let member = match proc.kind {
            ProcKind::PropGet => ClassMemberSlot::PropertyAccessor {
                name,
                vba_type: proc.ret_type.clone(),
                kind: AccessorKind::Get,
            },
            ProcKind::PropLet => ClassMemberSlot::PropertyAccessor {
                name,
                vba_type: proc.params.first().map(|p| p.vba_type.clone()).unwrap_or_default(),
                kind: AccessorKind::Let,
            },
            ProcKind::PropSet => ClassMemberSlot::PropertyAccessor {
                name,
                vba_type: proc.params.first().map(|p| p.vba_type.clone()).unwrap_or_default(),
                kind: AccessorKind::Set,
            },
            ProcKind::Sub | ProcKind::Function => ClassMemberSlot::Method {
                name,
                ret_type: proc.ret_type.clone(),
                params: proc.params.iter().map(|p| (p.vba_type.clone(), p.flags.by_val)).collect(),
            },
        };
        events.push((proc.name_span.start, member));
    }
    events.sort_by_key(|(pos, _)| *pos);
    let members = events.into_iter().map(|(_, m)| m).collect();
    (class_name_from_source(src), ExternalClass { members })
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
