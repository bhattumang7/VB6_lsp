//! Analysis behaviour, verified through the `vb6-engine` Session.
//!
//! These rewrite the original tree-sitter symbol-table / parser / converter unit
//! tests (`src/analysis/*`, `src/parser/*`). The old assertions inspected the
//! legacy AST and symbol-table structures directly; the engine exposes a
//! different model, so each test here keeps the *intent* of its predecessor —
//! "this declaration is recognized with this name/kind", "this name resolves",
//! "this scope rule holds", "a re-parse reflects the edit" — and asserts it
//! against the engine's public queries (`document_symbols`, `hover`,
//! `definition`, `diagnostics`, `update_file`).
//!
//! Pure position/range geometry from the old `analysis/position.rs` is covered
//! by the engine's own `session::line_index` unit tests and is not duplicated.

use vb6_engine::session::{Session, SymbolKind};

fn sess(src: &str) -> Session {
    Session::from_sources(vec![("m.bas".to_string(), src.as_bytes().to_vec())])
}

fn offset(src: &str, needle: &str) -> u32 {
    src.find(needle).expect("needle present in source") as u32
}

/// (name, kind) for every top-level document symbol in module 0.
fn symbols(s: &Session) -> Vec<(String, SymbolKind)> {
    s.document_symbols(0).into_iter().map(|x| (x.name, x.kind)).collect()
}

fn has(s: &Session, name: &str, kind: SymbolKind) -> bool {
    symbols(s).iter().any(|(n, k)| n == name && *k == kind)
}

// ── builder.rs: declaration recognition ─────────────────────────────────────────

#[test]
fn module_variable_declaration() {
    // was: test_variable_declaration
    let s = sess("Dim x As Integer\n");
    assert!(has(&s, "x", SymbolKind::Variable));
}

#[test]
fn sub_declaration_and_local() {
    // was: test_sub_declaration
    let src = "Sub Main()\n    Dim local As String\nEnd Sub\n";
    let s = sess(src);
    assert!(has(&s, "Main", SymbolKind::Sub));
    // Locals are not document symbols, but the binder must recognize the local:
    // hover on its declaration returns a signature.
    let h = s.hover(0, offset(src, "local")).expect("hover on local decl");
    assert!(h.text.contains("local"), "hover: {}", h.text);
    // And the local is correctly *not* surfaced as a module-level symbol.
    assert!(!symbols(&s).iter().any(|(n, _)| n == "local"));
}

#[test]
fn function_with_parameters() {
    // was: test_function_with_params (param names verified via the hover signature,
    // since the engine has no public "list parameters" API)
    let src = "Function Add(a As Integer, b As Integer) As Integer\n    Add = a + b\nEnd Function\n";
    let s = sess(src);
    assert!(has(&s, "Add", SymbolKind::Function));
    let h = s.hover(0, offset(src, "Add")).expect("hover on Add");
    assert!(h.text.contains('a') && h.text.contains('b'), "signature: {}", h.text);
    assert!(h.text.contains("Integer"), "signature: {}", h.text);
}

#[test]
fn enum_declaration_with_members() {
    // was: test_enum_declaration
    let src = "Public Enum Colors\n    Red = 1\n    Green = 2\n    Blue = 3\nEnd Enum\n";
    let s = sess(src);
    assert!(has(&s, "Colors", SymbolKind::Enum));
    let members = symbols(&s).into_iter().filter(|(_, k)| *k == SymbolKind::EnumMember).count();
    assert_eq!(members, 3);
}

#[test]
fn scope_hierarchy_module_vs_local() {
    // was: test_scope_hierarchy — module var is a document symbol, the local is not.
    let src = "Dim moduleVar As Integer\n\nSub Test()\n    Dim localVar As String\nEnd Sub\n";
    let s = sess(src);
    assert!(has(&s, "moduleVar", SymbolKind::Variable));
    assert!(!symbols(&s).iter().any(|(n, _)| n == "localVar"));
}

// ── symbol_table.rs / scope.rs: lookup & case-insensitivity ─────────────────────

#[test]
fn case_insensitive_resolution() {
    // was: test_case_insensitive_lookup / test_scope_case_insensitive_lookup —
    // a differently-cased use resolves to its declaration.
    let src = "Sub S()\n    Dim Counter As Long\n    counter = 1\nEnd Sub\n";
    let s = sess(src);
    let def = s.definition(0, offset(src, "counter")).expect("definition of `counter`");
    let span = def.span;
    let name = &src.as_bytes()[span.start as usize..(span.start + span.len) as usize];
    assert_eq!(name, b"Counter");
}

#[test]
fn declaration_lookup_by_symbol() {
    // was: test_create_symbol / test_scope_lookup — a module symbol is found by name.
    let s = sess("Public Const MAX_ITEMS As Long = 10\n");
    assert!(has(&s, "MAX_ITEMS", SymbolKind::Constant));
}

// ── parser/mod.rs & tree_sitter.rs: parsing produces the right declarations ──────

#[test]
fn parses_declarations_and_procedures() {
    // was: test_tree_sitter_parse — counts of variables and procedures.
    let src = "Option Explicit\nDim alpha As Integer\nPrivate beta As String\n\nSub Main()\nEnd Sub\n\nFunction Add(x As Integer) As Integer\nEnd Function\n";
    let s = sess(src);
    let syms = symbols(&s);
    let vars = syms.iter().filter(|(_, k)| *k == SymbolKind::Variable).count();
    let procs = syms
        .iter()
        .filter(|(_, k)| matches!(k, SymbolKind::Sub | SymbolKind::Function))
        .count();
    assert_eq!(vars, 2, "{syms:?}");
    assert_eq!(procs, 2, "{syms:?}");
}

#[test]
fn parser_smoke_valid_module_has_no_errors() {
    // was: test_parser_creation / test_basic_parse — valid source parses cleanly.
    let s = sess("Sub Main()\n    Dim x As Integer\nEnd Sub\n");
    assert!(s.diagnostics(0).is_empty());
    assert!(!symbols(&s).is_empty());
}

#[test]
fn incremental_update_reflects_edit() {
    // was: test_incremental_parse — re-feeding the file updates the model.
    let mut s = sess("Dim x As Integer");
    assert_eq!(symbols(&s).len(), 1);
    s.update_file("m.bas", b"Dim x As Integer\nDim y As String".to_vec());
    assert_eq!(symbols(&s).len(), 2);
}

// ── converter.rs: name/kind/visibility carried through ───────────────────────────

#[test]
fn convert_variable_sub_function() {
    // was: test_convert_variable / test_convert_sub / test_convert_function —
    // the three core declaration kinds round-trip to the right symbol kinds.
    let src = "Private mState As Long\n\nSub DoThing()\nEnd Sub\n\nFunction Calc() As Double\nEnd Function\n";
    let s = sess(src);
    assert!(has(&s, "mState", SymbolKind::Variable));
    assert!(has(&s, "DoThing", SymbolKind::Sub));
    assert!(has(&s, "Calc", SymbolKind::Function));
}

// ── diagnostics: undefined name under Option Explicit ────────────────────────────

#[test]
fn undefined_variable_diagnosed_under_option_explicit() {
    let s = sess("Option Explicit\nSub S()\n    x = 1\nEnd Sub\n");
    assert!(!s.diagnostics(0).is_empty());
}
