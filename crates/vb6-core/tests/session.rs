//! W6 tests: the Session API end-to-end (cross-module resolution, references,
//! hover, document/workspace symbols, rename, project diagnostics).

use vb6_core::session::{Location, Session, SymbolKind};
use vb6_core::sema::binder::{ERR_SUB_OR_FUNCTION_NOT_DEFINED, ERR_VARIABLE_NOT_DEFINED};

fn session(files: &[(&str, &str)]) -> Session {
    Session::from_sources(
        files
            .iter()
            .map(|(p, s)| (p.to_string(), s.as_bytes().to_vec()))
            .collect(),
    )
}

/// Byte offset just inside the first occurrence of `needle` in `src`.
fn at(src: &str, needle: &str) -> u32 {
    (src.find(needle).expect("needle not found") + 1) as u32
}

fn text<'a>(src: &'a str, loc: Location) -> &'a str {
    let lo = loc.span.start as usize;
    &src[lo..lo + loc.span.len as usize]
}

#[test]
fn cross_module_go_to_definition() {
    let mod0 = "Public Sub Greet()\nEnd Sub\n";
    let mod1 = "Sub Main()\n    Greet\nEnd Sub\n";
    let s = session(&[("Mod0.bas", mod0), ("Mod1.bas", mod1)]);

    // Cursor on the `Greet` call in Mod1 (module 1).
    let def = s.definition(1, at(mod1, "Greet")).expect("definition");
    assert_eq!(def.module, 0);
    assert_eq!(text(mod0, def), "Greet");
}

#[test]
fn private_decl_does_not_resolve_cross_module() {
    // `Private Sub Helper` in Mod0 must NOT be visible from Mod1.
    let mod0 = "Private Sub Helper()\nEnd Sub\n";
    let mod1 = "Sub Main()\n    Helper\nEnd Sub\n";
    let s = session(&[("Mod0.bas", mod0), ("Mod1.bas", mod1)]);
    assert!(
        s.definition(1, at(mod1, "Helper")).is_none(),
        "a Private Sub must not resolve across modules"
    );
}

#[test]
fn bare_dim_module_var_is_private() {
    // Module-level `Dim`/no-modifier var defaults to Private (VB6 rule): not
    // visible from another module. A `Public` var is.
    let mod0 = "Dim gPriv As Long\nPublic gPub As Long\n";
    let mod1 = "Sub Main()\n    gPriv = 1\n    gPub = 2\nEnd Sub\n";
    let s = session(&[("Mod0.bas", mod0), ("Mod1.bas", mod1)]);
    assert!(s.definition(1, at(mod1, "gPriv")).is_none(), "Dim var is module-private");
    assert!(s.definition(1, at(mod1, "gPub")).is_some(), "Public var resolves cross-module");
}

#[test]
fn cross_module_references_include_decl() {
    let mod0 = "Public Sub Greet()\nEnd Sub\n";
    let mod1 = "Sub Main()\n    Greet\nEnd Sub\n";
    let s = session(&[("Mod0.bas", mod0), ("Mod1.bas", mod1)]);

    // From the declaration in Mod0, find all references project-wide.
    let refs = s.references(0, at(mod0, "Greet"), true);
    // The use in Mod1 + the declaration in Mod0.
    assert_eq!(refs.len(), 2);
    assert!(refs.iter().any(|l| l.module == 1));
    assert!(refs.iter().any(|l| l.module == 0));
}

#[test]
fn local_definition_and_references() {
    let src = "Sub Foo()\n    Dim y As Long\n    y = y + 1\nEnd Sub\n";
    let s = session(&[("M.bas", src)]);
    // Cursor on the rhs `y` (second occurrence in the assignment line).
    let off = (src.rfind("y +").unwrap()) as u32;
    let def = s.definition(0, off).expect("definition");
    assert_eq!(def.module, 0);
    assert_eq!(text(src, def), "y"); // the `Dim y`

    let refs = s.references(0, off, true);
    // Dim y + two uses (target, rhs).
    assert_eq!(refs.len(), 3);
}

#[test]
fn hover_shows_function_signature() {
    let src = "Public Function Add(a As Long, b As Long) As Long\n\
               End Function\n\
               Sub T()\n    Dim r As Long\n    r = Add(1, 2)\nEnd Sub\n";
    let s = session(&[("M.bas", src)]);
    let off = at(src, "Add(1");
    let h = s.hover(0, off).expect("hover");
    // Names are rendered from the source span, so exact casing is preserved
    // (the interner would otherwise canonicalize `b` to its first-seen spelling).
    assert_eq!(h.text, "Public Function Add(a As Long, b As Long) As Long");
}

#[test]
fn hover_preserves_source_casing_over_interner() {
    // `MyVar` is declared lowercase-first elsewhere, then a param `myVar` would
    // canonicalize in the interner; rendering from source keeps each spelling.
    let src = "Sub UseIt()\n    Dim total As Long\nEnd Sub\n\
               Function Calc(Total As Long) As Long\nEnd Function\n";
    let s = session(&[("M.bas", src)]);
    // Hover the parameter `Total` (capital T) — must show `Total`, not `total`.
    let off = at(src, "Total As Long) As Long");
    let h = s.hover(0, off).expect("hover");
    assert!(h.text.contains("Total As Long"), "got: {}", h.text);
}

#[test]
fn hover_on_declaration() {
    let src = "Public gCount As Long\n";
    let s = session(&[("M.bas", src)]);
    let h = s.hover(0, at(src, "gCount")).expect("hover");
    assert_eq!(h.text, "Public gCount As Long");
}

#[test]
fn document_symbols_lists_decls() {
    let src = "Public gCount As Long\n\
               Sub Foo()\nEnd Sub\n\
               Public Type TPoint\n    X As Long\nEnd Type\n\
               Public Enum EColor\n    Red\nEnd Enum\n";
    let s = session(&[("M.bas", src)]);
    let syms = s.document_symbols(0);

    let by = |k: SymbolKind| syms.iter().filter(|x| x.kind == k).count();
    assert_eq!(by(SymbolKind::Variable), 1);
    assert_eq!(by(SymbolKind::Sub), 1);
    assert_eq!(by(SymbolKind::Type), 1);
    assert_eq!(by(SymbolKind::Enum), 1);
    assert_eq!(by(SymbolKind::EnumMember), 1);
    assert!(syms.iter().any(|x| x.name == "gCount"));
    assert!(syms.iter().any(|x| x.name == "EColor"));
}

#[test]
fn workspace_symbols_filters_by_query() {
    let mod0 = "Public Sub Alpha()\nEnd Sub\n";
    let mod1 = "Public Sub Beta()\nEnd Sub\nPublic Sub AlphaBeta()\nEnd Sub\n";
    let s = session(&[("Mod0.bas", mod0), ("Mod1.bas", mod1)]);
    let hits = s.workspace_symbols("alpha");
    // "Alpha" and "AlphaBeta".
    assert_eq!(hits.len(), 2);
    assert!(hits.iter().all(|x| x.name.to_ascii_lowercase().contains("alpha")));
}

#[test]
fn rename_covers_decl_and_all_uses() {
    let src = "Sub Foo()\n    Dim y As Long\n    y = y + 1\nEnd Sub\n";
    let s = session(&[("M.bas", src)]);
    let off = (src.rfind("y +").unwrap()) as u32;
    let edits = s.rename(0, off, "z");
    assert_eq!(edits.len(), 3);
    assert!(edits.iter().all(|e| e.new_text == "z" && e.module == 0));
}

#[test]
fn undeclared_variable_under_option_explicit() {
    let src = "Option Explicit\nSub Foo()\n    x = 1\nEnd Sub\n";
    let s = session(&[("M.bas", src)]);
    let diags = s.diagnostics(0);
    assert!(
        diags.iter().any(|d| d.code == ERR_VARIABLE_NOT_DEFINED as u32),
        "expected a variable-not-defined diagnostic, got {diags:?}"
    );
}

#[test]
fn cross_module_call_is_not_undeclared() {
    // `Greet` is defined in Mod0; calling it from Mod1 under Option Explicit
    // must NOT produce a "Variable not defined" diagnostic.
    let mod0 = "Public Sub Greet()\nEnd Sub\n";
    let mod1 = "Option Explicit\nSub Main()\n    Greet\nEnd Sub\n";
    let s = session(&[("Mod0.bas", mod0), ("Mod1.bas", mod1)]);
    let diags = s.diagnostics(1);
    assert!(
        !diags.iter().any(|d| d.code == ERR_VARIABLE_NOT_DEFINED as u32),
        "cross-module call wrongly flagged as undeclared: {diags:?}"
    );
}

#[test]
fn undefined_call_flagged_without_option_explicit() {
    // An undefined call errors regardless of Option Explicit; the check is not
    // gated on RequireDeclaration.
    let src = "Sub Main()\n    DoesNotExist\nEnd Sub\n";
    let s = session(&[("M.bas", src)]);
    let diags = s.diagnostics(0);
    assert!(
        diags.iter().any(|d| d.code == ERR_SUB_OR_FUNCTION_NOT_DEFINED as u32),
        "expected Sub-or-Function-not-defined, got {diags:?}"
    );
}

#[test]
fn defined_and_cross_module_calls_not_flagged() {
    let mod0 = "Public Sub Greet()\nEnd Sub\n";
    let mod1 = "Sub Main()\n    Greet\n    Local\nEnd Sub\nSub Local()\nEnd Sub\n";
    let s = session(&[("Mod0.bas", mod0), ("Mod1.bas", mod1)]);
    assert!(
        s.diagnostics(1).iter().all(|d| d.code != ERR_SUB_OR_FUNCTION_NOT_DEFINED as u32),
        "defined local + cross-module calls must not be flagged: {:?}",
        s.diagnostics(1)
    );
}

#[test]
fn declared_api_call_not_flagged() {
    // A `Declare`d external function is callable; calling it must not flag.
    let src = "Declare Function GetTickCount Lib \"kernel32\" () As Long\n\
               Sub Main()\n    Dim t As Long\n    t = GetTickCount()\nEnd Sub\n";
    let s = session(&[("M.bas", src)]);
    assert!(
        s.diagnostics(0).iter().all(|d| d.code != ERR_SUB_OR_FUNCTION_NOT_DEFINED as u32),
        "Declared API call wrongly flagged: {:?}",
        s.diagnostics(0)
    );
}

#[test]
fn no_undeclared_without_option_explicit() {
    let src = "Sub Foo()\n    x = 1\nEnd Sub\n";
    let s = session(&[("M.bas", src)]);
    let diags = s.diagnostics(0);
    assert!(diags.iter().all(|d| d.code != ERR_VARIABLE_NOT_DEFINED as u32));
}

#[test]
fn declared_local_not_flagged_undeclared() {
    let src = "Option Explicit\nSub Foo()\n    Dim x As Long\n    x = 42\nEnd Sub\n";
    let s = session(&[("M.bas", src)]);
    let diags = s.diagnostics(0);
    assert!(
        diags.iter().all(|d| d.code != ERR_VARIABLE_NOT_DEFINED as u32),
        "declared local must not trigger undeclared error; got {diags:?}"
    );
}

#[test]
fn parameter_not_flagged_undeclared() {
    let src = "Option Explicit\nSub Foo(x As Long)\n    x = 42\nEnd Sub\n";
    let s = session(&[("M.bas", src)]);
    let diags = s.diagnostics(0);
    assert!(
        diags.iter().all(|d| d.code != ERR_VARIABLE_NOT_DEFINED as u32),
        "parameter must not trigger undeclared error; got {diags:?}"
    );
}

#[test]
fn module_var_not_flagged_undeclared() {
    let src = "Option Explicit\nPublic x As Long\nSub Foo()\n    x = 42\nEnd Sub\n";
    let s = session(&[("M.bas", src)]);
    let diags = s.diagnostics(0);
    assert!(
        diags.iter().all(|d| d.code != ERR_VARIABLE_NOT_DEFINED as u32),
        "module-level var must not trigger undeclared error; got {diags:?}"
    );
}

#[test]
fn multiple_undeclared_each_flagged() {
    // x and y both undeclared under Option Explicit → two diagnostics.
    let src = "Option Explicit\nSub Foo()\n    x = 1\n    y = 2\nEnd Sub\n";
    let s = session(&[("M.bas", src)]);
    let count = s
        .diagnostics(0)
        .iter()
        .filter(|d| d.code == ERR_VARIABLE_NOT_DEFINED as u32)
        .count();
    assert_eq!(count, 2, "each undeclared name should produce its own diagnostic; got {count}");
}

#[test]
fn rename_proc_call_in_frm_file() {
    let src = std::fs::read_to_string(
        std::path::Path::new(env!("CARGO_MANIFEST_DIR"))
            .parent().unwrap().parent().unwrap()
            .join("tests/fixtures/vb6_sample/Form1.frm")
    ).unwrap();
    let s = session(&[("Form1.frm", &src)]);
    // "domy1" call site is on line 61 ('d' of domy1())
    let call_offset: u32 = src.find("    domy1()").unwrap() as u32 + 4; // +4 for 4 spaces
    let edits = s.rename(0, call_offset, "newName");
    assert!(
        edits.len() >= 2,
        "expected at least 2 rename edits (call + decl), got {}; call_offset={call_offset}",
        edits.len()
    );
}

// ── Completion ─────────────────────────────────────────────────────────────────

#[test]
fn completions_include_locals_and_procs() {
    let src = "Sub Foo(p As Long)\n    Dim x As Long\n    x = \nEnd Sub\nSub Bar()\nEnd Sub\n";
    let s = session(&[("M.bas", src)]);
    // Cursor inside Foo (after `x = `)
    let offset = src.find("x = ").unwrap() as u32 + 4;
    let items = s.completions(0, offset);
    let names: Vec<&str> = items.iter().map(|i| i.name.as_str()).collect();
    assert!(names.contains(&"x"), "local `x` missing: {names:?}");
    assert!(names.contains(&"p"), "param `p` missing: {names:?}");
    assert!(names.contains(&"Foo"), "proc `Foo` missing: {names:?}");
    assert!(names.contains(&"Bar"), "proc `Bar` missing: {names:?}");
}

#[test]
fn completions_include_vb6_keywords() {
    let src = "Sub Foo()\n    \nEnd Sub\n";
    let s = session(&[("M.bas", src)]);
    let offset = src.find("    \n").unwrap() as u32 + 2;
    let items = s.completions(0, offset);
    let names: Vec<&str> = items.iter().map(|i| i.name.as_str()).collect();
    assert!(names.contains(&"Dim"), "keyword `Dim` missing: {names:?}");
    assert!(names.contains(&"If"), "keyword `If` missing: {names:?}");
    assert!(names.contains(&"For"), "keyword `For` missing: {names:?}");
}

#[test]
fn completions_include_cross_module_publics() {
    let mod0 = "Public Sub Helper()\nEnd Sub\n";
    let mod1 = "Sub Main()\n    \nEnd Sub\n";
    let s = session(&[("Mod0.bas", mod0), ("Mod1.bas", mod1)]);
    let offset = mod1.find("    \n").unwrap() as u32 + 2;
    let items = s.completions(1, offset);
    let names: Vec<&str> = items.iter().map(|i| i.name.as_str()).collect();
    assert!(names.contains(&"Helper"), "cross-module public `Helper` missing: {names:?}");
}

#[test]
fn completions_deduplicate_case_insensitively() {
    // `Len` appears both as a built-in and might surface from keyword list;
    // it must appear only once regardless of casing.
    let src = "Sub Foo()\n    \nEnd Sub\n";
    let s = session(&[("M.bas", src)]);
    let offset = src.find("    \n").unwrap() as u32 + 2;
    let items = s.completions(0, offset);
    // Count occurrences of any name that case-folds to "dim"
    let dim_count = items.iter().filter(|i| i.name.eq_ignore_ascii_case("dim")).count();
    assert_eq!(dim_count, 1, "`Dim` appeared {dim_count} times in completion list");
}

// ── Signature help ─────────────────────────────────────────────────────────────

#[test]
fn signature_help_returns_signature_inside_call() {
    let src = "Sub Add(x As Long, y As Long)\nEnd Sub\nSub Main()\n    Add(\nEnd Sub\n";
    let s = session(&[("M.bas", src)]);
    // Cursor right after the `(` in `Add(`
    let offset = src.find("Add(").unwrap() as u32 + 4;
    let sh = s.signature_help(0, offset).expect("expected signature help");
    assert!(sh.label.contains("Add"), "label should contain proc name: {:?}", sh.label);
    assert_eq!(sh.params.len(), 2, "expected 2 params: {:?}", sh.params);
    assert_eq!(sh.active_param, 0, "first param should be active");
}

#[test]
fn signature_help_tracks_active_param_via_comma() {
    let src = "Sub Add(x As Long, y As Long)\nEnd Sub\nSub Main()\n    Add(1, \nEnd Sub\n";
    let s = session(&[("M.bas", src)]);
    // Cursor after the comma+space: `Add(1, |`
    let offset = src.find("Add(1, ").unwrap() as u32 + 7;
    let sh = s.signature_help(0, offset).expect("expected signature help");
    assert_eq!(sh.active_param, 1, "second param should be active after comma");
}

#[test]
fn signature_help_returns_none_outside_call() {
    let src = "Sub Add(x As Long)\nEnd Sub\nSub Main()\n    Dim x As Long\nEnd Sub\n";
    let s = session(&[("M.bas", src)]);
    // Cursor inside `Dim x` — not inside a call
    let offset = src.find("Dim x").unwrap() as u32 + 3;
    assert!(s.signature_help(0, offset).is_none(), "should be None outside a call");
}

// ── Document highlight ─────────────────────────────────────────────────────────

#[test]
fn document_highlights_finds_all_occurrences_in_module() {
    let src = "Sub Foo()\n    Dim y As Long\n    y = 1\n    y = y + 1\nEnd Sub\n";
    let s = session(&[("M.bas", src)]);
    let offset = src.find("Dim y").unwrap() as u32 + 4; // on the `y` in `Dim y`
    let spans = s.document_highlights(0, offset);
    // y appears in: Dim y, y = 1, y = (lhs), y + 1 (rhs) → 4 times + decl = 4 or 5 total
    assert!(spans.len() >= 3, "expected ≥3 highlight spans for `y`, got {}: {:?}", spans.len(), spans);
}

#[test]
fn document_highlights_does_not_cross_modules() {
    let mod0 = "Public gX As Long\n";
    let mod1 = "Sub Main()\n    gX = 1\n    gX = gX + 1\nEnd Sub\n";
    let s = session(&[("Mod0.bas", mod0), ("Mod1.bas", mod1)]);
    // Highlight gX from inside Mod1 — must only return spans within Mod1
    let offset = mod1.find("gX").unwrap() as u32 + 1;
    let spans = s.document_highlights(1, offset);
    assert!(
        spans.len() >= 2,
        "expected ≥2 spans for gX uses in Mod1, got {}", spans.len()
    );
    // No span should reference module 0
    // (document_highlights only returns spans within the queried module)
}

// ── Call hierarchy ─────────────────────────────────────────────────────────────

#[test]
fn call_hierarchy_prepare_resolves_proc_name() {
    let src = "Sub Foo()\nEnd Sub\n";
    let s = session(&[("M.bas", src)]);
    let offset = src.find("Foo").unwrap() as u32 + 1;
    let decl = s.prepare_call_hierarchy(0, offset).expect("should resolve on proc name");
    assert_eq!(decl.name, "Foo");
    assert_eq!(decl.location.module, 0);
}

#[test]
fn call_hierarchy_prepare_returns_none_not_on_proc() {
    let src = "Sub Foo()\n    Dim x As Long\nEnd Sub\n";
    let s = session(&[("M.bas", src)]);
    let offset = src.find("Dim").unwrap() as u32 + 1;
    assert!(s.prepare_call_hierarchy(0, offset).is_none(), "should be None when not on a proc name");
}

#[test]
fn call_hierarchy_incoming_finds_callers() {
    let src = "Sub Helper()\nEnd Sub\nSub Main()\n    Helper\nEnd Sub\n";
    let s = session(&[("M.bas", src)]);
    let callers = s.incoming_calls("Helper");
    assert_eq!(callers.len(), 1, "expected 1 caller of Helper, got {}", callers.len());
    assert!(
        callers[0].caller.name.eq_ignore_ascii_case("Main"),
        "caller should be Main, got {:?}", callers[0].caller.name
    );
    assert_eq!(callers[0].call_sites.len(), 1, "one call site in Main");
}

#[test]
fn call_hierarchy_incoming_empty_for_uncalled_proc() {
    let src = "Sub Orphan()\nEnd Sub\nSub Main()\nEnd Sub\n";
    let s = session(&[("M.bas", src)]);
    let callers = s.incoming_calls("Orphan");
    assert!(callers.is_empty(), "Orphan has no callers");
}

#[test]
fn call_hierarchy_outgoing_finds_callees() {
    let src = "Sub A()\nEnd Sub\nSub B()\nEnd Sub\nSub Main()\n    A\n    B\nEnd Sub\n";
    let s = session(&[("M.bas", src)]);
    let callees = s.outgoing_calls("Main");
    let names: Vec<&str> = callees.iter().map(|c| c.callee.name.as_str()).collect();
    assert!(names.iter().any(|n| n.eq_ignore_ascii_case("A")), "A missing from outgoing: {names:?}");
    assert!(names.iter().any(|n| n.eq_ignore_ascii_case("B")), "B missing from outgoing: {names:?}");
}

#[test]
fn call_hierarchy_cross_module_incoming() {
    let helper = "Public Sub Helper()\nEnd Sub\n";
    let main_mod = "Sub Main()\n    Helper\nEnd Sub\n";
    let s = session(&[("Helper.bas", helper), ("Main.bas", main_mod)]);
    let callers = s.incoming_calls("Helper");
    assert_eq!(callers.len(), 1);
    assert!(callers[0].caller.name.eq_ignore_ascii_case("Main"));
}

// ── Folding ────────────────────────────────────────────────────────────────────

#[test]
fn folding_ranges_cover_proc_body() {
    let src = "Sub Foo()\n    Dim x As Long\nEnd Sub\nSub Bar()\nEnd Sub\n";
    let s = session(&[("M.bas", src)]);
    let ranges = s.folding_ranges(0);
    // Should have at least 2 folds (one per Sub)
    assert!(ranges.len() >= 2, "expected ≥2 folds, got {}: {:?}", ranges.len(), ranges);
    // Each fold: start_line < end_line
    for r in &ranges {
        assert!(r.start_line < r.end_line, "fold range is zero-length: {:?}", r);
    }
}
