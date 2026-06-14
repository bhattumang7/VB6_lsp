//! Coverage for the engine↔LSP glue (`src/engine_glue.rs`).
//!
//! These conversion functions are otherwise only reached through the live LSP /
//! MCP servers, so the analysis tests (which call the engine `Session` directly)
//! never exercise them. Here we drive each conversion against a real `Session`.

use tower_lsp::lsp_types as lsp;
use vb6_engine::frontend::ast::Span;
use vb6_engine::session::Session;
use vb6_lsp::engine_glue;

const URI: &str = "file:///test.bas";

fn session(src: &str) -> Session {
    Session::from_sources(vec![(URI.to_string(), src.as_bytes().to_vec())])
}

#[test]
fn doc_key_round_trips_uri() {
    let uri = lsp::Url::parse(URI).unwrap();
    assert_eq!(engine_glue::doc_key(&uri), URI);
}

#[test]
fn to_cp1252_encodes_single_byte() {
    // Windows-1252: every char is one byte; “é” → 0xE9.
    assert_eq!(engine_glue::to_cp1252("Aé"), vec![0x41, 0xE9]);
}

#[test]
fn offset_at_and_span_range_round_trip() {
    let src = "Sub Main()\n    Dim count As Long\nEnd Sub\n";
    let s = session(src);

    // `count` starts at line 1, character 8 (after "    Dim ").
    let pos = lsp::Position { line: 1, character: 8 };
    let off = engine_glue::offset_at(&s, 0, pos).expect("offset for in-range position");
    assert_eq!(off as usize, src.find("count").unwrap());

    // Convert a known engine span back to an LSP range and check it lands on the
    // same line/character.
    let span = Span { start: off, len: 5 };
    let range = engine_glue::span_range(&s, 0, span);
    assert_eq!(range.start, pos);
    assert_eq!(range.end, lsp::Position { line: 1, character: 13 });
}

#[test]
fn offset_at_out_of_range_is_handled() {
    let s = session("Sub Main()\nEnd Sub\n");
    // A wildly out-of-range line still returns *some* offset (clamped), never panics.
    let pos = lsp::Position { line: 9999, character: 0 };
    let _ = engine_glue::offset_at(&s, 0, pos);
}

#[test]
fn definition_maps_to_lsp_location() {
    let src = "Sub Main()\n    Dim count As Long\n    count = count + 1\nEnd Sub\n";
    let s = session(src);
    let use_off = (src.rfind("count").unwrap()) as u32;
    let def = s.definition(0, use_off).expect("definition of count");
    let loc = engine_glue::location(&s, def).expect("lsp location");
    assert_eq!(loc.uri.as_str(), URI);
    assert_eq!(loc.range.start, lsp::Position { line: 1, character: 8 });
}

#[test]
fn hover_wraps_signature_in_code_fence() {
    let src = "Sub Main()\n    Dim count As Long\n    count = 1\nEnd Sub\n";
    let s = session(src);
    let off = (src.find("count").unwrap()) as u32;
    let h = s.hover(0, off).expect("engine hover");
    let lh = engine_glue::hover(&s, 0, h);
    match lh.contents {
        lsp::HoverContents::Markup(m) => {
            assert!(m.value.contains("```vb"), "hover markup: {}", m.value);
            assert!(m.value.contains("count"), "hover markup: {}", m.value);
        }
        other => panic!("expected markup hover, got {other:?}"),
    }
    assert!(lh.range.is_some());
}

#[test]
fn document_symbols_lists_declarations() {
    let s = session("Public gValue As Long\n\nSub Main()\nEnd Sub\n");
    let syms = engine_glue::document_symbols(&s, 0);
    let names: Vec<_> = syms.iter().map(|s| s.name.as_str()).collect();
    assert!(names.contains(&"Main"), "{names:?}");
    assert!(names.contains(&"gValue"), "{names:?}");
}

#[test]
fn workspace_symbols_filters_by_query() {
    let s = session("Sub Alpha()\nEnd Sub\nSub Beta()\nEnd Sub\n");
    let syms = engine_glue::workspace_symbols(&s, "alph");
    assert!(syms.iter().any(|s| s.name == "Alpha"), "{syms:?}");
    assert!(!syms.iter().any(|s| s.name == "Beta"), "{syms:?}");
}

#[test]
fn diagnostics_for_maps_engine_diagnostics() {
    let s = session("Option Explicit\nSub S()\n    x = 1\nEnd Sub\n");
    let diags = engine_glue::diagnostics_for(&s, URI);
    assert!(!diags.is_empty(), "expected an undefined-variable diagnostic");
    let d = &diags[0];
    assert_eq!(d.severity, Some(lsp::DiagnosticSeverity::ERROR));
    assert_eq!(d.source.as_deref(), Some("vb6-lsp"));
}

#[test]
fn formatting_and_workspace_edit_produce_edits() {
    // Lowercase keyword + no indentation → the formatter emits edits.
    let s = session("sub Main()\nDim x As Long\nEnd Sub\n");
    let edits = engine_glue::formatting(&s, 0);
    assert!(!edits.is_empty(), "expected formatting edits");

    // Feed engine format edits through the workspace-edit grouping.
    let we = engine_glue::workspace_edit(&s, s.format(0));
    let changes = we.changes.expect("workspace edit changes");
    let uri = lsp::Url::parse(URI).unwrap();
    assert!(changes.get(&uri).map(|v| !v.is_empty()).unwrap_or(false));
}

#[test]
fn semantic_tokens_are_delta_encoded() {
    let s = session("Sub Main()\n    Dim count As Long\n    count = 1\nEnd Sub\n");
    let toks = engine_glue::semantic_tokens(&s, 0);
    assert!(!toks.data.is_empty(), "expected semantic tokens");
}

#[test]
fn code_actions_runs_over_a_range() {
    let s = session("Sub S()\n    x = 1\nEnd Sub\n");
    let range = lsp::Range {
        start: lsp::Position { line: 1, character: 4 },
        end: lsp::Position { line: 1, character: 9 },
    };
    // Should not panic; may be empty depending on available quick-fixes.
    let _ = engine_glue::code_actions(&s, 0, range);
}
