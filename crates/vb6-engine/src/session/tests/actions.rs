use crate::session::{CodeAction, Session, TextEdit};

fn apply(src: &str, edits: &[TextEdit]) -> String {
    let mut bytes = src.as_bytes().to_vec();
    let mut es: Vec<&TextEdit> = edits.iter().collect();
    es.sort_by_key(|e| std::cmp::Reverse(e.span.start));
    for e in es {
        let s = e.span.start as usize;
        let end = s + e.span.len as usize;
        bytes.splice(s..end, e.new_text.bytes());
    }
    String::from_utf8(bytes).unwrap()
}

fn actions_at(src: &str, needle: &str) -> (Session, Vec<CodeAction>) {
    let s = Session::from_sources(vec![("m.bas".to_string(), src.as_bytes().to_vec())]);
    let off = src.find(needle).expect("needle in source") as u32;
    let acts = s.code_actions(0, off, off);
    (s, acts)
}

#[test]
fn declare_undefined_variable() {
    let src = "Option Explicit\nSub Foo()\n    x = 1\nEnd Sub\n";
    let (_s, acts) = actions_at(src, "x =");
    let a = acts
        .iter()
        .find(|a| a.title.contains("Declare"))
        .expect("declare action");
    let out = apply(src, &a.edits);
    assert!(out.contains("    Dim x As Variant\n    x = 1"), "got:\n{out}");
}

#[test]
fn create_missing_sub() {
    let src = "Sub Foo()\n    Bar\nEnd Sub\n";
    let (_s, acts) = actions_at(src, "Bar");
    let a = acts
        .iter()
        .find(|a| a.title.contains("Create Sub"))
        .expect("create sub action");
    let out = apply(src, &a.edits);
    assert!(out.contains("Private Sub Bar()\nEnd Sub"), "got:\n{out}");
}

#[test]
fn expand_single_line_if() {
    let src = "If x > 1 Then y = 2\n";
    let (_s, acts) = actions_at(src, "If");
    let a = acts
        .iter()
        .find(|a| a.kind == crate::session::CodeActionKind::RefactorRewrite)
        .expect("toggle action");
    let out = apply(src, &a.edits);
    assert_eq!(out, "If x > 1 Then\n    y = 2\nEnd If\n");
}

#[test]
fn expand_preserves_cp1252_string_literal() {
    // Source bytes are Windows-1252: 0x92 is a right single quote (’).
    let mut src = b"If x Then y = \"a".to_vec();
    src.push(0x92);
    src.extend_from_slice(b"b\"\n");
    let s = Session::from_sources(vec![("m.bas".to_string(), src.clone())]);
    let acts = s.code_actions(0, 0, 0);
    let a = acts
        .iter()
        .find(|a| a.kind == crate::session::CodeActionKind::RefactorRewrite)
        .expect("toggle action");
    // The relocated string text must carry the correct Unicode quote, not the
    // Latin-1 mis-decode (U+0092).
    let moved = &a.edits[0].new_text;
    assert!(moved.contains('\u{2019}'), "got: {moved:?}");
    assert!(!moved.contains('\u{0092}'), "got: {moved:?}");
}

#[test]
fn collapse_block_if() {
    let src = "If x > 1 Then\n    y = 2\nEnd If\n";
    let (_s, acts) = actions_at(src, "If");
    let a = acts
        .iter()
        .find(|a| a.kind == crate::session::CodeActionKind::RefactorRewrite)
        .expect("toggle action");
    let out = apply(src, &a.edits);
    assert_eq!(out, "If x > 1 Then y = 2\n");
}
