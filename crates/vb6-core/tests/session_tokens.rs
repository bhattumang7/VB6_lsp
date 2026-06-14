//! Semantic-token (syntax highlighting) classification tests.

use vb6_core::session::{SemTokenKind, Session};

fn classify(src: &str) -> Vec<(SemTokenKind, String)> {
    let s = Session::from_sources(vec![("M.bas".into(), src.as_bytes().to_vec())]);
    s.semantic_tokens(0)
        .into_iter()
        .map(|t| {
            let lo = t.span.start as usize;
            (t.kind, src[lo..lo + t.span.len as usize].to_string())
        })
        .collect()
}

#[test]
fn classifies_declarations_uses_keywords_and_literals() {
    let src = "Public Function Add(a As Long) As Long\n\
               \x20\x20\x20\x20Dim total As Long\n\
               \x20\x20\x20\x20total = a + 1 ' a note\n\
               End Function\n";
    let t = classify(src);
    let has = |k: SemTokenKind, txt: &str| t.iter().any(|(kk, s)| *kk == k && s == txt);

    assert!(has(SemTokenKind::Function, "Add"), "decl name -> Function: {t:?}");
    assert!(has(SemTokenKind::Parameter, "a"), "param -> Parameter");
    assert!(has(SemTokenKind::Variable, "total"), "local -> Variable");
    assert!(has(SemTokenKind::Keyword, "Public"), "Public -> Keyword");
    assert!(has(SemTokenKind::Keyword, "Dim"), "Dim -> Keyword");
    assert!(has(SemTokenKind::Number, "1"), "1 -> Number");
    assert!(t.iter().any(|(k, _)| *k == SemTokenKind::Comment), "comment present");

    // `a` and `total` appear twice each (declaration + use), all classified.
    assert_eq!(t.iter().filter(|(k, s)| *k == SemTokenKind::Parameter && s == "a").count(), 2);
    assert_eq!(t.iter().filter(|(k, s)| *k == SemTokenKind::Variable && s == "total").count(), 2);
}

#[test]
fn classifies_string_literals() {
    let src = "Sub F()\n    Dim s As String\n    s = \"hi\" ' note\nEnd Sub\n";
    let t = classify(src);
    assert!(
        t.iter().any(|(k, txt)| *k == SemTokenKind::String && txt.contains("hi")),
        "string literal classified: {t:?}"
    );
    assert!(t.iter().any(|(k, _)| *k == SemTokenKind::Comment));
}

#[test]
fn tokens_are_in_source_order() {
    let s = Session::from_sources(vec![(
        "M.bas".into(),
        b"Public gN As Long\nSub F()\n    gN = 1\nEnd Sub\n".to_vec(),
    )]);
    let toks = s.semantic_tokens(0);
    assert!(toks.windows(2).all(|w| w[0].span.start <= w[1].span.start), "ordered by offset");
    assert!(!toks.is_empty());
}

#[test]
fn cross_module_call_classified_as_function() {
    let mut s = Session::from_sources(vec![(
        "Mod1.bas".into(),
        b"Sub Main()\n    Greet\nEnd Sub\n".to_vec(),
    )]);
    s.update_file("Mod0.bas", b"Public Sub Greet()\nEnd Sub\n".to_vec());
    let m1 = s.module_of("Mod1.bas").unwrap();
    let toks = s.semantic_tokens(m1);
    // The `Greet` call resolves cross-module -> Function.
    let src = "Sub Main()\n    Greet\nEnd Sub\n";
    assert!(
        toks.iter().any(|t| {
            let lo = t.span.start as usize;
            t.kind == SemTokenKind::Function && &src[lo..lo + t.span.len as usize] == "Greet"
        }),
        "cross-module call should classify as Function"
    );
}
