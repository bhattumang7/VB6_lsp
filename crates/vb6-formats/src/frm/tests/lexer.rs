use super::*;

#[test]
fn lex_simple_property() {
    let toks = lex_line("   BackColor       =   &H80000005&");
    assert_eq!(toks[0], Token::Word("BackColor"));
    assert_eq!(toks[1], Token::Equals);
    assert_eq!(toks[2], Token::Word("&H80000005&"));
}

#[test]
fn lex_quoted_value() {
    let toks = lex_line(r#"Caption         =   "Hello World""#);
    assert_eq!(toks[0], Token::Word("Caption"));
    assert_eq!(toks[1], Token::Equals);
    assert_eq!(toks[2], Token::Quoted("Hello World"));
}

#[test]
fn lex_comment_stripped() {
    let toks = lex_line("  Begin  ' this is a comment");
    assert_eq!(toks.len(), 1);
    assert_eq!(toks[0], Token::Word("Begin"));
}

#[test]
fn lex_begin_line() {
    let toks = lex_line("   Begin VB.Form Form1");
    assert_eq!(toks[0], Token::Word("Begin"));
    assert_eq!(toks[1], Token::Word("VB.Form"));
    assert_eq!(toks[2], Token::Word("Form1"));
}

#[test]
fn lex_object_line() {
    // In .frm files the progid is in double quotes: Object = "...#ver#lcid"
    let toks = lex_line(
        r#"Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX""#,
    );
    assert_eq!(toks[0], Token::Word("Object"));
    assert_eq!(toks[1], Token::Equals);
    // The entire progid is inside outer double-quotes → Quoted token
    match &toks[2] {
        Token::Quoted(s) => assert!(s.contains("831FDD16")),
        other => panic!("expected Quoted, got {:?}", other),
    }
    assert_eq!(toks[3], Token::Semi);
    assert_eq!(toks[4], Token::Quoted("MSCOMCTL.OCX"));
}

#[test]
fn lex_object_line_unquoted() {
    // In .vbp files the progid is bare: Object={GUID}#ver#lcid ; filename
    let toks = lex_line(
        "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0; MSCOMCTL.OCX",
    );
    match &toks[0] {
        Token::Curly(s) => assert!(s.contains("831FDD16")),
        other => panic!("expected Curly, got {:?}", other),
    }
    // #2.0#0 follows as Word
    assert_eq!(toks[1], Token::Word("#2.0#0"));
}

#[test]
fn lex_dollar_string() {
    let toks = lex_line(r#"Caption         =   $"Hello""#);
    assert_eq!(toks[2], Token::DollarStr("Hello"));
}

#[test]
fn lex_escaped_quote_in_string() {
    let toks = lex_line(r#"Caption = "say ""hi""  ""#);
    assert_eq!(toks[2], Token::Quoted(r#"say ""hi""  "#));
}

#[test]
fn lex_unterminated_curly_runs_to_end_of_line() {
    let toks = lex_line("{abc def");
    assert_eq!(toks, vec![Token::Curly("abc def")]);
}

#[test]
fn lex_empty_curly_and_following_tokens() {
    let toks = lex_line("{}=x");
    assert_eq!(toks, vec![Token::Curly(""), Token::Equals, Token::Word("x")]);
}

#[test]
fn lex_word_stops_at_delimiters() {
    let toks = lex_line("a;b:c{d}e\"f\"g'h");
    assert_eq!(
        toks,
        vec![
            Token::Word("a"),
            Token::Semi,
            Token::Word("b"),
            Token::Colon,
            Token::Word("c"),
            Token::Curly("d"),
            Token::Word("e"),
            Token::Quoted("f"),
            Token::Word("g"),
        ]
    );
}
