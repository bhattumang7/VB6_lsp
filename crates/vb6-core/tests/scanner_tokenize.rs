//! Structural tests for the VB6 scanner (tokenizer).

use vb6_core::frontend::scanner::{Scanner, ScannerContext};
use vb6_core::frontend::token::{Kw, TokenKind};

fn tokenize(src: &[u8]) -> Vec<TokenKind> {
    let mut c = ScannerContext::new(1, 1, 0x0409);
    c.intern_keywords();
    let mut s = Scanner::new(&mut c, src);
    let mut kinds = Vec::new();
    loop {
        let t = s.next_token();
        if t.is_stmt_end() {
            break;
        }
        kinds.push(t.kind);
    }
    kinds
}

#[test]
fn tokenize_basic_statement() {
    let kinds = tokenize(b"Dim x As Integer");
    assert_eq!(kinds, vec![
        TokenKind::Kw(Kw::Dim),
        TokenKind::Ident,
        TokenKind::Kw(Kw::As),
        TokenKind::Kw(Kw::Integer),
    ]);
}

#[test]
fn tokenize_comment() {
    let kinds = tokenize(b"Dim x ' this is a comment");
    // Scanner emits Apos as a statement-end token; tokenize() stops there.
    assert_eq!(kinds, vec![
        TokenKind::Kw(Kw::Dim),
        TokenKind::Ident,
    ]);
}

#[test]
fn tokenize_literals() {
    let kinds = tokenize(b"123 1.23 \"hello\" #3/15/2001#");
    assert_eq!(kinds, vec![
        TokenKind::IntLit,
        TokenKind::DblLit,
        TokenKind::StrLit,
        TokenKind::DateLit,
    ]);
}

#[test]
fn tokenize_operators() {
    let kinds = tokenize(b"+ - * / ^ , = , <> , < , > , <= , >=");
    assert_eq!(kinds, vec![
        TokenKind::Kw(Kw::Plus),
        TokenKind::Kw(Kw::Minus),
        TokenKind::Kw(Kw::Star),
        TokenKind::Kw(Kw::Slash),
        TokenKind::Kw(Kw::Caret),
        TokenKind::Kw(Kw::Comma),
        TokenKind::Kw(Kw::Eq),
        TokenKind::Kw(Kw::Comma),
        TokenKind::Kw(Kw::Ne),
        TokenKind::Kw(Kw::Comma),
        TokenKind::Kw(Kw::Lt),
        TokenKind::Kw(Kw::Comma),
        TokenKind::Kw(Kw::Gt),
        TokenKind::Kw(Kw::Comma),
        TokenKind::Kw(Kw::Le),
        TokenKind::Kw(Kw::Comma),
        TokenKind::Kw(Kw::Ge),
    ]);
}

#[test]
fn two_character_operators_allow_blanks_and_reversed_spellings() {
    let cases: [(&[u8], Kw); 12] = [
        (b"< =", Kw::Le),
        (b"<\t=", Kw::Le),
        (b"< >", Kw::Ne),
        (b"> =", Kw::Ge),
        (b"=<", Kw::Le),
        (b"= <", Kw::Le),
        (b"=>", Kw::Ge),
        (b"= >", Kw::Ge),
        (b"><", Kw::Ne),
        (b"> <", Kw::Ne),
        (b"<>", Kw::Ne),
        (b": =", Kw::ColonEq),
    ];
    for (src, want) in cases {
        assert_eq!(tokenize(src), vec![TokenKind::Kw(want)], "{:?}", String::from_utf8_lossy(src));
    }
    // A blank before a non-operator character leaves the first operator alone.
    assert_eq!(
        tokenize(b"< 1"),
        vec![TokenKind::Kw(Kw::Lt), TokenKind::IntLit]
    );
}
