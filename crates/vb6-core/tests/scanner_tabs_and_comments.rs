//! Tabs are whitespace between tokens and part of comment and string text; a
//! comment ending in a line-continuation `_` continues on the next line.

use vb6_core::frontend::ast::{AstLit, ExprArena, ExprNode};
use vb6_core::frontend::parser::Parser;
use vb6_core::frontend::scanner::{Scanner, ScannerContext};
use vb6_core::frontend::token::TokenKind;

fn ctx() -> ScannerContext {
    let mut c = ScannerContext::new(1, 1, 0x0409);
    c.intern_keywords();
    c
}

/// Every token as `(kind text, source text)` up to Eof.
fn tokens(src: &str) -> Vec<(String, String)> {
    let mut c = ctx();
    let mut sc = Scanner::new(&mut c, src.as_bytes());
    let mut out = Vec::new();
    loop {
        let t = sc.next_token();
        let text = src[t.span.start as usize..(t.span.start + t.span.len) as usize].to_string();
        let kind = match t.kind {
            TokenKind::Eof => break,
            TokenKind::Eol => "Eol".to_string(),
            TokenKind::Ident => "Ident".to_string(),
            TokenKind::Kw(k) => format!("{k:?}"),
            other => format!("{other:?}"),
        };
        out.push((kind, text));
    }
    out
}

fn tok(kind: &str, text: &str) -> (String, String) {
    (kind.to_string(), text.to_string())
}

fn diags(src: &str) -> Vec<(u32, u32, u32)> {
    let mut c = ctx();
    let mut arena = ExprArena::new();
    let mut parser = Parser::new(&mut c, src.as_bytes());
    parser.parse_module(&mut arena);
    parser.diagnostics.items().iter().map(|d| (d.code, d.span.start, d.span.len)).collect()
}

#[test]
fn tab_between_tokens_is_whitespace() {
    assert_eq!(
        tokens("\tnTotal\t=\tnFirst\n"),
        vec![tok("Ident", "nTotal"), tok("Eq", "="), tok("Ident", "nFirst"), tok("Eol", "")]
    );
}

#[test]
fn tab_inside_a_comment_is_comment_text() {
    let src = "'**\t\t Ref email\nnTotal = 1\n";
    assert_eq!(
        tokens(src),
        vec![tok("Apos", "'**\t\t Ref email"), tok("Eol", ""), tok("Ident", "nTotal"), tok("Eq", "="),
             tok("IntLit", "1"), tok("Eol", "")]
    );
}

#[test]
fn tab_inside_a_string_literal_is_kept() {
    let mut c = ctx();
    let mut arena = ExprArena::new();
    let src = "Sub T()\nsText = \"a\tb\"\nEnd Sub\n";
    let mut parser = Parser::new(&mut c, src.as_bytes());
    let nodes = parser.parse_module(&mut arena);
    assert_eq!(parser.diagnostics.items().len(), 0);
    let ExprNode::ProcDecl { body, .. } = arena.get(nodes[0]) else { panic!("expected ProcDecl") };
    let ExprNode::Block { stmts } = arena.get(*body) else { panic!("expected Block") };
    let ExprNode::Assign { value, .. } = arena.get(stmts[0]) else { panic!("expected Assign") };
    assert!(matches!(arena.get(*value), ExprNode::Literal { lit: AstLit::Str(s) } if &**s == "a\tb"));
}

#[test]
fn tab_indented_procedure_parses_cleanly() {
    assert_eq!(diags("Sub T()\n\tIf nFirst = 1 Then\n\t\tnTotal = 2\n\tEnd If\nEnd Sub\n"), vec![]);
}

#[test]
fn comment_ending_in_continuation_swallows_the_next_line() {
    let src = "' sText = \"[\" & _\n      CStr(nValue) & \"]\"\nnTotal = 1\n";
    assert_eq!(
        tokens(src),
        vec![tok("Apos", "' sText = \"[\" & _\n      CStr(nValue) & \"]\""), tok("Eol", ""),
             tok("Ident", "nTotal"), tok("Eq", "="), tok("IntLit", "1"), tok("Eol", "")]
    );
}

#[test]
fn comment_continuation_chains_and_crlf() {
    let src = "' a _\r\n b _\r\n c\r\nnTotal = 1\r\n";
    assert_eq!(
        tokens(src),
        vec![tok("Apos", "' a _\r\n b _\r\n c"), tok("Eol", ""), tok("Ident", "nTotal"), tok("Eq", "="),
             tok("IntLit", "1"), tok("Eol", "")]
    );
}

#[test]
fn trailing_blanks_after_comment_continuation() {
    let src = "' a _ \t\n b\n";
    assert_eq!(tokens(src), vec![tok("Apos", "' a _ \t\n b"), tok("Eol", "")]);
}

#[test]
fn rem_comment_continues_too() {
    let src = "Rem a _\n b\nnTotal = 1\n";
    assert_eq!(
        tokens(src),
        vec![tok("Apos", "Rem a _\n b"), tok("Eol", ""), tok("Ident", "nTotal"), tok("Eq", "="),
             tok("IntLit", "1"), tok("Eol", "")]
    );
}

#[test]
fn trailing_comment_after_code_continues_too() {
    let src = "nTotal = 1 ' a _\n b\n";
    assert_eq!(
        tokens(src),
        vec![tok("Ident", "nTotal"), tok("Eq", "="), tok("IntLit", "1"), tok("Apos", "' a _\n b"), tok("Eol", "")]
    );
}

#[test]
fn underscore_ending_a_word_is_not_a_continuation() {
    let src = "' see m_\nnTotal = 1\n";
    assert_eq!(
        tokens(src),
        vec![tok("Apos", "' see m_"), tok("Eol", ""), tok("Ident", "nTotal"), tok("Eq", "="), tok("IntLit", "1"),
             tok("Eol", "")]
    );
}

#[test]
fn bare_underscore_comment_continues() {
    let src = "'_\n b\n";
    assert_eq!(tokens(src), vec![tok("Apos", "'_\n b"), tok("Eol", "")]);
}

#[test]
fn comment_continuation_at_end_of_file_stops_there() {
    let src = "' a _";
    assert_eq!(tokens(src), vec![tok("Apos", "' a _")]);
}
