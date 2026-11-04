//! `Rem` starts a comment only where a statement can start: at the start of a
//! line or after a colon. In the middle of a statement it is a syntax error;
//! the apostrophe comment is allowed anywhere.

mod common;
use common::{offset_of, parse};
use vb6_syntax::frontend::ast::ExprNode;

fn in_sub(stmt: &str) -> (common::Parsed, String) {
    let src = format!("Sub S()\n    {stmt}\nEnd Sub\n");
    (parse(&src), src)
}

#[test]
fn rem_in_the_middle_of_a_statement_is_a_syntax_error_at_rem() {
    for text in [
        "x = 1 Rem c", "Dim a As Long Rem c", "Exit Sub Rem c", "Stop Rem c", "Foo Rem c", "Foo 1 Rem c",
        "Set x = y Rem c", "Erase x Rem c", "Return Rem c", "Call Foo Rem c", "Resume Next Rem c",
    ] {
        let (p, src) = in_sub(text);
        assert!(
            p.found.iter().any(|d| d.0 == 0x9c6f && d.1 == offset_of(&src, "Rem", 0)),
            "{text}: {:?}",
            p.found
        );
    }
}

#[test]
fn rem_after_do_and_loop_conditions_is_a_syntax_error() {
    for text in ["Do Rem c\nLoop", "Do\nLoop Rem c", "Do While x Rem c\nLoop"] {
        let src = format!("Sub S()\n{text}\nEnd Sub\n");
        let p = parse(&src);
        assert!(p.found.iter().any(|d| d.1 == offset_of(&src, "Rem", 0)), "{text}: {:?}", p.found);
    }
}

#[test]
fn rem_where_a_statement_starts_is_a_comment() {
    for text in ["Rem c", "x = 1: Rem c", "Rem", "    Rem c"] {
        let (p, _) = in_sub(text);
        assert_eq!(p.found, vec![], "{text}");
    }
    let (p, _) = in_sub("x = 1: Rem c");
    assert_eq!(p.proc_body(0).len(), 1);
    assert!(matches!(p.node(p.proc_body(0)[0]), ExprNode::Assign { .. }));
}

#[test]
fn an_apostrophe_comment_is_allowed_anywhere() {
    for text in ["x = 1 ' c", "Dim a As Long ' c", "Exit Sub ' c", "Foo 1 ' c"] {
        let (p, _) = in_sub(text);
        assert_eq!(p.found, vec![], "{text}");
    }
}
