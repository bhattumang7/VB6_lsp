//! Words after a `.` or `!`: `Rem` is a name there, not a comment, and a member
//! named like a graphics statement takes coordinate pairs only.

mod common;
use common::{offset_of, parse, Parsed};
use vb6_syntax::frontend::ast::ExprNode;

fn body(stmt: &str) -> (Parsed, String) {
    let src = format!("Sub S(o As Object)\n    {stmt}\nEnd Sub\n");
    (parse(&src), src)
}

#[test]
fn rem_after_a_dot_or_bang_is_a_member_name() {
    for text in ["x = o.Rem", "x = o!Rem", "o.Rem = 1", "o.Rem 1", "x = o.a.Rem.b"] {
        let (p, _) = body(text);
        assert_eq!(p.found, vec![], "{text}");
        assert_eq!(p.proc_body(0).len(), 1, "{text}: the rest of the line is not a comment");
    }
}

#[test]
fn the_member_name_rem_is_kept() {
    let (p, _) = body("x = o.Rem");
    let ExprNode::Assign { value, .. } = p.node(p.proc_body(0)[0]) else { panic!("assignment") };
    let ExprNode::MemberAccess { member, bang, .. } = p.node(*value) else { panic!("member access") };
    assert_eq!((p.name(*member), *bang), ("Rem", false));
}

#[test]
fn rem_not_after_a_member_operator_is_still_a_comment() {
    for text in ["Rem c", "x = 1: Rem c", "x = 5. ' c"] {
        let (p, _) = body(text);
        assert_eq!(p.found, vec![], "{text}");
    }
}

#[test]
fn graphics_words_cannot_be_assigned_as_members_except_line() {
    for word in ["Scale", "PSet", "Print", "LineInput", "Circle"] {
        let (p, src) = body(&format!("o.{word} = 1"));
        assert_eq!(p.found[0], (0x9c70, offset_of(&src, "=", 0), 1), "{word}");
    }
    let (p, _) = body("o.Line = 1");
    assert_eq!(p.found, vec![]);
}

#[test]
fn graphics_words_as_called_members_take_coordinate_pairs_only() {
    for word in ["Scale", "PSet", "Line", "LineInput", "Circle"] {
        let (p, src) = body(&format!("o.{word} 1"));
        assert_eq!(p.found[0], (0x9c70, offset_of(&src, "1", 0), 1), "{word}");
    }
    for ok in ["o.PSet (1, 2)", "o.Circle (1, 2), 3", "o.Line (0, 0)-(1, 1)", "o.Print 1", "o.Scale (1, 1)-(2, 2)"] {
        let (p, _) = body(ok);
        assert_eq!(p.found, vec![], "{ok}");
    }
}

#[test]
fn inside_with_the_graphics_words_are_ordinary_members() {
    for word in ["Scale", "PSet", "Print", "LineInput", "Circle", "Line"] {
        let src = format!("Sub S(o As Object)\n    With o\n        .{word} = 1\n        .{word} 1\n    End With\nEnd Sub\n");
        assert_eq!(parse(&src).found, vec![], "{word}");
    }
}

#[test]
fn else_cannot_be_a_named_argument_of_a_statement_call() {
    let (p, src) = body("o.f Else:=1");
    assert!(p.found.iter().any(|d| d.1 == offset_of(&src, "Else", 0)), "{:?}", p.found);
}
