//! `ReDim` needs bounds; a line number only starts a line; the name of a
//! `Type`, `Enum` or `Event` takes no type character; a type character on a
//! variable or parameter excludes an `As` clause.

mod common;
use common::{offset_of, parse};
use vb6_syntax::frontend::ast::ExprNode;

fn in_sub(stmt: &str) -> (common::Parsed, String) {
    let src = format!("Sub S()\n    {stmt}\nEnd Sub\n");
    (parse(&src), src)
}

#[test]
fn redim_with_bounds_keeps_them() {
    let (p, _) = in_sub("ReDim a(1 To 2, 3)");
    assert_eq!(p.found, vec![]);
    let items = p.block(p.proc_body(0)[0]);
    let ExprNode::ReDimItem { bounds, .. } = p.node(items[0]) else { panic!("expected ReDimItem") };
    let ExprNode::ArgList { args } = p.node(bounds.expect("bounds")) else { panic!("expected bounds") };
    assert_eq!(args.len(), 2);
}

#[test]
fn redim_without_bounds_is_a_syntax_error() {
    for text in ["ReDim a()", "ReDim a(,)", "ReDim Preserve a()", "ReDim a"] {
        let (p, src) = in_sub(text);
        assert!(!p.found.is_empty(), "{text}");
        assert!(p.found.iter().all(|d| d.1 >= offset_of(&src, "ReDim", 0)), "{text}");
    }
}

#[test]
fn a_line_number_after_a_colon_is_a_syntax_error() {
    let (p, src) = in_sub("x = 1: 10 y = 2");
    assert_eq!(p.found[0], (0x9c70, offset_of(&src, "10", 0), 2));
    let (p, _) = in_sub("10 y = 2");
    assert_eq!(p.found, vec![]);
}

#[test]
fn type_enum_and_event_names_take_no_type_character() {
    for (text, name) in [("Type T%\nm As Long\nEnd Type\n", "T%"), ("Enum E$\nA\nEnd Enum\n", "E$")] {
        let p = parse(text);
        assert_eq!(p.found, vec![(0x9c6f, 5, 2)], "{name}");
    }
}

#[test]
fn a_type_character_excludes_as() {
    for text in ["Dim a% As Integer", "Static a$ As String"] {
        let (p, src) = in_sub(text);
        assert_eq!(p.found.len(), 1, "{text}");
        assert_eq!(p.found[0].1, offset_of(&src, "As", 0), "{text}");
    }
    let (p, _) = in_sub("Dim a%, b$");
    assert_eq!(p.found, vec![]);
}
