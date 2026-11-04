//! A `Def` statement may end with a comma; the length after `String *` may be
//! named by any word that is not reserved; `New Object` names no class.

mod common;
use common::{offset_of, parse};
use vb6_syntax::frontend::ast::ExprNode;

#[test]
fn def_statement_may_end_with_a_comma() {
    for text in ["DefInt A-Z ,\n", "DefInt A,\n", "DefInt A-C, E,\n"] {
        assert_eq!(parse(text).found, vec![], "{text}");
    }
}

#[test]
fn fixed_length_may_be_a_word_that_is_not_reserved() {
    for word in ["Step"] {
        let text = format!("Type T\n  a As String * {word}\nEnd Type\n");
        let p = parse(&text);
        assert_eq!(p.found, vec![], "{word}");
        let ExprNode::TypeDecl { members, .. } = p.node(p.top[0]) else { panic!("expected Type") };
        assert_eq!(members.len(), 1, "{word}");
    }
}

#[test]
fn fixed_length_may_not_be_a_reserved_word() {
    for word in ["Me", "True"] {
        let text = format!("Type T\n  a As String * {word}\nEnd Type\n");
        assert!(!parse(&text).found.is_empty(), "{word}");
    }
}

#[test]
fn new_object_in_a_parameter_names_no_class() {
    let text = "Property Set P(v As New Object)\nEnd Property\n";
    let p = parse(text);
    assert!(p.found.iter().any(|d| d.0 == 0x9c6f && d.1 == offset_of(text, "Object", 0)), "{:?}", p.found);
}
