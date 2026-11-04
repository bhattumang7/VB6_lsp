//! Name spellings the compiler rejects: a leading underscore, a bracket that
//! holds a bracket, characters no token starts with, a bracketed name as a
//! declared name, a bracketed reserved word called as a statement, and a bang
//! member called with arguments. Bracketed names stay fine as values, labels
//! and members.

mod common;
use common::{parse, Parsed};
use vb6_syntax::frontend::ast::ExprNode;

fn body(stmt: &str) -> Parsed {
    parse(&format!("Sub S(o As Object)\n    {stmt}\nEnd Sub\n"))
}

#[test]
fn a_name_cannot_start_with_an_underscore() {
    for text in ["_a = 1", "__a = 1", "x = _a", "x = o._a", "Dim _a As Long", "_a 1"] {
        assert!(!body(text).found.is_empty(), "{text}");
    }
    for ok in ["a_b = 1", "a_ = 1", "a__ = 1"] {
        assert_eq!(body(ok).found, vec![], "{ok}");
    }
}

#[test]
fn a_line_continuation_underscore_is_still_one() {
    assert_eq!(body("x = 1 + _\n        2").found, vec![]);
}

#[test]
fn characters_no_token_starts_with_are_errors() {
    for text in ["a~b = 1", "a`b = 1", "a|b = 1", "a{b = 1", "a}b = 1", "x = ~1"] {
        assert!(!body(text).found.is_empty(), "{text}");
    }
}

#[test]
fn a_bracket_holding_a_bracket_is_an_error() {
    for text in ["[[a] = 1", "x = [[a]", "x = o.[[a]"] {
        assert!(!body(text).found.is_empty(), "{text}");
    }
    assert_eq!(body("x = [a b]").found, vec![]);
}

#[test]
fn a_bracketed_name_is_not_a_declared_name() {
    for text in [
        "Sub [a]()\nEnd Sub\n",
        "Dim [a] As Long\n",
        "Sub S([a] As Long)\nEnd Sub\n",
        "Type T\n    [a] As Long\nEnd Type\n",
    ] {
        let p = parse(text);
        assert!(!p.found.is_empty(), "{text}");
        assert!(p.found.iter().any(|d| d.1 as usize == text.find('[').unwrap()), "{text}: {:?}", p.found);
    }
}

#[test]
fn a_bracketed_name_is_fine_as_a_value_a_label_and_a_member() {
    for text in ["x = [a b]", "[a b] = 1", "x = o.[a b]", "o.[a b] = 1"] {
        assert_eq!(body(text).found, vec![], "{text}");
    }
    assert_eq!(parse("Sub S()\n    GoTo [a b]\n[a b]:\nEnd Sub\n").found, vec![]);
}

#[test]
fn a_bracketed_reserved_word_called_with_arguments_is_a_syntax_error() {
    for text in ["[if] 1", "[Dim] 1"] {
        let p = body(text);
        assert!(p.found.iter().any(|d| d.0 == 0x9c70), "{text}: {:?}", p.found);
    }
    assert_eq!(body("[a] 1").found, vec![]);
    assert_eq!(body("[if] = 1").found, vec![]);
}

#[test]
fn a_bang_member_called_with_arguments_is_a_syntax_error() {
    let p = body("o!b 1");
    assert!(p.found.iter().any(|d| d.0 == 0x9c70), "{:?}", p.found);
    assert_eq!(body("x = o!b").found, vec![]);
    assert_eq!(body("o!b = 1").found, vec![]);
}

#[test]
fn bracketed_names_keep_their_text_as_the_name() {
    let p = body("x = [a b]");
    let ExprNode::Assign { value, .. } = p.node(p.proc_body(0)[0]) else { panic!("assignment") };
    let ExprNode::NameRef { sym, .. } = p.node(*value) else { panic!("name") };
    assert_eq!(p.name(*sym), "a b");
}
