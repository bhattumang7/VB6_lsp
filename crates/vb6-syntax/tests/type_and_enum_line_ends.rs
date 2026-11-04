//! The header line of a `Type` or `Enum` and the line that closes it end like
//! any statement: a colon may follow, a comment may follow, anything else is
//! an error.

mod common;
use common::{offset_of, parse};
use vb6_syntax::frontend::ast::ExprNode;

#[test]
fn a_colon_after_the_name_is_fine() {
    for text in ["Type T: \n  a As Long\nEnd Type\n", "Enum E: \n  A\nEnd Enum\n", "Type T :\n  a As Long\nEnd Type\n"] {
        let p = parse(text);
        assert_eq!(p.found, vec![], "{text}");
        assert!(matches!(p.node(p.top[0]), ExprNode::TypeDecl { .. } | ExprNode::EnumDecl { .. }), "{text}");
    }
}

#[test]
fn an_apostrophe_comment_after_the_name_or_the_end_is_fine() {
    for text in ["Type T ' c\n  a As Long\nEnd Type ' c\n", "Enum E ' c\n  A\nEnd Enum ' c\n"] {
        assert_eq!(parse(text).found, vec![], "{text}");
    }
}

#[test]
fn another_word_after_the_name_is_expected_end_of_statement() {
    for (text, junk) in [
        ("Type T junk\n  a As Long\nEnd Type\n", "junk"),
        ("Enum E junk\n  A\nEnd Enum\n", "junk"),
        ("Type T As Long\n  a As Long\nEnd Type\n", "As"),
        ("Enum E E\n  A\nEnd Enum\n", "E\n"),
    ] {
        let p = parse(text);
        // The second word of the header line, found from the end of the name.
        let at = if junk == "E\n" { text.find("E E").unwrap() as u32 + 2 } else { offset_of(text, junk, 0) };
        assert!(p.found.iter().any(|d| d.0 == 0x9c6f && d.1 == at), "{text}: {:?}", p.found);
    }
}

#[test]
fn rem_or_another_word_after_the_end_line_is_expected_end_of_statement() {
    for text in [
        "Type T\n  a As Long\nEnd Type Rem c\n",
        "Enum E\n  A\nEnd Enum Rem c\n",
        "Type T\n  a As Long\nEnd Type junk\n",
        "Enum E\n  A\nEnd Enum junk\n",
    ] {
        let p = parse(text);
        let at = if text.contains("Rem") { offset_of(text, "Rem", 0) } else { offset_of(text, "junk", 0) };
        assert!(p.found.iter().any(|d| d.0 == 0x9c6f && d.1 == at), "{text}: {:?}", p.found);
    }
}

#[test]
fn a_colon_after_the_end_line_starts_the_next_statement() {
    let p = parse("Type T\n  a As Long\nEnd Type: Dim b As Long\n");
    assert_eq!(p.found, vec![]);
}
