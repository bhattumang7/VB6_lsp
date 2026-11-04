//! A colon that starts a module-level line is accepted before `Option`, a
//! `Def` statement, `Type`, `Enum` and `Attribute`, and rejected before the
//! declaration of a procedure, a `Declare` or an `Event`.

mod common;
use common::{offset_of, parse};

#[test]
fn a_leading_colon_before_a_procedure_declare_or_event_is_expected_statement_or_end() {
    for (text, word) in [
        (": Sub S()\nEnd Sub\n", "Sub"),
        (": Function F() As Long\nEnd Function\n", "Function"),
        (": Property Get P() As Long\nEnd Property\n", "Property"),
        (": Declare Function f Lib \"k\" () As Long\n", "Declare"),
        (": Public Sub S()\nEnd Sub\n", "Public"),
        (": Private Function F() As Long\nEnd Function\n", "Private"),
    ] {
        let p = parse(text);
        let at = offset_of(text, word, 0);
        assert!(p.found.iter().any(|d| d.0 == 0x9c6f && d.1 == at), "{text}: {:?}", p.found);
    }
}

#[test]
fn a_leading_colon_before_the_other_declarations_is_fine() {
    for text in [
        ": Option Explicit\n",
        ": DefInt A-Z\n",
        ": Type T\n  a As Long\nEnd Type\n",
        ": Enum E\n  A\nEnd Enum\n",
        ": Attribute VB_Name = \"m\"\n",
    ] {
        assert_eq!(parse(text).found, vec![], "{text}");
    }
}

#[test]
fn a_colon_inside_a_line_does_not_matter_after_a_declaration_ended() {
    let p = parse("Dim a As Long\nSub S()\nEnd Sub\n");
    assert_eq!(p.found, vec![]);
}
