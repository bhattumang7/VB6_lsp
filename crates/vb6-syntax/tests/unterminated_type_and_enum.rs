//! A `Type` or `Enum` block that runs into the end of the file without its
//! `End` line is an error, reported where the file ends.

mod common;
use common::parse;

#[test]
fn type_without_end_type_at_the_end_of_the_file_is_an_error() {
    for text in ["Type T\n  a As Long\n", "Type T\n  a As Long ' End Type\n"] {
        let p = parse(text);
        assert_eq!(p.found.len(), 1, "{text:?}: {:?}", p.found);
        assert_eq!(p.found[0].0, 0x9c76, "{text:?}");
        assert_eq!(p.found[0].1 as usize, text.len(), "{text:?}");
    }
}

#[test]
fn enum_without_end_enum_at_the_end_of_the_file_is_an_error() {
    for text in ["Enum E\n  A\n", "Enum E\n  A ' End Enum\n"] {
        let p = parse(text);
        assert_eq!(p.found.len(), 1, "{text:?}: {:?}", p.found);
        assert_eq!(p.found[0].0, 0xdeff, "{text:?}");
        assert_eq!(p.found[0].1 as usize, text.len(), "{text:?}");
    }
}

#[test]
fn terminated_blocks_are_clean() {
    assert_eq!(parse("Type T\n  a As Long\nEnd Type\n").found, vec![]);
    assert_eq!(parse("Enum E\n  A\nEnd Enum\n").found, vec![]);
}

#[test]
fn a_type_member_with_new_before_a_built_in_type_is_expected_identifier() {
    for (text, word) in [
        ("Type T\n  a As New Long\nEnd Type\n", "Long"),
        ("Type T\n  b As New String * 3\nEnd Type\n", "String"),
    ] {
        let p = parse(text);
        let at = text.find(word).unwrap() as u32;
        assert!(p.found.iter().any(|d| d.0 == 0x9c6f && d.1 == at), "{text:?}: {:?}", p.found);
    }
}
