//! Enum member forms.

mod common;
use common::answers::{assert_none_wrong, disagreements};

const TEMPLATES: &str = include_str!("fixtures/enum_members/templates.txt");
const ANSWERS: &str = include_str!("fixtures/enum_members/oracle.tsv");

#[test]
fn enum_members_match_the_compiler() {
    let (checked, wrong) = disagreements(TEMPLATES, ANSWERS, |_, word, class, message| {
        // The compiler binds each line as it parses it, so a name it cannot find on an earlier
        // line is reported before a syntax error on a later one.
        // A value that does not fit the enumeration type is the conversion of a constant, which
        // belongs to the binder, not the parser.
        (class == "semantic" && word.contains("\\n")) || (message == "Overflow" && word.starts_with("A = "))
    });
    assert!(checked > 70, "fixture is unexpectedly small: {checked}");
    assert_none_wrong(checked, &wrong);
}
