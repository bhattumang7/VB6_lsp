//! Parameter lists of every procedure and event form: modifiers, types, defaults, arrays and separators.

mod common;
use common::answers::{assert_none_wrong, disagreements};

const TEMPLATES: &str = include_str!("fixtures/parameter_lists/templates.txt");
const ANSWERS: &str = include_str!("fixtures/parameter_lists/oracle.tsv");

#[test]
fn parameter_lists_match_the_compiler() {
    let (checked, wrong) = disagreements(TEMPLATES, ANSWERS, |_, word, class, message| {
        // The compiler binds each line as it parses it, so a name it cannot find on an earlier
        // line is reported before a syntax error on a later one.
        // An out-of-range default value is the conversion of a constant to the parameter type,
        // which belongs to the binder, not the parser.
        (class == "semantic" && word.contains("\\n")) || (message == "Overflow" && word.contains('='))
    });
    assert!(checked > 2000, "fixture is unexpectedly small: {checked}");
    assert_none_wrong(checked, &wrong);
}
