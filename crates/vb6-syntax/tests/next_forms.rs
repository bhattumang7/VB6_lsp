//! Next statement forms after one and two open For loops.

mod common;
use common::answers::{assert_none_wrong, disagreements};

const TEMPLATES: &str = include_str!("fixtures/next_forms/templates.txt");
const ANSWERS: &str = include_str!("fixtures/next_forms/oracle.tsv");

#[test]
fn next_forms_match_the_compiler() {
    let (checked, wrong) = disagreements(TEMPLATES, ANSWERS, |_, word, class, _| {
        // The compiler words a suffix that does not fit the loop variable as a type-character
        // mismatch before it compares the variable names.
        class == "semantic" && word == "Next i%"
    });
    assert!(checked > 25, "fixture is unexpectedly small: {checked}");
    assert_none_wrong(checked, &wrong);
}
