//! A string or a bracketed name on a line that ends in a continuation mark.

mod common;
use common::answers::{assert_none_wrong, disagreements};

const TEMPLATES: &str = include_str!("fixtures/string_continuation/templates.txt");
const ANSWERS: &str = include_str!("fixtures/string_continuation/oracle.tsv");

#[test]
fn strings_and_names_ending_at_a_continuation_mark_match_the_compiler() {
    let (checked, wrong) = disagreements(TEMPLATES, ANSWERS, |_, word, class, _| class == "semantic" && word.contains("\\n"));
    assert!(checked > 15, "fixture is unexpectedly small: {checked}");
    assert_none_wrong(checked, &wrong);
}
