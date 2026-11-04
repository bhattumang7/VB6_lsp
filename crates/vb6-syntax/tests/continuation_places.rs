//! A line-continuation mark inside strings, comments, numbers, names and between statement parts.

mod common;
use common::answers::{assert_none_wrong, disagreements};

const TEMPLATES: &str = include_str!("fixtures/continuation_places/templates.txt");
const ANSWERS: &str = include_str!("fixtures/continuation_places/oracle.tsv");

#[test]
fn continuation_marks_in_every_place_match_the_compiler() {
    let (checked, wrong) = disagreements(TEMPLATES, ANSWERS, |_, word, class, _| class == "semantic" && word.contains("\\n"));
    assert!(checked > 200, "fixture is unexpectedly small: {checked}");
    assert_none_wrong(checked, &wrong);
}
