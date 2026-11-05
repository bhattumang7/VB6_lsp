//! A last line that ends in a continuation mark, in standard and class modules and inside a procedure.

mod common;
use common::answers::{assert_none_wrong, disagreements};

const TEMPLATES: &str = include_str!("fixtures/end_of_text_continuation/templates.txt");
const ANSWERS: &str = include_str!("fixtures/end_of_text_continuation/oracle.tsv");

#[test]
fn end_of_text_continuation_matches_the_compiler() {
    let (checked, wrong) = disagreements(TEMPLATES, ANSWERS, |_, word, class, _| class == "semantic" && word.contains("\\n"));
    assert!(checked > 40, "fixture is unexpectedly small: {checked}");
    assert_none_wrong(checked, &wrong);
}
