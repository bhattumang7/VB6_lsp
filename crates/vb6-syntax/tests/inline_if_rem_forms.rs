//! A Rem comment after each kind of statement written in a single-line If.

mod common;
use common::answers::{assert_none_wrong, disagreements};

const TEMPLATES: &str = include_str!("fixtures/inline_if_rem_forms/templates.txt");
const ANSWERS: &str = include_str!("fixtures/inline_if_rem_forms/oracle.tsv");

#[test]
fn inline_if_rem_forms_match_the_compiler() {
    let (checked, wrong) = disagreements(TEMPLATES, ANSWERS, |_, word, class, _| class == "semantic" && word.contains("\\n"));
    assert!(checked > 70, "fixture is unexpectedly small: {checked}");
    assert_none_wrong(checked, &wrong);
}
