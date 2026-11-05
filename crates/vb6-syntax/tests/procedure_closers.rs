//! Procedure bodies and closing lines: End and Exit of every kind against every procedure kind.

mod common;
use common::answers::{assert_none_wrong, disagreements};

const TEMPLATES: &str = include_str!("fixtures/procedure_closers/templates.txt");
const ANSWERS: &str = include_str!("fixtures/procedure_closers/oracle.tsv");

#[test]
fn procedure_closers_match_the_compiler() {
    let (checked, wrong) = disagreements(TEMPLATES, ANSWERS, |_, word, class, _| class == "semantic" && word.contains("\\n"));
    assert!(checked > 1500, "fixture is unexpectedly small: {checked}");
    assert_none_wrong(checked, &wrong);
}
