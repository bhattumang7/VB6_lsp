//! The declared type of a WithEvents variable, in class and standard modules.

mod common;
use common::answers::{assert_none_wrong, disagreements};

const TEMPLATES: &str = include_str!("fixtures/withevents_types/templates.txt");
const ANSWERS: &str = include_str!("fixtures/withevents_types/oracle.tsv");

#[test]
fn with_events_variable_types_match_the_compiler() {
    let (checked, wrong) = disagreements(TEMPLATES, ANSWERS, |_, word, class, _| class == "semantic" && word.contains("\\n"));
    assert!(checked > 60, "fixture is unexpectedly small: {checked}");
    assert_none_wrong(checked, &wrong);
}
