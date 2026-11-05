//! Event, Implements and RaiseEvent forms.

mod common;
use common::answers::{assert_none_wrong, disagreements};

const TEMPLATES: &str = include_str!("fixtures/event_and_implements_forms/templates.txt");
const ANSWERS: &str = include_str!("fixtures/event_and_implements_forms/oracle.tsv");

#[test]
fn event_and_implements_forms_match_the_compiler() {
    let (checked, wrong) = disagreements(TEMPLATES, ANSWERS, |_, word, class, _| class == "semantic" && word.contains("\\n"));
    assert!(checked > 150, "fixture is unexpectedly small: {checked}");
    assert_none_wrong(checked, &wrong);
}
