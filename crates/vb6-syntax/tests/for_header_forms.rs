//! For and For Each header forms.

mod common;
use common::answers::{assert_none_wrong, disagreements};

const TEMPLATES: &str = include_str!("fixtures/for_header_forms/templates.txt");
const ANSWERS: &str = include_str!("fixtures/for_header_forms/oracle.tsv");

#[test]
fn for_header_forms_match_the_compiler() {
    let (checked, wrong) = disagreements(TEMPLATES, ANSWERS, |_, word, class, _| class == "semantic" && word.contains("\\n"));
    assert!(checked > 60, "fixture is unexpectedly small: {checked}");
    assert_none_wrong(checked, &wrong);
}
