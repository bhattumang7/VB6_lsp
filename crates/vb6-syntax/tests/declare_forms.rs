//! Declare statements: Lib, Alias, calling convention, parameter and return forms.

mod common;
use common::answers::{assert_none_wrong, disagreements};

const TEMPLATES: &str = include_str!("fixtures/declare_forms/templates.txt");
const ANSWERS: &str = include_str!("fixtures/declare_forms/oracle.tsv");

#[test]
fn declare_forms_match_the_compiler() {
    let (checked, wrong) = disagreements(TEMPLATES, ANSWERS, |_, word, class, _| class == "semantic" && word.contains("\\n"));
    assert!(checked > 100, "fixture is unexpectedly small: {checked}");
    assert_none_wrong(checked, &wrong);
}
