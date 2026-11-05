//! Bare Scale, Line, Circle, PSet and Print statements, and the same as methods of an object.

mod common;
use common::answers::{assert_none_wrong, disagreements};

const TEMPLATES: &str = include_str!("fixtures/graphics_statement_forms/templates.txt");
const ANSWERS: &str = include_str!("fixtures/graphics_statement_forms/oracle.tsv");

#[test]
fn graphics_statement_forms_match_the_compiler() {
    let (checked, wrong) = disagreements(TEMPLATES, ANSWERS, |_, word, class, _| class == "semantic" && word.contains("\\n"));
    assert!(checked > 140, "fixture is unexpectedly small: {checked}");
    assert_none_wrong(checked, &wrong);
}
