//! Case clause forms in Select Case.

mod common;
use common::answers::{assert_none_wrong, disagreements};

const TEMPLATES: &str = include_str!("fixtures/case_clause_forms/templates.txt");
const ANSWERS: &str = include_str!("fixtures/case_clause_forms/oracle.tsv");

#[test]
fn case_clause_forms_match_the_compiler() {
    let (checked, wrong) = disagreements(TEMPLATES, ANSWERS, |_, word, class, _| class == "semantic" && word.contains("\\n"));
    assert!(checked > 40, "fixture is unexpectedly small: {checked}");
    assert_none_wrong(checked, &wrong);
}
