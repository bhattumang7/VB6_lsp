//! Const statements: names, As clauses, values and lists, in a module and in a procedure.

mod common;
use common::answers::{assert_none_wrong, disagreements};

const TEMPLATES: &str = include_str!("fixtures/const_forms/templates.txt");
const ANSWERS: &str = include_str!("fixtures/const_forms/oracle.tsv");

#[test]
fn const_forms_match_the_compiler() {
    let (checked, wrong) = disagreements(TEMPLATES, ANSWERS, |_, word, class, _| class == "semantic" && word.contains("\\n"));
    assert!(checked > 90, "fixture is unexpectedly small: {checked}");
    assert_none_wrong(checked, &wrong);
}
