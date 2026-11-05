//! The TypeOf ... Is operator: its operand and the object type it names.

mod common;
use common::answers::{assert_none_wrong, disagreements};

const TEMPLATES: &str = include_str!("fixtures/typeof_forms/templates.txt");
const ANSWERS: &str = include_str!("fixtures/typeof_forms/oracle.tsv");

#[test]
fn typeof_forms_match_the_compiler() {
    let (checked, wrong) = disagreements(TEMPLATES, ANSWERS, |_, word, class, _| class == "semantic" && word.contains("\\n"));
    assert!(checked > 40, "fixture is unexpectedly small: {checked}");
    assert_none_wrong(checked, &wrong);
}
