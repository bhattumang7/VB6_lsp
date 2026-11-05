//! Mid assignment, RSet, LSet, Input and Name statements.

mod common;
use common::answers::{assert_none_wrong, disagreements};

const TEMPLATES: &str = include_str!("fixtures/mid_lset_input_name_forms/templates.txt");
const ANSWERS: &str = include_str!("fixtures/mid_lset_input_name_forms/oracle.tsv");

#[test]
fn mid_lset_input_name_forms_match_the_compiler() {
    let (checked, wrong) = disagreements(TEMPLATES, ANSWERS, |_, word, class, _| class == "semantic" && word.contains("\\n"));
    assert!(checked > 45, "fixture is unexpectedly small: {checked}");
    assert_none_wrong(checked, &wrong);
}
