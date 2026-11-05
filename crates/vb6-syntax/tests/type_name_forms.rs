//! The type after As: plain, qualified, built-in, with brackets and with fixed lengths, in five positions.

mod common;
use common::answers::{assert_none_wrong, disagreements};

const TEMPLATES: &str = include_str!("fixtures/type_name_forms/templates.txt");
const ANSWERS: &str = include_str!("fixtures/type_name_forms/oracle.tsv");

#[test]
fn type_name_forms_match_the_compiler() {
    let (checked, wrong) = disagreements(TEMPLATES, ANSWERS, |_, word, class, _| class == "semantic" && word.contains("\\n"));
    assert!(checked > 190, "fixture is unexpectedly small: {checked}");
    assert_none_wrong(checked, &wrong);
}
