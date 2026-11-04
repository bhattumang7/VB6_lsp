//! Every keyword followed by each type suffix, in statement, value, declaration and member positions.

mod common;
use common::answers::{assert_none_wrong, disagreements};

const TEMPLATES: &str = include_str!("fixtures/keyword_suffixes/templates.txt");
const ANSWERS: &str = include_str!("fixtures/keyword_suffixes/oracle.tsv");

#[test]
fn keywords_with_type_suffixes_match_the_compiler() {
    let (checked, wrong) = disagreements(TEMPLATES, ANSWERS, |_, word, class, _| class == "semantic" && word.contains("\\n"));
    assert!(checked > 3000, "fixture is unexpectedly small: {checked}");
    assert_none_wrong(checked, &wrong);
}
