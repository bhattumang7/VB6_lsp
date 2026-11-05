//! Type member forms.

mod common;
use common::answers::{assert_none_wrong, disagreements};

const TEMPLATES: &str = include_str!("fixtures/type_members/templates.txt");
const ANSWERS: &str = include_str!("fixtures/type_members/oracle.tsv");

#[test]
fn type_members_match_the_compiler() {
    let (checked, wrong) = disagreements(TEMPLATES, ANSWERS, |_, word, class, _| class == "semantic" && word.contains("\\n"));
    assert!(checked > 80, "fixture is unexpectedly small: {checked}");
    assert_none_wrong(checked, &wrong);
}
