//! Public and global members of a private class module.

mod common;
use common::answers::{assert_none_wrong, disagreements};

const TEMPLATES: &str = include_str!("fixtures/class_public_members/templates.txt");
const ANSWERS: &str = include_str!("fixtures/class_public_members/oracle.tsv");

#[test]
fn public_members_of_a_class_module_match_the_compiler() {
    let (checked, wrong) = disagreements(TEMPLATES, ANSWERS, |_, word, class, _| class == "semantic" && word.contains("\\n"));
    assert!(checked > 25, "fixture is unexpectedly small: {checked}");
    assert_none_wrong(checked, &wrong);
}
