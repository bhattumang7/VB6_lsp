//! Sequences of declaration modifiers before every declaration kind, in standard and class modules.

mod common;
use common::answers::{assert_none_wrong, disagreements};

const TEMPLATES: &str = include_str!("fixtures/modifier_sequences/templates.txt");
const ANSWERS: &str = include_str!("fixtures/modifier_sequences/oracle.tsv");

#[test]
fn modifier_sequences_before_declarations_match_the_compiler() {
    let (checked, wrong) = disagreements(TEMPLATES, ANSWERS, |_, word, class, _| class == "semantic" && word.contains("\\n"));
    assert!(checked > 2000, "fixture is unexpectedly small: {checked}");
    assert_none_wrong(checked, &wrong);
}
