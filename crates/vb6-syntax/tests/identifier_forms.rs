//! Identifier spellings (brackets, type characters, underscores, digits,
//! punctuation, non-ASCII letters and length limits) in declaration,
//! assignment, call, label and member positions, with the compiler's answer.

mod common;
use common::answers::{assert_none_wrong, disagreements};

const TEMPLATES: &str = include_str!("fixtures/identifier_forms/templates.txt");
const ANSWERS: &str = include_str!("fixtures/identifier_forms/oracle.tsv");

#[test]
fn identifier_spellings_match_the_compiler() {
    // Unresolved names, undefined labels and unclosed blocks are found after the semantic checks.
    let (checked, wrong) = disagreements(TEMPLATES, ANSWERS, |_, _, _, _| false);
    assert!(checked > 600, "fixture is unexpectedly small: {checked}");
    assert_none_wrong(checked, &wrong);
}
