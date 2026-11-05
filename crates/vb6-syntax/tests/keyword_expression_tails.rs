//! Every keyword at the start of an expression followed by each of a set of tails.

mod common;
use common::answers::{assert_none_wrong, disagreements};

const TEMPLATES: &str = include_str!("fixtures/keyword_expression_tails/templates.txt");
const ANSWERS: &str = include_str!("fixtures/keyword_expression_tails/oracle.tsv");

#[test]
fn keyword_expression_tails_match_the_compiler() {
    let (checked, wrong) = disagreements(TEMPLATES, ANSWERS, |_, word, class, message| {
        // The operand of a stray `Next` or `Case` is looked up before the block is paired.
        (class == "semantic" && message == "Invalid or unqualified reference" && (word == "Next .a" || word == "Case .a"))
            || (class == "semantic" && word.contains("\n"))
    });
    assert!(checked > 5000, "fixture is unexpectedly small: {checked}");
    assert_none_wrong(checked, &wrong);
}
