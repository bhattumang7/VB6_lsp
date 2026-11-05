//! Runs of one to four statement-level tokens inside a procedure.

mod common;
use common::answers::{assert_none_wrong, disagreements};

const TEMPLATES: &str = include_str!("fixtures/statement_token_runs/templates.txt");
const ANSWERS: &str = include_str!("fixtures/statement_token_runs/oracle.tsv");

#[test]
fn statement_token_runs_match_the_compiler() {
    let (checked, wrong) = disagreements(TEMPLATES, ANSWERS, |_, word, class, message| {
        // The operand of a stray block closer is looked up before the block is paired.
        class == "semantic"
            && message == "Invalid or unqualified reference"
            && ["Case ", "Next ", "Loop "].iter().any(|w| word.starts_with(w))
    });
    assert!(checked > 4500, "fixture is unexpectedly small: {checked}");
    assert_none_wrong(checked, &wrong);
}
