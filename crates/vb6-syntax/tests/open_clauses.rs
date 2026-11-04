//! Runs of Open clauses (mode, access, sharing, lock, length, channel) in every order.

mod common;
use common::answers::{assert_none_wrong, disagreements};

const TEMPLATES: &str = include_str!("fixtures/open_clauses/templates.txt");
const ANSWERS: &str = include_str!("fixtures/open_clauses/oracle.tsv");

#[test]
fn open_clauses_match_the_compiler() {
    let (checked, wrong) = disagreements(TEMPLATES, ANSWERS, |_, word, class, _| class == "semantic" && word.contains("\\n"));
    assert!(checked > 2000, "fixture is unexpectedly small: {checked}");
    assert_none_wrong(checked, &wrong);
}
