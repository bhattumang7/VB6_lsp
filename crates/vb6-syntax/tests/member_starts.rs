//! A statement that starts with a dot or a bang, outside and inside With, after a colon and after Then.

mod common;
use common::answers::{assert_none_wrong, disagreements};

const TEMPLATES: &str = include_str!("fixtures/member_starts/templates.txt");
const ANSWERS: &str = include_str!("fixtures/member_starts/oracle.tsv");

#[test]
fn statement_starting_with_a_member_mark_matches_the_compiler() {
    let (checked, wrong) = disagreements(TEMPLATES, ANSWERS, |_, _, _, _| false);
    assert!(checked > 50, "fixture is unexpectedly small: {checked}");
    assert_none_wrong(checked, &wrong);
}
