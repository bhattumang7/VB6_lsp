//! Words after a `.` or `!`: the fixture holds the compiler's answer for every
//! keyword as a member name in an expression, as an assignment target, as a
//! called member, after a bang and in the middle of a chain.

mod common;
use common::answers::{assert_none_wrong, disagreements};

const TEMPLATES: &str = include_str!("fixtures/member_names/templates.txt");
const ANSWERS: &str = include_str!("fixtures/member_names/oracle.tsv");

#[test]
fn member_names_match_the_compiler() {
    let (checked, wrong) = disagreements(TEMPLATES, ANSWERS, |_, _, _, _| false);
    assert!(checked > 800, "fixture is unexpectedly small: {checked}");
    assert_none_wrong(checked, &wrong);
}
