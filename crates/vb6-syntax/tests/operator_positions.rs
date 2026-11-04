//! Operators and operator-like words in binary, unary, trailing, doubled and
//! nested positions, with the compiler's answer for each.

mod common;
use common::answers::{assert_none_wrong, disagreements};

const TEMPLATES: &str = include_str!("fixtures/operator_positions/templates.txt");
const ANSWERS: &str = include_str!("fixtures/operator_positions/oracle.tsv");

#[test]
fn operators_in_every_position_match_the_compiler() {
    let (checked, wrong) = disagreements(TEMPLATES, ANSWERS, |_, _, _, _| false);
    assert!(checked > 500, "fixture is unexpectedly small: {checked}");
    assert_none_wrong(checked, &wrong);
}
