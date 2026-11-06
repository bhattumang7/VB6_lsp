//! Graphics and print statements by receiver form, bracket flag, position and operand shape,
//! written bare in a standard module and in a class module. The fixture has the compiler's
//! answer for each.

mod common;
use common::answers::{assert_none_wrong, disagreements};

const TEMPLATES: &str = include_str!("fixtures/graphics_receiver_forms/templates.txt");
const ANSWERS: &str = include_str!("fixtures/graphics_receiver_forms/oracle.tsv");

#[test]
fn graphics_statements_match_the_compiler() {
    let skip = |_: &str, _: &str, _: &str, _: &str| false;
    let (checked, wrong) = disagreements(TEMPLATES, ANSWERS, skip);
    assert!(checked > 600, "fixture is unexpectedly small: {checked}");
    assert_none_wrong(checked, &wrong);
}
