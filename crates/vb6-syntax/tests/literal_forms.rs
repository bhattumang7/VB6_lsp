//! Numeric, string and date literal forms in several positions, with the
//! compiler's answer for each.

mod common;
use common::answers::{assert_none_wrong, disagreements};

const TEMPLATES: &str = include_str!("fixtures/literal_forms/templates.txt");
const ANSWERS: &str = include_str!("fixtures/literal_forms/oracle.tsv");

#[test]
fn literal_forms_match_the_compiler() {
    let (checked, wrong) = disagreements(TEMPLATES, ANSWERS, |_, _, _, _| false);
    assert!(checked > 400, "fixture is unexpectedly small: {checked}");
    assert_none_wrong(checked, &wrong);
}
