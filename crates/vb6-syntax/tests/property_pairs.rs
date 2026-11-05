//! Two property or plain procedures of the same name, in every pairing and order.

mod common;
use common::answers::{assert_none_wrong, disagreements};

const TEMPLATES: &str = include_str!("fixtures/property_pairs/templates.txt");
const ANSWERS: &str = include_str!("fixtures/property_pairs/oracle.tsv");

#[test]
fn property_pairs_match_the_compiler() {
    let (checked, wrong) = disagreements(TEMPLATES, ANSWERS, |_, word, class, _| class == "semantic" && word.contains("\\n"));
    assert!(checked > 700, "fixture is unexpectedly small: {checked}");
    assert_none_wrong(checked, &wrong);
}
