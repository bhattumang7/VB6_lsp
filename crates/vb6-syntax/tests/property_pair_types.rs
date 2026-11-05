//! Property procedures of one name: every pairing of value types, index modes, names and counts.

mod common;
use common::answers::{assert_none_wrong, disagreements};

const TEMPLATES: &str = include_str!("fixtures/property_pair_types/templates.txt");
const ANSWERS: &str = include_str!("fixtures/property_pair_types/oracle.tsv");

#[test]
fn property_pair_value_and_index_types_match_the_compiler() {
    let (checked, wrong) = disagreements(TEMPLATES, ANSWERS, |_, word, class, _| class == "semantic" && word.contains("\\n"));
    assert!(checked > 450, "fixture is unexpectedly small: {checked}");
    assert_none_wrong(checked, &wrong);
}
