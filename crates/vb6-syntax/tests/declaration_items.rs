//! Variable, constant and member declarations: names, bounds, As clauses, New, WithEvents, lists.

mod common;
use common::answers::{assert_none_wrong, disagreements};

const TEMPLATES: &str = include_str!("fixtures/declaration_items/templates.txt");
const ANSWERS: &str = include_str!("fixtures/declaration_items/oracle.tsv");

#[test]
fn declaration_items_match_the_compiler() {
    let (checked, wrong) = disagreements(TEMPLATES, ANSWERS, |_, word, class, _| class == "semantic" && word.contains("\\n"));
    assert!(checked > 3000, "fixture is unexpectedly small: {checked}");
    assert_none_wrong(checked, &wrong);
}
