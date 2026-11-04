//! Runs of module-level items (Option, Def, declarations, procedures, Attribute, conditional directives) in standard and class modules.

mod common;
use common::answers::{assert_none_wrong, disagreements};

const TEMPLATES: &str = include_str!("fixtures/module_item_order/templates.txt");
const ANSWERS: &str = include_str!("fixtures/module_item_order/oracle.tsv");

#[test]
fn module_item_order_match_the_compiler() {
    let (checked, wrong) = disagreements(TEMPLATES, ANSWERS, |_, word, class, _| class == "semantic" && word.contains("\\n"));
    assert!(checked > 1500, "fixture is unexpectedly small: {checked}");
    assert_none_wrong(checked, &wrong);
}
