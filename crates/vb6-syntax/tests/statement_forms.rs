//! Statement and declaration forms the compiler was asked about, with its
//! answer: the parser must find a form wrong exactly when the compiler does.
//! Forms the compiler objects to only for what names or types mean count as
//! right here.

mod common;
use common::answers::{assert_none_wrong, disagreements};

const TEMPLATES: &str = include_str!("fixtures/statement_forms/templates.txt");
const ANSWERS: &str = include_str!("fixtures/statement_forms/oracle.tsv");

#[test]
fn statement_forms_match_the_compiler() {
    let (checked, wrong) = disagreements(TEMPLATES, ANSWERS, |_, _, _| false);
    assert!(checked > 300, "fixture is unexpectedly small: {checked}");
    assert_none_wrong(checked, &wrong);
}
