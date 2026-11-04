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
    // The parser reports a `Next` variable that is not the loop's control variable itself;
    // the compiler reports it as a semantic error. Both reject, and the parse is fine either way.
    let skip = |_: &str, _: &str, class: &str, message: &str| {
        class == "semantic" && message == "Invalid Next control variable reference"
    };
    let (checked, wrong) = disagreements(TEMPLATES, ANSWERS, skip);
    assert!(checked > 300, "fixture is unexpectedly small: {checked}");
    assert_none_wrong(checked, &wrong);
}
