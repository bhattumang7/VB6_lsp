//! Calls written with every keyword as the callee, in expression and
//! statement position, with every number of arguments and with empty
//! arguments. The fixture holds the compiler's answer for each: the parser
//! must find a form wrong exactly when the compiler finds the form wrong.

mod common;
use common::answers::{assert_none_wrong, disagreements};

const TEMPLATES: &str = include_str!("fixtures/intrinsic_calls/templates.txt");
const ANSWERS: &str = include_str!("fixtures/intrinsic_calls/oracle.tsv");

#[test]
fn calls_with_any_callee_match_the_compiler() {
    // A block opener whose templates do not close the block: the compiler's semantic
    // error comes first and hides the missing closer, so its answer says nothing
    // about the parse.
    let skip = |_: &str, word: &str, class: &str| {
        class == "semantic" && ["With", "While", "Case"].contains(&word)
    };
    let (checked, wrong) = disagreements(TEMPLATES, ANSWERS, skip);
    assert!(checked > 2500, "fixture is unexpectedly small: {checked}");
    assert_none_wrong(checked, &wrong);
}
