//! Conditional-compilation directives: #Const, #If, #ElseIf, #Else, #End If, and what may follow them.

mod common;
use common::answers::{assert_none_wrong, disagreements};

const TEMPLATES: &str = include_str!("fixtures/conditional_compilation_forms/templates.txt");
const ANSWERS: &str = include_str!("fixtures/conditional_compilation_forms/oracle.tsv");

#[test]
fn conditional_compilation_forms_match_the_compiler() {
    let (checked, wrong) = disagreements(TEMPLATES, ANSWERS, |_, word, class, message| {
        // The value of a condition is evaluated while the directive is read: a string or an
        // object where a number belongs is the evaluator's complaint, not the grammar's.
        class == "semantic" && matches!(message, "Type mismatch" | "Invalid use of object") && word.starts_with("#If ")
    });
    assert!(checked > 200, "fixture is unexpectedly small: {checked}");
    assert_none_wrong(checked, &wrong);
}
