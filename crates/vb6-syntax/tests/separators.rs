//! Whitespace, tabs, line continuations, colons and comments between the words
//! of multi-word statements and around operators, with the compiler's answer for each.

mod common;
use common::answers::{assert_none_wrong, disagreements};

const TEMPLATES: &str = include_str!("fixtures/separators/templates.txt");
const ANSWERS: &str = include_str!("fixtures/separators/oracle.tsv");

#[test]
fn separators_between_words_match_the_compiler() {
    let (checked, wrong) = disagreements(TEMPLATES, ANSWERS, |_, word, class, _| {
        // The compiler binds each line as it parses it, so a name it cannot find on an earlier
        // line is reported before a syntax error on a later one; such answers say nothing about
        // the syntax.
        class == "semantic" && word.contains("\\n")
    });
    assert!(checked > 500, "fixture is unexpectedly small: {checked}");
    assert_none_wrong(checked, &wrong);
}
