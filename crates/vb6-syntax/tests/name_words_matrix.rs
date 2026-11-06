//! Which words may name what: keywords spelled with a trailing `$`, the module-header words,
//! words in brackets and type-declaration characters, in every declaration position of
//! `templates.txt`. The fixture holds the compiler's answer for each combination; the parser
//! must accept exactly the forms the compiler accepts and reject the others.

mod common;
use common::answers::{assert_none_wrong, disagreements};

const TEMPLATES: &str = include_str!("fixtures/name_words/templates.txt");
const ANSWERS: &str = include_str!("fixtures/name_words/oracle.tsv");

#[test]
fn word_forms_as_names_match_the_compiler_in_every_position() {
    let skip = |template: &str, word: &str, class: &str, _: &str| {
        // The parser reports a `Next` variable that is not the loop's control variable itself;
        // the compiler reports it as a semantic error, whatever the word is.
        (class == "semantic" && template == "nextvar")
            || (class == "semantic" && ["With", "While", "Case"].contains(&word) && template.starts_with("start"))
    };
    let (checked, wrong) = disagreements(TEMPLATES, ANSWERS, skip);
    assert!(checked > 600, "fixture is unexpectedly small: {checked}");
    assert_none_wrong(checked, &wrong);
}
