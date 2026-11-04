//! Where a reserved word may be used as a name. The fixture holds the
//! compiler's answer for every keyword in every declaration position of
//! `templates.txt`; the parser must accept exactly the forms the compiler
//! accepts (or rejects only for what the names mean) and reject the others.

mod common;
use common::answers::{assert_none_wrong, disagreements};

const TEMPLATES: &str = include_str!("fixtures/keyword_names/templates.txt");
const ANSWERS: &str = include_str!("fixtures/keyword_names/oracle.tsv");

#[test]
fn keywords_as_names_match_the_compiler_in_every_position() {
    // A block opener whose templates do not close the block: the compiler's semantic
    // error comes first and hides the missing closer, so its answer says nothing
    // about the parse.
    let skip = |template: &str, word: &str, class: &str, _: &str| {
        // The parser reports a `Next` variable that is not the loop's control variable itself;
        // the compiler reports it as a semantic error, whatever the word is.
        (class == "semantic" && template == "nextvar")
            || (class == "semantic" && ["With", "While", "Case"].contains(&word) && template.starts_with("start"))
    };
    let (checked, wrong) = disagreements(TEMPLATES, ANSWERS, skip);
    assert!(checked > 5000, "fixture is unexpectedly small: {checked}");
    assert_none_wrong(checked, &wrong);
}
