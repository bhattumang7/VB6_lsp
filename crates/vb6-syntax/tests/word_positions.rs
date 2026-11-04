//! Every keyword in positions that take a word: a named argument, a jump
//! target, the word after `Exit`, `Option`, `Call`, a With-member and a label.
//! The fixture has the compiler's answer for each.

mod common;
use common::answers::{assert_none_wrong, disagreements};

const TEMPLATES: &str = include_str!("fixtures/word_positions/templates.txt");
const ANSWERS: &str = include_str!("fixtures/word_positions/oracle.tsv");

#[test]
fn words_in_these_positions_match_the_compiler() {
    // The parser reports the missing closer of a block opened by the word itself; the compiler
    // reports its semantic error first.
    let skip = |_: &str, word: &str, class: &str, _: &str| {
        class == "semantic" && ["With", "While", "Case"].contains(&word)
    };
    let (checked, wrong) = disagreements(TEMPLATES, ANSWERS, skip);
    assert!(checked > 2000, "fixture is unexpectedly small: {checked}");
    assert_none_wrong(checked, &wrong);
}
