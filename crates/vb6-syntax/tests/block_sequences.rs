//! Runs of block openers, closers, jumps and plain statements, one per line, in every pairing and sampled triples.

mod common;
use common::answers::{assert_none_wrong, disagreements};

const TEMPLATES: &str = include_str!("fixtures/block_sequences/templates.txt");
const ANSWERS: &str = include_str!("fixtures/block_sequences/oracle.tsv");

#[test]
fn block_sequences_match_the_compiler() {
    let (checked, wrong) = disagreements(TEMPLATES, ANSWERS, |_, word, class, _| class == "semantic" && word.contains("\\n"));
    assert!(checked > 1500, "fixture is unexpectedly small: {checked}");
    assert_none_wrong(checked, &wrong);
}
