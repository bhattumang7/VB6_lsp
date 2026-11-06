//! Which texts between `#` marks are date or time literals. The text goes to the runtime's
//! date parser (English, United States), so the fixture holds the compiler's answer for
//! separators, signs, numbers of every size, month and weekday names, times, and blanks.

mod common;
use common::answers::{assert_none_wrong, disagreements};

const TEMPLATES: &str = include_str!("fixtures/date_literal_forms/templates.txt");
const ANSWERS: &str = include_str!("fixtures/date_literal_forms/oracle.tsv");

#[test]
fn date_literal_forms_match_the_compiler() {
    let (checked, wrong) = disagreements(TEMPLATES, ANSWERS, |_, _, _, _| false);
    assert!(checked > 1000, "fixture is unexpectedly small: {checked}");
    assert_none_wrong(checked, &wrong);
}
