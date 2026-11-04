//! A label, by name or by line number, is defined once per procedure.

mod common;
use common::{offset_of, parse};

const DUPLICATE_DECLARATION: u32 = 0x9c9f;

#[test]
fn a_second_definition_of_a_label_is_reported_at_the_second_one() {
    for (body, needle, nth) in [
        ("L:\n    L:", "L:", 1),
        ("L:\n    l:", "l:", 0),
        ("10\n    10", "10", 1),
        ("10 x = 1\n    10 x = 2", "10", 1),
        ("L: x = 1\n    L: x = 2", "L:", 1),
        ("L:\n    GoTo L\n    L:", "L:", 1),
    ] {
        let src = format!("Sub S()\n    Dim x\n    {body}\nEnd Sub\n");
        let found = parse(&src).found;
        assert_eq!(found.len(), 1, "{body}: {found:?}");
        assert_eq!(found[0].0, DUPLICATE_DECLARATION, "{body}");
        let want = if body.starts_with("L:\n    l:") { offset_of(&src, "l:", 0) } else { offset_of(&src, needle, nth) };
        assert_eq!(found[0].1, want, "{body}");
    }
}

#[test]
fn a_name_and_a_number_do_not_clash_and_procedures_are_separate() {
    assert_eq!(parse("Sub S()\n    L:\n    10\nEnd Sub\n").found, vec![]);
    assert_eq!(parse("Sub S()\n    L:\nEnd Sub\nSub T()\n    L:\nEnd Sub\n").found, vec![]);
}
