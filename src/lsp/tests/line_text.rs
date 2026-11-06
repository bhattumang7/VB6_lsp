use super::*;

#[test]
fn a_line_is_found_by_the_number_the_editor_gives_it() {
    for sep in ["\n", "\r", "\r\n"] {
        let content = format!("one{sep}Icon = \"F.frx\":0000{sep}three{sep}");
        assert_eq!(line_text(&content, 0), Some("one"), "{sep:?}");
        assert_eq!(line_text(&content, 1), Some("Icon = \"F.frx\":0000"), "{sep:?}");
        assert_eq!(line_text(&content, 2), Some("three"), "{sep:?}");
        assert_eq!(line_text(&content, 3), None, "{sep:?}");
    }
}

#[test]
fn a_line_feed_and_a_carriage_return_are_two_breaks() {
    let content = "a\n\rb";
    assert_eq!(line_text(content, 1), Some(""));
    assert_eq!(line_text(content, 2), Some("b"));
}
