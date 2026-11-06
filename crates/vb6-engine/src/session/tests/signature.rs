use super::*;

#[test]
fn the_call_around_the_cursor_is_found_with_its_argument_index() {
    let src = b"x = Foo(1, Bar(2), ";
    assert_eq!(find_call_context(src, src.len() as u32), Some(("Foo".to_string(), 2)));
}

#[test]
fn the_search_does_not_cross_a_line_break_of_any_kind() {
    for sep in ["\n", "\r", "\r\n"] {
        let text = format!("x = Foo({sep}y = 1, ");
        assert_eq!(find_call_context(text.as_bytes(), text.len() as u32), None, "{sep:?}");
    }
}
