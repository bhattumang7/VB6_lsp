use super::*;

#[test]
fn single_line() {
    let idx = LineIndex::new(b"Dim x As Long");
    assert_eq!(idx.position(0), Position { line: 0, character: 0 });
    assert_eq!(idx.position(4), Position { line: 0, character: 4 });
}

#[test]
fn multi_line_lf() {
    let src = b"Sub Foo()\n    Dim y\nEnd Sub\n";
    let idx = LineIndex::new(src);
    // "Dim y" — the 'D' is at byte offset 14 (after "Sub Foo()\n    ").
    let d = src.iter().position(|&b| b == b'D').unwrap() as u32;
    assert_eq!(idx.position(d), Position { line: 1, character: 4 });
    // Start of line 2 ("End Sub").
    let e = "Sub Foo()\n    Dim y\n".len() as u32;
    assert_eq!(idx.position(e), Position { line: 2, character: 0 });
}

#[test]
fn crlf_keeps_cr_in_line() {
    let src = b"AB\r\nCD";
    let idx = LineIndex::new(src);
    // 'C' is line 1, char 0 (the \r\n both belong to line 0's terminator).
    assert_eq!(idx.position(4), Position { line: 1, character: 0 });
    // The \r is line 0, char 2.
    assert_eq!(idx.position(2), Position { line: 0, character: 2 });
}

#[test]
fn a_lone_carriage_return_ends_a_line() {
    let src = b"AB\rCD\rEF";
    let idx = LineIndex::new(src);
    assert_eq!(idx.line_count(), 3);
    assert_eq!(idx.position(3), Position { line: 1, character: 0 });
    assert_eq!(idx.position(4), Position { line: 1, character: 1 });
    assert_eq!(idx.position(6), Position { line: 2, character: 0 });
    assert_eq!(idx.position(7), Position { line: 2, character: 1 });
    // The carriage return is the last character of its own line.
    assert_eq!(idx.position(2), Position { line: 0, character: 2 });
    assert_eq!(idx.offset(Position { line: 2, character: 1 }), 7);
}

#[test]
fn a_line_feed_then_a_carriage_return_are_two_breaks() {
    let src = b"A\n\rB";
    let idx = LineIndex::new(src);
    assert_eq!(idx.line_count(), 3);
    assert_eq!(idx.position(2), Position { line: 1, character: 0 });
    assert_eq!(idx.position(3), Position { line: 2, character: 0 });
}

#[test]
fn two_carriage_returns_before_a_line_feed_are_two_breaks() {
    let src = b"A\r\r\nB";
    let idx = LineIndex::new(src);
    assert_eq!(idx.line_count(), 3);
    assert_eq!(idx.position(2), Position { line: 1, character: 0 });
    assert_eq!(idx.position(4), Position { line: 2, character: 0 });
}

#[test]
fn a_final_break_leaves_an_empty_last_line() {
    for src in [&b"A\n"[..], b"A\r", b"A\r\n"] {
        let idx = LineIndex::new(src);
        assert_eq!(idx.line_count(), 2, "{src:?}");
        assert_eq!(idx.position(src.len() as u32), Position { line: 1, character: 0 }, "{src:?}");
    }
}

#[test]
fn round_trip() {
    let src = b"alpha\nbeta gamma\rdelta\r\nomega\n";
    let idx = LineIndex::new(src);
    for off in 0..=src.len() as u32 {
        let pos = idx.position(off);
        assert_eq!(idx.offset(pos), off.min(src.len() as u32), "offset {off}");
    }
}

#[test]
fn column_assumption_cp1252() {
    // 0xE9 = 'é' in Windows-1252, a single byte = single UTF-16 unit.
    // Column math is pure byte distance, so the identifier after it lands
    // at the expected character index.
    let src = b"Dim caf\xE9 As Long";
    let idx = LineIndex::new(src);
    // "As" begins after "Dim café " = 9 bytes.
    let a = src.windows(2).position(|w| w == b"As").unwrap() as u32;
    assert_eq!(idx.position(a), Position { line: 0, character: a });
}

#[test]
fn offset_past_end_clamps() {
    let idx = LineIndex::new(b"abc");
    assert_eq!(idx.position(99), Position { line: 0, character: 3 });
}

#[test]
fn span_to_range() {
    let src = b"Sub Foo()\nEnd Sub\n";
    let idx = LineIndex::new(src);
    let span = Span { start: 4, len: 3 }; // "Foo"
    let r = idx.range(span);
    assert_eq!(r.start, Position { line: 0, character: 4 });
    assert_eq!(r.end, Position { line: 0, character: 7 });
}

#[test]
fn a_span_in_a_carriage_return_file_maps_to_its_own_line() {
    let src = b"Sub Foo()\rEnd Sub\r";
    let idx = LineIndex::new(src);
    let r = idx.range(Span { start: 10, len: 3 }); // "End"
    assert_eq!(r.start, Position { line: 1, character: 0 });
    assert_eq!(r.end, Position { line: 1, character: 3 });
}
