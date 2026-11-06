//! Physical lines, pieces, logical lines and the limits on them, as the line reader sees a text.

use std::borrow::Cow;
use vb6_syntax::frontend::diagnostics::{ERR_OUT_OF_MEMORY, ERR_TOO_MANY_LINE_CONTINUATIONS};
use vb6_syntax::frontend::line_reader::{
    physical_line_starts, physical_lines, read, source_offset, LineTable, PhysicalLine, ReadProblem, MAX_LINES,
};

fn line(start: usize, end: usize, next: usize) -> PhysicalLine {
    PhysicalLine { start, end, next }
}

#[test]
fn a_line_ends_at_lf_cr_or_cr_lf() {
    let lines: Vec<_> = physical_lines(b"a\r\nbc\rd\n\ne").collect();
    assert_eq!(lines, vec![line(0, 1, 3), line(3, 5, 6), line(6, 7, 8), line(8, 8, 9), line(9, 10, 10)]);
}

#[test]
fn every_other_pairing_of_breaks_is_two_breaks() {
    let cases: [(&[u8], Vec<PhysicalLine>); 8] = [
        (b"", vec![]),
        (b"a", vec![line(0, 1, 1)]),
        (b"a\n", vec![line(0, 1, 2)]),
        (b"\r", vec![line(0, 0, 1)]),
        (b"\r\n", vec![line(0, 0, 2)]),
        (b"\n\r", vec![line(0, 0, 1), line(1, 1, 2)]),
        (b"\r\r\n", vec![line(0, 0, 1), line(1, 1, 3)]),
        (b"a\n\n", vec![line(0, 1, 2), line(2, 2, 3)]),
    ];
    for (text, expected) in cases {
        assert_eq!(physical_lines(text).collect::<Vec<_>>(), expected, "{text:?}");
    }
}

#[test]
fn line_starts_count_every_break_once() {
    assert_eq!(physical_line_starts(b""), vec![0]);
    assert_eq!(physical_line_starts(b"ab"), vec![0]);
    assert_eq!(physical_line_starts(b"ab\r\ncd\ref\ngh"), vec![0, 4, 7, 10]);
    assert_eq!(physical_line_starts(b"a\n"), vec![0, 2]);
    assert_eq!(physical_line_starts(b"a\r"), vec![0, 2]);
    assert_eq!(physical_line_starts(b"a\r\n"), vec![0, 3]);
    assert_eq!(physical_line_starts(b"\n\r"), vec![0, 1, 2]);
    assert_eq!(physical_line_starts(b"\r\r\n"), vec![0, 1, 3]);
}

#[test]
fn a_short_text_is_read_as_it_is() {
    let src = b"Sub S()\r\n    i = 1\r\nEnd Sub\r\n";
    let r = read(src);
    assert!(matches!(r.text, Cow::Borrowed(_)));
    assert_eq!(&*r.text, &src[..]);
    assert_eq!(r.separators, Vec::<u32>::new());
    assert_eq!(r.problem, None);
}

#[test]
fn a_long_line_gets_a_nul_where_its_next_piece_starts() {
    let mut src = vec![b'a'; 1023];
    src.extend_from_slice(b"bcd\r\nxyz");
    let r = read(&src);
    assert_eq!(r.separators, vec![1023]);
    assert_eq!(r.text.len(), src.len() + 1);
    assert_eq!(&r.text[1020..1028], b"aaa\0bcd\r");
    assert_eq!(source_offset(&r.separators, 1022), 1022);
    assert_eq!(source_offset(&r.separators, 1023), 1023);
    assert_eq!(source_offset(&r.separators, 1024), 1023);
    assert_eq!(source_offset(&r.separators, 1026), 1025);
    // A line of exactly 1023 columns is one piece.
    let exact = vec![b'a'; 1023];
    assert_eq!(read(&exact).separators, Vec::<u32>::new());
    // The next line starts again at column 0.
    let mut src = vec![b'a'; 1500];
    src.extend_from_slice(b"\r\n");
    src.extend(vec![b'b'; 1023]);
    src.extend_from_slice(b"\r\n");
    src.extend(vec![b'c'; 1024]);
    let r = read(&src);
    assert_eq!(r.separators, vec![1023, 1500 + 2 + 1023 + 2 + 1023 + 1]);
}

#[test]
fn a_tab_fills_to_the_next_multiple_of_8_and_may_start_a_piece() {
    // 1015 characters, then a tab (fills column 1015 to 1016): it fits.
    let mut src = vec![b'a'; 1015];
    src.push(b'\t');
    src.push(b'x');
    assert_eq!(read(&src).separators, Vec::<u32>::new());
    // 1016 characters, then a tab (fills 1016 to 1024): the tab starts the next piece.
    let mut src = vec![b'a'; 1016];
    src.push(b'\t');
    src.push(b'x');
    assert_eq!(read(&src).separators, vec![1016]);
    // In the new piece the tab starts at column 0 again, so a second tab fills 8 columns.
    let mut src = vec![b'a'; 1016];
    src.extend_from_slice(b"\t\tx");
    assert_eq!(read(&src).separators, vec![1016]);
    // Tabs inside the line count toward the columns: 127 tabs are 1016 columns.
    let mut src = vec![b'\t'; 127];
    src.extend_from_slice(b"ab");
    assert_eq!(read(&src).separators, Vec::<u32>::new());
    let mut src = vec![b'\t'; 128];
    src.push(b'a');
    assert_eq!(read(&src).separators, vec![127]);
}

#[test]
fn a_last_ctrl_z_is_read_as_a_nul() {
    for (src, nul_at) in [
        (&b"End Sub\x1a"[..], Some(7)),
        (b"End Sub\r\n\x1a", Some(9)),
        (b"End Sub  \x1a", Some(9)),
        // A carriage return that is the last byte leaves its line the last one.
        (b"End Sub\r\n\x1a\r", Some(9)),
        (b"End Sub\x1a\r", Some(7)),
        // Not the last line.
        (b"End Sub\x1a\r\n", None),
        (b"End Sub\x1a\n", None),
        (b"\x1a\x1a", Some(1)),
        (b"\x1aEnd Sub", None),
        (b"\x1a\r\nEnd Sub", None),
    ] {
        let r = read(src);
        let mut expected = src.to_vec();
        if let Some(at) = nul_at {
            expected[at] = 0;
        }
        assert_eq!(&*r.text, &expected[..], "{src:?}");
    }
}

#[test]
fn a_line_continued_into_the_end_of_the_text_is_not_read() {
    for (src, kept) in [
        (&b"a _"[..], 0),
        (b"a _\r\n", 0),
        (b"x\r\na _", 3),
        (b"x\r\na _\r\nb _\r\n", 3),
        (b"x\r\na _  \t", 3),
        // A further line, even an empty one, ends the continuation.
        (b"a _\r\n\r\n", 7),
        (b"a _\r\nb", 6),
        // Not a mark: no blank before the underscore.
        (b"a_", 2),
        (b"_", 1),
    ] {
        let r = read(src);
        assert_eq!(&*r.text, &src[..kept], "{src:?}");
    }
}

#[test]
fn a_nul_ends_the_text_of_a_piece_for_the_continuation_mark() {
    // The text of the piece is `a _`: the line is continued into the end of the text.
    assert_eq!(&*read(b"x\na _\0 abc").text, b"x\n");
    // An end-of-text 0x1A read as a NUL after a mark does the same.
    assert_eq!(&*read(b"x\na _\x1a").text, b"x\n");
}

#[test]
fn lines_are_numbered_by_piece() {
    let mut src = vec![b'a'; 1100];
    src.extend_from_slice(b"\r\nb\rc\n");
    let t = LineTable::new(&src);
    assert_eq!(t.len(), 4);
    assert_eq!(t.line_of(0), 1);
    assert_eq!(t.line_of(1022), 1);
    assert_eq!(t.line_of(1023), 2);
    assert_eq!(t.line_of(1099), 2);
    // The line break belongs to the line it ends.
    assert_eq!(t.line_of(1100), 2);
    assert_eq!(t.line_of(1102), 3);
    assert_eq!(t.line_of(1104), 4);
    assert_eq!(t.line_of(1105), 4);
}

#[test]
fn the_end_of_a_text_that_ends_in_a_break_is_on_an_empty_last_line() {
    for src in [&b"a\nb\n"[..], b"a\nb\r", b"a\nb\r\n"] {
        let t = LineTable::new(src);
        let end = src.len() as u32;
        assert_eq!(t.line_of(end), 3, "{src:?}");
        assert_eq!(t.logical_line_of(end), 3, "{src:?}");
        assert_eq!(t.logical_line_start(end), end, "{src:?}");
        // The last line's own break still belongs to it.
        assert_eq!(t.line_of(end - 1), 2, "{src:?}");
    }
    // Without a final break the end is in the last line.
    let t = LineTable::new(b"a\nb");
    assert_eq!(t.line_of(3), 2);
    assert_eq!(t.logical_line_of(3), 2);
    assert_eq!(t.logical_line_start(3), 2);
}

#[test]
fn an_error_in_a_continued_statement_is_on_its_first_line() {
    let src = b"a\r\nb + _\r\n  c _\r\n  d\r\ne";
    let t = LineTable::new(src);
    let at = |needle: &str| src.windows(needle.len()).position(|w| w == needle.as_bytes()).unwrap() as u32;
    assert_eq!(t.logical_line_of(at("a")), 1);
    assert_eq!(t.logical_line_of(at("b")), 2);
    assert_eq!(t.logical_line_of(at("c")), 2);
    assert_eq!(t.logical_line_of(at("d")), 2);
    assert_eq!(t.logical_line_of(at("e")), 5);
    assert_eq!(t.logical_line_start(at("d")), at("b"));
    assert_eq!(t.logical_line_start(at("e")), at("e"));
    // Each physical line keeps its own number.
    assert_eq!(t.line_of(at("c")), 3);
    assert_eq!(t.line_of(at("d")), 4);
}

#[test]
fn the_pieces_of_a_long_line_belong_to_one_logical_line() {
    let mut src = b"x\r\n".to_vec();
    src.extend(vec![b'a'; 2100]);
    src.extend_from_slice(b"\r\ny");
    let t = LineTable::new(&src);
    // The long line is lines 2, 3 and 4; the next one is line 5.
    assert_eq!(t.line_of(3), 2);
    assert_eq!(t.line_of(3 + 1023), 3);
    assert_eq!(t.line_of(3 + 2046), 4);
    assert_eq!(t.logical_line_of(3 + 2046), 2);
    assert_eq!(t.logical_line_start(3 + 2046), 3);
    let y = src.len() as u32 - 1;
    assert_eq!(t.line_of(y), 5);
    assert_eq!(t.logical_line_of(y), 5);
}

#[test]
fn a_26th_piece_is_a_problem_found_at_its_start() {
    let mut src = Vec::new();
    for _ in 0..25 {
        src.extend_from_slice(b"a _\r\n");
    }
    src.extend_from_slice(b"b\r\n");
    assert_eq!(read(&src).problem, Some(ReadProblem { code: ERR_TOO_MANY_LINE_CONTINUATIONS, offset: 125 }));
    // 25 pieces are fine.
    let mut src = Vec::new();
    for _ in 0..24 {
        src.extend_from_slice(b"a _\r\n");
    }
    src.extend_from_slice(b"b\r\n");
    assert_eq!(read(&src).problem, None);
    // The pieces of long lines count the same.
    let mut src = Vec::new();
    for _ in 0..24 {
        src.extend_from_slice(b"a _\r\n");
    }
    src.extend(vec![b'b'; 1024]);
    assert_eq!(read(&src).problem, Some(ReadProblem { code: ERR_TOO_MANY_LINE_CONTINUATIONS, offset: 120 + 1023 }));
    // Only the first problem is kept, and the next logical line starts the count again.
    let mut src = Vec::new();
    for _ in 0..30 {
        src.extend_from_slice(b"a _\r\n");
    }
    src.extend_from_slice(b"b\r\nc\r\n");
    assert_eq!(read(&src).problem.map(|p| p.offset), Some(125));
}

fn comment_lines(n: usize) -> Vec<u8> {
    b"' c\r\n".repeat(n)
}

#[test]
fn a_module_holds_65530_lines() {
    assert_eq!(MAX_LINES, 65530);
    assert_eq!(read(&comment_lines(65530)).problem, None);
    let src = comment_lines(65531);
    assert_eq!(read(&src).problem, Some(ReadProblem { code: ERR_OUT_OF_MEMORY, offset: 65530 * 5 }));
    // A last line without a break counts, and so does an empty line.
    let mut src = comment_lines(65529);
    src.extend_from_slice(b"' c");
    assert_eq!(read(&src).problem, None);
    let mut src = comment_lines(65530);
    src.extend_from_slice(b"' c");
    assert_eq!(read(&src).problem, Some(ReadProblem { code: ERR_OUT_OF_MEMORY, offset: 65530 * 5 }));
    let mut src = comment_lines(65530);
    src.extend_from_slice(b"\r\n");
    assert_eq!(read(&src).problem, Some(ReadProblem { code: ERR_OUT_OF_MEMORY, offset: 65530 * 5 }));
    // Every kind of break counts as one.
    assert_eq!(read(&b"' c\r".repeat(65530)).problem, None);
    assert_eq!(read(&b"' c\r".repeat(65531)).problem.map(|p| p.code), Some(ERR_OUT_OF_MEMORY));
    assert_eq!(read(&b"' c\n".repeat(65531)).problem.map(|p| p.code), Some(ERR_OUT_OF_MEMORY));
}

#[test]
fn a_continued_line_and_a_long_line_count_once() {
    // 40000 statements of two lines each are 80000 lines but 40000 logical lines.
    assert_eq!(read(&b"x = 1 + _\r\n 2\r\n".repeat(40000)).problem, None);
    // 65531 logical lines of two lines each.
    let src = b"x = 1 + _\r\n 2\r\n".repeat(65531);
    assert_eq!(read(&src).problem, Some(ReadProblem { code: ERR_OUT_OF_MEMORY, offset: 65530 * 15 }));
    // Pieces of a long line do not count.
    let mut line = b"' ".to_vec();
    line.extend(vec![b'c'; 1100]);
    line.extend_from_slice(b"\r\n");
    assert_eq!(read(&line.repeat(65530)).problem, None);
}

#[test]
fn header_lines_are_not_counted() {
    let class_header = b"VERSION 1.0 CLASS\r\nBEGIN\r\n  MultiUse = -1  'True\r\nEND\r\nAttribute VB_Name = \"Class1\"\r\nAttribute VB_GlobalNameSpace = False\r\n";
    let mut src = class_header.to_vec();
    src.extend(comment_lines(65530));
    assert_eq!(read(&src).problem, None);
    let mut src = class_header.to_vec();
    src.extend(comment_lines(65531));
    assert_eq!(
        read(&src).problem,
        Some(ReadProblem { code: ERR_OUT_OF_MEMORY, offset: (class_header.len() + 65530 * 5) as u32 })
    );
    // An attribute line anywhere is a header line.
    let mut src = b"Sub S()\r\nAttribute S.VB_Description = \"x\"\r\n".to_vec();
    src.extend(comment_lines(65528));
    src.extend_from_slice(b"End Sub\r\n");
    assert_eq!(read(&src).problem, None);
    let mut src = b"Sub S()\r\nAttribute S.VB_Description = \"x\"\r\n".to_vec();
    src.extend(comment_lines(65529));
    src.extend_from_slice(b"End Sub\r\n");
    assert_eq!(read(&src).problem.map(|p| p.code), Some(ERR_OUT_OF_MEMORY));
    // `Begin` and `Version` after code are ordinary lines.
    let mut src = b"x = 1\r\nVERSION 1.0 CLASS\r\n".to_vec();
    src.extend(comment_lines(65528));
    assert_eq!(read(&src).problem, None);
    src.extend_from_slice(b"' c\r\n");
    assert_eq!(read(&src).problem.map(|p| p.code), Some(ERR_OUT_OF_MEMORY));
}

#[test]
fn nested_blocks_of_a_designer_header_close_in_pairs() {
    let header = b"VERSION 5.00\r\nBegin VB.Form Form1\r\n   Begin VB.Button B\r\n      BeginProperty Font\r\n         Name = \"x\"\r\n      EndProperty\r\n   End\r\nEnd\r\nAttribute VB_Name = \"Form1\"\r\n";
    let mut src = header.to_vec();
    src.extend(comment_lines(65530));
    assert_eq!(read(&src).problem, None);
    src.extend_from_slice(b"' c\r\n");
    assert_eq!(read(&src).problem.map(|p| p.code), Some(ERR_OUT_OF_MEMORY));
}
