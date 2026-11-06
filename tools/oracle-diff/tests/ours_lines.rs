//! The line a diagnostic is given in the report: counted the way the compiler counts, and for an
//! error inside a continued statement the first line of that statement.

use oracle_diff::mutate::header_lines;
use oracle_diff::oracle::ModuleKind;
use oracle_diff::ours::{diagnostics, OurDiagnostic};

fn lines_of(source: &[u8]) -> Vec<Option<u32>> {
    diagnostics(source, ModuleKind::Bas).iter().map(|d| d.line).collect()
}

#[test]
fn every_kind_of_line_break_counts_one_line() {
    // The bad statement is on the 3rd line whatever ends the lines.
    for sep in ["\n", "\r", "\r\n"] {
        let src = format!("Sub S(){sep}x = 1{sep}x = 1 +* 2{sep}End Sub{sep}");
        assert_eq!(lines_of(src.as_bytes())[0], Some(3), "{sep:?}");
    }
    // A line feed and a carriage return are two breaks, so are two carriage returns before a line feed.
    assert_eq!(lines_of(b"Sub S()\n\rx = 1\n\rx = 1 +* 2\n\rEnd Sub\n\r")[0], Some(5));
    assert_eq!(lines_of(b"Sub S()\r\r\nx = 1\r\r\nx = 1 +* 2\r\r\nEnd Sub\r\r\n")[0], Some(5));
}

#[test]
fn every_piece_of_a_long_line_is_a_line() {
    let src = format!("Sub S()\n' {}\nx = 1 +* 2\nEnd Sub\n", "c".repeat(1100));
    // The comment takes two lines (2 and 3), the bad statement is on the 4th.
    assert_eq!(lines_of(src.as_bytes())[0], Some(4));
    // An exactly full line takes one.
    let src = format!("Sub S()\n' {}\nx = 1 +* 2\nEnd Sub\n", "c".repeat(1021));
    assert_eq!(lines_of(src.as_bytes())[0], Some(3));
}

#[test]
fn an_error_in_a_continued_statement_is_on_its_first_line() {
    let src = b"Sub S()\r\nx = 1 + _\r\n  2 )\r\nEnd Sub\r\n";
    let found = diagnostics(src, ModuleKind::Bas);
    assert_eq!(found.iter().map(|d| d.line).collect::<Vec<_>>(), vec![Some(2)]);
    assert_eq!(found[0].code, 0x9c6f);
    // The error is on the third line of the statement.
    let src = b"Sub S()\r\nx = 1 + _\r\n  2 + _\r\n  3 )\r\nEnd Sub\r\n";
    assert_eq!(lines_of(src), vec![Some(2)]);
    // The next statement is numbered by its own line.
    let src = b"Sub S()\r\nx = 1 + _\r\n  2\r\nx = 1 +* 2\r\nEnd Sub\r\n";
    assert_eq!(lines_of(src)[0], Some(4));
}

#[test]
fn a_limit_found_while_the_text_is_read_has_no_line() {
    let mut src = String::from("Sub S()\nx = 1 +* 2\nx = 1");
    for _ in 0..25 {
        src.push_str(" + _\n 1");
    }
    src.push_str("\nEnd Sub\n");
    let found: Vec<OurDiagnostic> = diagnostics(src.as_bytes(), ModuleKind::Bas);
    assert_eq!(found[0].code, 0x9c00);
    assert_eq!(found[0].line, None);
    assert_eq!(found[0].message, "Too many line continuations");
    assert_eq!(found[1].line, Some(2));
}

#[test]
fn the_header_is_counted_by_the_same_breaks() {
    assert_eq!(header_lines(b"Attribute VB_Name = \"M\"\r\nOption Explicit\r\n"), 1);
    assert_eq!(header_lines(b"Attribute VB_Name = \"M\"\rOption Explicit\r"), 1);
    assert_eq!(header_lines(b"Attribute VB_Name = \"M\"\nOption Explicit\n"), 1);
    assert_eq!(header_lines(b"Option Explicit\n"), 0);
}
