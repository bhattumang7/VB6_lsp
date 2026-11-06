//! Line breaks (LF, CR, CR LF), the end of the text (a last Ctrl-Z, a line continued into the
//! end), a NUL in a string, and the limits that stop the reading of a text.

mod common;
use common::{body_shapes, parse};
use vb6_syntax::frontend::ast::ExprArena;
use vb6_syntax::frontend::line_reader::{read, LineLimits};
use vb6_syntax::frontend::parser::Parser;
use vb6_syntax::frontend::scanner::ScannerContext;

const ERR_EXPECTED: u32 = 0x9c6f;
const ERR_TOO_MANY_PIECES: u32 = 0x9c00;
const ERR_OUT_OF_MEMORY: u32 = 7;

#[test]
fn a_nul_ends_a_string_and_the_text_of_its_piece() {
    // The string holds what comes before the NUL; the rest of the line is not read.
    for (line, expected) in [
        ("    x = \"a\0b\"", "a"),
        ("    x = \"a\0\" + \"x\"", "a"),
        ("    x = \"a\0", "a"),
        ("    x = \"\0\"", ""),
        ("    x = \"ab\0   )", "ab"),
    ] {
        let p = parse(&format!("Sub S()\n{line}\nEnd Sub\n"));
        assert_eq!(p.found, vec![], "{line:?}");
        assert_eq!(body_shapes(&p), vec![format!("(= x Str({expected:?}))")], "{line:?}");
    }
    // A NUL right after a closed string ends the line text just the same.
    let p = parse("Sub S()\n    x = \"ab\"\0\n    y = 2\nEnd Sub\n");
    assert_eq!(p.found, vec![]);
    assert_eq!(body_shapes(&p), vec!["(= x Str(\"ab\"))", "(= y Int(2))"]);
}

#[test]
fn a_last_ctrl_z_is_not_part_of_the_text() {
    for src in [
        "Sub S()\r\nEnd Sub\r\n\x1a",
        "Sub S()\r\nEnd Sub\x1a",
        "Sub S()\r\nEnd Sub  \x1a",
        "Sub S()\r\nEnd Sub\r\n\x1a\r",
        "Sub S()\r\nEnd Sub\x1a\r",
        "Sub S()\nEnd Sub\n\x1a",
        "Sub S()\rEnd Sub\r\x1a",
        "Sub S()\r\nEnd Sub\r\n' note\x1a",
    ] {
        let p = parse(src);
        assert_eq!(p.found, vec![], "{src:?}");
        assert_eq!(p.top.len(), 1, "{src:?}");
        assert_eq!(p.proc_body(0).len(), 0, "{src:?}");
    }
}

#[test]
fn a_ctrl_z_anywhere_else_is_an_invalid_character() {
    // Two at the end: the second is dropped, the first is not.
    let src = "Sub S()\r\nEnd Sub\r\n\x1a\x1a";
    let p = parse(src);
    assert_eq!(p.found.len(), 1);
    assert_eq!(p.found[0].0, ERR_EXPECTED);
    assert_eq!(p.found[0].1 as usize, src.find('\x1a').unwrap());
    // Followed by a line break, or by more text.
    for src in ["Sub S()\r\nEnd Sub\r\n\x1a\r\n", "Sub S()\r\n\x1a\r\n    i = 2\r\nEnd Sub\r\n", "Sub S()\r\nEnd Sub\x1a\r\n"] {
        let p = parse(src);
        assert!(!p.found.is_empty(), "{src:?}");
        assert_eq!(p.found[0].1 as usize, src.find('\x1a').unwrap(), "{src:?}");
    }
}

#[test]
fn a_ctrl_z_that_ends_a_statement_line_is_still_invalid_inside_the_text() {
    let p = parse("Sub S()\n    i = 1\x1a\nEnd Sub\n");
    assert_eq!(p.found.len(), 1);
    assert_eq!(p.found[0].1, 17);
}

#[test]
fn text_after_the_end_of_a_procedure_is_still_reported() {
    // Not a Ctrl-Z: ordinary text after `End Sub` on the same line.
    let p = parse("Sub S()\nEnd Sub x\n");
    assert_eq!(p.found.len(), 1);
    assert_eq!(p.found[0].0, 0x9c70);
}

#[test]
fn every_kind_of_line_break_ends_a_statement() {
    for sep in ["\n", "\r", "\r\n", "\n\r", "\r\r\n"] {
        let src = format!("Sub S(){sep}    i = 1{sep}    j = 2{sep}End Sub{sep}");
        let p = parse(&src);
        assert_eq!(p.found, vec![], "{sep:?}");
        assert_eq!(body_shapes(&p), vec!["(= i Int(1))", "(= j Int(2))"], "{sep:?}");
    }
    // A mix of them.
    let p = parse("Sub S()\r    i = 1\n    j = 2\r\n    k = 3\n\rEnd Sub");
    assert_eq!(p.found, vec![]);
    assert_eq!(body_shapes(&p), vec!["(= i Int(1))", "(= j Int(2))", "(= k Int(3))"]);
}

#[test]
fn a_lone_carriage_return_ends_a_comment_and_a_continued_line_goes_on() {
    let p = parse("Sub S()\r    i = 1 ' note\r    j = 2 + _\r        3\rEnd Sub\r");
    assert_eq!(p.found, vec![]);
    assert_eq!(body_shapes(&p), vec!["(= i Int(1))", "(= j (Add Int(2) Int(3)))"]);
}

#[test]
fn a_line_continued_into_the_end_of_the_text_is_dropped() {
    // The text ends in a line that goes on: that line is not read, so the procedure is open.
    let p = parse("Sub S()\r\nEnd Sub\r\nPrivate j As Long _");
    assert_eq!(p.found, vec![]);
    assert_eq!(p.top.len(), 1);
    let p = parse("Sub S()\r\n    i = 1 _\r\n");
    assert_eq!(p.found.len(), 1);
    assert_eq!(p.found[0].0, 0x9c8f);
}

#[test]
fn the_text_after_a_continuation_mark_before_a_nul_is_continuation_too() {
    // The text of the line is `    i = 1 _`: the next line is part of the statement.
    let p = parse("Sub S()\n    i = 1 _\0 junk\n    + 2\nEnd Sub\n");
    assert_eq!(p.found, vec![]);
    assert_eq!(body_shapes(&p), vec!["(= i (Add Int(1) Int(2)))"]);
}

#[test]
fn a_limit_broken_while_the_text_is_read_is_the_first_error_and_the_text_is_still_parsed() {
    // A statement in error comes first in the text, a logical line of 26 pieces after it.
    let mut src = String::from("Sub S()\n    i = 1 +* 2\n    j = 1");
    for _ in 0..25 {
        src.push_str(" + _\n 1");
    }
    src.push_str("\nEnd Sub\n");
    let p = parse(&src);
    let start_of_26th = src.rfind(" 1\nEnd Sub").unwrap() as u32;
    // The `*` after the `+` and the `2` after it are found as the statement is read.
    assert_eq!(
        p.found,
        vec![(ERR_TOO_MANY_PIECES, start_of_26th, 0), (ERR_EXPECTED, 8 + 11, 1), (ERR_EXPECTED, 8 + 13, 1)]
    );
    assert_eq!(body_shapes(&p).len(), 2);
}

#[test]
fn a_module_of_more_than_65530_lines_is_out_of_memory_at_the_first_line_over() {
    let mut src = String::from("Sub S()\n");
    src.push_str(&"' c\n".repeat(65528));
    src.push_str("End Sub\n");
    assert_eq!(parse(&src).found, vec![]);
    let mut src = String::from("Sub S()\n");
    src.push_str(&"' c\n".repeat(65529));
    src.push_str("End Sub\n");
    let p = parse(&src);
    let last = src.rfind("End Sub").unwrap() as u32;
    assert_eq!(p.found, vec![(ERR_OUT_OF_MEMORY, last, 0)]);
    assert_eq!(p.top.len(), 1);
}

#[test]
fn a_header_the_text_does_not_show_is_not_counted() {
    // 20 blank lines stand for a designer section; 65530 comment lines follow.
    let src = format!("{}{}", "\n".repeat(20), "' c\n".repeat(65530));
    let run = |header: Option<usize>| {
        let mut ctx = ScannerContext::new(1, 1, 0x0409);
        ctx.intern_keywords();
        let mut arena = ExprArena::new();
        let mut parser = Parser::new(&mut ctx, src.as_bytes());
        if let Some(lines) = header {
            parser.set_header_lines(lines);
        }
        parser.parse_module(&mut arena);
        parser.diagnostics.items().iter().map(|d| (d.code, d.span.start, d.span.len)).collect::<Vec<_>>()
    };
    assert_eq!(run(None), vec![(ERR_OUT_OF_MEMORY, 20 + 65510 * 4, 0)]);
    assert_eq!(run(Some(20)), vec![]);
    assert_eq!(run(Some(10)), vec![(ERR_OUT_OF_MEMORY, 20 + 65520 * 4, 0)]);
}

#[test]
fn a_piece_found_one_too_many_in_the_last_counted_line_comes_before_the_line_limit() {
    let mut src = "' c\r\n".repeat(65530).into_bytes();
    for _ in 0..25 {
        src.extend_from_slice(b"a _\r\n");
    }
    src.extend_from_slice(b"b\r\n");
    let r = read(&src);
    let problem = r.problem.expect("a problem");
    assert_eq!(problem.code, ERR_TOO_MANY_PIECES);
    assert_eq!(problem.offset as usize, 65530 * 5 + 25 * 5);
    // With the header lines left out the line is not too much any more, the pieces still are.
    assert_eq!(r.limits.problem(5).map(|p| p.code), Some(ERR_TOO_MANY_PIECES));
    assert_eq!(LineLimits::default().problem(0), None);
}
