//! A physical line is read in pieces of at most 1023 columns (a tab fills the columns up to the
//! next multiple of 8, counted from the start of its piece). The pieces of a line are joined for
//! the scanner but no token spans two of them; nothing of a long line is lost. Each piece counts
//! toward the 25 pieces a logical line may hold.

mod common;
use common::{body_shapes, offset_of, parse};

const ERR_EXPECTED: u32 = 0x9c6f;
const ERR_TOO_MANY_PIECES: u32 = 0x9c00;
const ERR_SYNTAX: u32 = 0x9c70;

/// `prefix` followed by blanks up to `columns` characters.
fn pad_to(prefix: &str, columns: usize) -> String {
    assert!(prefix.len() <= columns);
    format!("{prefix}{}", " ".repeat(columns - prefix.len()))
}

fn sub_of(line: &str) -> String {
    format!("Sub S()\n{line}\nEnd Sub\n")
}

/// Where the line of `sub_of` starts.
const LINE_START: u32 = 8;

/// `Sub S()`, `i = 1 + _`, `extra` lines of `1 + _`, then `last`, then `End Sub`.
fn continued(extra: usize, last: &str) -> String {
    let mut src = String::from("Sub S()\n    i = 1 + _\n");
    for _ in 0..extra {
        src.push_str("    1 + _\n");
    }
    src.push_str(last);
    src.push_str("\nEnd Sub\n");
    src
}

#[test]
fn a_token_split_by_the_end_of_a_piece_is_two_tokens() {
    // `+` is the last column of the first piece, or the first column of the second.
    for (end_of_prefix, label) in [(1022usize, "last column of the first piece"), (1023, "first column of the second")] {
        let src = sub_of(&format!("{}+ 2", pad_to("    i = 1 +", end_of_prefix)));
        let p = parse(&src);
        assert_eq!(p.found, vec![], "{label}");
        assert_eq!(body_shapes(&p), vec!["(= i (Add Int(1) Int(2)))"], "{label}");
    }
}

#[test]
fn a_name_split_by_the_end_of_a_piece_reads_as_two_names() {
    // `ii` with the first `i` in column 1023 and the second in column 1024.
    let src = sub_of(&format!("{}i", pad_to("    j = i", 1023)));
    let p = parse(&src);
    assert_eq!(body_shapes(&p), vec!["(= j i)"]);
    assert_eq!(p.found, vec![(ERR_EXPECTED, LINE_START + 1023, 1)]);
    // At the start of a statement the two names are a call with one argument.
    let src = sub_of(&format!("{}ii", " ".repeat(1022)));
    let p = parse(&src);
    assert_eq!(p.found, vec![]);
    assert_eq!(body_shapes(&p), vec!["(callstmt i [i])"]);
    // A name wholly inside the second piece is one name.
    let src = sub_of(&format!("{}ii", pad_to("    j = ", 1023)));
    let p = parse(&src);
    assert_eq!(p.found, vec![]);
    assert_eq!(body_shapes(&p), vec!["(= j ii)"]);
}

#[test]
fn what_follows_the_end_of_a_piece_is_read() {
    let src = sub_of(&format!("{})", pad_to("    i = 1", 1023)));
    let p = parse(&src);
    assert_eq!(body_shapes(&p), vec!["(= i Int(1))"]);
    assert_eq!(p.found, vec![(ERR_EXPECTED, LINE_START + 1023, 1)]);
}

#[test]
fn a_string_ends_at_the_end_of_its_piece() {
    // The closing quote is the first column of the second piece: the string is not closed in
    // its piece, and the quote opens another string.
    let body = "a".repeat(1014);
    let src = sub_of(&format!("    x = \"{body}\""));
    assert_eq!(src.lines().nth(1).map(str::len), Some(1024));
    let p = parse(&src);
    assert_eq!(body_shapes(&p), vec![format!("(= x Str({body:?}))")]);
    assert_eq!(p.found, vec![(ERR_EXPECTED, LINE_START + 1023, 1)]);
}

#[test]
fn a_comment_runs_to_the_end_of_the_physical_line_over_every_piece() {
    let src = format!("Sub S()\n    i = 1 ' {}\n    j = 2\nEnd Sub\n", "c".repeat(3000));
    let p = parse(&src);
    assert_eq!(p.found, vec![]);
    assert_eq!(body_shapes(&p), vec!["(= i Int(1))", "(= j Int(2))"]);
}

#[test]
fn a_nul_ends_the_text_of_its_piece_and_the_next_piece_is_read() {
    // The NUL is in the first piece; the second holds ` )`.
    let src = sub_of(&format!("{} )", pad_to("    i = 1\0", 1023)));
    let p = parse(&src);
    assert_eq!(body_shapes(&p), vec!["(= i Int(1))"]);
    assert_eq!(p.found, vec![(ERR_EXPECTED, LINE_START + 1024, 1)]);
    // With nothing in the second piece, nothing is wrong.
    let src = sub_of(&pad_to("    i = 1\0", 1023));
    let p = parse(&src);
    assert_eq!(p.found, vec![]);
    assert_eq!(body_shapes(&p), vec!["(= i Int(1))"]);
}

#[test]
fn a_continuation_mark_counts_in_the_last_piece_of_a_line_only() {
    // The mark is the 1025th column: the second piece is ` _`, and the statement goes on.
    let src = format!("Sub S()\n{}_\n    1\nEnd Sub\n", pad_to("    i = 1 +", 1024));
    let p = parse(&src);
    assert_eq!(p.found, vec![]);
    assert_eq!(body_shapes(&p), vec!["(= i (Add Int(1) Int(1)))"]);
    // The underscore is the first column of the second piece, with no blank before it in the
    // piece: it is not a mark.
    let src = format!("Sub S()\n{}_\n    1\nEnd Sub\n", pad_to("    i = 1 +", 1023));
    let p = parse(&src);
    assert_eq!(p.found[0], (ERR_SYNTAX, LINE_START + 1023, 1));
    // An underscore ending the first piece of a line that goes on is not a mark either.
    let src = sub_of(&format!("{}_ 2", pad_to("    i = 1 +", 1022)));
    let p = parse(&src);
    assert_eq!(p.found[0], (ERR_EXPECTED, LINE_START + 1022, 1));
}

#[test]
fn a_logical_line_holds_25_pieces() {
    // 25 physical lines of at most 1023 columns are accepted.
    let p = parse(&continued(23, "    1"));
    assert_eq!(p.found, vec![]);
    let last = format!("    1{}", " ".repeat(1018));
    assert_eq!(last.len(), 1023);
    let p = parse(&continued(23, &last));
    assert_eq!(p.found, vec![]);
    // A 26th physical line is one piece too many, found at the start of that line.
    let src = continued(24, "    1");
    let p = parse(&src);
    let start_of_26th = src.match_indices("    1\nEnd Sub").next().unwrap().0 as u32;
    assert_eq!(p.found[0], (ERR_TOO_MANY_PIECES, start_of_26th, 0));
    // So is a 1024th column on the last of 25 lines: the piece it starts is the 26th.
    let src = continued(23, &format!("{last}x"));
    let p = parse(&src);
    let start_of_26th = src.find("x\nEnd Sub").unwrap() as u32;
    assert_eq!(p.found[0], (ERR_TOO_MANY_PIECES, start_of_26th, 0));
}

#[test]
fn a_tab_that_does_not_fit_in_the_piece_starts_the_next_piece() {
    // A tab at column 1016 (counting from 0) fills up to 1024: it does not fit.
    let fits = continued(23, &format!("    1{}\t", " ".repeat(1010)));
    assert_eq!(parse(&fits).found, vec![]);
    let at_limit = continued(23, &format!("    1{}\t", " ".repeat(1009)));
    assert_eq!(parse(&at_limit).found, vec![]);
    let src = continued(23, &format!("    1{}\t", " ".repeat(1011)));
    let p = parse(&src);
    let tab = src.find('\t').unwrap() as u32;
    assert_eq!(p.found[0], (ERR_TOO_MANY_PIECES, tab, 0));
}

#[test]
fn tab_stops_in_a_string_count_from_the_start_of_the_piece() {
    // The string is in the second piece: its tab is at column 2, so it fills 6 columns.
    let src = sub_of(&format!("{}\"a\tb\"", pad_to("    x = ", 1023)));
    let p = parse(&src);
    assert_eq!(p.found, vec![]);
    assert_eq!(body_shapes(&p), vec![format!("(= x Str({:?}))", "a      b")]);
    // In the first piece the stops count from the start of the line.
    let p = parse(&sub_of("    x = \"ab\tc\""));
    assert_eq!(body_shapes(&p), vec![format!("(= x Str({:?}))", "ab     c")]);
}

#[test]
fn token_spans_are_offsets_in_the_text_whatever_is_split() {
    // The statement after a split line starts where the text says.
    let src = format!("Sub S()\n{}+ 2\n    j = 3\nEnd Sub\n", pad_to("    i = 1 +", 1022));
    let p = parse(&src);
    assert_eq!(p.found, vec![]);
    assert_eq!(body_shapes(&p), vec!["(= i (Add Int(1) Int(2)))", "(= j Int(3))"]);
    let j = offset_of(&src, "j = 3", 0);
    assert_eq!(&src.as_bytes()[j as usize..j as usize + 5], b"j = 3");
}
