//! A physical line is read up to its 1023rd byte, and a logical line joins at most 24
//! continued lines.

mod common;
use common::parse;

#[test]
fn a_line_is_cut_after_1023_bytes() {
    // `s = "` and `"` take 6 bytes plus 4 of indent: a 1013-byte body makes 1023 in all.
    let ok = format!("Sub S()\n    s = \"{}\"\nEnd Sub\n", "a".repeat(1013));
    assert_eq!(parse(&ok).found, vec![]);
    for extra in [1usize, 2] {
        let long = format!("Sub S()\n    s = \"{}\"\nEnd Sub\n", "a".repeat(1013 + extra));
        assert!(!parse(&long).found.is_empty(), "{extra}");
    }
    // Comment text past the cut does no harm.
    let comment = format!("Sub S()\n    x = 1 ' {}\nEnd Sub\n", "c".repeat(1100));
    assert_eq!(parse(&comment).found, vec![]);
}

#[test]
fn twenty_four_continuations_are_joined_and_a_twenty_fifth_is_an_error() {
    let joined = |n: usize| format!("Sub S()\n    x = 1{}\nEnd Sub\n", " + _\n 1".repeat(n));
    assert_eq!(parse(&joined(24)).found, vec![]);
    let p = parse(&joined(25));
    assert_eq!(p.found.first().map(|f| f.0), Some(0x9c00));
}

#[test]
fn a_full_symbol_table_is_out_of_memory() {
    // 32767 names fill the table of names; the next new one cannot be kept.
    let mut src = String::from("Sub S()\n");
    for chunk in (0..33000).collect::<Vec<usize>>().chunks(60) {
        let names: Vec<String> = chunk.iter().map(|i| format!("q{i}")).collect();
        src.push_str(&format!(" x = {}\n", names.join(" + ")));
    }
    src.push_str("End Sub\n");
    let p = parse(&src);
    assert!(p.found.iter().any(|f| f.0 == 7), "{:?}", p.found.first());
}
