//! `Def` letter lists: a range written backwards covers the same letters as the forward one,
//! and the list goes on after a comma only while a name follows.

mod common;
use common::{offset_of, parse};
use vb6_syntax::frontend::ast::ExprNode;

const DUPLICATE_DEFTYPE: u32 = 0x9c7a;
const UNEXPECTED_TOKEN: u32 = 0x9c6f;

/// The letters of every top-level `Def` statement as the parser records them: `(from, to)`
/// names, with an empty `to` when no range was written.
fn def_statements(src: &str) -> Vec<Vec<(String, String)>> {
    let p = parse(src);
    p.top
        .iter()
        .filter_map(|id| match p.node(*id) {
            ExprNode::DefType { ranges, .. } => Some(
                ranges
                    .iter()
                    .map(|(lo, hi)| (p.name(*lo).to_string(), if *hi == 0 { String::new() } else { p.name(*hi).to_string() }))
                    .collect(),
            ),
            _ => None,
        })
        .collect()
}

fn pair(a: &str, b: &str) -> (String, String) {
    (a.to_string(), b.to_string())
}

#[test]
fn a_range_written_backwards_covers_its_letters_for_the_overlap_check() {
    for later in ["Y", "X", "Z", "X-Z", "Z-X"] {
        let src = format!("DefInt Z-X\nDefLng {later}\n");
        assert_eq!(parse(&src).found, vec![(DUPLICATE_DEFTYPE, offset_of(&src, "DefLng", 0), 6)], "{later}");
    }
}

#[test]
fn letters_outside_a_backwards_range_do_not_overlap() {
    for src in ["DefInt Z-X\nDefLng W\n", "DefInt Z-X\nDefLng A-W\n", "DefInt z-x\nDefLng a-w\n"] {
        assert_eq!(parse(src).found, vec![], "{src}");
    }
}

#[test]
fn a_backwards_range_overlaps_a_forward_one_written_later() {
    let src = "DefInt D-B\nDefStr C\n";
    assert_eq!(parse(src).found, vec![(DUPLICATE_DEFTYPE, offset_of(src, "DefStr", 0), 6)]);
}

#[test]
fn a_backwards_range_is_kept_as_written_in_the_tree() {
    let got = def_statements("DefInt Z-X, B\n");
    assert_eq!(got, vec![vec![pair("Z", "X"), pair("B", "")]]);
}

#[test]
fn a_comma_before_a_colon_ends_the_list() {
    let src = "DefInt A, : DefLng B\n";
    assert_eq!(parse(src).found, vec![]);
    assert_eq!(def_statements(src), vec![vec![pair("A", "")], vec![pair("B", "")]]);
}

#[test]
fn a_comma_before_a_comment_or_the_end_of_the_line_ends_the_list() {
    for src in ["DefInt A,\n", "DefInt A, ' note\n", "DefInt A-C, E,\n", "DefInt A,:DefLng B\n"] {
        assert_eq!(parse(src).found, vec![], "{src}");
    }
}

#[test]
fn a_comma_before_anything_but_a_name_leaves_the_statement_unfinished() {
    for src in ["DefInt A,5\n", "DefInt A,,B\n", "DefInt A, \"x\"\n"] {
        let found = parse(src).found;
        assert_eq!(found.len(), 1, "{src}");
        assert_eq!(found[0].0, UNEXPECTED_TOKEN, "{src}");
    }
}
