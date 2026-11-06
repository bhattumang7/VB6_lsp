use oracle_diff::matrix::{instantiate, parse_templates, parse_words, Fate, Template};
use oracle_diff::mutate::{header_lines, mutants, pieces, Rng, VOCABULARY};
use oracle_diff::oracle::ModuleKind;

const MODULE: &[u8] = b"Sub S()\r\n    x = 1 + 2 ' note\r\nEnd Sub\r\n";

fn texts(source: &[u8]) -> Vec<String> {
    pieces(source)
        .iter()
        .filter(|p| !p.is_eol)
        .map(|p| String::from_utf8_lossy(&source[p.start..p.end]).into_owned())
        .collect()
}

#[test]
fn pieces_are_the_code_tokens_without_comments() {
    assert_eq!(
        texts(MODULE),
        ["Sub", "S", "(", ")", "x", "=", "1", "+", "2", "End", "Sub"]
    );
}

#[test]
fn pieces_mark_line_breaks() {
    let breaks = pieces(MODULE).iter().filter(|p| p.is_eol).count();
    assert_eq!(breaks, 3);
}

#[test]
fn the_same_seed_gives_the_same_mutants() {
    let a = mutants(MODULE, 7, 12);
    let b = mutants(MODULE, 7, 12);
    assert_eq!(a.len(), 12);
    for (x, y) in a.iter().zip(&b) {
        assert_eq!((&x.description, x.line, &x.source), (&y.description, y.line, &y.source));
    }
}

#[test]
fn a_different_seed_gives_different_mutants() {
    let a = mutants(MODULE, 1, 12);
    let b = mutants(MODULE, 2, 12);
    assert!(a.iter().zip(&b).any(|(x, y)| x.source != y.source));
}

#[test]
fn every_mutant_differs_from_its_parent_on_a_real_line() {
    for m in mutants(MODULE, 3, 40) {
        assert_ne!(m.source, MODULE.to_vec(), "{}", m.description);
        assert!((1..=3).contains(&m.line), "{} line {}", m.description, m.line);
        assert!(!m.description.is_empty());
    }
}

#[test]
fn module_headers_are_never_mutated_and_are_not_counted_in_lines() {
    let class = b"VERSION 1.0 CLASS\r\nBEGIN\r\n  MultiUse = -1  'True\r\nEND\r\nAttribute VB_Name = \"C\"\r\nSub S()\r\nEnd Sub\r\n";
    assert_eq!(header_lines(class), 5);
    assert_eq!(header_lines(MODULE), 0);
    let header_len = class.len() - b"Sub S()\r\nEnd Sub\r\n".len();
    for m in mutants(class, 5, 60) {
        assert!(m.source.starts_with(&class[..header_len]), "{}", m.description);
    }
}

#[test]
fn only_the_leading_block_is_header_and_later_attribute_or_end_lines_are_code() {
    let std_module = b"Attribute VB_Name = \"M\"\r\nSub S()\r\nEnd\r\nEnd Sub\r\nAttribute VB_Foo = 1\r\n";
    assert_eq!(header_lines(std_module), 1);
    let blank_between = b"\r\nAttribute VB_Name = \"M\"\r\n\r\nSub S()\r\nEnd Sub\r\n";
    assert_eq!(header_lines(blank_between), 2);
    let no_header = b"Sub S()\r\nEnd\r\nEnd Sub\r\n";
    assert_eq!(header_lines(no_header), 0);
    let class = b"VERSION 1.0 CLASS\r\nBEGIN\r\n  MultiUse = -1  'True\r\nEND\r\nAttribute VB_Name = \"C\"\r\nSub S()\r\nEnd\r\nEnd Sub\r\n";
    assert_eq!(header_lines(class), 5);
}

#[test]
fn rng_is_deterministic_and_bounded() {
    let mut a = Rng::new(9);
    let mut b = Rng::new(9);
    for _ in 0..50 {
        let x = a.below(10);
        assert_eq!(x, b.below(10));
        assert!(x < 10);
    }
    assert!(VOCABULARY.contains(&"Then"));
}

#[test]
fn templates_parse_with_their_kind_and_text() {
    let text = "# comment\n=== one bas\nSub S()\n    Dim {k}\nEnd Sub\n=== two cls\nEvent {k}()\n";
    let t = parse_templates(text).unwrap();
    assert_eq!(
        t,
        vec![
            Template { name: "one".into(), kind: ModuleKind::Bas, text: "Sub S()\n    Dim {k}\nEnd Sub\n".into() },
            Template { name: "two".into(), kind: ModuleKind::Cls, text: "Event {k}()\n".into() },
        ]
    );
    assert_eq!(instantiate(&t[0], "Abs"), "Sub S()\n    Dim Abs\nEnd Sub\n");
}

#[test]
fn bad_template_files_are_rejected() {
    assert!(parse_templates("text first\n=== a bas\n").is_err());
    assert!(parse_templates("=== a txt\n").is_err());
    assert!(parse_templates("=== a\n").is_err());
    assert!(parse_templates("=== a bas\n=== a bas\n").is_err());
}

#[test]
fn words_skip_blank_and_comment_lines() {
    assert_eq!(parse_words("# keywords\r\nAbs\r\n\r\n  Len  \r\n#\r\n#If\r\n"), ["Abs", "Len", "#If"]);
}

#[test]
fn fates_round_trip_through_their_columns() {
    for fate in [
        Fate::Accepted,
        Fate::Semantic("Type mismatch".into()),
        Fate::Syntax("Syntax error".into()),
        Fate::Unanswered("exit -1".into()),
    ] {
        let column = fate.column();
        let (class, message) = column.split_once('\t').unwrap();
        assert_eq!(Fate::parse(class, message), Some(fate));
    }
    assert_eq!(Fate::parse("other", ""), None);
}

#[test]
fn only_syntax_rejections_are_rejections_of_the_form() {
    assert_eq!(Fate::Accepted.parses(), Some(true));
    assert_eq!(Fate::Semantic("m".into()).parses(), Some(true));
    assert_eq!(Fate::Syntax("m".into()).parses(), Some(false));
    assert_eq!(Fate::Unanswered("m".into()).parses(), None);
}
