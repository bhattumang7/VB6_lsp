//! A `Declare` alias that starts with `#` names an ordinal: the number scanner must find a
//! whole number at its start, and the whole text must convert to an Integer. Every form in
//! the fixture is replayed against the compiler's answer, and the three kinds of report are
//! checked exactly.

mod common;
use common::answers::{assert_none_wrong, disagreements};
use common::parse;
use vb6_syntax::frontend::diagnostics::{ERR_INVALID_ORDINAL, ERR_OVERFLOW, ERR_SYNTAX_ERROR};

const TEMPLATES: &str = include_str!("fixtures/declare_ordinal_forms/templates.txt");
const ANSWERS: &str = include_str!("fixtures/declare_ordinal_forms/oracle.tsv");

#[test]
fn ordinal_forms_match_the_compiler() {
    let (checked, wrong) = disagreements(TEMPLATES, ANSWERS, |_, _, _, _| false);
    assert!(checked > 600, "fixture is unexpectedly small: {checked}");
    assert_none_wrong(checked, &wrong);
}

/// The diagnostics for `Declare Sub Foo Lib "k" Alias "<alias>" ()` as `(code, offset of the
/// alias string, its length)`.
fn reported(alias: &str) -> Vec<(u32, u32, u32)> {
    let src = format!("Declare Sub Foo Lib \"k\" Alias \"{alias}\" ()\n");
    let at = src.find("Alias").unwrap() as u32 + 6;
    parse(&src).found.into_iter().map(|(c, s, l)| (c, s - at, l)).collect()
}

fn len(alias: &str) -> u32 {
    alias.chars().count() as u32 + 2
}

#[test]
fn text_that_converts_is_accepted() {
    for alias in [
        "#1,000", "#5+", "#5-", "#5 +", "#1,", "#1,,2", "#32,767", "#32768-", "#&17", "#&177777", "#&H8000", "#&HFFFF", "#1,5.",
        "#1,5.0", "#0,0", "#32767-", "#0005",
    ] {
        assert_eq!(reported(alias), vec![], "{alias}");
    }
}

#[test]
fn text_that_does_not_convert_is_an_invalid_ordinal_at_the_alias_string() {
    for alias in ["#32,768", "#99,999", "#32768", "#32768+", "#1 ,", "#1+1", "#1,x", "#5.7", "#5.", "#.5", "#,5", "#&8", "#&H10000", "#5$", "#5%"] {
        assert_eq!(reported(alias), vec![(ERR_INVALID_ORDINAL, 0, len(alias))], "{alias}");
    }
}

#[test]
fn an_exponent_mark_and_sign_without_digits_is_an_overflow() {
    for alias in ["#1E+", "#1E-", "#1.5E+"] {
        assert_eq!(reported(alias), vec![(ERR_OVERFLOW, 0, len(alias))], "{alias}");
    }
}

#[test]
fn an_exponent_mark_with_nothing_after_it_is_a_syntax_error() {
    for alias in ["#12e", "#1E", "#1.E", "#1.5E"] {
        assert_eq!(reported(alias), vec![(ERR_SYNTAX_ERROR, 0, len(alias))], "{alias}");
    }
}

#[test]
fn a_non_breaking_space_after_the_number_is_a_blank_and_other_high_characters_are_not() {
    // Bytes above 0x7F are read as the Windows ANSI characters they stand for.
    let src_with = |tail: u8| {
        let mut v = b"Declare Sub Foo Lib \"k\" Alias \"#5".to_vec();
        v.push(tail);
        v.extend_from_slice(b"\" ()\n");
        v
    };
    let codes = |src: Vec<u8>| {
        use vb6_syntax::frontend::ast::ExprArena;
        use vb6_syntax::frontend::parser::Parser;
        use vb6_syntax::frontend::scanner::ScannerContext;
        let mut ctx = ScannerContext::new(1, 1, 0x0409);
        ctx.intern_keywords();
        let mut arena = ExprArena::new();
        let mut parser = Parser::new(&mut ctx, &src);
        parser.parse_module(&mut arena);
        parser.diagnostics.items().iter().map(|d| d.code).collect::<Vec<u32>>()
    };
    assert_eq!(codes(src_with(0xA0)), Vec::<u32>::new());
    assert_eq!(codes(src_with(0xB2)), vec![ERR_INVALID_ORDINAL]);
    assert_eq!(codes(src_with(0xBD)), vec![ERR_INVALID_ORDINAL]);
}

#[test]
fn an_alias_that_does_not_start_with_a_number_sign_is_not_an_ordinal() {
    for alias in ["Foo", "x#5", "", "#"] {
        let expected = if alias == "#" { vec![ERR_INVALID_ORDINAL] } else { vec![] };
        assert_eq!(reported(alias).into_iter().map(|r| r.0).collect::<Vec<u32>>(), expected, "{alias}");
    }
}
