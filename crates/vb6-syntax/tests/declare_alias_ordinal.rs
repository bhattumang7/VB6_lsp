//! A `Declare` alias that starts with `#` names an ordinal: the rest must be a number that
//! fits sixteen bits; anything else is "Invalid ordinal" at the alias string.

mod common;
use common::parse;

const INVALID_ORDINAL: u32 = 0x1c4;

fn found(alias: &str, sub: bool) -> (String, Vec<(u32, u32, u32)>) {
    let src = if sub {
        format!("Declare Sub f Lib \"kernel32\" Alias \"{alias}\" ()\n")
    } else {
        format!("Declare Function f Lib \"kernel32\" Alias \"{alias}\" () As Long\n")
    };
    let f = parse(&src).found;
    (src, f)
}

#[test]
fn numeric_ordinals_are_accepted() {
    for alias in ["#5", "#0", "#32767", "#&H10", "#&HFFFF", "#&H8000", "#&O17", "#5 ", "#&h1f", "name", "Name2", "f"] {
        assert_eq!(found(alias, false).1, vec![], "{alias}");
        assert_eq!(found(alias, true).1, vec![], "sub {alias}");
    }
}

#[test]
fn other_ordinals_are_invalid_at_the_alias_string() {
    for alias in ["#abc", "#12abc", "#70000", "#32768", "#-1", "#+5", "# 5", "#5.0", "#", "#&HFFFFFFFF", "#&H10000", "#&O", "#&H"] {
        let (src, f) = found(alias, false);
        let at = src.find("\"#").unwrap() as u32;
        assert_eq!(f, vec![(INVALID_ORDINAL, at, alias.len() as u32 + 2)], "{alias}");
        assert_eq!(found(alias, true).1.len(), 1, "sub {alias}");
    }
}
