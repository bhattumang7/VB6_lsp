//! Date literals: where the separators may stand, how two numbers read as a date, and the
//! AM/PM designators around the time. The yes/no answers are in `date_literal_forms`; these
//! tests check the values.

mod common;
use common::parse;
use vb6_syntax::frontend::ast::{AstLit, ExprNode};
use vb6_syntax::frontend::scanner::current_year;

fn date(text: &str) -> Option<f64> {
    let p = parse(&format!("Sub S()\n x = #{text}#\nEnd Sub\n"));
    if !p.found.is_empty() {
        return None;
    }
    let body = p.proc_body(0);
    let ExprNode::Assign { value, .. } = p.node(body[0]) else { return None };
    match p.node(*value) {
        ExprNode::Literal { lit: AstLit::Date(d) } => Some(*d),
        _ => None,
    }
}

fn close(a: Option<f64>, b: f64) {
    let a = a.expect("a date literal");
    assert!((a - b).abs() < 1e-9, "{a} != {b}");
}

const JAN_1_2000: f64 = 36526.0;
const THREE_OH_FOUR: f64 = 3.0 / 24.0 + 4.0 / 1440.0;

#[test]
fn a_separator_stands_between_two_parts_of_the_date() {
    for text in [
        "/1/2/2000", "1/2/2000/", "1/2/", "1/", "/1", "-1", "1-", "1,", ",1/2/2000", "1/,2/2000", "1,/2/2000", "1/-2/2000", "1//2/2000",
        "-1/2/2000", "1/2/-2000", "1/1/-1", "-5", "+5", "1/2/2000 -1:00", "1/2/2000 - 3:04", "1/2/2000, 3:04", "1/2/2000-3:04", "1/2/2000,3 PM",
        "2000-01-02,3:04", "Jan 5, 2000, 3:04", "Jan 5 2000, 3:04",
    ] {
        assert_eq!(date(text), None, "{text}");
    }
}

#[test]
fn any_of_the_three_separators_joins_the_parts_of_a_date() {
    for text in ["1/2/2000", "1-2-2000", "1,2,2000", "1/2-2000", "1,2/2000", "1 2 2000", "1 / 2 / 2000", "1 - 2 - 2000", "1 , 2 , 2000"] {
        assert_eq!(date(text), Some(JAN_1_2000 + 1.0), "{text}");
    }
}

#[test]
fn a_time_after_the_date_needs_no_separator_and_a_date_after_the_time_neither() {
    close(date("1/2/2000 3:04"), JAN_1_2000 + 1.0 + THREE_OH_FOUR);
    close(date("3:04 1/2/2000"), JAN_1_2000 + 1.0 + THREE_OH_FOUR);
    close(date("1/2/2000\t3:04"), JAN_1_2000 + 1.0 + THREE_OH_FOUR);
}

#[test]
fn a_month_and_a_year_written_as_two_numbers() {
    assert_eq!(date("12/2000"), Some(36861.0));
    assert_eq!(date("1/2000"), Some(JAN_1_2000));
    assert_eq!(date("2000/12"), Some(36861.0));
    assert_eq!(date("2000/1"), Some(JAN_1_2000));
    assert_eq!(date("12-2000"), Some(36861.0));
    assert_eq!(date("12,2000"), Some(36861.0));
    assert_eq!(date("12 2000"), Some(36861.0));
    assert_eq!(date("1/100"), date("1/1/100"));
}

#[test]
fn two_small_numbers_are_month_and_day_or_day_and_month_or_month_and_year() {
    let year = current_year();
    assert_eq!(date("12/13"), date(&format!("12/13/{year}")));
    // A first number that cannot be a month is the day.
    assert_eq!(date("13/12"), date(&format!("12/13/{year}")));
    assert_eq!(date("31/12"), date(&format!("12/31/{year}")));
    // A number that is no day of the month is a two-digit year.
    assert_eq!(date("1/32"), Some(48214.0));
    assert_eq!(date("32/1"), Some(48214.0));
    assert_eq!(date("12/99"), date("12/1/1999"));
    assert_eq!(date("0/1"), Some(JAN_1_2000));
    assert_eq!(date("1/0"), Some(JAN_1_2000));
    assert_eq!(date("2/30"), date("2/1/2030"));
}

#[test]
fn two_numbers_with_no_month_among_them_are_not_a_date() {
    for text in ["13/13", "13/31", "31/31", "32/32", "99/13", "0/0", "0/13", "100/2000", "0/2000", "2000/13", "1/10000", "10000/1", "1/99999"] {
        assert_eq!(date(text), None, "{text}");
    }
}

#[test]
fn a_number_after_a_month_name_that_is_no_day_of_it_is_the_year() {
    assert_eq!(date("Feb 30"), date("Feb 1 2030"));
    assert_eq!(date("Jan 32"), Some(48214.0));
    assert_eq!(date("Jan 100"), date("Jan 1 100"));
    assert_eq!(date("Jan 9999"), date("1/1/9999"));
    assert_eq!(date("Jan 10000"), None);
    assert_eq!(date("Jan 32 2000"), None);
    assert_eq!(date("Feb 30 2000"), None);
}

#[test]
fn a_designator_may_open_the_text_and_another_may_follow_the_time() {
    close(date("AM 3:04"), THREE_OH_FOUR);
    close(date("PM 3:04"), 15.0 / 24.0 + 4.0 / 1440.0);
    close(date("AM 3:04 AM 1/2/2000"), JAN_1_2000 + 1.0 + THREE_OH_FOUR);
    close(date("PM 3:04 1/2/2000"), JAN_1_2000 + 1.0 + 15.0 / 24.0 + 4.0 / 1440.0);
    // Any word that reads as a designator will do for the one that opens the text.
    close(date("AMX 3:04"), THREE_OH_FOUR);
    close(date("A1 3:04"), THREE_OH_FOUR);
    assert_eq!(date("AMX 1/2/2000"), Some(JAN_1_2000 + 1.0));
    assert_eq!(date("AX 1/2/2000"), None);
    assert_eq!(date("AM AM 3:04 1/2/2000"), None);
    assert_eq!(date("AM AM 1/2/2000"), None);
    assert_eq!(date("AM PM 1/2/2000"), None);
}

#[test]
fn one_more_designator_may_stand_between_the_time_and_a_numeric_date() {
    let expected = JAN_1_2000 + 1.0 + THREE_OH_FOUR;
    for text in [
        "3:04 AM AM 1/2/2000", "3:04 AM PM 1/2/2000", "3:04 AM A 1/2/2000", "3:04 AM P 1/2/2000", "3:04 AM AMX 1/2/2000", "3:04 AM A1 1/2/2000",
        "3:04 AM AM 1 2 2000", "3:04 AM AM 1-2-2000", "3:04 AM AM1 1/2/2000", "3:04 AMAM 1/2/2000", "3:04 AMPM 1/2/2000", "3:04 AMPMAM 1/2/2000",
    ] {
        close(date(text), expected);
    }
    close(date("3 AM AM 1/2/2000"), JAN_1_2000 + 1.0 + 3.0 / 24.0);
}

#[test]
fn a_designator_that_has_no_place_makes_the_text_invalid() {
    for text in [
        // The one that follows the time must be a designator by itself.
        "3:04 AMX 1/2/2000", "3:04 AMM 1/2/2000", "3:04 PMX 1/2/2000", "3:04 A1 1/2/2000", "3:04 AX 1/2/2000", "3:04 AP 1/2/2000",
        // Only one more is allowed, and only before a numeric date.
        "3 AM PM PM 1/2/2000", "3 AM AM AM 1/2/2000", "3:04 AM AM Jan 5 2000", "3:04 AM PM Jan", "3:04 AM AM", "3:04 AM AM 12", "3:04 AM AMX",
        // Not after the date, and not before a time that follows the date.
        "3:04 AM 1/2/2000 AM", "3:04 AM AM 1/2/2000 AM", "1/2/2000 3:04 AM AM", "1/2/2000 3 AM PM", "1/2/2000 AM 3:04", "1/2/2000 AMX 3:04",
        "3:04 AM AM 3:04", "3:04 AM AM 1/2/2000 4",
    ] {
        assert_eq!(date(text), None, "{text}");
    }
}
