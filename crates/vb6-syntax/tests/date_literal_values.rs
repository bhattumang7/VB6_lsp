//! Date and time literal values and validity.

mod common;
use common::parse;
use vb6_syntax::frontend::ast::{AstLit, ExprNode};

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

#[test]
fn numeric_dates_are_month_day_year() {
    assert_eq!(date("1/1/2000"), Some(36526.0));
    assert_eq!(date("2/29/2000"), Some(36585.0));
    assert_eq!(date("12/30/1899"), Some(0.0));
    assert_eq!(date("1/1/1900"), Some(2.0));
    assert_eq!(date("1-1-2000"), Some(36526.0));
}

#[test]
fn leading_year_means_year_month_day() {
    assert_eq!(date("2000/1/2"), Some(36527.0));
}

#[test]
fn two_digit_years_expand() {
    assert_eq!(date("1/1/29"), Some(47119.0));
    assert_eq!(date("1/1/30"), Some(47484.0));
    assert_eq!(date("1/1/49"), Some(54424.0));
    assert_eq!(date("1/1/50"), Some(18264.0));
    assert_eq!(date("1/1/050"), Some(18264.0));
}

#[test]
fn month_names_give_real_values() {
    assert_eq!(date("Jan 1, 2000"), Some(36526.0));
    assert_eq!(date("January 1 2000"), Some(36526.0));
    assert_eq!(date("1 January 2000"), Some(36526.0));
    assert_eq!(date("1-Jan-2000"), Some(36526.0));
    assert_eq!(date("December, 24, 2000"), Some(36884.0));
    assert_eq!(date("January 2000"), Some(36526.0));
}

#[test]
fn weekday_names_make_the_text_invalid() {
    assert_eq!(date("Saturday, January 1, 2000"), None);
    assert_eq!(date("Mon Jan 1 2000"), None);
}

#[test]
fn a_month_name_may_be_cut_short_to_three_letters_or_more() {
    assert_eq!(date("Sept 1 2000"), Some(36770.0));
    assert_eq!(date("Septe 1 2000"), Some(36770.0));
    assert_eq!(date("Janu 1 2000"), Some(36526.0));
    assert_eq!(date("Ja 1 2000"), None);
    assert_eq!(date("Jann 1 2000"), None);
}

#[test]
fn doubled_separators_are_not_a_date() {
    assert_eq!(date("1//1//2000"), None);
    assert_eq!(date("1--1--2000"), None);
    assert_eq!(date("1,,1,2000"), None);
}

#[test]
fn before_the_epoch_the_time_is_a_magnitude() {
    assert_eq!(date("1/1/1850 6:00 AM"), Some(-18260.25));
}

#[test]
fn a_first_field_that_cannot_be_a_month_is_the_day() {
    assert_eq!(date("13/1/2000"), Some(36538.0));
    assert_eq!(date("31/12/1999"), date("12/31/1999"));
}

#[test]
fn impossible_dates_are_not_literals() {
    for t in ["2/30/2000", "2/29/1900", "13/13/2000", "0/1/2000", "1/32/2000", "1/1/10000", "1/1/99999"] {
        assert_eq!(date(t), None, "{t}");
    }
}

#[test]
fn times_of_day() {
    close(date("5 PM"), 17.0 / 24.0);
    close(date("12:00 AM"), 0.0);
    close(date("12:00 PM"), 0.5);
    close(date("1:30"), 1.5 / 24.0);
    close(date("23:59:59"), 86399.0 / 86400.0);
    close(date("1/1/2000 5:30 PM"), 36526.0 + (17.5 / 24.0));
}

#[test]
fn impossible_times_are_not_literals() {
    for t in ["25:00", "12:60", "12:00:60", "12:30 A.M.", "12:30 am pm", "1/1/2000 PM 1:00", "1/1/2000 AM", "2000.1.1", "1.1.2000"] {
        assert_eq!(date(t), None, "{t}");
    }
}

#[test]
fn a_meridiem_may_be_a_letter_and_follows_any_hour() {
    close(date("12:30p"), 12.5 / 24.0);
    close(date("12:30 a"), 0.5 / 24.0);
    close(date("12:30 p"), 12.5 / 24.0);
    close(date("13:30 PM"), 13.5 / 24.0);
    close(date("23:59 PM"), 23.0 / 24.0 + 59.0 / 1440.0);
    close(date("0:00 AM"), 0.0);
    close(date("13 pm"), 13.0 / 24.0);
    close(date("0 am"), 0.0);
    assert_eq!(date("AM 1/1/2000"), Some(36526.0));
}

#[test]
fn a_period_between_numbers_separates_time_parts() {
    close(date("1.2"), 62.0 / 1440.0);
    close(date("12:30.5"), 12.5 / 24.0 + 5.0 / 86400.0);
    assert_eq!(date("Jan. 1, 2000"), Some(36526.0));
}
