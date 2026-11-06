//! The text of a project's conditional-compilation arguments (`DEBUG = 1 : Level = -2`):
//! which texts are accepted, what values they define, and which are refused.

use vb6_syntax::frontend::cond_compile::{parse_cond_comp_args, CcValue, ConstantDeclarationsError};

const INVALID: Result<Vec<(String, CcValue)>, ConstantDeclarationsError> = Err(ConstantDeclarationsError::InvalidSyntax);

fn ints(pairs: &[(&str, i16)]) -> Result<Vec<(String, CcValue)>, ConstantDeclarationsError> {
    Ok(pairs.iter().map(|(n, v)| (n.to_string(), CcValue::Integer(*v))).collect())
}

fn accepted(text: &str, pairs: &[(&str, i16)]) {
    assert_eq!(parse_cond_comp_args(text), ints(pairs), "{text:?}");
}

fn refused(text: &str) {
    assert_eq!(parse_cond_comp_args(text), INVALID, "{text:?}");
}

#[test]
fn a_sign_alone_is_the_value_zero() {
    accepted("Foo=-", &[("foo", 0)]);
    accepted("Foo=+", &[("foo", 0)]);
    accepted("Foo = -", &[("foo", 0)]);
    accepted("Foo=-:Bar=+", &[("foo", 0), ("bar", 0)]);
    accepted("Foo=- : Bar=5", &[("foo", 0), ("bar", 5)]);
}

#[test]
fn blanks_between_a_sign_and_its_digits_are_skipped() {
    accepted("Foo = - 1", &[("foo", -1)]);
    accepted("Foo=- 1", &[("foo", -1)]);
    accepted("Foo=+ 7", &[("foo", 7)]);
    accepted("Foo=-\t\u{b}\u{c}  12", &[("foo", -12)]);
}

#[test]
fn a_sign_followed_only_by_blanks_to_the_end_is_refused() {
    refused("Foo=-  ");
    refused("Foo=+\t");
    refused("Foo= ");
    refused("Foo=");
}

#[test]
fn only_one_sign_is_read() {
    for bad in ["Foo=--1", "Foo=+-1", "Foo=-+1", "Foo=- -1", "Foo=++1", "Foo=-x", "Foo=1-"] {
        refused(bad);
    }
}

#[test]
fn an_unsigned_value_has_at_most_five_digits() {
    accepted("Foo=12345", &[("foo", 12345)]);
    accepted("Foo=00001", &[("foo", 1)]);
    accepted("Foo=0", &[("foo", 0)]);
    refused("Foo=000001");
    refused("Foo=0000000000");
    refused("Foo=1234567");
    refused("Foo=99999");
}

#[test]
fn once_a_sign_has_been_read_later_unsigned_values_may_have_more_digits() {
    accepted("A=-1:Foo=000001", &[("a", -1), ("foo", 1)]);
    accepted("A=-1:Foo=0000001", &[("a", -1), ("foo", 1)]);
    accepted("A=-1:Foo=00000000001", &[("a", -1), ("foo", 1)]);
    accepted("A=-1:Foo=0000000005", &[("a", -1), ("foo", 5)]);
    accepted("A=+1:Foo=000001", &[("a", 1), ("foo", 1)]);
    accepted("Foo=-:Bar=000001", &[("foo", 0), ("bar", 1)]);
    accepted("Foo=+:Bar=000001", &[("foo", 0), ("bar", 1)]);
    accepted("Foo=-1:Bar=1:Baz=000001", &[("foo", -1), ("bar", 1), ("baz", 1)]);
    accepted("Foo=1:Bar=-000001:Baz=000001", &[("foo", 1), ("bar", -1), ("baz", 1)]);
    // Without an earlier sign the same text is refused.
    refused("Foo=1:Bar=000001");
    refused("Foo=1:Bar=2:Baz=000001");
}

#[test]
fn a_signed_value_may_have_up_to_eleven_digits() {
    accepted("Foo=-0000000001", &[("foo", -1)]);
    accepted("Foo=-00000000001", &[("foo", -1)]);
    accepted("Foo=+00000000007", &[("foo", 7)]);
    refused("Foo=-000000000001");
    refused("Foo=+000000000001");
    refused("Foo=-0000000000000000001");
}

#[test]
fn a_run_of_twelve_digits_is_refused_after_an_earlier_sign_too() {
    refused("A=-1:Foo=000000000001");
    refused("A=-1:Foo=1234567890");
    refused("A=-1:Foo=0000065541");
    refused("A=-1:Foo=0000099999");
}

#[test]
fn long_digit_runs_wrap_in_thirty_two_bits_before_the_range_check() {
    accepted("A=-1:Foo=4294967301", &[("a", -1), ("foo", 5)]);
    accepted("A=-1:Foo=4294967296", &[("a", -1), ("foo", 0)]);
    accepted("A=-1:Foo=4294934528", &[("a", -1), ("foo", -32768)]);
    accepted("Foo=-4294967301", &[("foo", -5)]);
    refused("A=-1:Foo=2147483648");
    refused("A=-1:Foo=4294934527");
}

#[test]
fn values_are_sixteen_bit_integers() {
    accepted("A=32767:B=-32768:C=+32767", &[("a", 32767), ("b", -32768), ("c", 32767)]);
    refused("A=32768");
    refused("A=-32769");
    refused("A=65535");
}

#[test]
fn a_name_cannot_start_with_an_underscore() {
    refused("_Foo=1");
    refused("_=1");
    refused("__a=1");
    refused("A=1:_B=2");
    accepted("F_oo=1", &[("f_oo", 1)]);
    accepted("Foo_=1", &[("foo_", 1)]);
    accepted("a_1_=1", &[("a_1_", 1)]);
}

#[test]
fn a_name_starts_with_a_letter_and_goes_on_with_letters_digits_and_underscores() {
    for bad in ["1A=1", "A$=1", "A%=1", "A.B=1", "[A]=1", "A-B=1", "Fo o=1", "-5=1", "A!=1", "A#=1"] {
        refused(bad);
    }
    accepted("A1=1", &[("a1", 1)]);
}

#[test]
fn names_may_hold_the_characters_from_a_with_ring_to_y_with_diaeresis() {
    // The multiplication sign, the feminine ordinal indicator, a superscript two and the
    // no-break space are all name characters.
    accepted("\u{d7}a=1", &[("\u{d7}a", 1)]);
    accepted("a\u{d7}=1", &[("a\u{d7}", 1)]);
    accepted("\u{aa}a=1", &[("\u{aa}a", 1)]);
    accepted("a\u{b2}=1", &[("a\u{b2}", 1)]);
    accepted("\u{a0}Foo=1", &[("\u{a0}foo", 1)]);
    accepted("Foo\u{a0}=1", &[("foo\u{a0}", 1)]);
    accepted("\u{c9}a=1", &[("\u{e9}a", 1)]);
    accepted("\u{81}a=1", &[("\u{81}a", 1)]);
    accepted("a\u{ff}=1", &[("a\u{ff}", 1)]);
}

#[test]
fn characters_outside_the_latin_1_range_are_not_name_characters() {
    for bad in ["\u{20ac}a=1", "a\u{20ac}=1", "\u{160}a=1", "a\u{152}=1", "\u{178}=1", "a\u{2013}=1", "\u{7f}a=1", "a\u{1}=1", "a\u{1f600}=1"] {
        refused(bad);
    }
}

#[test]
fn blanks_are_the_tab_line_end_vertical_tab_form_feed_and_space() {
    accepted(" \t\u{b}\u{c}Foo \t=\t\u{b}1 \u{c}\t: Bar=2 ", &[("foo", 1), ("bar", 2)]);
    accepted("Foo=1\r\n", &[("foo", 1)]);
    accepted("\r\nFoo=1", &[("foo", 1)]);
}

#[test]
fn other_white_space_is_not_a_blank() {
    // The no-break space after a value, and wider spaces, are not skipped.
    refused("Foo=1\u{a0}");
    refused("Foo=\u{a0}1");
    refused("Foo=1\u{2003}");
    refused("Foo=\u{3000}1");
    refused("\u{2003}Foo=1");
}

#[test]
fn a_name_given_twice_ignoring_case_is_refused() {
    refused("Foo=1:Foo=2");
    refused("foo=1:FOO=2");
    refused("Foo=1:Bar=2:fOO=3");
    refused("Dim=1:dim=2");
}

#[test]
fn names_that_differ_only_in_the_case_of_latin_1_letters_are_the_same_name() {
    refused("\u{c9}a=1:\u{e9}a=2");
    refused("\u{e9}a=1:\u{c9}A=2");
    refused("\u{d0}a=1:\u{f0}a=2");
    refused("\u{d8}=1:\u{f8}=2");
    refused("\u{dd}=1:\u{fd}=2");
}

#[test]
fn some_letters_compare_equal_to_two_letters() {
    refused("\u{df}a=1:SSa=2");
    refused("ssa=1:\u{df}A=2");
    refused("\u{c6}x=1:aEx=2");
    refused("\u{e6}x=1:AEX=2");
    refused("\u{fe}a=1:tha=2");
    refused("\u{de}a=1:THa=2");
    refused("a\u{df}b=1:assb=2");
}

#[test]
fn letters_that_only_look_alike_are_different_names() {
    accepted("\u{aa}a=1:aa=2", &[("\u{aa}a", 1), ("aa", 2)]);
    accepted("\u{b5}a=1:ua=2", &[("\u{b5}a", 1), ("ua", 2)]);
    accepted("\u{df}a=1:sa=2", &[("\u{df}a", 1), ("sa", 2)]);
    accepted("\u{c9}a=1:ea=2", &[("\u{e9}a", 1), ("ea", 2)]);
    accepted("\u{d7}a=1:\u{f7}a=2", &[("\u{d7}a", 1), ("\u{f7}a", 2)]);
}

#[test]
fn text_of_1024_or_more_characters_is_too_long() {
    let name = "A".repeat(1021);
    // 1023 characters.
    accepted(&format!("{name}=1"), &[(&name.to_ascii_lowercase(), 1)]);
    // 1024 characters.
    let long = format!("{name}=1 ");
    assert_eq!(long.chars().count(), 1024);
    assert_eq!(parse_cond_comp_args(&long), Err(ConstantDeclarationsError::DataTooLarge));
    assert_eq!(parse_cond_comp_args(&format!("{name}==1")), Err(ConstantDeclarationsError::DataTooLarge));
    assert_eq!(parse_cond_comp_args(&" ".repeat(1024)), Err(ConstantDeclarationsError::DataTooLarge));
    assert_eq!(parse_cond_comp_args(&"x".repeat(5000)), Err(ConstantDeclarationsError::DataTooLarge));
    // A text that is wrong and short is wrong in the other way.
    assert_eq!(parse_cond_comp_args(&"x".repeat(1023)), INVALID);
    accepted(&" ".repeat(1023), &[]);
}

#[test]
fn the_length_is_counted_in_sixteen_bit_units() {
    // 512 characters outside the basic plane are 1024 units; 511 are 1022.
    assert_eq!(parse_cond_comp_args(&"\u{1f600}".repeat(512)), Err(ConstantDeclarationsError::DataTooLarge));
    assert_eq!(parse_cond_comp_args(&"\u{1f600}".repeat(511)), INVALID);
}

#[test]
fn names_may_be_long() {
    let name = "A".repeat(400);
    accepted(&format!("{name}=1"), &[(&name.to_ascii_lowercase(), 1)]);
}

#[test]
fn empty_and_blank_text_define_nothing() {
    accepted("", &[]);
    accepted("   ", &[]);
    accepted("\t\u{b}\u{c}\r\n", &[]);
}

#[test]
fn separators_and_empty_parts() {
    for bad in ["A=1:", "A=1::B=2", ":A=1", "=1", "A=1: ", "A=1:\t", "A", "A=1;B=2", "A=1,B=2", "A=1:B", "A:B=1", "A=1=2"] {
        refused(bad);
    }
    accepted("A=1 : B=2", &[("a", 1), ("b", 2)]);
    accepted("  A = 1  :  B = 2  ", &[("a", 1), ("b", 2)]);
}

#[test]
fn a_value_is_digits_only() {
    for bad in ["A=1.5", "A=&H10", "A=True", "A=x", "A=1+2", "A=B", "A=\"x\"", "A=1_", "A=1x", "A=1 2", "A=\u{660}"] {
        refused(bad);
    }
}
