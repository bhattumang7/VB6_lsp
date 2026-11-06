//! The values constant expressions work out: conversions to a declared type, the operators, the
//! built-in functions, numbers and dates written as text.

use vb6_syntax::frontend::const_value::{
    call_function, format_date, number_in_text, parse_date, BinaryOp, CompareMode, ConstValue as V, EvalError as E,
    NumberInText, Operand, TargetType as T,
};

fn plain(v: V) -> Operand {
    Operand::plain(v)
}

fn variant(v: V) -> Operand {
    Operand { value: v, variant: true }
}

fn bin(op: BinaryOp, l: V, r: V) -> Result<V, E> {
    Operand::binary(op, &plain(l), &plain(r), CompareMode::Binary).map(|o| o.value)
}

fn s(text: &str) -> V {
    V::Str(text.to_string())
}

#[test]
fn numbers_are_read_from_text_as_conversion_reads_them() {
    let value = |t: &str| number_in_text(t).map(|n| n.value);
    for (text, want) in [
        ("7", 7.0),
        (" 7 ", 7.0),
        ("\t7\r\n", 7.0),
        ("1.5", 1.5),
        (".5", 0.5),
        ("5.", 5.0),
        ("+5", 5.0),
        ("-5", -5.0),
        ("- 5", -5.0),
        ("5-", -5.0),
        ("5 -", -5.0),
        ("5+", 5.0),
        ("(5)", -5.0),
        ("(1,000)", -1000.0),
        ("1,000", 1000.0),
        ("1,,0", 10.0),
        ("1,", 1.0),
        ("1,000.5", 1000.5),
        ("5,.5", 5.5),
        ("1e3", 1000.0),
        ("1D2", 100.0),
        ("1E+2", 100.0),
        ("1E-2", 0.01),
        ("1e-400", 0.0),
        ("&H10", 16.0),
        ("&h1f", 31.0),
        ("&O17", 15.0),
        ("&HFFFF", -1.0),
        ("&H8000", -32768.0),
        ("&H7FFF", 32767.0),
        ("&H10000", 65536.0),
        ("&HFFFFFF", 16_777_215.0),
        ("&HFFFFFFFF", -1.0),
        ("&H80000000", -2_147_483_648.0),
        ("&H7FFFFFFF", 2_147_483_647.0),
        ("&H100000000", 4_294_967_296.0),
        ("&O177777", -1.0),
        ("&O100000", -32768.0),
        ("&O37777777777", -1.0),
        ("&O20000000000", -2_147_483_648.0),
    ] {
        assert_eq!(value(text), Some(want), "{text:?}");
    }
    for text in [
        "", " ", ".", "-", ",5", "(5", "5)", "-(5)", "--5", "5--", "(5-)", "+-5", "(-5)", "(5)-", "+5-", "$5", "$5-", "1.5.2", "1e", "1e+", "e5",
        "1e5.5", "1 e5", "1e 5", ". 5", "1d", "1 000", "1_0", "0x10", "&", "&H", "&HG", "&O8", "&H10&", "&H10%", "&H 10", "& H10", "&H-1", "-&H1",
        "+&H1", "&H1.5", "(&H10)", "&H10-", "True", "False", "Infinity", "NaN", "1.#INF", "5%", "5&", "5!", "5#", "5@", "5$", "%5",
    ] {
        assert_eq!(value(text), None, "{text:?}");
    }
    assert_eq!(value("1e400"), Some(f64::INFINITY));
}

#[test]
fn a_hexadecimal_number_wider_than_32_bits_keeps_its_low_bits_for_integer_and_byte() {
    assert_eq!(number_in_text("&H100000000"), Some(NumberInText { value: 4_294_967_296.0, wrapped: Some(0) }));
    assert_eq!(number_in_text("&HFFFFFFFFF"), Some(NumberInText { value: 68_719_476_735.0, wrapped: Some(-1) }));
    assert_eq!(number_in_text("&O40000000000"), Some(NumberInText { value: 4_294_967_296.0, wrapped: Some(0) }));
    assert_eq!(number_in_text("&HFFFFFFFF"), Some(NumberInText { value: -1.0, wrapped: None }));
    assert_eq!(s("&H100000000").convert(T::Integer), Ok(V::Integer(0)));
    assert_eq!(s("&H100000000").convert(T::Byte), Ok(V::Byte(0)));
    assert_eq!(s("&HFFFFFFFFF").convert(T::Integer), Ok(V::Integer(-1)));
    assert_eq!(s("&HFFFFFFFFF").convert(T::Byte), Err(E::Overflow));
    assert_eq!(s("&H100000000").convert(T::Long), Err(E::Overflow));
    assert_eq!(s("&HFFFFFFFF").convert(T::Long), Ok(V::Long(-1)));
    assert_eq!(s("&HFFFFFFFF").convert(T::Integer), Ok(V::Integer(-1)));
    assert_eq!(s("&H80000000").convert(T::Long), Ok(V::Long(i32::MIN)));
    assert_eq!(s("&H80000000").convert(T::Integer), Err(E::Overflow));
}

#[test]
fn conversion_to_a_declared_type_rounds_halves_to_even_and_checks_the_range() {
    assert_eq!(V::Double(2.5).convert(T::Integer), Ok(V::Integer(2)));
    assert_eq!(V::Double(3.5).convert(T::Integer), Ok(V::Integer(4)));
    assert_eq!(V::Double(-0.5).convert(T::Integer), Ok(V::Integer(0)));
    assert_eq!(V::Double(254.5).convert(T::Byte), Ok(V::Byte(254)));
    assert_eq!(V::Double(255.5).convert(T::Byte), Err(E::Overflow));
    assert_eq!(V::Double(-0.4).convert(T::Byte), Ok(V::Byte(0)));
    assert_eq!(V::Long(256).convert(T::Byte), Err(E::Overflow));
    assert_eq!(V::Long(-1).convert(T::Byte), Err(E::Overflow));
    assert_eq!(V::Long(32768).convert(T::Integer), Err(E::Overflow));
    assert_eq!(V::Long(-32769).convert(T::Integer), Err(E::Overflow));
    assert_eq!(V::Double(2_147_483_648.0).convert(T::Long), Err(E::Overflow));
    assert_eq!(V::Double(-2_147_483_648.0).convert(T::Long), Ok(V::Long(i32::MIN)));
    assert_eq!(V::Double(1e40).convert(T::Single), Err(E::Overflow));
    assert_eq!(V::Double(3e38).convert(T::Single), Ok(V::Single(3e38_f32)));
    assert_eq!(V::Double(1e20).convert(T::Currency), Err(E::Overflow));
    assert_eq!(V::Double(1.5).convert(T::Currency), Ok(V::Currency(15_000)));
    assert_eq!(V::Double(f64::INFINITY).convert(T::Double), Err(E::Overflow));
    assert_eq!(V::Double(f64::INFINITY).convert(T::Boolean), Ok(V::Boolean(true)));
    assert_eq!(V::Double(f64::INFINITY).convert(T::String), Err(E::Overflow));
    assert_eq!(V::Double(f64::INFINITY).convert(T::Variant), Err(E::Overflow));
    assert_eq!(V::Boolean(true).convert(T::Byte), Ok(V::Byte(255)));
    assert_eq!(V::Boolean(true).convert(T::Integer), Ok(V::Integer(-1)));
    assert_eq!(V::Boolean(false).convert(T::Byte), Ok(V::Byte(0)));
    assert_eq!(V::Empty.convert(T::Long), Ok(V::Long(0)));
    assert_eq!(V::Empty.convert(T::String), Ok(s("")));
    assert_eq!(V::Empty.convert(T::Boolean), Ok(V::Boolean(false)));
    assert_eq!(V::Integer(5).convert(T::Boolean), Ok(V::Boolean(true)));
    assert_eq!(V::Integer(0).convert(T::Boolean), Ok(V::Boolean(false)));
    assert_eq!(V::Double(1.5).convert(T::String), Ok(s("1.5")));
    assert_eq!(V::Long(7).convert(T::String), Ok(s("7")));
    assert_eq!(V::Boolean(true).convert(T::String), Ok(s("True")));
    assert_eq!(V::Currency(15_000).convert(T::String), Ok(s("1.5")));
    assert_eq!(V::Long(7).convert(T::Variant), Ok(V::Long(7)));
}

#[test]
fn null_and_nothing_convert_to_no_declared_type() {
    for t in [T::Boolean, T::Byte, T::Integer, T::Long, T::Single, T::Double, T::Currency, T::Date, T::String] {
        assert_eq!(V::Null.convert(t), Err(E::InvalidUseOfNull), "{t:?}");
    }
    for t in [T::Boolean, T::Byte, T::Integer, T::Long, T::Single, T::Double, T::Currency, T::Date, T::String] {
        assert_eq!(V::Nothing.convert(t), Err(E::InvalidUseOfObject), "{t:?}");
    }
    assert_eq!(V::Null.convert(T::Variant), Ok(V::Null));
    assert_eq!(V::Nothing.convert(T::Variant), Err(E::ObjectNotSet));
}

#[test]
fn text_converts_to_a_number_or_is_a_type_mismatch() {
    assert_eq!(s("7").convert(T::Integer), Ok(V::Integer(7)));
    assert_eq!(s("1.5").convert(T::Integer), Ok(V::Integer(2)));
    assert_eq!(s("1,000").convert(T::Integer), Ok(V::Integer(1000)));
    assert_eq!(s("1,000").convert(T::Byte), Err(E::Overflow));
    assert_eq!(s("a").convert(T::Integer), Err(E::TypeMismatch));
    assert_eq!(s("").convert(T::Integer), Err(E::TypeMismatch));
    assert_eq!(s("x").convert(T::Boolean), Err(E::TypeMismatch));
    assert_eq!(s("True").convert(T::Boolean), Ok(V::Boolean(true)));
    assert_eq!(s("false").convert(T::Boolean), Ok(V::Boolean(false)));
    assert_eq!(s("5").convert(T::Boolean), Ok(V::Boolean(true)));
    assert_eq!(s("0").convert(T::Boolean), Ok(V::Boolean(false)));
    assert_eq!(s("True").convert(T::Integer), Err(E::TypeMismatch));
    assert_eq!(s("1e400").convert(T::Double), Err(E::Overflow));
    assert_eq!(s("1.5").convert(T::Currency), Ok(V::Currency(15_000)));
}

#[test]
fn dates_are_read_from_text_and_shown_as_text() {
    let jan_1_2000 = 36_526.0;
    for (text, want) in [
        ("1/1/2000", jan_1_2000),
        ("1-1-2000", jan_1_2000),
        ("2000-01-01", jan_1_2000),
        ("January 1, 2000", jan_1_2000),
        ("1 Jan 2000", jan_1_2000),
        ("1-Jan-2000", jan_1_2000),
        ("13/1/2000", 36_538.0),
        ("1/1/00", jan_1_2000),
        ("1/1/30", 10_959.0),
        ("12/31/9999", 2_958_465.0),
        ("12:30", 12.5 / 24.0),
        ("12:30 PM", 12.5 / 24.0),
        ("12:30 AM", 0.5 / 24.0),
        ("1:00 PM", 13.0 / 24.0),
        ("1/1/2000 12:30", jan_1_2000 + 12.5 / 24.0),
    ] {
        let got = parse_date(text).unwrap_or_else(|| panic!("{text:?}"));
        assert!((got - want).abs() < 1e-9, "{text:?}: {got} != {want}");
    }
    for text in ["", "x", "13/13/2000", "2/30/2000", "1/1/2000/1", "25:00", "12:60", "1/1/10000", "1 2 3 4"] {
        assert_eq!(parse_date(text), None, "{text:?}");
    }
    assert_eq!(parse_date("2/29/2000"), Some(36_585.0));
    assert_eq!(parse_date("2/29/1900"), None);
    assert_eq!(format_date(36_526.0), "1/1/2000");
    assert_eq!(format_date(36_526.5), "1/1/2000 12:00:00 PM");
    assert_eq!(format_date(0.5), "12:00:00 PM");
    assert_eq!(format_date(0.0), "12/30/1899");
    assert_eq!(s("1/1/2000").convert(T::Date), Ok(V::Date(36_526.0)));
    assert_eq!(s("5").convert(T::Date), Ok(V::Date(5.0)));
    assert_eq!(s("x").convert(T::Date), Err(E::TypeMismatch));
    assert_eq!(V::Double(3_000_000.0).convert(T::Date), Err(E::Overflow));
    assert_eq!(V::Date(36_526.0).convert(T::String), Ok(s("1/1/2000")));
}

#[test]
fn arithmetic_has_the_wider_operand_type_and_overflows_instead_of_widening() {
    use BinaryOp::*;
    assert_eq!(bin(Add, V::Integer(1), V::Integer(2)), Ok(V::Integer(3)));
    assert_eq!(bin(Add, V::Integer(32767), V::Integer(1)), Err(E::Overflow));
    assert_eq!(bin(Add, V::Integer(32767), V::Long(1)), Ok(V::Long(32768)));
    assert_eq!(bin(Add, V::Long(i32::MAX), V::Long(1)), Err(E::Overflow));
    assert_eq!(bin(Add, V::Long(i32::MAX), V::Integer(1)), Err(E::Overflow));
    assert_eq!(bin(Mul, V::Integer(200), V::Integer(200)), Err(E::Overflow));
    assert_eq!(bin(Add, V::Byte(200), V::Byte(100)), Err(E::Overflow));
    assert_eq!(bin(Add, V::Byte(200), V::Integer(100)), Ok(V::Integer(300)));
    assert_eq!(bin(Sub, V::Byte(1), V::Byte(2)), Err(E::Overflow));
    assert_eq!(bin(Add, V::Boolean(true), V::Boolean(true)), Ok(V::Integer(-2)));
    assert_eq!(bin(Add, V::Single(1.5), V::Integer(1)), Ok(V::Single(2.5)));
    assert_eq!(bin(Mul, V::Single(3e38), V::Single(3e38)), Err(E::Overflow));
    assert_eq!(bin(Add, V::Double(1.5), V::Integer(1)), Ok(V::Double(2.5)));
    // A Long with a Single is a Double; an Integer with a Single is a Single.
    assert_eq!(bin(Mul, V::Long(2_000_000_000), V::Single(3e38)), Ok(V::Double(2_000_000_000.0 * f64::from(3e38_f32))));
    assert_eq!(bin(Add, V::Long(1), V::Single(0.5)), Ok(V::Double(1.5)));
    assert_eq!(bin(Mul, V::Integer(2), V::Single(3e38)), Err(E::Overflow));
    assert_eq!(bin(Add, V::Currency(5_000), V::Integer(1)), Ok(V::Currency(15_000)));
    assert_eq!(bin(Mul, V::Currency(5_000), V::Currency(5_000)), Ok(V::Currency(2_500)));
    assert_eq!(bin(Add, s("1"), s("2")), Ok(s("12")));
    assert_eq!(bin(Add, s("1"), V::Integer(2)), Ok(V::Double(3.0)));
    assert_eq!(bin(Add, s("a"), V::Integer(2)), Err(E::TypeMismatch));
    assert_eq!(bin(Add, V::Empty, V::Integer(2)), Ok(V::Integer(2)));
    assert_eq!(bin(Cat, V::Integer(1), V::Integer(2)), Ok(s("12")));
    assert_eq!(bin(Cat, s("a"), V::Null), Ok(s("a")));
    assert_eq!(bin(Cat, V::Null, V::Null), Ok(V::Null));
    assert_eq!(bin(Add, V::Null, V::Integer(2)), Ok(V::Null));
    assert_eq!(bin(Add, V::Null, s("a")), Err(E::TypeMismatch));
}

#[test]
fn a_value_held_in_a_variant_widens_instead_of_overflowing() {
    let add = |l: Operand, r: Operand| Operand::binary(BinaryOp::Add, &l, &r, CompareMode::Binary);
    assert_eq!(add(variant(V::Integer(32767)), plain(V::Integer(1))), Ok(variant(V::Long(32768))));
    assert_eq!(add(variant(V::Long(i32::MAX)), plain(V::Integer(1))), Ok(variant(V::Double(2_147_483_648.0))));
    assert_eq!(add(plain(V::Integer(32767)), variant(V::Integer(1))), Ok(variant(V::Long(32768))));
    assert_eq!(add(variant(V::Integer(1)), plain(V::Integer(1))), Ok(variant(V::Integer(2))));
}

#[test]
fn division_and_powers() {
    use BinaryOp::*;
    assert_eq!(bin(Div, V::Integer(1), V::Integer(2)), Ok(V::Double(0.5)));
    assert_eq!(bin(Div, V::Single(1.0), V::Integer(2)), Ok(V::Single(0.5)));
    assert_eq!(bin(Div, V::Integer(1), V::Integer(0)), Err(E::DivisionByZero));
    assert_eq!(bin(Div, V::Integer(0), V::Integer(0)), Err(E::Overflow));
    assert_eq!(bin(Div, V::Double(1e300), V::Double(1e-10)), Err(E::Overflow));
    assert_eq!(bin(IDiv, V::Integer(7), V::Integer(2)), Ok(V::Integer(3)));
    assert_eq!(bin(IDiv, V::Integer(1), V::Integer(0)), Err(E::DivisionByZero));
    assert_eq!(bin(IDiv, V::Integer(5), V::Double(0.4)), Err(E::DivisionByZero));
    assert_eq!(bin(IDiv, V::Integer(-32768), V::Integer(-1)), Ok(V::Long(32768)));
    assert_eq!(bin(Mod, V::Integer(7), V::Integer(2)), Ok(V::Integer(1)));
    assert_eq!(bin(Mod, V::Integer(5), V::Integer(0)), Err(E::DivisionByZero));
    assert_eq!(bin(Mod, V::Long(i32::MIN), V::Long(-1)), Err(E::Overflow));
    assert_eq!(bin(Pow, V::Integer(2), V::Integer(2)), Ok(V::Double(4.0)));
    assert_eq!(bin(Pow, V::Integer(2), V::Integer(1024)), Err(E::Overflow));
    assert_eq!(bin(Pow, V::Integer(0), V::Integer(-1)), Err(E::InvalidProcedureCall));
    assert_eq!(bin(Pow, V::Integer(-8), V::Double(1.0 / 3.0)), Err(E::InvalidProcedureCall));
}

#[test]
fn double_products_may_run_to_infinity_until_the_value_is_used() {
    use BinaryOp::*;
    assert_eq!(bin(Mul, V::Double(1e300), V::Double(1e10)), Ok(V::Double(f64::INFINITY)));
    assert_eq!(bin(Add, V::Double(1e308), V::Double(1e308)), Ok(V::Double(f64::INFINITY)));
}

#[test]
fn comparison_and_logic() {
    use BinaryOp::*;
    assert_eq!(bin(Eq, V::Integer(1), V::Long(1)), Ok(V::Boolean(true)));
    assert_eq!(bin(Lt, V::Integer(1), V::Double(1.5)), Ok(V::Boolean(true)));
    assert_eq!(bin(Eq, s("a"), s("A")), Ok(V::Boolean(false)));
    assert_eq!(bin(Eq, s("a"), V::Empty), Ok(V::Boolean(false)));
    assert_eq!(bin(Eq, s(""), V::Empty), Ok(V::Boolean(true)));
    assert_eq!(bin(Eq, s("1"), V::Integer(1)), Ok(V::Boolean(true)));
    assert_eq!(bin(Eq, s("a"), V::Integer(1)), Err(E::TypeMismatch));
    assert_eq!(bin(Eq, V::Null, V::Integer(1)), Ok(V::Null));
    let text = |op, l: V, r: V| Operand::binary(op, &plain(l), &plain(r), CompareMode::Text).map(|o| o.value);
    assert_eq!(text(Eq, s("a"), s("A")), Ok(V::Boolean(true)));
    assert_eq!(text(Lt, s("a"), s("B")), Ok(V::Boolean(true)));
    assert_eq!(bin(Lt, s("a"), s("B")), Ok(V::Boolean(false)));
    assert_eq!(bin(And, V::Integer(12), V::Integer(10)), Ok(V::Integer(8)));
    assert_eq!(bin(Or, V::Integer(12), V::Integer(10)), Ok(V::Integer(14)));
    assert_eq!(bin(Xor, V::Integer(12), V::Integer(10)), Ok(V::Integer(6)));
    assert_eq!(bin(And, V::Boolean(true), V::Boolean(false)), Ok(V::Boolean(false)));
    assert_eq!(bin(And, V::Long(12), V::Integer(10)), Ok(V::Long(8)));
    assert_eq!(bin(And, V::Byte(12), V::Byte(10)), Ok(V::Byte(8)));
    assert_eq!(bin(And, V::Null, V::Integer(0)), Ok(V::Integer(0)));
    assert_eq!(bin(Or, V::Null, V::Integer(1)), Ok(V::Integer(1)));
    assert_eq!(bin(And, V::Null, V::Integer(1)), Ok(V::Null));
    assert_eq!(bin(Is, V::Nothing, V::Nothing), Ok(V::Boolean(true)));
    assert_eq!(bin(Is, V::Integer(1), V::Nothing), Err(E::TypeMismatch));
}

#[test]
fn unary_operators() {
    let neg = |v: V| plain(v).negate().map(|o| o.value);
    assert_eq!(neg(V::Integer(5)), Ok(V::Integer(-5)));
    assert_eq!(neg(V::Byte(5)), Ok(V::Integer(-5)));
    assert_eq!(neg(V::Boolean(true)), Ok(V::Integer(1)));
    assert_eq!(neg(V::Integer(-32768)), Err(E::Overflow));
    assert_eq!(neg(V::Null), Ok(V::Null));
    assert_eq!(neg(s("a")), Err(E::TypeMismatch));
    assert_eq!(variant(V::Integer(-32768)).negate().map(|o| o.value), Ok(V::Long(32768)));
    let not = |v: V| plain(v).not().map(|o| o.value);
    assert_eq!(not(V::Integer(0)), Ok(V::Integer(-1)));
    assert_eq!(not(V::Long(0)), Ok(V::Long(-1)));
    assert_eq!(not(V::Boolean(true)), Ok(V::Boolean(false)));
    assert_eq!(not(V::Null), Ok(V::Null));
}

#[test]
fn like_matches_patterns_and_rejects_malformed_ones() {
    use BinaryOp::*;
    let like = |t: &str, p: &str, mode| Operand::binary(Like, &plain(s(t)), &plain(s(p)), mode).map(|o| o.value);
    let b = CompareMode::Binary;
    assert_eq!(like("abc", "a*", b), Ok(V::Boolean(true)));
    assert_eq!(like("abc", "a?c", b), Ok(V::Boolean(true)));
    assert_eq!(like("a1", "a#", b), Ok(V::Boolean(true)));
    assert_eq!(like("b", "[a-c]", b), Ok(V::Boolean(true)));
    assert_eq!(like("d", "[!a-c]", b), Ok(V::Boolean(true)));
    assert_eq!(like("B", "b", b), Ok(V::Boolean(false)));
    assert_eq!(like("B", "b", CompareMode::Text), Ok(V::Boolean(true)));
    let t = CompareMode::Text;
    // An empty list consumes nothing; a list holding only `!` holds the character `!`.
    assert_eq!(like("", "[]", b), Ok(V::Boolean(true)));
    assert_eq!(like("a", "[]", b), Ok(V::Boolean(false)));
    assert_eq!(like("ab", "a[]b", b), Ok(V::Boolean(true)));
    assert_eq!(like("]", "[]]", b), Ok(V::Boolean(true)));
    assert_eq!(like("!", "[!]", b), Ok(V::Boolean(true)));
    assert_eq!(like("a", "[!]", b), Ok(V::Boolean(false)));
    assert_eq!(like("a", "[!!]", b), Ok(V::Boolean(true)));
    assert_eq!(like("!", "[!!]", b), Ok(V::Boolean(false)));
    assert_eq!(like("-", "[a-]", b), Ok(V::Boolean(true)));
    assert_eq!(like("e", "[a-c-e]", b), Ok(V::Boolean(true)));
    assert_eq!(like("B", "[a-c]", t), Ok(V::Boolean(true)));
    assert_eq!(like("A", "[!a-c]", t), Ok(V::Boolean(false)));
    // A malformed list is an error only where matching reaches it.
    assert_eq!(like("a", "[a", b), Err(E::InvalidPattern));
    assert_eq!(like("b", "[a", b), Err(E::InvalidPattern));
    assert_eq!(like("", "[a", b), Ok(V::Boolean(false)));
    assert_eq!(like("xb", "a[b", b), Ok(V::Boolean(false)));
    assert_eq!(like("ab", "a[b", b), Err(E::InvalidPattern));
    assert_eq!(like("a", "[!a", b), Ok(V::Boolean(false)));
    assert_eq!(like("b", "[!a", b), Err(E::InvalidPattern));
    assert_eq!(like("b", "[!a-c", b), Ok(V::Boolean(false)));
    assert_eq!(like("d", "[!a-c", b), Err(E::InvalidPattern));
    assert_eq!(like("a", "[z-a]", b), Err(E::InvalidPattern));
    assert_eq!(like("a", "a[z-a]", b), Ok(V::Boolean(false)));
    assert_eq!(like("ax", "a[z-a]", b), Err(E::InvalidPattern));
    assert_eq!(like("a", "[az-a]", b), Ok(V::Boolean(true)));
    assert_eq!(like("b", "[az-a]", b), Err(E::InvalidPattern));
    assert_eq!(like("a", "[!az-a]", b), Ok(V::Boolean(false)));
    assert_eq!(like("b", "[!az-a]", b), Err(E::InvalidPattern));
    assert_eq!(like("", "*[", b), Ok(V::Boolean(false)));
    assert_eq!(like("a", "*[", b), Err(E::InvalidPattern));
    assert_eq!(like("a", "a*[", b), Ok(V::Boolean(false)));
    assert_eq!(like("ab", "a*[", b), Err(E::InvalidPattern));
    assert_eq!(like("a", "[a][", b), Ok(V::Boolean(false)));
}

#[test]
fn only_a_few_functions_are_constant() {
    let call = |name: &str, v: V| call_function(name, &plain(v)).map(|o| o.value);
    assert_eq!(call("Len", s("abc")), Ok(V::Long(3)));
    assert_eq!(call("LenB", s("abc")), Ok(V::Long(6)));
    assert_eq!(call("Len", V::Integer(12)), Ok(V::Long(2)));
    assert_eq!(call("Int", V::Double(-2.5)), Ok(V::Double(-3.0)));
    assert_eq!(call("Fix", V::Double(-2.5)), Ok(V::Double(-2.0)));
    assert_eq!(call("Abs", V::Integer(-2)), Ok(V::Integer(2)));
    assert_eq!(call("Abs", V::Integer(-32768)), Err(E::Overflow));
    assert_eq!(call("Sgn", V::Double(-0.1)), Ok(V::Integer(-1)));
    assert_eq!(call("CLng", V::Double(2.5)), Ok(V::Long(2)));
    assert_eq!(call("CInt", s("7")), Ok(V::Integer(7)));
    assert_eq!(call("CByte", V::Integer(300)), Err(E::Overflow));
    assert_eq!(call("CStr", V::Integer(7)), Ok(s("7")));
    assert_eq!(call("CBool", V::Integer(2)), Ok(V::Boolean(true)));
    assert_eq!(call("CDate", s("1/1/2000")), Ok(V::Date(36_526.0)));
    assert_eq!(call("CDbl", V::Integer(2)), Ok(V::Double(2.0)));
    assert_eq!(call("CSng", V::Integer(2)), Ok(V::Single(2.0)));
    assert_eq!(call("CCur", V::Integer(2)), Ok(V::Currency(20_000)));
    assert_eq!(call_function("CVar", &plain(V::Integer(2))), Ok(variant(V::Integer(2))));
    assert!(vb6_syntax::frontend::const_value::is_constant_function("clng"));
    assert!(!vb6_syntax::frontend::const_value::is_constant_function("CDec"));
    assert!(!vb6_syntax::frontend::const_value::is_constant_function("Chr"));
    assert!(vb6_syntax::frontend::const_value::is_language_function("CDec"));
    assert!(vb6_syntax::frontend::const_value::is_language_function("Date"));
    assert!(!vb6_syntax::frontend::const_value::is_language_function("Chr"));
}

#[test]
fn every_error_has_its_message_code() {
    use vb6_syntax::frontend::diagnostics::message_for_code;
    for (err, code, message) in [
        (E::NotConstant, 0x9cb2, "Constant expression required"),
        (E::ExpectedFunctionOrVariable, 0x9c8c, "Expected Function or variable"),
        (E::ExpectedArray, 0x9c8a, "Expected array"),
        (E::InvalidQualifier, 0x9c9e, "Invalid qualifier"),
        (E::MemberNotFound, 0x1cd, "Method or data member not found"),
        (E::QualifierMustBeCollection, 0x9cb6, "Qualifier must be collection"),
        (E::Overflow, 6, "Overflow"),
        (E::DivisionByZero, 11, "Division by zero"),
        (E::TypeMismatch, 13, "Type mismatch"),
        (E::InvalidUseOfNull, 94, "Invalid use of Null"),
        (E::InvalidUseOfObject, 0x9cb9, "Invalid use of object"),
        (E::InvalidPattern, 93, "Invalid pattern string"),
        (E::InvalidProcedureCall, 5, "Invalid procedure call or argument"),
        (E::InvalidDataType, 0x9cae, "Invalid data type for constant"),
        (E::ObjectNotSet, 91, "Object variable or With block variable not set"),
    ] {
        assert_eq!(err.code(), code, "{err:?}");
        assert_eq!(message_for_code(code), Some(message), "{err:?}");
    }
}

#[test]
fn a_date_result_is_held_to_the_date_range_by_the_operands_types() {
    use BinaryOp::*;
    let max = V::Date(2_958_465.0);
    let min = V::Date(-657_434.0);
    // A static Date on the left: the result is a Date and must fit.
    assert_eq!(bin(Add, max.clone(), V::Byte(200)), Err(E::Overflow));
    assert_eq!(bin(Add, max.clone(), V::Integer(-5)), Ok(V::Date(2_958_460.0)));
    assert_eq!(bin(Sub, min.clone(), V::Integer(5)), Err(E::Overflow));
    assert_eq!(bin(Sub, max.clone(), V::Boolean(true)), Err(E::Overflow));
    assert_eq!(bin(Add, max.clone(), max.clone()), Err(E::Overflow));
    assert_eq!(bin(Sub, max.clone(), s("-5")), Err(E::Overflow));
    // The difference of two Dates is a Double.
    assert_eq!(bin(Sub, min.clone(), max.clone()), Ok(V::Double(-3_615_899.0)));
    // A number on the left of a Date gives a Date that is not held to the range.
    assert_eq!(bin(Add, V::Byte(200), max.clone()), Ok(V::Date(2_958_665.0)));
    assert_eq!(bin(Sub, V::Long(2_000_000_000), max.clone()), Ok(V::Date(1_997_041_535.0)));
    assert_eq!(bin(Add, V::Integer(-5), min.clone()), Ok(V::Date(-657_439.0)));
    // A String added to a Date is read as a Date first; subtracted from or by a number or a
    // String on the left, the result is a Date that is not held to the range.
    assert_eq!(bin(Add, s("5"), max.clone()), Err(E::Overflow));
    assert_eq!(bin(Add, max.clone(), s("3000000")), Err(E::TypeMismatch));
    assert_eq!(bin(Add, s("1/1/2000"), min.clone()), Ok(V::Date(-620_908.0)));
    assert_eq!(bin(Sub, s("5"), max.clone()), Ok(V::Date(-2_958_460.0)));
    // A Variant operand keeps a Variant: a Date while the value fits, a plain number after.
    let add = |l: Operand, r: Operand| Operand::binary(Add, &l, &r, CompareMode::Binary);
    assert_eq!(add(plain(max.clone()), variant(V::Integer(30_000))), Ok(variant(V::Double(2_988_465.0))));
    assert_eq!(add(variant(max.clone()), plain(V::Byte(200))), Ok(variant(V::Double(2_958_665.0))));
    assert_eq!(add(variant(max.clone()), plain(V::Integer(-5))), Ok(variant(V::Date(2_958_460.0))));
    // Such a number is a type mismatch as a Date, and a Date outside the range has no text.
    assert_eq!(variant(V::Double(2_958_665.0)).convert(T::Date), Err(E::TypeMismatch));
    assert_eq!(variant(V::Long(3_000_000)).convert(T::Date), Err(E::TypeMismatch));
    assert_eq!(plain(V::Long(3_000_000)).convert(T::Date), Err(E::Overflow));
    assert_eq!(s("3000000").convert(T::Date), Err(E::TypeMismatch));
    assert_eq!(V::Date(2_958_665.0).convert(T::String), Err(E::InvalidProcedureCall));
    assert_eq!(V::Date(2_958_665.0).convert(T::Long), Ok(V::Long(2_958_665)));
}

#[test]
fn currency_in_a_sum_or_difference_takes_a_floating_operand_as_currency() {
    use BinaryOp::*;
    assert_eq!(bin(Add, V::Currency(15_000), V::Double(0.25)), Ok(V::Currency(17_500)));
    assert_eq!(bin(Sub, V::Double(0.25), V::Currency(15_000)), Ok(V::Currency(-12_500)));
    assert_eq!(bin(Add, V::Currency(15_000), V::Double(0.00005)), Ok(V::Currency(15_000)));
    assert_eq!(bin(Add, V::Currency(15_000), V::Double(1e15)), Err(E::Overflow));
    assert_eq!(bin(Sub, V::Single(1e30), V::Currency(15_000)), Err(E::Overflow));
    assert_eq!(bin(Add, V::Currency(i64::MAX), V::Double(1.0)), Err(E::Overflow));
    // A product or quotient of Currency and a floating operand is a Double.
    assert_eq!(bin(Mul, V::Currency(15_000), V::Double(1e15)), Ok(V::Double(1.5e15)));
    assert_eq!(bin(Div, V::Currency(15_000), V::Double(1e-30)), Ok(V::Double(1.5e30)));
}

#[test]
fn currency_compared_with_a_floating_value_compares_at_four_decimals() {
    use BinaryOp::*;
    let t = |b: bool| Ok(V::Boolean(b));
    assert_eq!(bin(Eq, V::Currency(15_000), V::Double(1.50004)), t(true));
    assert_eq!(bin(Eq, V::Double(1.50004), V::Currency(15_000)), t(true));
    assert_eq!(bin(Eq, V::Currency(15_000), V::Single(1.50001)), t(true));
    assert_eq!(bin(Lt, V::Currency(15_000), V::Double(1.50006)), t(true));
    assert_eq!(bin(Gt, V::Double(1.49994), V::Currency(15_000)), t(false));
    assert_eq!(bin(Ne, V::Currency(15_000), V::Double(1.50001)), t(false));
    // A floating value too large for Currency is compared as the number it is.
    assert_eq!(bin(Lt, V::Currency(15_000), V::Double(1e20)), t(true));
    assert_eq!(bin(Gt, V::Currency(15_000), V::Double(-1e20)), t(true));
    assert_eq!(bin(Eq, V::Currency(i64::MAX), V::Single(1e20)), t(false));
}

#[test]
fn whole_number_division_converts_both_operands_before_a_null_is_seen() {
    use BinaryOp::*;
    assert_eq!(bin(IDiv, V::Single(3e38), V::Null), Err(E::Overflow));
    assert_eq!(bin(Mod, V::Null, V::Double(1e10)), Err(E::Overflow));
    assert_eq!(bin(IDiv, V::Integer(1), V::Null), Ok(V::Null));
    assert_eq!(bin(Mod, V::Null, V::Integer(1)), Ok(V::Null));
}
