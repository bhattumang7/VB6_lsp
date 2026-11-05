//! Numeric and string literal typing and range rules.

mod common;
use common::{offset_of, parse};
use vb6_syntax::frontend::ast::{AstLit, ExprNode};
use vb6_syntax::frontend::diagnostics::ERR_SYNTAX_ERROR;

const ERR_UNEXPECTED_TOKEN: u32 = 0x9c6f;

/// Parse `x = <text>` and return the literal on the right, asserting that
/// nothing was reported.
fn literal(text: &str) -> AstLit {
    let p = parse(&format!("Sub S()\n x = {text}\nEnd Sub\n"));
    assert_eq!(p.found, vec![], "diagnostics for {text}");
    let body = p.proc_body(0);
    assert_eq!(body.len(), 1);
    let ExprNode::Assign { value, .. } = p.node(body[0]) else { panic!("not an assignment") };
    match p.node(*value) {
        ExprNode::Literal { lit } => lit.clone(),
        _ => panic!("not a literal"),
    }
}

/// Parse `x = <text>` and return the diagnostics `(code, offset-in-text, len)`.
fn rejected(text: &str) -> Vec<(u32, u32, u32)> {
    let src = format!("Sub S()\n x = {text}\nEnd Sub\n");
    let base = offset_of(&src, text, 0);
    parse(&src).found.into_iter().map(|(c, s, l)| (c, s - base, l)).collect()
}

#[test]
fn hex_literal_that_fits_sixteen_bits_is_an_integer_bit_pattern() {
    assert_eq!(literal("&H7FFF"), AstLit::Int(32767));
    assert_eq!(literal("&H8000"), AstLit::Int(-32768));
    assert_eq!(literal("&HFFFF"), AstLit::Int(-1));
    assert_eq!(literal("&H0"), AstLit::Int(0));
}

#[test]
fn hex_literal_beyond_sixteen_bits_is_a_long() {
    assert_eq!(literal("&H10000"), AstLit::Long(65536));
    assert_eq!(literal("&H7FFFFFFF"), AstLit::Long(i32::MAX));
    assert_eq!(literal("&HFFFFFFFF"), AstLit::Long(-1));
}

#[test]
fn ampersand_suffix_makes_a_hex_literal_long() {
    assert_eq!(literal("&HFFFF&"), AstLit::Long(65535));
    assert_eq!(literal("&H1&"), AstLit::Long(1));
}

#[test]
fn percent_suffix_keeps_a_hex_literal_integer() {
    assert_eq!(literal("&HFFFE%"), AstLit::Int(-2));
    assert_eq!(literal("&H7F%"), AstLit::Int(127));
}

#[test]
fn octal_literals_follow_the_same_sizing() {
    assert_eq!(literal("&O177777"), AstLit::Int(-1));
    assert_eq!(literal("&O17"), AstLit::Int(15));
    assert_eq!(literal("&O200000"), AstLit::Long(65536));
    assert_eq!(literal("&O37777777777"), AstLit::Long(-1));
    assert_eq!(literal("&O7&"), AstLit::Long(7));
}

/// Parse `x = a <text>` and return the name read after the `&` operator.
fn name_after_concatenation(text: &str) -> String {
    let p = parse(&format!("Sub S()\n x = a {text}\nEnd Sub\n"));
    assert_eq!(p.found, vec![], "diagnostics for {text}");
    let body = p.proc_body(0);
    let ExprNode::Assign { value, .. } = p.node(body[0]) else { panic!("not an assignment") };
    let ExprNode::BinOp { op: vb6_syntax::frontend::ast::BinOpKind::Cat, rhs, .. } = p.node(*value) else { panic!("not a concatenation: {text}") };
    let ExprNode::NameRef { sym, .. } = p.node(*rhs) else { panic!("not a name: {text}") };
    p.name(*sym).to_string()
}

#[test]
fn radix_text_that_is_not_a_whole_literal_is_the_concatenation_operator_and_a_name() {
    for (text, name) in [
        ("&H", "H"),
        ("&O", "O"),
        ("&Hz", "Hz"),
        ("&Oz", "Oz"),
        ("&H1o", "H1o"),
        ("&H100000000", "H100000000"),
        ("&O40000000000", "O40000000000"),
        ("&O18", "O18"),
        ("&H1O", "H1O"),
    ] {
        assert_eq!(name_after_concatenation(text), name, "{text}");
    }
}

#[test]
fn an_eight_or_nine_after_octal_digits_makes_the_ampersand_an_operator() {
    let p = parse("Sub S()\n n = 5 &18\nEnd Sub\n");
    assert_eq!(p.found, vec![]);
    let body = p.proc_body(0);
    let ExprNode::Assign { value, .. } = p.node(body[0]) else { panic!("not an assignment") };
    let ExprNode::BinOp { op: vb6_syntax::frontend::ast::BinOpKind::Cat, rhs, .. } = p.node(*value) else { panic!("not a concatenation") };
    assert!(matches!(p.node(*rhs), ExprNode::Literal { lit: AstLit::Int(18) }));
}

#[test]
fn a_hex_literal_with_the_integer_suffix_must_fit_sixteen_bits() {
    assert_eq!(rejected("&H10000%"), vec![(ERR_SYNTAX_ERROR, 0, 8)]);
}

#[test]
fn whole_number_literals_are_sized_by_range() {
    assert_eq!(literal("32767"), AstLit::Int(32767));
    assert_eq!(literal("32768"), AstLit::Long(32768));
    assert_eq!(literal("2147483647"), AstLit::Long(i32::MAX));
}

#[test]
fn whole_number_beyond_long_is_a_double() {
    assert_eq!(literal("2147483648"), AstLit::Double(2147483648.0));
    assert_eq!(literal("99999999999999999999"), AstLit::Double(1e20));
}

#[test]
fn suffixed_whole_number_must_fit_its_type() {
    assert_eq!(literal("32767%"), AstLit::Int(32767));
    assert_eq!(literal("2147483647&"), AstLit::Long(i32::MAX));
    assert_eq!(rejected("32768%"), vec![(ERR_SYNTAX_ERROR, 0, 6)]);
    // A whole number too large for a Long is a Double, and a `&` after it is the operator.
    let p = parse("Sub S()\n x = 3000000000& \"a\"\nEnd Sub\n");
    assert_eq!(p.found, vec![]);
}

#[test]
fn floating_literals_beyond_their_range_are_syntax_errors() {
    assert_eq!(rejected("1E400"), vec![(ERR_SYNTAX_ERROR, 0, 5)]);
    assert_eq!(rejected("3.4E39!"), vec![(ERR_SYNTAX_ERROR, 0, 7)]);
    assert_eq!(literal("1E38!"), AstLit::Single(1e38));
    assert_eq!(literal("1.5E300"), AstLit::Double(1.5e300));
}

#[test]
fn currency_literal_is_exact() {
    assert_eq!(literal("123456789012345@"), AstLit::Currency(1_234_567_890_123_450_000));
    assert_eq!(literal("1.5@"), AstLit::Currency(15_000));
    // Halves round to the even unit.
    assert_eq!(literal("0.00005@"), AstLit::Currency(0));
    assert_eq!(literal("0.00015@"), AstLit::Currency(2));
    assert_eq!(literal("0.00025@"), AstLit::Currency(2));
    assert_eq!(literal("0.00035@"), AstLit::Currency(4));
    assert_eq!(literal("5E-5@"), AstLit::Currency(0));
    assert_eq!(literal("922337203685477.5807@"), AstLit::Currency(i64::MAX));
    assert_eq!(rejected("922337203685478@"), vec![(ERR_SYNTAX_ERROR, 0, 16)]);
}

#[test]
fn string_ends_at_the_end_of_its_line() {
    let src = "Sub S()\n x = \"abc\n y = 2\nEnd Sub\n";
    let p = parse(src);
    assert_eq!(p.found, vec![]);
    let body = p.proc_body(0);
    assert_eq!(body.len(), 2);
    let ExprNode::Assign { value, .. } = p.node(body[0]) else { panic!("not an assignment") };
    match p.node(*value) {
        ExprNode::Literal { lit: AstLit::Str(s) } => assert_eq!(&**s, "abc"),
        _ => panic!("not a string"),
    }
    let ExprNode::Assign { value, .. } = p.node(body[1]) else { panic!("not an assignment") };
    assert!(matches!(p.node(*value), ExprNode::Literal { lit: AstLit::Int(2) }));
}

#[test]
fn underscore_inside_a_closed_string_is_not_a_continuation() {
    let src = "Sub S()\n x = \"a _\"\n y = 2\nEnd Sub\n";
    let p = parse(src);
    assert_eq!(p.found, vec![]);
    assert_eq!(p.proc_body(0).len(), 2);
}

#[test]
fn reserved_words_cannot_start_an_expression() {
    for word in ["Then", "To", "As", "End", "If", "Dim", "Sub"] {
        let src = format!("Sub S()\n x = {word}\nEnd Sub\n");
        let at = offset_of(&src, &format!("= {word}"), 0) + 2;
        let found = parse(&src).found;
        assert_eq!(found[0], (ERR_UNEXPECTED_TOKEN, at, word.len() as u32), "{word}");
    }
}

#[test]
fn values_and_builtin_functions_start_expressions() {
    for word in ["True", "False", "Nothing", "Empty", "Null", "Me", "Date", "Len(a)", "Array(1)", "New C", "Not a", "TypeOf a Is C", "CDec(1)", "DoEvents()", "Abs(1)"] {
        let p = parse(&format!("Sub S()\n x = {word}\nEnd Sub\n"));
        assert_eq!(p.found, vec![], "{word}");
    }
}

#[test]
fn punctuation_cannot_start_an_expression() {
    for (text, at) in [(")", 0u32), (",", 0), ("1 + )", 4)] {
        let src = format!("Sub S()\n x = {text}\nEnd Sub\n");
        let base = offset_of(&src, &format!("= {text}"), 0) + 2;
        let found = parse(&src).found;
        assert_eq!(found[0], (ERR_UNEXPECTED_TOKEN, base + at, 1), "{text}");
    }
}

#[test]
fn an_exponent_mark_needs_digits() {
    for text in ["1E", "2D", "1Else", "1e x", "1e+x"] {
        let found = rejected(text);
        assert_eq!(found.first().map(|f| (f.0, f.1)), Some((ERR_SYNTAX_ERROR, 0)), "{text}");
    }
    assert_eq!(literal("1E5"), AstLit::Double(100000.0));
    assert_eq!(literal("1d2"), AstLit::Double(100.0));
    assert_eq!(literal("1E+2"), AstLit::Double(100.0));
}

#[test]
fn a_single_literal_is_rounded_through_a_double() {
    // Two roundings (text to Double, Double to Single) give one Single for both spellings.
    let a = literal("0.0008833812607917935!");
    let b = literal("0.000883381231687963!");
    assert_eq!(a, b);
}
