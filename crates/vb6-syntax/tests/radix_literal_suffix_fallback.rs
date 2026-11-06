//! A hexadecimal or octal literal cannot take `!`, `#`, `@` or `$`, nor `%` on a value beyond
//! sixteen bits. The text is then not a literal: the `&` is the concatenation operator and the
//! rest is read again as ordinary tokens. A decimal number cannot take `$` at all.

mod common;
use common::{offset_of, parse};
use vb6_syntax::frontend::ast::{AstLit, BinOpKind, ExprNode};
use vb6_syntax::frontend::diagnostics::ERR_SYNTAX_ERROR;
use vb6_syntax::frontend::token::TypeSuffix;

fn lit_of(p: &common::Parsed, id: vb6_syntax::frontend::ast::NodeId) -> AstLit {
    match p.node(id) {
        ExprNode::Literal { lit } => lit.clone(),
        _ => panic!("not a literal"),
    }
}

/// Parse `x = 1 <text>`, assert that it is the concatenation of the Integer 1 and a name with
/// the given type suffix, and that nothing was reported.
fn name_after_the_ampersand(text: &str, name: &str, suffix: TypeSuffix) {
    let p = parse(&format!("Sub S()\n x = 1 {text}\nEnd Sub\n"));
    assert_eq!(p.found, vec![], "diagnostics for {text}");
    let body = p.proc_body(0);
    assert_eq!(body.len(), 1, "{text}");
    let ExprNode::Assign { value, .. } = p.node(body[0]) else { panic!("not an assignment") };
    let ExprNode::BinOp { op: BinOpKind::Cat, lhs, rhs } = p.node(*value) else { panic!("not a concatenation: {text}") };
    assert_eq!(lit_of(&p, *lhs), AstLit::Int(1), "{text}");
    let ExprNode::NameRef { sym, suffix: got } = p.node(*rhs) else { panic!("not a name: {text}") };
    assert_eq!((p.name(*sym), *got), (name, suffix), "{text}");
}

/// Same, with a number as the right operand.
fn number_after_the_ampersand(text: &str, expected: AstLit) {
    let p = parse(&format!("Sub S()\n x = 1 {text}\nEnd Sub\n"));
    assert_eq!(p.found, vec![], "diagnostics for {text}");
    let body = p.proc_body(0);
    let ExprNode::Assign { value, .. } = p.node(body[0]) else { panic!("not an assignment") };
    let ExprNode::BinOp { op: BinOpKind::Cat, lhs, rhs } = p.node(*value) else { panic!("not a concatenation: {text}") };
    assert_eq!(lit_of(&p, *lhs), AstLit::Int(1), "{text}");
    assert_eq!(lit_of(&p, *rhs), expected, "{text}");
}

#[test]
fn hexadecimal_literal_with_a_float_or_string_suffix_is_an_ampersand_and_a_suffixed_name() {
    name_after_the_ampersand("&H1!", "H1", TypeSuffix::Single);
    name_after_the_ampersand("&H1#", "H1", TypeSuffix::Double);
    name_after_the_ampersand("&H1@", "H1", TypeSuffix::Currency);
    name_after_the_ampersand("&H1$", "H1", TypeSuffix::String);
    name_after_the_ampersand("&HFF!", "HFF", TypeSuffix::Single);
}

#[test]
fn octal_literal_with_a_float_or_string_suffix_is_an_ampersand_and_a_suffixed_name() {
    name_after_the_ampersand("&O17!", "O17", TypeSuffix::Single);
    name_after_the_ampersand("&O17#", "O17", TypeSuffix::Double);
    name_after_the_ampersand("&O17@", "O17", TypeSuffix::Currency);
    name_after_the_ampersand("&O17$", "O17", TypeSuffix::String);
}

#[test]
fn octal_literal_without_the_letter_and_a_suffix_is_an_ampersand_and_a_number() {
    number_after_the_ampersand("&17!", AstLit::Single(17.0));
    number_after_the_ampersand("&17#", AstLit::Double(17.0));
    number_after_the_ampersand("&17@", AstLit::Currency(170_000));
}

#[test]
fn integer_suffix_beyond_sixteen_bits_gives_the_ampersand_back() {
    name_after_the_ampersand("&H10000%", "H10000", TypeSuffix::Integer);
    name_after_the_ampersand("&O200000%", "O200000", TypeSuffix::Integer);
}

#[test]
fn integer_and_long_suffixes_that_fit_stay_part_of_the_literal() {
    for (text, lit) in [
        ("&HFFFF%", AstLit::Int(-1)),
        ("&H7F%", AstLit::Int(127)),
        ("&O177777%", AstLit::Int(-1)),
        ("&HFFFF&", AstLit::Long(65535)),
        ("&H10000&", AstLit::Long(65536)),
        ("&O17&", AstLit::Long(15)),
    ] {
        let p = parse(&format!("Sub S()\n x = {text}\nEnd Sub\n"));
        assert_eq!(p.found, vec![], "{text}");
        let ExprNode::Assign { value, .. } = p.node(p.proc_body(0)[0]) else { panic!("not an assignment") };
        assert_eq!(lit_of(&p, *value), lit, "{text}");
    }
}

#[test]
fn a_decimal_number_followed_by_a_dollar_sign_is_one_bad_token() {
    for text in ["1$", "1.5$", "15$", "1234$", "1E3$", "32768$", "99999999999$", ".5$"] {
        let src = format!("Sub S()\n x = {text}\nEnd Sub\n");
        let at = offset_of(&src, text, 0);
        let found = parse(&src).found;
        assert_eq!(found[0], (ERR_SYNTAX_ERROR, at, text.len() as u32), "{text}");
    }
}
