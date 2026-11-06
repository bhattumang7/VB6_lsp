//! A number that cannot be made into a literal is reported by what went wrong with it.
//! Outside a procedure a floating value beyond the Double range (or an exponent mark and sign
//! with no digits) is an overflow, a value that does not fit its suffix is an unreadable
//! expression, and an exponent mark with nothing after it is a syntax error. Inside a
//! procedure all of them are syntax errors.

use vb6_syntax::frontend::ast::ExprArena;
use vb6_syntax::frontend::diagnostics::{ERR_OVERFLOW, ERR_SYNTAX_ERROR};
use vb6_syntax::frontend::parser::Parser;
use vb6_syntax::frontend::scanner::ScannerContext;

/// `(code, start, len, message)` of every diagnostic for the source.
fn reported(src: &str) -> Vec<(u32, u32, u32, String)> {
    let mut ctx = ScannerContext::new(1, 1, 0x0409);
    ctx.intern_keywords();
    let mut arena = ExprArena::new();
    let mut parser = Parser::new(&mut ctx, src.as_bytes());
    parser.parse_module(&mut arena);
    parser
        .diagnostics
        .items()
        .iter()
        .map(|d| (d.code, d.span.start, d.span.len, d.message().unwrap_or_default()))
        .collect()
}

fn module_level(value: &str) -> Vec<(u32, u32, u32, String)> {
    let src = format!("Const c = {value}\n");
    let at = src.find(value).unwrap() as u32;
    reported(&src).into_iter().map(|(c, s, l, m)| (c, s - at, l, m)).collect()
}

fn in_procedure(value: &str) -> Vec<(u32, u32, u32, String)> {
    let src = format!("Sub S()\n Const c = {value}\nEnd Sub\n");
    let at = src.find(value).unwrap() as u32;
    reported(&src).into_iter().map(|(c, s, l, m)| (c, s - at, l, m)).collect()
}

#[test]
fn a_floating_value_beyond_the_double_range_is_an_overflow_at_module_level() {
    for text in ["1E400", "1E309", "1E400#", "1.5E400", "1D400", "1d400", ".5E400", "5.E400"] {
        assert_eq!(module_level(text), vec![(ERR_OVERFLOW, 0, text.len() as u32, "Overflow".to_string())], "{text}");
    }
}

#[test]
fn an_exponent_mark_and_sign_without_digits_is_an_overflow_at_module_level() {
    for text in ["1E+", "1E-"] {
        assert_eq!(module_level(text), vec![(ERR_OVERFLOW, 0, 3, "Overflow".to_string())], "{text}");
    }
}

#[test]
fn an_exponent_mark_without_digits_or_sign_is_a_syntax_error() {
    for text in ["1E", "1Ex", "12e"] {
        assert_eq!(module_level(text), vec![(ERR_SYNTAX_ERROR, 0, text.len() as u32, "Syntax error".to_string())], "{text}");
    }
}

#[test]
fn a_value_that_does_not_fit_its_suffix_is_an_unreadable_expression_at_module_level() {
    for text in ["3.4E39!", "1E39!", "32768%", "1E5%", "1E20@", "1E15@", "1$", "1E400$", "1E400@"] {
        assert_eq!(module_level(text), vec![(0x9c6f, 0, text.len() as u32, "Expected: expression".to_string())], "{text}");
    }
}

#[test]
fn values_at_the_edge_of_their_type_are_accepted() {
    for text in ["1E308", "1E-400", "3.4E38!", "1E-45!", "32767%", "99999999999999999999", "922337203685477@", "1E14@"] {
        assert_eq!(module_level(text), vec![], "{text}");
    }
}

#[test]
fn inside_a_procedure_every_one_of_them_is_a_syntax_error() {
    for text in ["1E400", "1E+", "1E", "3.4E39!", "32768%", "1$", "1E400$"] {
        assert_eq!(in_procedure(text), vec![(ERR_SYNTAX_ERROR, 0, text.len() as u32, "Syntax error".to_string())], "{text}");
    }
}
