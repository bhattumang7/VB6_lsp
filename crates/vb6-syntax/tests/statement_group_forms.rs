//! The brackets after the callee of a call statement written without `Call` hold plain
//! expressions: `ByVal` and named arguments are syntax errors there (a named argument may
//! stand in the brackets of something assigned to), and `As` cannot name the first
//! argument after a plain name.

mod common;
use common::{offset_of, parse};

fn in_sub(body: &str) -> (String, Vec<(u32, u32, u32)>) {
    let src = format!("Sub S()\n{body}\nEnd Sub\n");
    let found = parse(&src).found;
    (src, found)
}

#[test]
fn byval_and_named_arguments_in_the_brackets_of_a_call_statement_are_syntax_errors() {
    for (body, at) in [
        ("T(ByVal 1)", "ByVal"),
        ("T(a:=1)", "a:="),
        ("T(Text:=1)", "Text:="),
        ("o.Foo(a:=1)", "a:="),
        ("o.Foo.Bar(ByVal 2)", "ByVal"),
        ("o.Foo(1).Bar(a:=2)", "a:="),
        ("x(ByVal 1) = 2", "ByVal"),
    ] {
        let (src, found) = in_sub(body);
        assert_eq!(found.first().map(|f| (f.0, f.1)), Some((0x9c70, offset_of(&src, at, 0))), "{body}");
    }
}

#[test]
fn the_same_forms_are_fine_after_call_or_in_a_statement_argument_list() {
    for body in ["Call T(a:=1)", "Call T(ByVal 1)", "T 1, a:=2", "T ByVal 1, a:=ByVal 2", "x(a:=1) = 2", "s = T(a:=1)", "s = T(ByVal 1)"] {
        let (_, found) = in_sub(body);
        assert_eq!(found, vec![], "{body}");
    }
}

#[test]
fn as_cannot_name_the_first_argument_after_a_plain_name() {
    for body in ["T As:=1", "x = 1: T As:=1", "If x Then T As:=1"] {
        let (src, found) = in_sub(body);
        assert_eq!(found.first().map(|f| (f.0, f.1)), Some((0x9c70, offset_of(&src, "As", 0))), "{body}");
    }
    for body in ["T 1, As:=1", "o.Foo As:=1"] {
        let (_, found) = in_sub(body);
        assert_eq!(found, vec![], "{body}");
    }
}

#[test]
fn a_bracket_group_after_another_takes_neither_byval_nor_an_omitted_argument() {
    for (body, at) in [("x = Bar(1)(ByVal 2)", "ByVal"), ("x = o.Foo(1)(ByVal 2)", "ByVal"), ("x = Bar(1)(, 2)", ", 2")] {
        let (src, found) = in_sub(body);
        assert_eq!(found.first().map(|f| f.1), Some(offset_of(&src, at, 0)), "{body}");
    }
    for body in ["x = Bar(1)(a:=2)", "x = Bar(1)()", "x = o.Foo(1)(AddressOf Bar)", "x = Bar(, 2)(3)", "x = Bar(ByVal 1)"] {
        let (_, found) = in_sub(body);
        assert_eq!(found, vec![], "{body}");
    }
}
