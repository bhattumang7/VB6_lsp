//! A line-continuation mark must follow a blank; a statement cannot end in a
//! bang member access; the closing words of an Enum share one physical line.

mod common;
use common::{offset_of, parse};
use vb6_syntax::frontend::ast::{AstLit, BinOpKind, ExprNode};

fn in_sub(stmt: &str) -> String {
    format!("Sub S(o As Object)\n    {stmt}\nEnd Sub\n")
}

#[test]
fn a_continuation_mark_after_a_blank_joins_the_lines() {
    for text in ["x = 1 + _\n        2", "x = 1 +\t_\n        2", "x = 1 + _ \n        2"] {
        assert_eq!(parse(&in_sub(text)).found, vec![], "{text}");
    }
}

#[test]
fn a_continuation_mark_against_a_token_is_an_error_at_the_mark() {
    for text in ["x = 1 +_\n        2", "x = (1)_\n + 2"] {
        let src = in_sub(text);
        let p = parse(&src);
        assert!(p.found.iter().any(|d| d.1 == offset_of(&src, "_", 0)), "{text}: {:?}", p.found);
    }
}

#[test]
fn a_continuation_mark_at_the_start_of_a_line_is_an_error() {
    let src = in_sub("x = 1 + _\n_\n        2");
    let p = parse(&src);
    assert!(p.found.iter().any(|d| d.1 == offset_of(&src, "_", 1)), "{:?}", p.found);
}

#[test]
fn an_indented_continuation_line_is_still_a_continuation() {
    assert_eq!(parse(&in_sub("x = 1 + _\n        _\n        2")).found, vec![]);
}

#[test]
fn a_statement_ending_in_a_bang_member_is_a_syntax_error() {
    for text in ["o!p", "o!p!q", "o.p!q", "o!p(1)", "o!p 1", "Call o!p"] {
        let src = in_sub(text);
        let p = parse(&src);
        assert!(p.found.iter().any(|d| d.0 == 0x9c70), "{text}: {:?}", p.found);
    }
}

#[test]
fn a_bang_member_in_the_middle_or_assigned_is_fine() {
    for text in ["o!p = 1", "o!p.q", "o!p.q 1", "x = o!p", "x = o!p!q"] {
        assert_eq!(parse(&in_sub(text)).found, vec![], "{text}");
    }
}

#[test]
fn a_bang_statement_inside_with_is_also_an_error() {
    let src = "Sub S(o As Object)\n    With o\n        !p\n    End With\nEnd Sub\n";
    let p = parse(src);
    assert!(p.found.iter().any(|d| d.0 == 0x9c70 && d.1 >= offset_of(src, "!p", 0)), "{:?}", p.found);
}

#[test]
fn enum_closing_words_on_one_line_are_fine() {
    assert_eq!(parse("Enum E\n    A\nEnd Enum\n").found, vec![]);
    assert_eq!(parse("Enum E\n    A\nEnd   Enum\n").found, vec![]);
}

#[test]
fn enum_closing_words_split_by_a_continuation_are_invalid_inside_enum() {
    let src = "Enum E\n    A\nEnd _\n Enum\n";
    let p = parse(src);
    assert!(p.found.iter().any(|d| d.0 == 0xdeff), "{:?}", p.found);
}

#[test]
fn a_string_on_a_line_ending_in_a_mark_loses_the_blanks_before_the_mark() {
    for text in ["x = \"ab _
 + \"z\"", "x = \"ab	  _ 
 + \"z\""] {
        let p = parse(&in_sub(text));
        assert_eq!(p.found, vec![], "{text}");
        let ExprNode::Assign { value, .. } = p.node(p.proc_body(0)[0]) else { panic!("assignment") };
        let ExprNode::BinOp { op: BinOpKind::Add, lhs, rhs } = p.node(*value) else { panic!("sum") };
        assert!(matches!(p.node(*lhs), ExprNode::Literal { lit: AstLit::Str(s) } if &**s == "ab"), "{text}");
        assert!(matches!(p.node(*rhs), ExprNode::Literal { lit: AstLit::Str(s) } if &**s == "z"), "{text}");
    }
}

#[test]
fn blanks_that_do_not_lead_to_the_mark_stay_in_the_string() {
    let p = parse(&in_sub("x = \"a _x \"
"));
    assert_eq!(p.found, vec![]);
    let ExprNode::Assign { value, .. } = p.node(p.proc_body(0)[0]) else { panic!("assignment") };
    assert!(matches!(p.node(*value), ExprNode::Literal { lit: AstLit::Str(s) } if &**s == "a _x "));
}

#[test]
fn a_string_joined_to_the_next_line_leaves_that_line_as_more_operands() {
    for text in ["x = \"ab _
cd\"", "x = \"ab _
y = 2"] {
        assert!(!parse(&in_sub(text)).found.is_empty(), "{text}");
    }
}

#[test]
fn a_bracketed_name_cannot_span_lines() {
    for text in ["x = [ab\ncd]", "x = [ab _\ncd]"] {
        let src = in_sub(text);
        let p = parse(&src);
        assert!(p.found.iter().any(|d| d.1 == offset_of(&src, "[ab", 0)), "{text}: {:?}", p.found);
    }
}
