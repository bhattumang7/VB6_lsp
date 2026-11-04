//! `&` followed by an octal digit is an octal literal without the `O`, wherever
//! it stands; a date whose first field cannot be a month is read day first.

mod common;
use common::{parse, Parsed};
use vb6_syntax::frontend::ast::{AstLit, ExprNode, NodeId};

fn rhs(src: &str) -> (Parsed, NodeId) {
    let p = parse(&format!("Sub S()\n    x = {src}\nEnd Sub\n"));
    let ExprNode::Assign { value, .. } = p.node(p.proc_body(0)[0]) else { panic!("assignment") };
    let value = *value;
    (p, value)
}

fn literal(src: &str) -> AstLit {
    let (p, v) = rhs(src);
    assert_eq!(p.found, vec![], "{src}");
    let ExprNode::Literal { lit } = p.node(v) else { panic!("{src}: expected a literal") };
    lit.clone()
}

#[test]
fn octal_literal_without_the_letter() {
    assert_eq!(literal("&7"), AstLit::Int(7));
    assert_eq!(literal("&07"), AstLit::Int(7));
    assert_eq!(literal("&17777"), AstLit::Int(0o17777));
    assert_eq!(literal("&177777"), AstLit::Int(-1));
    assert_eq!(literal("&200000"), AstLit::Long(65536));
    assert_eq!(literal("&7&"), AstLit::Long(7));
    assert_eq!(literal("&7%"), AstLit::Int(7));
}

#[test]
fn octal_literal_equals_the_form_with_the_letter() {
    assert_eq!(literal("&17"), literal("&O17"));
}

#[test]
fn an_octal_literal_after_an_operand_is_two_operands_in_a_row() {
    for text in ["x = a &7", "x = a &1", "x = a &7 + 1"] {
        let src = format!("Sub S()\n    {text}\nEnd Sub\n");
        let p = parse(&src);
        assert!(!p.found.is_empty(), "{text}");
        assert!(p.found.iter().any(|d| d.1 as usize == src.find('&').unwrap()), "{text}: {:?}", p.found);
    }
    assert_eq!(parse("Sub S()\n    x = a & 7\nEnd Sub\n").found, vec![]);
    assert_eq!(parse("Sub S()\n    x = 1 + &7\nEnd Sub\n").found, vec![]);
}

#[test]
fn ampersand_before_eight_or_nine_is_still_the_operator() {
    for text in ["x = &8", "x = &9", "x = &"] {
        assert!(!parse(&format!("Sub S()\n    {text}\nEnd Sub\n")).found.is_empty(), "{text}");
    }
}

#[test]
fn a_date_with_a_first_field_that_cannot_be_a_month_is_day_first() {
    assert_eq!(literal("#13/1/2000#"), literal("#1/13/2000#"));
    assert_eq!(literal("#31/12/1999#"), literal("#12/31/1999#"));
}

#[test]
fn dates_that_stay_month_first_or_invalid_are_unchanged() {
    for text in ["#1/2/2000#", "#12/31/1999#"] {
        assert_eq!(parse(&format!("Sub S()\n    x = {text}\nEnd Sub\n")).found, vec![], "{text}");
    }
    for text in ["#13/13/2000#", "#0/1/2000#", "#2/30/2000#", "#32/1/2000#"] {
        assert!(!parse(&format!("Sub S()\n    x = {text}\nEnd Sub\n")).found.is_empty(), "{text}");
    }
}
