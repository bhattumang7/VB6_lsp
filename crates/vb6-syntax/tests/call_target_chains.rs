//! Which member chains can follow `Call`: a `!` access cannot be called, and a plain name cannot
//! take two bracket groups in a row, while `Me` and the object of a `With` can.

mod common;
use common::{offset_of, parse, parse_kind};
use vb6_syntax::frontend::ast::ExprNode;
use vb6_syntax::frontend::parser::ModuleKind;

const SYNTAX_ERROR: u32 = 0x9c70;

fn in_sub(body: &str) -> String {
    format!("Sub S(o As Object, arr() As Long)\n{body}\nEnd Sub\n")
}

#[test]
fn a_bang_anywhere_in_the_chain_is_a_syntax_error_over_the_chain() {
    for chain in ["o.a!b", "o!a.b", "o!a(1).b", "o!a", "o!a(1)"] {
        assert_one_error_within(&in_sub(&format!("Call {chain}")), chain);
    }
    assert_one_error_within(&in_sub("With o\nCall .a!b\nEnd With"), ".a!b");
    assert_one_error_within(&in_sub("With o\nCall !a\nEnd With"), "!a");
}

/// The source has exactly one diagnostic, a syntax error that lies inside `chain`.
fn assert_one_error_within(src: &str, chain: &str) {
    let found = parse(src).found;
    let start = offset_of(src, chain, 0);
    assert_eq!(found.len(), 1, "{chain}: {found:?}");
    let (code, at, len) = found[0];
    assert_eq!(code, SYNTAX_ERROR, "{chain}");
    assert!(at >= start && at + len <= start + chain.len() as u32, "{chain}: {found:?}");
}

#[test]
fn two_bracket_groups_in_a_row_after_a_plain_name_are_a_syntax_error() {
    for chain in ["Q(1)(2)", "o.a(1)(2)", "o.a(1) (2)", "arr(1)(2)", "o.a(1).b(2)(3)"] {
        assert_one_error_within(&in_sub(&format!("Call {chain}")), chain);
    }
}

#[test]
fn me_and_the_object_of_a_with_take_bracket_groups_in_a_row() {
    for chain in [".a(1)(2)", ".a(1)(2)(3)", ".a(1).b(2)(3)", ".a.b(1)(2)"] {
        let src = in_sub(&format!("With o\nCall {chain}\nEnd With"));
        assert_eq!(parse(&src).found, vec![], "{chain}");
    }
    let src = in_sub("Call Me.a(1)(2)");
    assert_eq!(parse_kind(&src, ModuleKind::Class).found, vec![]);
}

#[test]
fn a_chain_with_one_bracket_group_per_member_is_called() {
    let src = in_sub("Call o.a(1).b(2)");
    let p = parse(&src);
    assert_eq!(p.found, vec![]);
    let ExprNode::CallStmt { callee, args } = p.node(p.proc_body(0)[0]) else { panic!("not a call statement") };
    let ExprNode::Call { func, args: last } = p.node(*callee) else { panic!("callee is not a call") };
    assert!(matches!(p.node(*func), ExprNode::MemberAccess { bang: false, .. }));
    let ExprNode::ArgList { args: last } = p.node(*last) else { panic!("arg list") };
    assert_eq!(last.len(), 1);
    let ExprNode::ArgList { args } = p.node(*args) else { panic!("arg list") };
    assert!(args.is_empty());
}
