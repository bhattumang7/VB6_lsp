//! A call whose callee is itself a call, or a `!` member, is not a call statement.

mod common;
use common::{offset_of, parse};
use vb6_syntax::frontend::ast::ExprNode;

const SYNTAX_ERROR: u32 = 0x9c70;

fn sub(body: &str) -> String {
    format!("Sub S()\n Dim d As Object\n{body}\nEnd Sub\n")
}

#[test]
fn a_call_of_a_call_is_rejected_where_the_statement_should_end() {
    // (statement, offset within it of the first rejected token, its length)
    let cases: [(&str, u32, u32); 3] = [("F(0)(1)", 7, 0), ("o.Items(1)(2)", 13, 0), ("F(0)(1), 2", 7, 1)];
    for (stmt, rel, len) in cases {
        let src = sub(&format!(" {stmt}"));
        let p = parse(&src);
        assert_eq!(p.found[0], (SYNTAX_ERROR, offset_of(&src, stmt, 0) + rel, len), "{stmt}: {:?}", p.found);
    }
}

#[test]
fn a_bang_member_with_a_bracketed_first_argument_is_rejected() {
    let src = sub(" d!x (1), 2");
    let p = parse(&src);
    assert_eq!(p.found[0], (SYNTAX_ERROR, offset_of(&src, "2
", 0), 1), "{:?}", p.found);
}

#[test]
fn a_single_call_with_a_continued_argument_is_still_a_call_statement() {
    let p = parse(&sub(" Foo (1), 2"));
    assert_eq!(p.diags, 0);
    let b = p.proc_body(0);
    let ExprNode::CallStmt { args, .. } = p.node(*b.last().unwrap()) else { panic!("not a call statement") };
    let ExprNode::ArgList { args } = p.node(*args) else { panic!("no argument list") };
    assert_eq!(args.len(), 2);
}

#[test]
fn a_plain_call_statement_with_one_group_is_accepted() {
    assert_eq!(parse("Sub S()\n Foo (1)\nEnd Sub\n").found, vec![]);
}
