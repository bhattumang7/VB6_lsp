//! Items of the `Print` family: positioning directives and progress on bad
//! items; `Seek` as a function.

mod common;
use common::parse;
use vb6_syntax::frontend::ast::{ExprNode, FileIoKind};

fn clean(body: &str) {
    let p = parse(&format!("Sub S()\n{body}\nEnd Sub\n"));
    assert_eq!(p.found, vec![], "{body}");
    assert_eq!(p.proc_body(0).len(), 1, "{body}");
}

#[test]
fn tab_and_spc_are_print_items() {
    for body in [
        "Print Tab(3);",
        "Print \"a\"; Spc(1); \"x\"",
        "Print #1, Tab(5); a; Spc(2)",
        "Debug.Print Tab(3)",
        "Write #1, Tab(3)",
        "Print Tab; \"x\"",
        "o.Print Tab(3)",
        "Form1.Print \"b\", Spc(2)",
        "o.Print 1; Tab(3)",
    ] {
        clean(body);
    }
}

#[test]
fn tab_keeps_its_argument_in_the_tree() {
    let p = parse("Sub S()\n Print #1, Tab(5); a\nEnd Sub\n");
    assert_eq!(p.found, vec![]);
    let body = p.proc_body(0);
    let ExprNode::FileIoStmt { kind: FileIoKind::Print, args, .. } = p.node(body[0]) else {
        panic!("not a Print")
    };
    assert_eq!(args.len(), 2);
    let ExprNode::Call { func, args: call_args } = p.node(args[0]) else { panic!("not a call") };
    let ExprNode::NameRef { sym, .. } = p.node(*func) else { panic!("not a name") };
    assert_eq!(p.name(*sym), "Tab");
    let ExprNode::ArgList { args: a } = p.node(*call_args) else { panic!("no arglist") };
    assert_eq!(a.len(), 1);
}

#[test]
fn tab_is_not_an_ordinary_expression() {
    let p = parse("Sub S()\n a = Spc(2)\nEnd Sub\n");
    assert_eq!(p.found[0].0, 0x9c6f);
}

#[test]
fn reserved_word_in_a_print_list_terminates() {
    for body in ["o.Print Then", "o.Print 1; End", "Print #1, Sub", "Print If"] {
        let p = parse(&format!("Sub S()\n{body}\nEnd Sub\n"));
        assert!(!p.found.is_empty(), "{body}");
    }
}

#[test]
fn seek_is_a_function_in_expressions() {
    clean("a = Seek(1)");
    clean("Seek #1, 5");
}
