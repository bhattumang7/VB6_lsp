//! `Call` takes its callee with brackets or alone; a statement call with empty
//! brackets, or with several bracketed arguments, needs `Call`.

mod common;
use common::{offset_of, parse, Parsed};
use vb6_syntax::frontend::ast::{ExprNode, NodeId};

fn name_of(p: &Parsed, id: NodeId) -> String {
    let ExprNode::NameRef { sym, .. } = p.node(id) else { panic!("expected a name") };
    p.name(*sym).to_string()
}

fn only(src: &str) -> (Parsed, NodeId, String) {
    let full = format!("Sub S()\n    {src}\nEnd Sub\n");
    let p = parse(&full);
    let body = p.proc_body(0);
    (p, body[0], full)
}

#[test]
fn call_with_brackets_keeps_the_arguments() {
    let (p, stmt, _) = only("Call Foo(1, a)");
    assert_eq!(p.found, vec![]);
    let ExprNode::CallStmt { callee, args } = p.node(stmt) else { panic!("expected CallStmt") };
    // The brackets belong to the callee expression; the statement itself has no loose arguments.
    assert!(matches!(p.node(*args), ExprNode::ArgList { args } if args.is_empty()));
    let ExprNode::Call { func, args } = p.node(*callee) else { panic!("expected a bracketed call") };
    assert_eq!(name_of(&p, *func), "Foo");
    let ExprNode::ArgList { args } = p.node(*args) else { panic!("expected arguments") };
    assert_eq!(args.len(), 2);
    assert_eq!(name_of(&p, args[1]), "a");
}

#[test]
fn call_without_brackets_has_no_arguments() {
    let (p, stmt, _) = only("Call Foo");
    assert_eq!(p.found, vec![]);
    let ExprNode::CallStmt { callee, args } = p.node(stmt) else { panic!("expected CallStmt") };
    assert_eq!(name_of(&p, *callee), "Foo");
    assert!(matches!(p.node(*args), ExprNode::ArgList { args } if args.is_empty()));
}

#[test]
fn call_with_loose_arguments_is_an_error_at_the_first() {
    for (text, at) in [("Call Foo 1", "1"), ("Call Foo, 1", ", 1"), ("Call Foo(1), 2", ", 2"), ("Call Foo(1) 2", "2")] {
        let (p, _, full) = only(text);
        assert!(!p.found.is_empty(), "{text}");
        assert!(p.found.iter().any(|d| d.1 == offset_of(&full, at, 0)), "{text}: {:?}", p.found);
    }
}

#[test]
fn call_of_a_reserved_word_is_a_syntax_error_at_the_word() {
    let (p, _, full) = only("Call Dim");
    assert_eq!(p.found[0], (0x9c70, offset_of(&full, "Dim", 0), 3));
}

#[test]
fn statement_calls_with_empty_or_several_bracketed_arguments_need_call() {
    for text in ["Foo()", "Foo(1, 2)", "Foo(AddressOf Bar)", "o.Foo()"] {
        let (p, _, _) = only(text);
        assert!(!p.found.is_empty(), "{text}");
        assert_eq!(p.found[0].0, 0x9c70, "{text}");
    }
    for text in ["Foo(1)", "Foo 1, 2", "Foo (1), 2", "PSet (1, 2)"] {
        let (p, _, _) = only(text);
        assert_eq!(p.found, vec![], "{text}");
    }
}
