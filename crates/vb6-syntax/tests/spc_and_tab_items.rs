//! In a `Print`-family list `Spc` takes exactly one bracketed expression, `Tab` that or
//! nothing, and the item ends at the closing bracket.

mod common;
use common::{offset_of, parse, Parsed};
use vb6_syntax::frontend::ast::{AstLit, ExprNode, FileIoKind, NodeId};

fn print_args(p: &Parsed) -> Vec<NodeId> {
    let b = p.proc_body(0);
    match p.node(b[0]) {
        ExprNode::FileIoStmt { kind: FileIoKind::Print, args, .. } => args.clone(),
        _ => panic!("not a Print statement"),
    }
}

fn is_name(p: &Parsed, id: NodeId, want: &str) -> bool {
    matches!(p.node(id), ExprNode::NameRef { sym, .. } if p.name(*sym) == want)
}

fn call_with_one_int(p: &Parsed, id: NodeId, callee: &str, want: i32) -> bool {
    let ExprNode::Call { func, args } = p.node(id) else { return false };
    let ExprNode::ArgList { args } = p.node(*args) else { return false };
    is_name(p, *func, callee)
        && args.len() == 1
        && matches!(p.node(args[0]), ExprNode::Literal { lit: AstLit::Int(n) } if *n == want)
}

#[test]
fn spc_and_tab_with_one_expression() {
    let p = parse("Sub S()\n Print Spc(3); Tab(5)\nEnd Sub\n");
    assert_eq!(p.diags, 0);
    let a = print_args(&p);
    assert_eq!(a.len(), 2);
    assert!(call_with_one_int(&p, a[0], "Spc", 3));
    assert!(call_with_one_int(&p, a[1], "Tab", 5));
}

#[test]
fn bare_tab_is_a_name() {
    let p = parse("Sub S()\n Print Tab; 1\nEnd Sub\n");
    assert_eq!(p.diags, 0);
    let a = print_args(&p);
    assert_eq!(a.len(), 2);
    assert!(is_name(&p, a[0], "Tab"));
}

#[test]
fn the_item_ends_at_the_closing_bracket() {
    let p = parse("Sub S()\n Print Tab(1)(2)\nEnd Sub\n");
    assert_eq!(p.diags, 0);
    let a = print_args(&p);
    assert_eq!(a.len(), 2);
    assert!(call_with_one_int(&p, a[0], "Tab", 1));
    let ExprNode::Paren { inner } = p.node(a[1]) else { panic!("second item is not bracketed") };
    assert!(matches!(p.node(*inner), ExprNode::Literal { lit: AstLit::Int(2) }));
}

#[test]
fn spc_and_tab_reject_other_forms() {
    // (item text, offset within it of the first rejected token)
    let cases: [(&str, u32); 5] = [("Spc", 3), ("Spc 5", 4), ("Spc()", 4), ("Spc(1, 2)", 5), ("Tab()", 4)];
    for (item, rel) in cases {
        let src = format!("Sub S()\n Print {item}\nEnd Sub\n");
        let p = parse(&src);
        assert_eq!(p.found[0].1, offset_of(&src, "Print ", 0) + 6 + rel, "{item}: {:?}", p.found);
    }
    let src = "Sub S()\n Print Tab(1, 2)\nEnd Sub\n";
    let p = parse(src);
    assert_eq!(p.found[0].1, offset_of(src, ", 2", 0));
}
