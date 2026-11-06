//! A `.` or `!` after blanks on the same line does not continue the member
//! chain before it: in a call statement it begins a With-member argument, and
//! anywhere else it is a syntax error. A line continuation does not count as
//! blanks. A blank after the `.` is fine; after the `!` it is not.

mod common;
use common::{offset_of, parse, Parsed};
use vb6_syntax::frontend::ast::{AstLit, BinOpKind, ExprNode, NodeId};

fn name_of(p: &Parsed, id: NodeId) -> String {
    let ExprNode::NameRef { sym, .. } = p.node(id) else { panic!("expected a name") };
    p.name(*sym).to_string()
}

/// `base.member` / `base!member`: the base node, the member name and the bang flag.
fn member(p: &Parsed, id: NodeId) -> (NodeId, String, bool) {
    let ExprNode::MemberAccess { base, member, bang, .. } = p.node(id) else { panic!("expected member access") };
    (*base, p.name(*member).to_string(), *bang)
}

fn with_member(p: &Parsed, id: NodeId, want: &str) {
    let (base, name, bang) = member(p, id);
    assert!(matches!(p.node(base), ExprNode::WithContext));
    assert_eq!((name.as_str(), bang), (want, false));
}

fn call_args(p: &Parsed, id: NodeId) -> (NodeId, Vec<NodeId>) {
    let ExprNode::CallStmt { callee, args } = p.node(id) else { panic!("expected a call statement") };
    let ExprNode::ArgList { args } = p.node(*args) else { panic!("expected an argument list") };
    (*callee, args.clone())
}

fn body_of(src: &str) -> (Parsed, Vec<NodeId>) {
    let p = parse(&format!("Sub S()\n{src}\nEnd Sub\n"));
    let body = p.proc_body(0);
    (p, body)
}

#[test]
fn a_dot_after_blanks_starts_a_with_member_argument() {
    let (p, body) = body_of("    o.Foo .x0");
    assert_eq!(p.found, vec![]);
    assert_eq!(body.len(), 1);
    let (callee, args) = call_args(&p, body[0]);
    let (base, name, bang) = member(&p, callee);
    assert_eq!((name.as_str(), bang, name_of(&p, base).as_str()), ("Foo", false, "o"));
    assert_eq!(args.len(), 1);
    with_member(&p, args[0], "x0");
}

#[test]
fn operators_after_a_with_member_argument_stay_inside_that_argument() {
    let (p, body) = body_of("    o.Foo .x0 & \", \" & .y0");
    assert_eq!(p.found, vec![]);
    let (_, args) = call_args(&p, body[0]);
    assert_eq!(args.len(), 1);
    let ExprNode::BinOp { op: BinOpKind::Cat, lhs, rhs } = p.node(args[0]) else { panic!("expected &") };
    with_member(&p, *rhs, "y0");
    let ExprNode::BinOp { op: BinOpKind::Cat, lhs: first, rhs: sep } = p.node(*lhs) else { panic!("expected &") };
    with_member(&p, *first, "x0");
    assert!(matches!(p.node(*sep), ExprNode::Literal { lit: AstLit::Str(s) } if &**s == ", "));
}

#[test]
fn with_member_arguments_separated_by_commas() {
    let (p, body) = body_of("    o.Foo 1, .y0");
    assert_eq!(p.found, vec![]);
    let (_, args) = call_args(&p, body[0]);
    assert_eq!(args.len(), 2);
    assert!(matches!(p.node(args[0]), ExprNode::Literal { lit: AstLit::Int(1) }));
    with_member(&p, args[1], "y0");
}

#[test]
fn a_name_followed_by_a_dot_member_is_a_call_with_that_argument() {
    let (p, body) = body_of("    o .Foo");
    assert_eq!(p.found, vec![]);
    let (callee, args) = call_args(&p, body[0]);
    assert_eq!(name_of(&p, callee), "o");
    assert_eq!(args.len(), 1);
    with_member(&p, args[0], "Foo");
}

#[test]
fn a_comparison_inside_the_with_member_argument() {
    let (p, body) = body_of("    o.Item .Name = 1");
    assert_eq!(p.found, vec![]);
    let (callee, args) = call_args(&p, body[0]);
    assert_eq!(member(&p, callee).1, "Item");
    assert_eq!(args.len(), 1);
    let ExprNode::BinOp { op: BinOpKind::Eq, lhs, rhs } = p.node(args[0]) else { panic!("expected =") };
    with_member(&p, *lhs, "Name");
    assert!(matches!(p.node(*rhs), ExprNode::Literal { lit: AstLit::Int(1) }));
}

#[test]
fn a_bang_after_blanks_starts_a_with_key_argument() {
    let (p, body) = body_of("    o.Foo !Key");
    assert_eq!(p.found, vec![]);
    let (_, args) = call_args(&p, body[0]);
    assert_eq!(args.len(), 1);
    let (base, name, bang) = member(&p, args[0]);
    assert!(matches!(p.node(base), ExprNode::WithContext));
    assert_eq!((name.as_str(), bang), ("Key", true));
}

#[test]
fn adjacent_dot_and_bang_still_chain() {
    let (p, body) = body_of("    x = o.Item!Name");
    assert_eq!(p.found, vec![]);
    let ExprNode::Assign { value, .. } = p.node(body[0]) else { panic!("assignment") };
    let (item, name, bang) = member(&p, *value);
    assert_eq!((name.as_str(), bang), ("Name", true));
    let (o, item_name, bang) = member(&p, item);
    assert_eq!((item_name.as_str(), bang, name_of(&p, o).as_str()), ("Item", false, "o"));
}

#[test]
fn a_blank_after_the_dot_is_allowed() {
    let (p, body) = body_of("    x = o. Name");
    assert_eq!(p.found, vec![]);
    let ExprNode::Assign { value, .. } = p.node(body[0]) else { panic!("assignment") };
    let (base, name, bang) = member(&p, *value);
    assert_eq!((name.as_str(), bang, name_of(&p, base).as_str()), ("Name", false, "o"));
}

#[test]
fn a_continuation_line_starting_with_a_dot_continues_the_chain() {
    let (p, body) = body_of("    x = o _\n        .Name");
    assert_eq!(p.found, vec![]);
    let ExprNode::Assign { value, .. } = p.node(body[0]) else { panic!("assignment") };
    let (base, name, bang) = member(&p, *value);
    assert_eq!((name.as_str(), bang, name_of(&p, base).as_str()), ("Name", false, "o"));
}

#[test]
fn a_continuation_before_a_dot_after_a_call_continues_the_chain() {
    let (p, body) = body_of("    x = o.Foo(1) _\n        .Bar");
    assert_eq!(p.found, vec![]);
    let ExprNode::Assign { value, .. } = p.node(body[0]) else { panic!("assignment") };
    let (call, name, _) = member(&p, *value);
    assert_eq!(name, "Bar");
    assert!(matches!(p.node(call), ExprNode::Call { .. }));
}

#[test]
fn a_continued_call_statement_chains_instead_of_taking_an_argument() {
    let (p, body) = body_of("    o.Foo _\n        .x0");
    assert_eq!(p.found, vec![]);
    assert_eq!(body.len(), 1);
    let (base, name, _) = member(&p, body[0]);
    assert_eq!(name, "x0");
    assert_eq!(member(&p, base).1, "Foo");
}

fn assert_leftover_is_reported(src: &str, leftover: &str, name: &str) {
    let (p, body) = body_of(src);
    let full = format!("Sub S()\n{src}\nEnd Sub\n");
    assert_eq!(p.found, vec![(0x9c6f, offset_of(&full, leftover, 0), leftover.len() as u32)], "{src}");
    assert_eq!(body.len(), 1);
    let ExprNode::Assign { value, .. } = p.node(body[0]) else { panic!("assignment") };
    assert_eq!(name_of(&p, *value), name, "{src}");
}

#[test]
fn a_dot_after_blanks_in_an_expression_is_an_error() {
    assert_leftover_is_reported("    x = o .Name", ".", "o");
}

#[test]
fn a_dot_after_a_tab_in_an_expression_is_an_error() {
    assert_leftover_is_reported("    x = o\t.Name", ".", "o");
}

#[test]
fn a_bang_after_blanks_in_an_expression_is_an_error() {
    let src = "Sub S()\n    x = o.Item !Name\nEnd Sub\n";
    let p = parse(src);
    assert_eq!(p.found, vec![(0x9c6f, offset_of(src, "!", 0), 1)]);
    let body = p.proc_body(0);
    assert_eq!(body.len(), 1);
    let ExprNode::Assign { value, .. } = p.node(body[0]) else { panic!("assignment") };
    let (base, name, bang) = member(&p, *value);
    assert_eq!((name.as_str(), bang, name_of(&p, base).as_str()), ("Item", false, "o"));
}

#[test]
fn a_dot_after_blanks_following_a_call_or_a_group_is_an_error() {
    let (p, _) = body_of("    x = o.Item(1) .Name");
    assert_eq!(p.found.len(), 1);
    assert_eq!(p.found[0].0, 0x9c6f);
    let (p, _) = body_of("    x = (o) .Name");
    assert_eq!(p.found.len(), 1);
    assert_eq!(p.found[0].0, 0x9c6f);
}

#[test]
fn a_blank_between_a_bang_and_its_key_is_a_syntax_error() {
    let src = "Sub S()\n    x = a(1)! Name\nEnd Sub\n";
    let p = parse(src);
    assert_eq!(p.found, vec![(0x9c70, offset_of(src, "Name", 0), 4)]);
}

#[test]
fn nothing_joins_to_a_parenthesised_group() {
    for (src, at) in [("x = (o) .Name", ".Name"), ("x = (o).Name", ".Name"), ("x = (o)(1)", "(1)")] {
        let full = format!("Sub S()\n    {src}\nEnd Sub\n");
        let p = parse(&full);
        assert_eq!(p.found, vec![(0x9c6f, offset_of(&full, at, 0), 1)], "{src}");
    }
}
