//! In a single-line `If`, a call statement without arguments ends at the
//! `Else` that follows it.

mod common;
use common::{parse, Parsed};
use vb6_syntax::frontend::ast::{AstLit, ExprNode, NodeId};

fn name_of(p: &Parsed, id: NodeId) -> String {
    let ExprNode::NameRef { sym, .. } = p.node(id) else { panic!("expected a name") };
    p.name(*sym).to_string()
}

fn member_of_me(p: &Parsed, id: NodeId) -> String {
    let ExprNode::MemberAccess { base, member, bang, .. } = p.node(id) else { panic!("expected a member") };
    assert!(matches!(p.node(*base), ExprNode::Me));
    assert!(!*bang);
    p.name(*member).to_string()
}

/// The two parts of the single-line `If` that is the only statement of the procedure.
fn parts(src: &str) -> (Parsed, NodeId, Vec<NodeId>, Vec<NodeId>) {
    let p = parse(&format!("Sub S()\n{src}\nEnd Sub\n"));
    assert_eq!(p.found, vec![], "{src}");
    let body = p.proc_body(0);
    assert_eq!(body.len(), 1);
    let ExprNode::If { cond, then_body, else_body } = p.node(body[0]) else { panic!("expected If") };
    let (cond, then_stmts, else_stmts) = (*cond, p.block(*then_body), p.block(else_body.expect("an Else part")));
    (p, cond, then_stmts, else_stmts)
}

#[test]
fn a_bare_call_then_else_with_an_assignment() {
    let (p, cond, then_stmts, else_stmts) = parts("If c Then Probe Else x = False");
    assert_eq!(name_of(&p, cond), "c");
    assert_eq!(then_stmts.len(), 1);
    assert_eq!(name_of(&p, then_stmts[0]), "Probe");
    assert_eq!(else_stmts.len(), 1);
    let ExprNode::Assign { target, value } = p.node(else_stmts[0]) else { panic!("assignment") };
    assert_eq!(name_of(&p, *target), "x");
    assert!(matches!(p.node(*value), ExprNode::Literal { lit: AstLit::Bool(false) }));
}

#[test]
fn bare_member_calls_on_both_sides() {
    let (p, _, then_stmts, else_stmts) = parts("If c Then Me.LockIn Else Me.LockRelease");
    assert_eq!(member_of_me(&p, then_stmts[0]), "LockIn");
    assert_eq!(member_of_me(&p, else_stmts[0]), "LockRelease");
}

#[test]
fn a_bare_call_then_else_with_a_call_with_arguments() {
    let (p, _, then_stmts, else_stmts) = parts("If c Then Probe Else Probe 1, 2");
    assert_eq!(name_of(&p, then_stmts[0]), "Probe");
    let ExprNode::CallStmt { callee, args } = p.node(else_stmts[0]) else { panic!("call") };
    assert_eq!(name_of(&p, *callee), "Probe");
    let ExprNode::ArgList { args } = p.node(*args) else { panic!("args") };
    assert_eq!(args.len(), 2);
    assert!(matches!(p.node(args[0]), ExprNode::Literal { lit: AstLit::Int(1) }));
    assert!(matches!(p.node(args[1]), ExprNode::Literal { lit: AstLit::Int(2) }));
}

#[test]
fn a_call_with_arguments_then_a_bare_call_after_else() {
    let (p, _, then_stmts, else_stmts) = parts("If c Then Probe 1 Else Probe");
    let ExprNode::CallStmt { args, .. } = p.node(then_stmts[0]) else { panic!("call") };
    let ExprNode::ArgList { args } = p.node(*args) else { panic!("args") };
    assert_eq!(args.len(), 1);
    assert_eq!(name_of(&p, else_stmts[0]), "Probe");
}

#[test]
fn a_bare_call_before_else_with_a_colon_separated_then_part() {
    let (p, _, then_stmts, else_stmts) = parts("If c Then a = 1: Probe Else Probe");
    assert_eq!(then_stmts.len(), 2);
    assert!(matches!(p.node(then_stmts[0]), ExprNode::Assign { .. }));
    assert_eq!(name_of(&p, then_stmts[1]), "Probe");
    assert_eq!(name_of(&p, else_stmts[0]), "Probe");
}

#[test]
fn a_qualified_call_with_a_block_argument_after_else() {
    let (p, _, then_stmts, else_stmts) = parts("If c Then a.B Else a.C d");
    assert!(matches!(p.node(then_stmts[0]), ExprNode::MemberAccess { .. }));
    let ExprNode::CallStmt { callee, args } = p.node(else_stmts[0]) else { panic!("call") };
    assert!(matches!(p.node(*callee), ExprNode::MemberAccess { .. }));
    let ExprNode::ArgList { args } = p.node(*args) else { panic!("args") };
    assert_eq!(args.len(), 1);
    assert_eq!(name_of(&p, args[0]), "d");
}
