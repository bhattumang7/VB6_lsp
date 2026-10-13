//! A parenthesised expression takes member access and the calls that follow it
//! like any object value: `(New Foo).Bar`, `(o).M(1)`, `(a)!Key`. A group
//! that is not followed by `.` or `!` is unchanged.

use vb6_core::frontend::ast::{ExprArena, ExprNode, NodeId};
use vb6_core::frontend::parser::{ModuleKind, Parser};
use vb6_core::frontend::scanner::ScannerContext;

/// (diagnostics, the body statements of `Sub T` rendered).
fn body(stmts_src: &str) -> (usize, Vec<String>) {
    let src = format!("Sub T()\n{stmts_src}\nEnd Sub\n");
    let mut ctx = ScannerContext::new(1, 1, 0x0409);
    ctx.intern_keywords();
    let mut arena = ExprArena::new();
    let mut parser = Parser::with_module_kind(&mut ctx, src.as_bytes(), ModuleKind::Standard);
    let nodes = parser.parse_module(&mut arena);
    let diags = parser.diagnostics.items().len();
    drop(parser);
    let ExprNode::ProcDecl { body, .. } = arena.get(nodes[0]) else { panic!("expected ProcDecl") };
    let ExprNode::Block { stmts } = arena.get(*body) else { panic!("expected Block") };
    (diags, stmts.iter().map(|s| render(&ctx, &arena, *s)).collect())
}

fn render(ctx: &ScannerContext, arena: &ExprArena, id: NodeId) -> String {
    let name = |s: u32| ctx.symbol(s as usize).name.clone();
    let list = |a: NodeId| match arena.get(a) {
        ExprNode::ArgList { args } => args.iter().map(|x| render(ctx, arena, *x)).collect::<Vec<_>>().join(" "),
        _ => panic!("expected ArgList"),
    };
    match arena.get(id) {
        ExprNode::NameRef { sym, .. } => name(*sym),
        ExprNode::Literal { lit } => format!("{lit:?}"),
        ExprNode::Paren { inner } => format!("(paren {})", render(ctx, arena, *inner)),
        ExprNode::New { type_spec } => format!("(new {})", render(ctx, arena, *type_spec)),
        ExprNode::UserType { name: n, child } => match child {
            Some(c) => format!("{}.{}", name(*n), render(ctx, arena, *c)),
            None => name(*n),
        },
        ExprNode::BinOp { op, lhs, rhs } => {
            format!("({op:?} {} {})", render(ctx, arena, *lhs), render(ctx, arena, *rhs))
        }
        ExprNode::MemberAccess { base, member, bang } => {
            format!("({} {} {})", if *bang { "!" } else { "." }, render(ctx, arena, *base), name(*member))
        }
        ExprNode::Call { func, args } => format!("(call {} [{}])", render(ctx, arena, *func), list(*args)),
        ExprNode::CallStmt { callee, args } => format!("(stmt {} [{}])", render(ctx, arena, *callee), list(*args)),
        ExprNode::Assign { target, value } => format!("(= {} {})", render(ctx, arena, *target), render(ctx, arena, *value)),
        ExprNode::SetAssign { target, value } => {
            format!("(set {} {})", render(ctx, arena, *target), render(ctx, arena, *value))
        }
        _ => panic!("unexpected node"),
    }
}

#[test]
fn new_in_parentheses_takes_a_member_statement() {
    assert_eq!(body("(New clsTrip).Save"), (0, vec!["(. (paren (new clsTrip)) Save)".to_string()]));
}

#[test]
fn new_in_parentheses_takes_a_member_call_with_arguments() {
    assert_eq!(
        body("(New clsTrip).Save 1, \"a\""),
        (0, vec!["(stmt (. (paren (new clsTrip)) Save) [Int(1) Str(\"a\")])".to_string()])
    );
}

#[test]
fn new_in_parentheses_takes_a_member_in_an_expression() {
    assert_eq!(
        body("x = (New clsTrip).Save"),
        (0, vec!["(= x (. (paren (new clsTrip)) Save))".to_string()])
    );
}

#[test]
fn qualified_new_in_parentheses() {
    assert_eq!(
        body("x = (New Proj.clsTrip).Save"),
        (0, vec!["(= x (. (paren (new Proj.clsTrip)) Save))".to_string()])
    );
}

#[test]
fn group_member_followed_by_a_call() {
    assert_eq!(
        body("x = (o).M(1)"),
        (0, vec!["(= x (call (. (paren o) M) [Int(1)]))".to_string()])
    );
}

#[test]
fn group_member_chain() {
    assert_eq!(
        body("x = (o).Child.Name"),
        (0, vec!["(= x (. (. (paren o) Child) Name))".to_string()])
    );
}

#[test]
fn group_bang_member() {
    assert_eq!(body("x = (a)!Key"), (0, vec!["(= x (! (paren a) Key))".to_string()]));
}

#[test]
fn set_target_value_takes_a_member() {
    assert_eq!(
        body("Set o = (New clsTrip).Child"),
        (0, vec!["(set o (. (paren (new clsTrip)) Child))".to_string()])
    );
}

#[test]
fn plain_group_is_unchanged() {
    assert_eq!(body("x = (a)"), (0, vec!["(= x (paren a))".to_string()]));
}

#[test]
fn group_in_arithmetic_is_unchanged() {
    assert_eq!(
        body("x = (p) + (q)"),
        (0, vec!["(= x (Add (paren p) (paren q)))".to_string()])
    );
}
