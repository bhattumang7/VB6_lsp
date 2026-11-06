//! A parenthesised expression is a value on its own: a member access or bang after the
//! closing bracket (`(New Foo).Bar`, `(o).M(1)`, `(a)!Key`) is a syntax error, and no
//! statement starts with a group.

use vb6_core::frontend::ast::{ExprArena, ExprNode, NodeId};
use vb6_core::frontend::parser::{ModuleKind, Parser};
use vb6_core::frontend::scanner::ScannerContext;

/// The first diagnostic as (code, start, len), and the body statements of `Sub T` rendered.
fn body_with_first_error(stmts_src: &str) -> (Option<(u32, u32, u32)>, Vec<String>) {
    let src = format!("Sub T()
{stmts_src}
End Sub
");
    let mut ctx = ScannerContext::new(1, 1, 0x0409);
    ctx.intern_keywords();
    let mut arena = ExprArena::new();
    let mut parser = Parser::with_module_kind(&mut ctx, src.as_bytes(), ModuleKind::Standard);
    let nodes = parser.parse_module(&mut arena);
    let first = parser.diagnostics.items().first().map(|d| (d.code, d.span.start, d.span.len));
    drop(parser);
    let ExprNode::ProcDecl { body, .. } = arena.get(nodes[0]) else { panic!("expected ProcDecl") };
    let ExprNode::Block { stmts } = arena.get(*body) else { panic!("expected Block") };
    (first, stmts.iter().map(|s| render(&ctx, &arena, *s)).collect())
}

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
        ExprNode::MemberAccess { base, member, bang, .. } => {
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
fn a_group_cannot_take_a_member() {
    // (statement, text whose offset within the statement is where the error is reported)
    for (stmt, at) in [
        ("x = (o).M(1)", ".M"),
        ("x = (o).Child.Name", ".Child"),
        ("x = (a)!Key", "!Key"),
        ("x = (New clsTrip).Save", ".Save"),
        ("x = (New Proj.clsTrip).Save", ".Save"),
        ("Set o = (New clsTrip).Child", ".Child"),
    ] {
        let (found, rendered) = body_with_first_error(stmt);
        let src = format!("Sub T()\n{stmt}\nEnd Sub\n");
        let want = src.find(at).unwrap() as u32;
        assert_eq!(found.map(|f| f.1), Some(want), "{stmt}");
        assert_eq!(rendered.len(), 1, "{stmt}");
    }
}

#[test]
fn a_group_keeps_its_value_when_a_member_follows() {
    assert_eq!(body("x = (o).M(1)"), (1, vec!["(= x (paren o))".to_string()]));
    assert_eq!(body("x = (a)!Key"), (1, vec!["(= x (paren a))".to_string()]));
    assert_eq!(body("x = (New clsTrip).Save"), (1, vec!["(= x (paren (new clsTrip)))".to_string()]));
    assert_eq!(body("Set o = (New clsTrip).Child"), (1, vec!["(set o (paren (new clsTrip)))".to_string()]));
}

#[test]
fn a_statement_cannot_start_with_a_group() {
    for stmt in ["(New clsTrip).Save", "(New clsTrip).Save 1, \"a\""] {
        let (found, rendered) = body_with_first_error(stmt);
        assert_eq!(found.map(|f| f.1), Some("Sub T()\n".len() as u32), "{stmt}");
        assert_eq!(rendered, Vec::<String>::new(), "{stmt}");
    }
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
