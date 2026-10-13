//! `Is` is an ordinary comparison operator: same precedence as `=`, `<>`,
//! `Like`, left-associative, above `Not`/`And`/`Or`. There is no `Is Not`
//! operator; `a Is Not b` is `a Is (Not b)`.

use vb6_core::frontend::ast::{BinOpKind, ExprArena, ExprNode, NodeId, UnOpKind};
use vb6_core::frontend::parser::Parser;
use vb6_core::frontend::scanner::ScannerContext;

/// Parse `v = <expr>` inside a Sub; return the diagnostics count and the
/// value expression rendered as an S-expression.
fn rhs(expr: &str) -> (usize, String) {
    let src = format!("Sub T()\nv = {expr}\nEnd Sub\n");
    let mut ctx = ScannerContext::new(1, 1, 0x0409);
    ctx.intern_keywords();
    let mut arena = ExprArena::new();
    let mut parser = Parser::new(&mut ctx, src.as_bytes());
    let nodes = parser.parse_module(&mut arena);
    let errors = parser.diagnostics.items().len();
    drop(parser);
    let ExprNode::ProcDecl { body, .. } = arena.get(nodes[0]) else { panic!("expected ProcDecl") };
    let ExprNode::Block { stmts } = arena.get(*body) else { panic!("expected Block body") };
    assert_eq!(stmts.len(), 1, "one statement");
    let ExprNode::Assign { value, .. } = arena.get(stmts[0]) else { panic!("expected Assign") };
    (errors, render(&ctx, &arena, *value))
}

fn render(ctx: &ScannerContext, arena: &ExprArena, id: NodeId) -> String {
    match arena.get(id) {
        ExprNode::NameRef { sym, .. } => ctx.symbol(*sym as usize).name.clone(),
        ExprNode::Nothing => "Nothing".into(),
        ExprNode::Literal { lit } => format!("{lit:?}"),
        ExprNode::Paren { inner } => format!("(paren {})", render(ctx, arena, *inner)),
        ExprNode::UnOp { op, operand } => {
            let o = match op { UnOpKind::Not => "Not", UnOpKind::Neg => "Neg", UnOpKind::Pos => "Pos" };
            format!("({o} {})", render(ctx, arena, *operand))
        }
        ExprNode::BinOp { op, lhs, rhs } => {
            let o = match op {
                BinOpKind::Is => "Is",
                BinOpKind::Eq => "=",
                BinOpKind::And => "And",
                BinOpKind::Or => "Or",
                BinOpKind::Like => "Like",
                other => panic!("unexpected operator {other:?}"),
            };
            format!("({o} {} {})", render(ctx, arena, *lhs), render(ctx, arena, *rhs))
        }
        _ => panic!("unexpected node in expression"),
    }
}

#[test]
fn not_applies_to_the_whole_is_comparison() {
    assert_eq!(rhs("Not m_Obj Is Nothing"), (0, "(Not (Is m_Obj Nothing))".into()));
}

#[test]
fn is_inside_or_right_operand() {
    assert_eq!(rhs("m_Count = 0 Or m_Obj Is Nothing"), (0, "(Or (= m_Count Int(0)) (Is m_Obj Nothing))".into()));
}

#[test]
fn is_inside_and_operands_with_not() {
    assert_eq!(
        rhs("Not m_Left Is Nothing And m_Right Is Nothing"),
        (0, "(And (Not (Is m_Left Nothing)) (Is m_Right Nothing))".into())
    );
}

#[test]
fn is_is_left_associative_with_other_comparisons() {
    assert_eq!(rhs("m_Left Is m_Right = m_Flag"), (0, "(= (Is m_Left m_Right) m_Flag)".into()));
    assert_eq!(rhs("m_Flag = m_Left Is m_Right"), (0, "(Is (= m_Flag m_Left) m_Right)".into()));
}

#[test]
fn is_not_is_is_applied_to_a_not_operand() {
    assert_eq!(rhs("m_Left Is Not m_Right"), (0, "(Is m_Left (Not m_Right))".into()));
}

#[test]
fn parenthesised_is_under_not() {
    assert_eq!(rhs("Not (m_Obj Is Nothing)"), (0, "(Not (paren (Is m_Obj Nothing)))".into()));
}
