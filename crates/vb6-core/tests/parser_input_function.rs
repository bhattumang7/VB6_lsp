//! The `Input`/`InputB` functions (and `$` forms) take an optional `#` before
//! the file number; with or without it the call is the same.

use vb6_core::frontend::ast::{ExprArena, ExprNode, NodeId};
use vb6_core::frontend::parser::Parser;
use vb6_core::frontend::scanner::ScannerContext;

/// (diagnostics, rendered value of `v = <expr>`).
fn rhs(expr: &str) -> (usize, String) {
    let src = format!("Sub T()\nv = {expr}\nEnd Sub\n");
    let mut ctx = ScannerContext::new(1, 1, 0x0409);
    ctx.intern_keywords();
    let mut arena = ExprArena::new();
    let mut parser = Parser::new(&mut ctx, src.as_bytes());
    let nodes = parser.parse_module(&mut arena);
    let diags = parser.diagnostics.items().len();
    drop(parser);
    let ExprNode::ProcDecl { body, .. } = arena.get(nodes[0]) else { panic!("expected ProcDecl") };
    let ExprNode::Block { stmts } = arena.get(*body) else { panic!("expected Block") };
    assert_eq!(stmts.len(), 1);
    let ExprNode::Assign { value, .. } = arena.get(stmts[0]) else { panic!("expected Assign") };
    (diags, render(&ctx, &arena, *value))
}

fn render(ctx: &ScannerContext, arena: &ExprArena, id: NodeId) -> String {
    match arena.get(id) {
        ExprNode::NameRef { sym, .. } => ctx.symbol(*sym as usize).name.clone(),
        ExprNode::Literal { lit } => format!("{lit:?}"),
        ExprNode::Call { func, args } => {
            let ExprNode::ArgList { args } = arena.get(*args) else { panic!("expected ArgList") };
            let a: Vec<_> = args.iter().map(|x| render(ctx, arena, *x)).collect();
            format!("(call {} [{}])", render(ctx, arena, *func), a.join(" "))
        }
        ExprNode::BinOp { lhs, rhs, .. } => format!("(op {} {})", render(ctx, arena, *lhs), render(ctx, arena, *rhs)),
        _ => panic!("unexpected node"),
    }
}

#[test]
fn input_with_hash_file_number() {
    assert_eq!(rhs("Input(LOF(nFile), #nFile)"), (0, "(call Input [(call LOF [nFile]) nFile])".into()));
}

#[test]
fn input_without_hash_is_the_same_call() {
    assert_eq!(rhs("Input(LOF(nFile), nFile)"), (0, "(call Input [(call LOF [nFile]) nFile])".into()));
}

#[test]
fn input_with_expression_arguments() {
    assert_eq!(rhs("Input(nEnd - nStart, #nFile)"), (0, "(call Input [(op nEnd nStart) nFile])".into()));
}

#[test]
fn input_dollar_and_inputb_forms() {
    assert_eq!(rhs("Input$(10, #nFile)").1, "(call Input$ [Int(10) nFile])");
    assert_eq!(rhs("InputB(10, #nFile)").1, "(call InputB [Int(10) nFile])");
    assert_eq!(rhs("InputB$(10, #nFile)").1, "(call InputB$ [Int(10) nFile])");
    assert_eq!(rhs("Input$(10, #nFile)").0, 0);
}

#[test]
fn input_call_nested_in_another_call() {
    assert_eq!(
        rhs("Split(Input(LOF(nFile), #nFile), sSep)"),
        (0, "(call Split [(call Input [(call LOF [nFile]) nFile]) sSep])".into())
    );
}

#[test]
fn hash_is_not_accepted_on_other_function_arguments() {
    let src = "Sub T()\nv = Len(#nFile)\nEnd Sub\n";
    let mut ctx = ScannerContext::new(1, 1, 0x0409);
    ctx.intern_keywords();
    let mut arena = ExprArena::new();
    let mut parser = Parser::new(&mut ctx, src.as_bytes());
    parser.parse_module(&mut arena);
    assert!(parser.diagnostics.has_errors(), "a # before an argument of any other function is an error");
}
