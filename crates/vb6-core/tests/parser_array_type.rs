//! `As T()` is an array of `T`: the type-spec is an `ArrayType` around the
//! element type, and binds to an array type.

use vb6_core::frontend::ast::{ExprArena, ExprNode, NodeId};
use vb6_core::frontend::parser::{ModuleKind, Parser};
use vb6_core::frontend::scanner::ScannerContext;
use vb6_core::sema::{bind, VbaType};

fn parse(src: &str) -> (ScannerContext, ExprArena, Vec<NodeId>, usize) {
    let mut ctx = ScannerContext::new(1, 1, 0x0409);
    ctx.intern_keywords();
    let mut arena = ExprArena::new();
    let mut parser = Parser::with_module_kind(&mut ctx, src.as_bytes(), ModuleKind::Class);
    let top = parser.parse_module(&mut arena);
    let diags = parser.diagnostics.items().len();
    drop(parser);
    (ctx, arena, top, diags)
}

#[test]
fn function_returning_an_array_of_a_builtin() {
    let (_ctx, arena, top, diags) = parse("Function F() As Long()\nEnd Function\n");
    assert_eq!(diags, 0);
    let ExprNode::ProcDecl { ret_type, .. } = arena.get(top[0]) else { panic!("ProcDecl") };
    let ExprNode::ArrayType { elem } = arena.get(ret_type.unwrap()) else { panic!("ArrayType") };
    assert!(matches!(arena.get(*elem), ExprNode::BuiltinType { kind: 3 }));
}

#[test]
fn function_returning_an_array_of_a_class() {
    let (ctx, arena, top, diags) = parse("Function F() As clsTrip()\nEnd Function\n");
    assert_eq!(diags, 0);
    let ExprNode::ProcDecl { ret_type, .. } = arena.get(top[0]) else { panic!("ProcDecl") };
    let ExprNode::ArrayType { elem } = arena.get(ret_type.unwrap()) else { panic!("ArrayType") };
    let ExprNode::UserType { name, child } = arena.get(*elem) else { panic!("UserType") };
    assert_eq!(ctx.symbol(*name as usize).name, "clsTrip");
    assert!(child.is_none());
}

#[test]
fn array_return_type_binds_to_an_array() {
    let mut ctx = ScannerContext::new(1, 1, 0x0409);
    ctx.intern_keywords();
    let mut arena = ExprArena::new();
    let src = "Function F() As Long()\nEnd Function\nFunction G() As Long\nEnd Function\n";
    let mut parser = Parser::with_module_kind(&mut ctx, src.as_bytes(), ModuleKind::Class);
    let top = parser.parse_module(&mut arena);
    let spans = std::mem::take(&mut parser.node_spans);
    let vis = std::mem::take(&mut parser.decl_public);
    drop(parser);
    let m = bind(&ctx, &arena, &top, &spans, &vis);
    assert_eq!(m.procs[0].ret_type, VbaType::Array(Box::new(VbaType::Long)));
    assert_eq!(m.procs[1].ret_type, VbaType::Long);
}
