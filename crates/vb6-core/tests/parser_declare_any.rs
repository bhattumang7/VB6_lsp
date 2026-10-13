//! `As Any` is accepted on `Declare` parameters (as an `AnyType` node, bound
//! to `VbaType::Any`) and nowhere else.

use vb6_core::frontend::ast::{ExprArena, ExprNode, NodeId};
use vb6_core::frontend::parser::Parser;
use vb6_core::frontend::scanner::ScannerContext;

fn parse(src: &str) -> (Vec<NodeId>, ExprArena, Vec<(u32, u32, u32)>) {
    let mut ctx = ScannerContext::new(1, 1, 0x0409);
    ctx.intern_keywords();
    let mut arena = ExprArena::new();
    let mut parser = Parser::new(&mut ctx, src.as_bytes());
    let nodes = parser.parse_module(&mut arena);
    let diags = parser.diagnostics.items().iter().map(|d| (d.code, d.span.start, d.span.len)).collect();
    (nodes, arena, diags)
}

/// `(flags, is AnyType)` of each parameter of the first top-level declaration.
fn params(arena: &ExprArena, node: NodeId) -> Vec<(u16, bool)> {
    let list = match arena.get(node) {
        ExprNode::DeclareDecl { params: Some(p), .. } | ExprNode::ProcDecl { params: Some(p), .. } => *p,
        _ => panic!("expected a declaration with parameters"),
    };
    let ExprNode::ArgList { args } = arena.get(list) else { panic!("expected ArgList") };
    args.iter()
        .map(|a| match arena.get(*a) {
            ExprNode::ParamDef { flags, type_node: Some(t), .. } => (*flags, matches!(arena.get(*t), ExprNode::AnyType)),
            _ => panic!("expected a typed ParamDef"),
        })
        .collect()
}

#[test]
fn declare_accepts_byval_and_byref_as_any() {
    let (nodes, arena, diags) =
        parse("Declare Sub CopyMem Lib \"kernel32\" Alias \"RtlMoveMemory\" (ByRef pDest As Any, ByVal pSrc As Any, ByVal nLen As Long)\n");
    assert_eq!(diags, vec![]);
    assert_eq!(nodes.len(), 1);
    assert_eq!(params(&arena, nodes[0]), vec![(0x4, true), (0x2, true), (0x2, false)]);
}

#[test]
fn as_any_on_a_procedure_parameter_is_rejected_at_any() {
    let src = "Sub Copy(ByVal pSrc As Any)\nEnd Sub\n";
    let (_, _, diags) = parse(src);
    let at = src.find("Any").unwrap() as u32;
    assert_eq!(diags.first().map(|d| (d.1, d.2)), Some((at, 3)), "first diagnostic must be at Any: {diags:?}");
}

#[test]
fn as_any_on_a_variable_is_rejected_at_any() {
    let src = "Dim m_Value As Any\n";
    let (_, _, diags) = parse(src);
    let at = src.find("Any").unwrap() as u32;
    assert_eq!(diags.first().map(|d| (d.1, d.2)), Some((at, 3)), "first diagnostic must be at Any: {diags:?}");
}
