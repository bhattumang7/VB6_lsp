//! `ReDim` of a member array (`ReDim [Preserve] udt.Field(bounds)`): the
//! `ReDimItem` records the base it belongs to and the member name.

use vb6_core::frontend::ast::{ExprArena, ExprNode, NodeId};
use vb6_core::frontend::parser::Parser;
use vb6_core::frontend::scanner::ScannerContext;
use vb6_core::session::Session;

/// (diagnostics, rendered ReDim items of the single statement in `Sub T`).
fn redim(stmt: &str) -> (usize, Vec<String>) {
    let src = format!("Sub T()\n{stmt}\nEnd Sub\n");
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
    let ExprNode::Block { stmts: items } = arena.get(stmts[0]) else { panic!("expected the ReDim Block") };
    (diags, items.iter().map(|i| render(&ctx, &arena, *i)).collect())
}

fn render(ctx: &ScannerContext, arena: &ExprArena, id: NodeId) -> String {
    let name = |s: u32| ctx.symbol(s as usize).name.clone();
    match arena.get(id) {
        ExprNode::ReDimItem { preserve, base, name: n, bounds, type_node } => format!(
            "(redim{} {} {} [{}] {})",
            if *preserve { "-preserve" } else { "" },
            base.map(|b| render(ctx, arena, b)).unwrap_or_else(|| "-".into()),
            name(*n),
            bounds.map(|b| render(ctx, arena, b)).unwrap_or_default(),
            type_node.map(|t| render(ctx, arena, t)).unwrap_or_else(|| "-".into()),
        ),
        ExprNode::ArgList { args } => args.iter().map(|a| render(ctx, arena, *a)).collect::<Vec<_>>().join(" "),
        ExprNode::NameRef { sym, .. } => name(*sym),
        ExprNode::MemberAccess { base, member, .. } => format!("(. {} {})", render(ctx, arena, *base), name(*member)),
        ExprNode::RangeTo { lo, hi } => format!("({} To {})", render(ctx, arena, *lo), render(ctx, arena, *hi)),
        ExprNode::Literal { lit } => format!("{lit:?}"),
        ExprNode::BuiltinType { kind } => format!("type{kind}"),
        _ => panic!("unexpected node"),
    }
}

#[test]
fn preserve_member_array_of_a_variable() {
    assert_eq!(redim("ReDim Preserve oAction.ID(nCount)"), (0, vec!["(redim-preserve oAction ID [nCount] -)".into()]));
}

#[test]
fn member_array_through_a_chain_with_bounds_and_type() {
    assert_eq!(
        redim("ReDim m_Data.Route.Stops(1 To nLast) As Long"),
        (0, vec!["(redim (. m_Data Route) Stops [(Int(1) To nLast)] type3)".into()])
    );
}

#[test]
fn plain_array_has_no_base() {
    assert_eq!(redim("ReDim aValues(10)"), (0, vec!["(redim - aValues [Int(10)] -)".into()]));
}

#[test]
fn list_mixing_plain_and_member_arrays() {
    assert_eq!(
        redim("ReDim aValues(10), oAction.Counters(nCount)"),
        (0, vec!["(redim - aValues [Int(10)] -)".into(), "(redim oAction Counters [nCount] -)".into()])
    );
}

#[test]
fn the_base_variable_is_a_reference() {
    let src = "Private oAction As tAction\nSub T()\n    ReDim Preserve oAction.ID(3)\nEnd Sub\n";
    let s = Session::from_sources(vec![("M.bas".to_string(), src.as_bytes().to_vec())]);
    let decl = src.find("oAction").unwrap() as u32;
    let sites: Vec<u32> = s.references(0, decl, false).iter().map(|l| l.span.start).collect();
    assert_eq!(sites, vec![src.rfind("oAction").unwrap() as u32]);
}
