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

// ── An element's member array: `ReDim a(i).b(n)` ─────────────────────────────

#[test]
fn redim_of_an_array_elements_member_array() {
    use vb6_core::frontend::ast::{ExprArena, ExprNode};
    use vb6_core::frontend::parser::{ModuleKind, Parser};
    use vb6_core::frontend::scanner::ScannerContext;
    let src = "Sub T()\n    ReDim Preserve a(i).Counters(0)\n    ReDim a(1).b(2).c(3, 4)\nEnd Sub\n";
    let mut ctx = ScannerContext::new(1, 1, 0x0409);
    ctx.intern_keywords();
    let mut arena = ExprArena::new();
    let mut parser = Parser::with_module_kind(&mut ctx, src.as_bytes(), ModuleKind::Standard);
    let top = parser.parse_module(&mut arena);
    assert_eq!(parser.diagnostics.items().len(), 0);
    drop(parser);
    let name = |s: u32| ctx.symbol(s as usize).name.clone();
    let ExprNode::ProcDecl { body, .. } = arena.get(top[0]) else { panic!("ProcDecl") };
    let ExprNode::Block { stmts } = arena.get(*body) else { panic!("Block") };
    // First statement: base is the call `a(i)`, member array `Counters` with one bound.
    let ExprNode::Block { stmts: first } = arena.get(stmts[0]) else { panic!("Block") };
    let ExprNode::ReDimItem { preserve, base, name: n, bounds, .. } = arena.get(first[0]) else { panic!("item") };
    assert!(*preserve);
    assert_eq!(name(*n), "Counters");
    let ExprNode::Call { func, args } = arena.get(base.unwrap()) else { panic!("call") };
    assert!(matches!(arena.get(*func), ExprNode::NameRef { sym, .. } if name(*sym).eq_ignore_ascii_case("a")));
    let ExprNode::ArgList { args } = arena.get(*args) else { panic!("args") };
    assert_eq!(args.len(), 1);
    let ExprNode::ArgList { args: b } = arena.get(bounds.unwrap()) else { panic!("bounds") };
    assert_eq!(b.len(), 1);
    // Second: `a(1).b(2).c(3, 4)` - two indexed links, two bounds.
    let ExprNode::Block { stmts: second } = arena.get(stmts[1]) else { panic!("Block") };
    let ExprNode::ReDimItem { preserve, base, name: n, bounds, .. } = arena.get(second[0]) else { panic!("item") };
    assert!(!*preserve);
    assert!(name(*n).eq_ignore_ascii_case("c"));
    let ExprNode::Call { func, .. } = arena.get(base.unwrap()) else { panic!("call") };
    assert!(matches!(arena.get(*func), ExprNode::MemberAccess { member, .. } if name(*member).eq_ignore_ascii_case("b")));
    let ExprNode::ArgList { args: b } = arena.get(bounds.unwrap()) else { panic!("bounds") };
    assert_eq!(b.len(), 2);
}
