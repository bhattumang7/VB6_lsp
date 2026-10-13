//! `WithEvents` on variable declarations: recorded per variable on the
//! `DimItem`, accepted on every declaration form, rejected only in standard
//! modules.

use vb6_core::frontend::ast::{ExprArena, ExprNode, NodeId};
use vb6_core::frontend::parser::{ModuleKind, Parser};
use vb6_core::frontend::scanner::ScannerContext;

/// "Only valid in object module".
const ERR_OBJECT_MODULE_ONLY: u32 = 0xdee1;

struct Parsed {
    ctx: ScannerContext,
    nodes: Vec<NodeId>,
    arena: ExprArena,
    diags: Vec<(u32, u32, u32)>,
}

fn parse(src: &str, kind: ModuleKind) -> Parsed {
    let mut ctx = ScannerContext::new(1, 1, 0x0409);
    ctx.intern_keywords();
    let mut arena = ExprArena::new();
    let mut parser = Parser::with_module_kind(&mut ctx, src.as_bytes(), kind);
    let nodes = parser.parse_module(&mut arena);
    let diags = parser.diagnostics.items().iter().map(|d| (d.code, d.span.start, d.span.len)).collect();
    drop(parser);
    Parsed { ctx, nodes, arena, diags }
}

/// The declaration a `DimItem` describes, with names spelled out.
#[derive(Debug, PartialEq)]
struct Item {
    name: String,
    is_const: bool,
    is_static: bool,
    is_new: bool,
    with_events: bool,
    has_bounds: bool,
    user_type: Option<String>,
    has_init: bool,
}

/// Every `DimItem` of the single top-level declaration statement.
fn items(p: &Parsed) -> Vec<Item> {
    assert_eq!(p.nodes.len(), 1, "one declaration statement");
    let ExprNode::Block { stmts } = p.arena.get(p.nodes[0]) else {
        panic!("declaration statement must be a Block");
    };
    let name = |sym: u32| p.ctx.symbol(sym as usize).name.clone();
    stmts
        .iter()
        .map(|id| match p.arena.get(*id) {
            ExprNode::DimItem { name: n, is_const, is_static, is_new, with_events, bounds, type_node, init } => Item {
                name: name(*n),
                is_const: *is_const,
                is_static: *is_static,
                is_new: *is_new,
                with_events: *with_events,
                has_bounds: bounds.is_some(),
                user_type: type_node.map(|t| match p.arena.get(t) {
                    ExprNode::UserType { name: tn, child: None } => name(*tn),
                    _ => panic!("expected a plain UserType"),
                }),
                has_init: init.is_some(),
            },
            _ => panic!("expected DimItem"),
        })
        .collect()
}

fn var(name: &str, with_events: bool, ty: &str) -> Item {
    Item {
        name: name.into(),
        is_const: false,
        is_static: false,
        is_new: false,
        with_events,
        has_bounds: false,
        user_type: Some(ty.into()),
        has_init: false,
    }
}

#[test]
fn private_with_events_marks_the_variable() {
    let p = parse("Private WithEvents m_Clock As Clock\n", ModuleKind::Class);
    assert_eq!(p.diags, vec![]);
    assert_eq!(items(&p), vec![var("m_Clock", true, "Clock")]);
}

#[test]
fn dim_with_events_marks_the_variable() {
    let p = parse("Dim WithEvents m_Clock As Clock\n", ModuleKind::Class);
    assert_eq!(p.diags, vec![]);
    assert_eq!(items(&p), vec![var("m_Clock", true, "Clock")]);
}

#[test]
fn public_with_events_in_form_marks_the_variable() {
    let p = parse("Public WithEvents m_Clock As Clock\n", ModuleKind::Form);
    assert_eq!(p.diags, vec![]);
    assert_eq!(items(&p), vec![var("m_Clock", true, "Clock")]);
}

#[test]
fn with_events_applies_per_variable_in_a_list() {
    let p = parse("Private WithEvents m_First As Clock, m_Second As Clock, WithEvents m_Third As Timer\n", ModuleKind::Class);
    assert_eq!(p.diags, vec![]);
    assert_eq!(items(&p), vec![var("m_First", true, "Clock"), var("m_Second", false, "Clock"), var("m_Third", true, "Timer")]);
}

#[test]
fn plain_declaration_is_not_with_events() {
    let p = parse("Private m_Clock As Clock\n", ModuleKind::Class);
    assert_eq!(p.diags, vec![]);
    assert_eq!(items(&p), vec![var("m_Clock", false, "Clock")]);
}

#[test]
fn with_events_in_standard_module_is_rejected_and_still_recorded() {
    let src = "Private WithEvents m_First As Clock, WithEvents m_Second As Clock\n";
    let p = parse(src, ModuleKind::Standard);
    let first = src.find("WithEvents").unwrap() as u32;
    let second = src.rfind("WithEvents").unwrap() as u32;
    let len = "WithEvents".len() as u32;
    assert_eq!(p.diags, vec![(ERR_OBJECT_MODULE_ONLY, first, len), (ERR_OBJECT_MODULE_ONLY, second, len)]);
    assert_eq!(items(&p), vec![var("m_First", true, "Clock"), var("m_Second", true, "Clock")]);
}
