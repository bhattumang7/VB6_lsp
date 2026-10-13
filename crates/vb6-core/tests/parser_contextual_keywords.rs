//! Contextual (non-reserved) keywords such as `Name`, `Text`, `Access`,
//! `Error` and `Line` can name variables and types; reserved words cannot.

use vb6_core::frontend::ast::{ExprArena, ExprNode, NodeId, ProcKind};
use vb6_core::frontend::parser::{ModuleKind, Parser};
use vb6_core::frontend::scanner::ScannerContext;

struct Parsed {
    ctx: ScannerContext,
    nodes: Vec<NodeId>,
    arena: ExprArena,
    diags: Vec<(u32, u32, u32)>,
}

fn parse(src: &str) -> Parsed {
    let mut ctx = ScannerContext::new(1, 1, 0x0409);
    ctx.intern_keywords();
    let mut arena = ExprArena::new();
    let mut parser = Parser::with_module_kind(&mut ctx, src.as_bytes(), ModuleKind::Class);
    let nodes = parser.parse_module(&mut arena);
    let diags = parser.diagnostics.items().iter().map(|d| (d.code, d.span.start, d.span.len)).collect();
    drop(parser);
    Parsed { ctx, nodes, arena, diags }
}

fn type_text(p: &Parsed, t: Option<NodeId>) -> String {
    match t.map(|t| p.arena.get(t)) {
        Some(ExprNode::UserType { name, child: None }) => p.ctx.symbol(*name as usize).name.clone(),
        Some(ExprNode::StringType { fixed_len: None }) => "String".into(),
        Some(ExprNode::BuiltinType { kind }) => format!("type{kind}"),
        _ => panic!("unexpected type node"),
    }
}

/// `[(name, type)]` of the DimItems of a single top-level declaration.
fn vars(p: &Parsed) -> Vec<(String, String)> {
    assert_eq!(p.nodes.len(), 1);
    let ExprNode::Block { stmts } = p.arena.get(p.nodes[0]) else { panic!("expected Block") };
    stmts
        .iter()
        .map(|s| match p.arena.get(*s) {
            ExprNode::DimItem { name, is_const: false, type_node, init: None, .. } => {
                (p.ctx.symbol(*name as usize).name.clone(), type_text(p, *type_node))
            }
            _ => panic!("expected a variable DimItem"),
        })
        .collect()
}

#[test]
fn public_variables_named_by_contextual_keywords() {
    for (src, name, ty) in [
        ("Public Name As String\n", "Name", "String"),
        ("Public Text As String\n", "Text", "String"),
        ("Public Access As NodeAccess\n", "Access", "NodeAccess"),
        ("Private Line As Long\n", "Line", "type3"),
    ] {
        let p = parse(src);
        assert_eq!(p.diags, vec![], "{src:?}");
        assert_eq!(vars(&p), vec![(name.to_string(), ty.to_string())], "{src:?}");
    }
}

#[test]
fn contextual_keyword_as_a_type_name() {
    let p = parse("Dim m_Err As Error\n");
    assert_eq!(p.diags, vec![]);
    assert_eq!(vars(&p), vec![("m_Err".to_string(), "Error".to_string())]);
}

#[test]
fn typeof_is_with_a_contextual_keyword_type() {
    let p = parse("Sub T()\nIf TypeOf oControl Is Line Then nCount = 1\nEnd Sub\n");
    assert_eq!(p.diags, vec![]);
    let typeof_types: Vec<String> = (0..p.arena.len())
        .filter_map(|i| match p.arena.get(NodeId(i as u32)) {
            ExprNode::TypeOf { type_spec, .. } => Some(type_text(&p, Some(*type_spec))),
            _ => None,
        })
        .collect();
    assert_eq!(typeof_types, vec!["Line".to_string()]);
}

#[test]
fn property_keyword_after_public_still_declares_a_property() {
    let p = parse("Public Property Get Name() As String\nEnd Property\n");
    assert_eq!(p.diags, vec![]);
    assert!(matches!(p.arena.get(p.nodes[0]), ExprNode::ProcDecl { kind: ProcKind::PropGet, .. }));
}

#[test]
fn reserved_word_cannot_name_a_variable() {
    let src = "Public If As Long\n";
    let p = parse(src);
    assert!(p.diags.iter().any(|d| d.1 == src.find("If").unwrap() as u32), "{:?}", p.diags);
}

#[test]
fn reserved_word_cannot_name_a_type() {
    let src = "Dim m_Value As If\n";
    let p = parse(src);
    assert!(p.diags.iter().any(|d| d.1 == src.find("If").unwrap() as u32), "{:?}", p.diags);
}
