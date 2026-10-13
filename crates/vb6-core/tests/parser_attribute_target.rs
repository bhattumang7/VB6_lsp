//! `Attribute Member.VB_UserMemId = 0` keeps the member it applies to;
//! module-level attributes have no target.

use vb6_core::frontend::ast::{AstLit, ExprArena, ExprNode, NodeId};
use vb6_core::frontend::parser::{ModuleKind, Parser};
use vb6_core::frontend::scanner::ScannerContext;

/// Every Attribute node as (target, name, values) in arena order.
fn attributes(src: &str) -> (usize, Vec<(Option<String>, String, Vec<AstLit>)>) {
    let mut ctx = ScannerContext::new(1, 1, 0x0409);
    ctx.intern_keywords();
    let mut arena = ExprArena::new();
    let mut parser = Parser::with_module_kind(&mut ctx, src.as_bytes(), ModuleKind::Class);
    let _ = parser.parse_module(&mut arena);
    let diags = parser.diagnostics.items().len();
    drop(parser);
    let name = |s: u32| ctx.symbol(s as usize).name.clone();
    let mut out = Vec::new();
    for i in 0..arena.len() as u32 {
        if let ExprNode::Attribute { target, name: n, values } = arena.get(NodeId(i)) {
            let lits = values
                .iter()
                .map(|v| match arena.get(*v) {
                    ExprNode::Literal { lit } => lit.clone(),
                    _ => panic!("attribute value is a literal"),
                })
                .collect();
            out.push((target.map(name), name(*n), lits));
        }
    }
    (diags, out)
}

#[test]
fn module_attribute_has_no_target() {
    let (diags, attrs) = attributes("Attribute VB_Name = \"clsTrips\"\n");
    assert_eq!(diags, 0);
    assert_eq!(attrs, vec![(None, "VB_Name".to_string(), vec![AstLit::Str("clsTrips".into())])]);
}

#[test]
fn procedure_attribute_names_its_member() {
    let src = "Public Function Item(i As Long) As Long\nAttribute Item.VB_UserMemId = 0\nEnd Function\n";
    let (diags, attrs) = attributes(src);
    assert_eq!(diags, 0);
    assert_eq!(
        attrs,
        vec![(Some("Item".to_string()), "VB_UserMemId".to_string(), vec![AstLit::Int(0)])]
    );
}

#[test]
fn variable_attribute_names_its_member() {
    let src = "Public Value As Long\nAttribute Value.VB_VarUserMemId = 0\n";
    let (diags, attrs) = attributes(src);
    assert_eq!(diags, 0);
    assert_eq!(
        attrs,
        vec![(Some("Value".to_string()), "VB_VarUserMemId".to_string(), vec![AstLit::Int(0)])]
    );
}

#[test]
fn predeclared_id_module_attribute() {
    let (diags, attrs) = attributes("Attribute VB_PredeclaredId = True\n");
    assert_eq!(diags, 0);
    assert_eq!(attrs, vec![(None, "VB_PredeclaredId".to_string(), vec![AstLit::Bool(true)])]);
}
