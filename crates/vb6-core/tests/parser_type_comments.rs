//! Comment lines (`'` and `Rem`) inside a `Type ... End Type` block are blank
//! lines: the members are exactly the declared fields.

use vb6_core::frontend::ast::{ExprArena, ExprNode, NodeId};
use vb6_core::frontend::parser::Parser;
use vb6_core::frontend::scanner::ScannerContext;

/// `(type name, [(member, type)])` of the only top-level node, plus diagnostics.
fn type_decl(src: &str) -> (Vec<(u32, u32, u32)>, String, Vec<(String, String)>) {
    let mut ctx = ScannerContext::new(1, 1, 0x0409);
    ctx.intern_keywords();
    let mut arena = ExprArena::new();
    let mut parser = Parser::new(&mut ctx, src.as_bytes());
    let nodes = parser.parse_module(&mut arena);
    let diags = parser.diagnostics.items().iter().map(|d| (d.code, d.span.start, d.span.len)).collect();
    drop(parser);
    assert_eq!(nodes.len(), 1, "one top-level node");
    let name = |sym: u32| ctx.symbol(sym as usize).name.clone();
    let ExprNode::TypeDecl { name: tn, members } = arena.get(nodes[0]) else { panic!("expected TypeDecl") };
    let ty = |arena: &ExprArena, t: Option<NodeId>| match t.map(|t| arena.get(t)) {
        Some(ExprNode::BuiltinType { kind }) => format!("type{kind}"),
        Some(ExprNode::StringType { fixed_len: None }) => "String".into(),
        Some(ExprNode::UserType { name: n, child: None }) => name(*n),
        _ => panic!("unexpected member type"),
    };
    let members = members
        .iter()
        .map(|m| match arena.get(*m) {
            ExprNode::DimItem { name: n, is_const: false, is_static: false, is_new: false, with_events: false, bounds: None, type_node, init: None } => {
                (name(*n), ty(&arena, *type_node))
            }
            _ => panic!("expected a plain field DimItem"),
        })
        .collect();
    (diags, name(*tn), members)
}

fn fields() -> Vec<(String, String)> {
    vec![("sFieldName".into(), "String".into()), ("vValue".into(), "type12".into()), ("nKind".into(), "type3".into())]
}

#[test]
fn comment_lines_before_between_and_after_members() {
    let src = "Private Type tField\n    'The name of this field\n    sFieldName As String\n\n    ' The value written back\n    vValue As Variant\n    nKind As Long\n    ' last line\nEnd Type\n";
    assert_eq!(type_decl(src), (vec![], "tField".into(), fields()));
}

#[test]
fn rem_lines_are_comments_too() {
    let src = "Type tField\nRem The name\n    sFieldName As String\n    vValue As Variant\n    Rem kind\n    nKind As Long\nEnd Type\n";
    assert_eq!(type_decl(src), (vec![], "tField".into(), fields()));
}

#[test]
fn trailing_comment_after_a_member() {
    let src = "Type tField\n    sFieldName As String ' name\n    vValue As Variant   'value\n    nKind As Long\nEnd Type\n";
    assert_eq!(type_decl(src), (vec![], "tField".into(), fields()));
}

#[test]
fn comment_lines_with_crlf() {
    let src = "Type tField\r\n    ' name\r\n    sFieldName As String\r\n    vValue As Variant\r\n    nKind As Long\r\nEnd Type\r\n";
    assert_eq!(type_decl(src), (vec![], "tField".into(), fields()));
}
