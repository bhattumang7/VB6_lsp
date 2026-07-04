//! `bind_with_classes`: cross-module member-type resolution for
//! `Dim o As New ClassName` / `o.Field`, given the class's Public field list
//! from a separately-bound module (see `ExternalClass`).

use std::collections::HashMap;

use vb6_sema::frontend::ast::{ExprArena, ExprNode, NodeId};
use vb6_sema::frontend::parser::Parser;
use vb6_sema::frontend::scanner::ScannerContext;
use vb6_sema::sema::{bind_with_classes, ExternalClass, VbaType};

fn bind_with(src: &str, classes: &HashMap<String, ExternalClass>) -> vb6_sema::sema::BoundModule {
    let mut ctx = ScannerContext::new(1, 1, 0x0409);
    ctx.intern_keywords();
    let mut arena = ExprArena::new();
    let mut parser = Parser::new(&mut ctx, src.as_bytes());
    let top = parser.parse_module(&mut arena);
    let spans = std::mem::take(&mut parser.node_spans);
    let vis = std::mem::take(&mut parser.decl_public);
    drop(parser);
    bind_with_classes(&ctx, &arena, &top, &spans, &vis, classes)
}

#[test]
fn member_access_on_known_class_resolves_field_type() {
    let mut classes = HashMap::new();
    classes.insert(
        "class1".to_string(),
        ExternalClass { fields: vec![("F".to_string(), VbaType::Long)] },
    );

    let src = "Attribute VB_Name = \"Module1\"\r\n\
               Sub Main()\r\n\
               \x20   Dim o As New Class1\r\n\
               \x20   Dim x As Long\r\n\
               \x20   o.F = 1\r\n\
               \x20   x = o.F\r\n\
               End Sub\r\n";
    let m = bind_with(src, &classes);

    // Every `ExprNode::MemberAccess` (`o.F`) in the arena must type as Long,
    // not fall back to Variant (the lookup-miss fallback). `bind_with` doesn't
    // expose the arena it built internally, so re-parse the same source to
    // walk it alongside the returned `BoundModule`.
    let mut ctx2 = ScannerContext::new(1, 1, 0x0409);
    ctx2.intern_keywords();
    let mut arena2 = ExprArena::new();
    let mut parser2 = Parser::new(&mut ctx2, src.as_bytes());
    let _top2 = parser2.parse_module(&mut arena2);
    let member_ids: Vec<NodeId> = (0..arena2.len() as u32)
        .map(NodeId)
        .filter(|id| matches!(arena2.get(*id), ExprNode::MemberAccess { .. }))
        .collect();
    assert_eq!(member_ids.len(), 2, "expected exactly two o.F member accesses");
    for id in member_ids {
        assert_eq!(m.types.get(&id.0), Some(&VbaType::Long), "o.F must resolve to Long, not fall back to Variant");
    }
}

#[test]
fn member_access_on_unknown_class_falls_back_to_variant() {
    let classes: HashMap<String, ExternalClass> = HashMap::new(); // no class registered
    let src = "Attribute VB_Name = \"Module1\"\r\n\
               Sub Main()\r\n\
               \x20   Dim o As New Class1\r\n\
               \x20   Dim x As Long\r\n\
               \x20   x = o.F\r\n\
               End Sub\r\n";
    let m = bind_with(src, &classes);

    let mut ctx2 = ScannerContext::new(1, 1, 0x0409);
    ctx2.intern_keywords();
    let mut arena2 = ExprArena::new();
    let mut parser2 = Parser::new(&mut ctx2, src.as_bytes());
    let _top2 = parser2.parse_module(&mut arena2);
    let member_id = (0..arena2.len() as u32)
        .map(NodeId)
        .find(|id| matches!(arena2.get(*id), ExprNode::MemberAccess { .. }))
        .expect("expected one member access");
    assert_eq!(m.types.get(&member_id.0), Some(&VbaType::Variant));
}
