//! Qualified type names (`Project.Class`, `Module.Type`) in `As`, `As New`,
//! `New` expressions and `ReDim`: the `UserType` node spans the qualifier (or
//! the plain name) and carries the member name as a `NameRef` child with its
//! own span.

use vb6_core::frontend::ast::{ExprArena, ExprNode, NodeId, NodeSpans};
use vb6_core::frontend::parser::{ModuleKind, Parser};
use vb6_core::frontend::scanner::ScannerContext;

struct Parsed {
    src: String,
    ctx: ScannerContext,
    arena: ExprArena,
    spans: NodeSpans,
    top: Vec<NodeId>,
    diags: usize,
}

fn parse(src: &str) -> Parsed {
    let mut ctx = ScannerContext::new(1, 1, 0x0409);
    ctx.intern_keywords();
    let mut arena = ExprArena::new();
    let mut parser = Parser::with_module_kind(&mut ctx, src.as_bytes(), ModuleKind::Class);
    let top = parser.parse_module(&mut arena);
    let spans = std::mem::take(&mut parser.node_spans);
    let diags = parser.diagnostics.items().len();
    drop(parser);
    Parsed { src: src.to_string(), ctx, arena, spans, top, diags }
}

impl Parsed {
    fn text(&self, id: NodeId) -> &str {
        let s = self.spans.get(id);
        &self.src[s.start as usize..(s.start + s.len) as usize]
    }

    fn name(&self, sym: u32) -> String {
        self.ctx.symbol(sym as usize).name.clone()
    }

    /// `(qualifier-or-name, child name, qualifier text, child text)` of a UserType.
    fn user_type(&self, id: NodeId) -> (String, Option<String>, String, Option<String>) {
        let ExprNode::UserType { name, child } = self.arena.get(id) else { panic!("expected UserType") };
        let child_name = child.map(|c| {
            let ExprNode::NameRef { sym, .. } = self.arena.get(c) else { panic!("child is a NameRef") };
            self.name(*sym)
        });
        (self.name(*name), child_name, self.text(id).to_string(), child.map(|c| self.text(c).to_string()))
    }

    /// Statements of the single block a module-level declaration list or a Sub body holds.
    fn block_stmts(&self, id: NodeId) -> Vec<NodeId> {
        let ExprNode::Block { stmts } = self.arena.get(id) else { panic!("expected Block") };
        stmts.clone()
    }

    fn sub_body(&self) -> Vec<NodeId> {
        let ExprNode::ProcDecl { body, .. } = self.arena.get(self.top[0]) else { panic!("expected ProcDecl") };
        self.block_stmts(*body)
    }
}

#[test]
fn dim_as_qualified_class() {
    let p = parse("Dim o As Proj.clsTrip\n");
    assert_eq!(p.diags, 0);
    let stmts = p.block_stmts(p.top[0]);
    let ExprNode::DimItem { name, is_new, type_node, bounds, .. } = p.arena.get(stmts[0]) else {
        panic!("expected DimItem")
    };
    assert_eq!(p.name(*name), "o");
    assert!(!*is_new);
    assert!(bounds.is_none());
    assert_eq!(
        p.user_type(type_node.unwrap()),
        ("Proj".to_string(), Some("clsTrip".to_string()), "Proj".to_string(), Some("clsTrip".to_string()))
    );
}

#[test]
fn dim_as_new_qualified_class() {
    let p = parse("Dim o As New Proj.clsTrip\n");
    assert_eq!(p.diags, 0);
    let stmts = p.block_stmts(p.top[0]);
    let ExprNode::DimItem { is_new, type_node, .. } = p.arena.get(stmts[0]) else { panic!("expected DimItem") };
    assert!(*is_new);
    assert_eq!(
        p.user_type(type_node.unwrap()),
        ("Proj".to_string(), Some("clsTrip".to_string()), "Proj".to_string(), Some("clsTrip".to_string()))
    );
}

#[test]
fn dim_as_unqualified_class_spans_the_name() {
    let p = parse("Dim o As clsTrip\n");
    assert_eq!(p.diags, 0);
    let stmts = p.block_stmts(p.top[0]);
    let ExprNode::DimItem { type_node, .. } = p.arena.get(stmts[0]) else { panic!("expected DimItem") };
    assert_eq!(
        p.user_type(type_node.unwrap()),
        ("clsTrip".to_string(), None, "clsTrip".to_string(), None)
    );
}

#[test]
fn new_expression_with_qualified_class() {
    let p = parse("Sub T()\nSet o = New Proj.clsTrip\nEnd Sub\n");
    assert_eq!(p.diags, 0);
    let stmts = p.sub_body();
    let ExprNode::SetAssign { value, .. } = p.arena.get(stmts[0]) else { panic!("expected SetAssign") };
    let ExprNode::New { type_spec } = p.arena.get(*value) else { panic!("expected New") };
    assert_eq!(
        p.user_type(*type_spec),
        ("Proj".to_string(), Some("clsTrip".to_string()), "Proj".to_string(), Some("clsTrip".to_string()))
    );
}

#[test]
fn redim_as_new_qualified_class() {
    let p = parse("Sub T()\nReDim a(3) As New Proj.clsTrip\nEnd Sub\n");
    assert_eq!(p.diags, 0);
    let stmts = p.sub_body();
    let inner = p.block_stmts(stmts[0]);
    let ExprNode::ReDimItem { name, base, type_node, .. } = p.arena.get(inner[0]) else {
        panic!("expected ReDimItem")
    };
    assert_eq!(p.name(*name), "a");
    assert!(base.is_none());
    assert_eq!(p.text(inner[0]), "a");
    assert_eq!(
        p.user_type(type_node.unwrap()),
        ("Proj".to_string(), Some("clsTrip".to_string()), "Proj".to_string(), Some("clsTrip".to_string()))
    );
}

#[test]
fn redim_member_array_spans_the_final_name() {
    let p = parse("Sub T()\nReDim t.Items(5)\nEnd Sub\n");
    assert_eq!(p.diags, 0);
    let stmts = p.sub_body();
    let inner = p.block_stmts(stmts[0]);
    let ExprNode::ReDimItem { name, base, .. } = p.arena.get(inner[0]) else { panic!("expected ReDimItem") };
    assert_eq!(p.name(*name), "Items");
    let base = base.expect("member array has a base");
    assert!(matches!(p.arena.get(base), ExprNode::NameRef { .. }));
    assert_eq!(p.text(base), "t");
    assert_eq!(p.text(inner[0]), "Items");
}
