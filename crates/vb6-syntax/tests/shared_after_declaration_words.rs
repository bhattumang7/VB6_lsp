//! `Shared` directly after `Dim`, `Public`, `Private` or `Global` is skipped and the
//! rest is an ordinary variable list; after `Static`, `Const` or `ReDim` it is an error.

mod common;
use common::{offset_of, parse, Parsed};
use vb6_syntax::frontend::ast::{ExprNode, NodeId};

fn dim_items(p: &Parsed) -> Vec<NodeId> {
    let mut out = Vec::new();
    let mut stack: Vec<NodeId> = p.top.iter().rev().copied().collect();
    while let Some(n) = stack.pop() {
        match p.node(n) {
            ExprNode::DimItem { .. } => out.push(n),
            ExprNode::Block { stmts } => stack.extend(stmts.iter().rev().copied()),
            _ => {}
        }
    }
    out
}

#[test]
fn shared_is_skipped_after_the_four_declaration_words() {
    for word in ["Dim", "Public", "Private", "Global"] {
        let p = parse(&format!("{word} Shared first, second As Long\n"));
        assert_eq!(p.diags, 0, "{word}");
        let items = dim_items(&p);
        assert_eq!(items.len(), 2, "{word}");
        for (item, want) in items.iter().zip(["first", "second"]) {
            let ExprNode::DimItem { name, is_const: false, is_static: false, bounds: None, .. } = p.node(*item) else { panic!("{word}: not a plain DimItem") };
            assert_eq!(p.name(*name), want, "{word}");
        }
    }
}

#[test]
fn shared_is_skipped_before_an_array_and_inside_a_procedure() {
    let p = parse("Dim Shared a() As Long\n");
    assert_eq!(p.diags, 0);
    let items = dim_items(&p);
    assert_eq!(items.len(), 1);
    assert!(matches!(p.node(items[0]), ExprNode::DimItem { bounds: Some(_), .. }));
    let p = parse("Sub S()\n Dim Shared x As Integer\nEnd Sub\n");
    assert_eq!(p.diags, 0);
}

#[test]
fn shared_is_an_error_after_static_const_and_redim() {
    for (src, word) in [
        ("Sub S()\n Static Shared x As Long\nEnd Sub\n", "Shared"),
        ("Const Shared = 1\n", "Shared"),
        ("Sub S()\n ReDim Shared a(3)\nEnd Sub\n", "Shared"),
        ("Sub S()\n ReDim Shared Preserve a(3)\nEnd Sub\n", "Shared"),
    ] {
        let p = parse(src);
        assert_eq!(p.found[0].1, offset_of(src, word, 0), "{src:?}: {:?}", p.found);
    }
}
