//! The length after `String *` is one whole-number literal or one constant
//! name. In a `Const` the `=` after it belongs to the constant's value.

mod common;
use common::{offset_of, parse, Parsed};
use vb6_syntax::frontend::ast::{AstLit, ExprNode, NodeId};

/// The `DimItem` nodes of the first top-level declaration statement.
fn first_decl_items(p: &Parsed) -> Vec<NodeId> {
    let decl = p
        .top
        .iter()
        .copied()
        .find(|n| matches!(p.node(*n), ExprNode::Block { .. }))
        .expect("a declaration");
    p.block(decl)
}

fn length_of(p: &Parsed, item: NodeId) -> NodeId {
    let ExprNode::DimItem { type_node, .. } = p.node(item) else { panic!("DimItem") };
    let ExprNode::StringType { fixed_len } = p.node(type_node.expect("a type")) else { panic!("StringType") };
    fixed_len.expect("a fixed length")
}

#[test]
fn const_with_a_fixed_length_string_keeps_its_value() {
    let p = parse("Public Const MACRO_EXT As String * 3 = \"pdm\"\n");
    assert_eq!(p.found, vec![]);
    let items = first_decl_items(&p);
    assert_eq!(items.len(), 1);
    let ExprNode::DimItem { name, is_const, init, .. } = p.node(items[0]) else { panic!("DimItem") };
    assert_eq!((p.name(*name), *is_const), ("MACRO_EXT", true));
    assert!(matches!(p.node(length_of(&p, items[0])), ExprNode::Literal { lit: AstLit::Int(3) }));
    assert!(matches!(p.node(init.unwrap()), ExprNode::Literal { lit: AstLit::Str(s) } if &**s == "pdm"));
}

#[test]
fn const_list_with_two_fixed_length_strings() {
    let p = parse("Const c As String * 3 = \"abc\", d As String * 2 = \"ab\"\n");
    assert_eq!(p.found, vec![]);
    let items = first_decl_items(&p);
    assert_eq!(items.len(), 2);
    for (item, (name, len, text)) in items.iter().zip([("c", 3, "abc"), ("d", 2, "ab")]) {
        let ExprNode::DimItem { name: n, init, .. } = p.node(*item) else { panic!("DimItem") };
        assert_eq!(p.name(*n), name);
        assert!(matches!(p.node(length_of(&p, *item)), ExprNode::Literal { lit: AstLit::Int(l) } if *l == len));
        assert!(matches!(p.node(init.unwrap()), ExprNode::Literal { lit: AstLit::Str(s) } if &**s == text));
    }
}

#[test]
fn length_may_be_a_constant_name() {
    let p = parse("Const N = 4\nDim s As String * N\nConst c As String * N = \"abcd\"\n");
    assert_eq!(p.found, vec![]);
    let decls: Vec<NodeId> =
        p.top.iter().copied().filter(|n| matches!(p.node(*n), ExprNode::Block { .. })).collect();
    assert_eq!(decls.len(), 3);
    for decl in &decls[1..] {
        let items = p.block(*decl);
        let ExprNode::NameRef { sym, .. } = p.node(length_of(&p, items[0])) else { panic!("name") };
        assert_eq!(p.name(*sym), "N");
    }
}

#[test]
fn length_literals_in_other_bases_and_with_suffixes() {
    for (text, want) in [("&H10", 16), ("&O17", 15), ("3%", 3), ("3&", 3)] {
        let p = parse(&format!("Dim s As String * {text}\n"));
        assert_eq!(p.found, vec![], "{text}");
        let items = first_decl_items(&p);
        let ok = match p.node(length_of(&p, items[0])) {
            ExprNode::Literal { lit: AstLit::Int(n) | AstLit::Long(n) } => *n == want,
            _ => false,
        };
        assert!(ok, "{text}");
    }
}

#[test]
fn length_range_is_checked_for_literals() {
    for (text, bad) in [("0", true), ("1", false), ("65526", false), ("65527", true)] {
        let src = format!("Dim s As String * {text}\n");
        let p = parse(&src);
        let want = if bad { vec![(0x9cb0, offset_of(&src, text, 0), text.len() as u32)] } else { vec![] };
        assert_eq!(p.found, want, "{text}");
    }
}

#[test]
fn an_operator_after_the_length_ends_the_declaration() {
    for tail in ["+ 1", "* 3", "^ 2", "Mod 3", "And 3"] {
        let src = format!("Dim s As String * 2 {tail}\n");
        let p = parse(&src);
        let at = src.find(tail).unwrap() as u32;
        assert_eq!(p.found, vec![(0x9c6f, at, tail.split(' ').next().unwrap().len() as u32)], "{tail}");
    }
}

#[test]
fn only_a_literal_or_a_name_can_be_the_length() {
    for bad in ["(2 + 1)", "-1", "3.0", "1E1", "\"3\"", "True", "Me"] {
        let src = format!("Dim s As String * {bad}\n");
        let p = parse(&src);
        assert!(!p.found.is_empty(), "{bad}");
        assert_eq!(p.found[0], (0x9c6f, offset_of(&src, bad, 0), p.found[0].2), "{bad}");
    }
}

#[test]
fn const_without_a_value_after_a_fixed_length_string_needs_equals() {
    let src = "Const c As String * 3\n";
    let p = parse(src);
    assert_eq!(p.found, vec![(0x9c6f, 21, 0), (0x9c6f, 21, 0)]);
}
