//! Size limits: identifier length, dimensions per array, fixed-string length.

mod common;
use common::{offset_of, parse};
use vb6_syntax::frontend::ast::{AstLit, ExprNode};

const IDENTIFIER_TOO_LONG: u32 = 0x9c6d;
const TOO_MANY_DIMENSIONS: u32 = 0x9c80;
const INVALID_FIXED_STRING_LENGTH: u32 = 0x9cb0;

fn first_dim_item(p: &common::Parsed) -> &ExprNode {
    let ExprNode::Block { stmts } = p.node(p.top[0]) else { panic!("expected a Dim statement") };
    p.node(stmts[0])
}

#[test]
fn identifier_of_255_bytes_is_fine_and_256_is_reported() {
    let ok = "a".repeat(255);
    let src = format!("Dim {ok} As Long\n");
    assert_eq!(parse(&src).found, vec![]);

    let long = "a".repeat(256);
    let src = format!("Dim {long} As Long\n");
    let p = parse(&src);
    assert_eq!(p.found, vec![(IDENTIFIER_TOO_LONG, 4, 256)]);
    // The declaration is kept with its full name.
    let ExprNode::DimItem { name, .. } = first_dim_item(&p) else { panic!("expected Dim") };
    assert_eq!(p.name(*name).len(), 256);
}

#[test]
fn every_use_of_a_long_identifier_is_reported() {
    let long = "b".repeat(300);
    let src = format!("Sub S()\n{long} = 1\nx = {long}\nEnd Sub\n");
    let p = parse(&src);
    assert_eq!(
        p.found,
        vec![
            (IDENTIFIER_TOO_LONG, offset_of(&src, &long, 0), 300),
            (IDENTIFIER_TOO_LONG, offset_of(&src, &long, 1), 300),
        ]
    );
}

#[test]
fn an_array_may_have_60_dimensions_but_not_61() {
    let dims = |n: usize| vec!["1"; n].join(",");
    let src = format!("Dim a({}) As Long\n", dims(60));
    assert_eq!(parse(&src).found, vec![]);
    let src = format!("Dim a({}) As Long\n", dims(61));
    let p = parse(&src);
    assert_eq!(p.found, vec![(TOO_MANY_DIMENSIONS, offset_of(&src, "1", 0), 1)]);
    let ExprNode::DimItem { bounds: Some(b), .. } = first_dim_item(&p) else { panic!("expected Dim") };
    let ExprNode::ArgList { args } = p.node(*b) else { panic!("expected ArgList") };
    assert_eq!(args.len(), 61);
}

#[test]
fn fixed_string_literal_length_must_be_in_range() {
    for ok in ["1", "10", "65526"] {
        assert_eq!(parse(&format!("Dim s As String * {ok}\n")).found, vec![], "for {ok}");
    }
    for bad in ["0", "65527", "70000"] {
        let src = format!("Dim s As String * {bad}\n");
        let p = parse(&src);
        assert_eq!(
            p.found,
            vec![(INVALID_FIXED_STRING_LENGTH, offset_of(&src, bad, 0), bad.len() as u32)],
            "for {bad}"
        );
    }
    // The length is kept in the tree.
    let p = parse("Dim s As String * 10\n");
    let ExprNode::DimItem { type_node: Some(t), .. } = first_dim_item(&p) else { panic!("expected Dim") };
    let ExprNode::StringType { fixed_len: Some(len) } = p.node(*t) else { panic!("expected String type") };
    assert!(matches!(p.node(*len), ExprNode::Literal { lit: AstLit::Int(10) }));
}
