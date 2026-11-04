//! Rules for `Type` and `Enum` blocks: both need members, a `Type` holds
//! only member declarations, and `End Type` needs an open `Type`.

mod common;
use common::{offset_of, parse};
use vb6_syntax::frontend::ast::ExprNode;

const END_TYPE_OUTSIDE_TYPE: u32 = 0x9c75;
const STATEMENT_IN_TYPE: u32 = 0x9c76;
const TYPE_WITHOUT_MEMBERS: u32 = 0x9cfc;
const ENUM_WITHOUT_MEMBERS: u32 = 0xdf43;

#[test]
fn type_without_members_is_reported_at_end_type() {
    let src = "Type T\nEnd Type\nSub S()\nEnd Sub\n";
    let p = parse(src);
    assert_eq!(p.found, vec![(TYPE_WITHOUT_MEMBERS, offset_of(src, "End Type", 0), 3)]);
    assert_eq!(p.top.len(), 2);
    let ExprNode::TypeDecl { members, .. } = p.node(p.top[0]) else { panic!("expected Type") };
    assert!(members.is_empty());

    // Comments and blank lines are not members.
    let src = "Type T\n' nothing\n\nEnd Type\n";
    assert_eq!(parse(src).found, vec![(TYPE_WITHOUT_MEMBERS, offset_of(src, "End Type", 0), 3)]);
}

#[test]
fn enum_without_members_is_reported_at_end_enum() {
    let src = "Enum E\nEnd Enum\nSub S()\nEnd Sub\n";
    let p = parse(src);
    assert_eq!(p.found, vec![(ENUM_WITHOUT_MEMBERS, offset_of(src, "End Enum", 0), 3)]);
    assert_eq!(p.top.len(), 2);
    let ExprNode::EnumDecl { members, .. } = p.node(p.top[0]) else { panic!("expected Enum") };
    assert!(members.is_empty());
}

#[test]
fn type_and_enum_with_members_are_clean() {
    let p = parse("Type T\nx As Long\ny As String\nEnd Type\nEnum E\nA\nB = 2\nEnd Enum\n");
    assert_eq!(p.found, vec![]);
    let ExprNode::TypeDecl { members, .. } = p.node(p.top[0]) else { panic!("expected Type") };
    assert_eq!(members.len(), 2);
    let ExprNode::EnumDecl { members, .. } = p.node(p.top[1]) else { panic!("expected Enum") };
    assert_eq!(members.len(), 2);
}

#[test]
fn a_member_may_be_named_like_a_keyword() {
    let p = parse("Type T\nType As String\nName As String\nText(3) As Long\nEnd Type\n");
    assert_eq!(p.found, vec![]);
    let ExprNode::TypeDecl { members, .. } = p.node(p.top[0]) else { panic!("expected Type") };
    assert_eq!(members.len(), 3);
}

#[test]
fn a_member_without_a_type_is_a_statement_in_the_type() {
    let src = "Type T\nx As Long\ny\nEnd Type\n";
    let p = parse(src);
    assert_eq!(p.found, vec![(0x9c76, src.find("y\nEnd").unwrap() as u32, 1)]);
}

#[test]
fn a_member_named_like_a_statement_cannot_take_bounds() {
    let src = "Type T\nm As Long\nLine(3) As Long\nEnd Type\n";
    let p = parse(src);
    assert_eq!(p.found, vec![(0x9c6f, src.find("Line").unwrap() as u32, 4)]);
}

#[test]
fn statements_inside_a_type_are_reported_and_skipped() {
    for line in ["Dim z As Long", "If a Then b", "Public z As Long", "x = 1"] {
        let src = format!("Type T\nm As Long\n{line}\nn As Long\nEnd Type\n");
        let p = parse(&src);
        assert_eq!(
            p.found,
            vec![(STATEMENT_IN_TYPE, offset_of(&src, line, 0), line.split(' ').next().unwrap().len() as u32)],
            "for {line:?}"
        );
        let ExprNode::TypeDecl { members, .. } = p.node(p.top[0]) else { panic!("expected Type") };
        let count = if line == "x = 1" { 3 } else { 2 };
        assert_eq!(members.len(), count, "for {line:?}");
    }
}

#[test]
fn end_type_with_no_type_is_reported_and_skipped() {
    let src = "End Type\nSub S()\nEnd Sub\n";
    let p = parse(src);
    assert_eq!(p.found, vec![(END_TYPE_OUTSIDE_TYPE, 0, 3)]);
    assert_eq!(p.top.len(), 1);
}
