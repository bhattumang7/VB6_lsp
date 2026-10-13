//! Members of an `Enum` block: names (including non-reserved keywords),
//! values, comments and `:` separators; anything else is reported.

use vb6_core::frontend::ast::{ExprArena, ExprNode};
use vb6_core::frontend::parser::{ModuleKind, Parser};
use vb6_core::frontend::scanner::ScannerContext;

/// (diagnostic labels at their spelling, member names) of the first Enum.
fn enum_of(src: &str) -> (Vec<(Option<&'static str>, String)>, Vec<String>) {
    let mut ctx = ScannerContext::new(1, 1, 0x0409);
    ctx.intern_keywords();
    let mut arena = ExprArena::new();
    let mut parser = Parser::with_module_kind(&mut ctx, src.as_bytes(), ModuleKind::Standard);
    let top = parser.parse_module(&mut arena);
    let diags: Vec<_> = parser
        .diagnostics
        .items()
        .iter()
        .map(|d| (d.label, src[d.span.start as usize..(d.span.start + d.span.len) as usize].to_string()))
        .collect();
    drop(parser);
    let ExprNode::EnumDecl { members, .. } = arena.get(top[0]) else { panic!("EnumDecl") };
    let names = members
        .iter()
        .map(|m| match arena.get(*m) {
            ExprNode::DimItem { name, .. } => ctx.symbol(*name as usize).name.clone(),
            _ => panic!("member is a DimItem"),
        })
        .collect();
    (diags, names)
}

#[test]
fn plain_members_values_comments_and_colons() {
    let src = "Enum Color\n    Red\n    ' a comment\n    Green = 5\n    Blue: Alpha\nEnd Enum\n";
    assert_eq!(
        enum_of(src),
        (vec![], vec!["Red".to_string(), "Green".to_string(), "Blue".to_string(), "Alpha".to_string()])
    );
}

#[test]
fn non_reserved_keyword_names_a_member() {
    let src = "Enum Field\n    Text\n    Key = 2\n    Name\nEnd Enum\n";
    let (diags, names) = enum_of(src);
    assert_eq!(diags, vec![]);
    assert_eq!(names, vec!["Text".to_string(), "Key".to_string(), "Name".to_string()]);
}

#[test]
fn junk_after_a_member_is_reported() {
    let src = "Enum Color\n    Red Green\n    Blue\nEnd Enum\n";
    let (diags, names) = enum_of(src);
    assert_eq!(diags, vec![(Some("end of statement"), "Green".to_string())]);
    assert_eq!(names, vec!["Red".to_string(), "Blue".to_string()]);
}

#[test]
fn a_line_that_is_not_a_member_is_reported() {
    let src = "Enum Color\n    Red\n    Dim x As Long\n    Blue\nEnd Enum\n";
    let (diags, names) = enum_of(src);
    assert_eq!(diags, vec![(Some("identifier"), "Dim".to_string())]);
    assert_eq!(names, vec!["Red".to_string(), "Blue".to_string()]);
}
