//! `New` takes only a class name; the variables of `Input #` and `Line Input #`
//! are assignment targets, and `Line Input` requires its `#`.

mod common;
use common::{offset_of, parse, Parsed};
use vb6_syntax::frontend::ast::{AstLit, ExprNode, FileIoKind, NodeId};
use vb6_syntax::frontend::token::TypeSuffix;

fn body(src: &str) -> (Parsed, Vec<NodeId>) {
    let p = parse(src);
    let b = p.proc_body(0);
    (p, b)
}

fn name_of(p: &Parsed, id: NodeId) -> String {
    match p.node(id) {
        ExprNode::NameRef { sym, suffix: TypeSuffix::None } => p.name(*sym).to_string(),
        _ => panic!("expected a plain NameRef"),
    }
}

#[test]
fn new_class_is_a_user_type() {
    let (p, b) = body("Sub S()\n x = New Collection\nEnd Sub\n");
    assert_eq!(p.diags, 0);
    let ExprNode::Assign { value, .. } = p.node(b[0]) else { panic!("not an assignment") };
    let ExprNode::New { type_spec } = p.node(*value) else { panic!("not New") };
    let ExprNode::UserType { name, child: None } = p.node(*type_spec) else { panic!("not a user type") };
    assert_eq!(p.name(*name), "Collection");
}

#[test]
fn new_qualified_class_nests_the_member_name() {
    let (p, b) = body("Sub S()\n x = New Lib.Widget\nEnd Sub\n");
    assert_eq!(p.diags, 0);
    let ExprNode::Assign { value, .. } = p.node(b[0]) else { panic!("not an assignment") };
    let ExprNode::New { type_spec } = p.node(*value) else { panic!("not New") };
    let ExprNode::UserType { name, child: Some(child) } = p.node(*type_spec) else { panic!("not a qualified type") };
    assert_eq!(p.name(*name), "Lib");
    assert_eq!(name_of(&p, *child), "Widget");
}

#[test]
fn new_rejects_a_built_in_type_object_and_a_second_new() {
    for word in ["Integer", "Long", "Single", "Double", "Currency", "Date", "Byte", "Boolean", "Variant", "String", "Object", "New"] {
        let src = format!("Sub S()\n x = New {word}\nEnd Sub\n");
        let p = parse(&src);
        let at = offset_of(&src, "New ", 0) + 4;
        assert_eq!(p.found[0], (0x9c70, at, word.len() as u32), "{word}");
    }
}

#[test]
fn new_rejection_at_module_level_names_the_expected_item() {
    let src = "Const c = New Integer\n";
    let p = parse(src);
    assert_eq!(p.found, vec![(0x9c6f, offset_of(src, "Integer", 0), 7)]);
}

#[test]
fn input_variables_are_targets() {
    let (p, b) = body("Sub S()\n Input #1, a, b(1), c.d, .e\nEnd Sub\n");
    assert_eq!(p.diags, 0);
    let ExprNode::FileIoStmt { kind: FileIoKind::Input, channel: Some(ch), args, .. } = p.node(b[0]) else { panic!("not Input") };
    assert!(matches!(p.node(*ch), ExprNode::Literal { lit: AstLit::Int(1) }));
    assert_eq!(args.len(), 4);
    assert_eq!(name_of(&p, args[0]), "a");
    assert!(matches!(p.node(args[1]), ExprNode::Call { .. }));
    assert!(matches!(p.node(args[2]), ExprNode::MemberAccess { .. }));
    assert!(matches!(p.node(args[3]), ExprNode::MemberAccess { .. }));
}

#[test]
fn input_rejects_a_variable_that_is_not_a_target() {
    // (text, offset of the first rejected token within it): a target followed by an
    // operator is rejected at the operator.
    let cases: [(&str, u32); 18] = [
        ("5", 0), ("a + 1", 2), ("(a)", 0), ("-a", 0), ("Not a", 0), ("New a", 0), ("Date", 0), ("Nothing", 0),
        ("True", 0), ("\"q\"", 0), ("1.5", 0), ("a & b", 2), ("Len(v)", 0), ("CInt(a)", 0), ("a = 2", 2),
        ("Empty", 0), ("Null", 0), ("Seek", 0),
    ];
    for (bad, rel) in cases {
        let src = format!("Sub S()
 Input #1, {bad}
End Sub
");
        let p = parse(&src);
        assert_eq!(p.found[0].1, offset_of(&src, "#1, ", 0) + 4 + rel, "{bad}: {:?}", p.found);
    }
}

#[test]
fn line_input_target_and_channel() {
    let (p, b) = body("Sub S()\n Line Input #1, target\nEnd Sub\n");
    assert_eq!(p.diags, 0);
    let ExprNode::FileIoStmt { kind: FileIoKind::LineInput, channel: Some(ch), channel_marks, args, open_mode: 0, open_flags: 0 } = p.node(b[0]) else { panic!("not Line Input") };
    assert!(matches!(p.node(*ch), ExprNode::Literal { lit: AstLit::Int(1) }));
    assert_eq!(channel_marks, &vec![true]);
    assert_eq!(args.len(), 1);
    assert_eq!(name_of(&p, args[0]), "target");
}

#[test]
fn line_input_requires_the_channel_mark() {
    let src = "Sub S()\n Line Input 1, s\nEnd Sub\n";
    let p = parse(src);
    assert_eq!(p.found[0], (0x9c70, offset_of(src, "1", 0), 1));
}

#[test]
fn line_input_rejects_a_variable_that_is_not_a_target() {
    let cases: [(&str, u32); 9] = [
        ("5", 0), ("v + v", 2), ("(v)", 0), ("-v", 0), ("\"a\"", 0), ("Date", 0), ("Nothing", 0), ("New v", 0), ("Len(v)", 0),
    ];
    for (bad, rel) in cases {
        let src = format!("Sub S()
 Line Input #1, {bad}
End Sub
");
        let p = parse(&src);
        assert_eq!(p.found[0].1, offset_of(&src, "#1, ", 0) + 4 + rel, "{bad}: {:?}", p.found);
    }
}
