//! The first operand of the `Mid` statement is an assignment target; the others are
//! plain expressions, never named, omitted or parenthesised ahead of a target.

mod common;
use common::{offset_of, parse, Parsed};
use vb6_syntax::frontend::ast::{AstLit, ExprNode, NodeId};

fn first_statement(src: &str) -> (Parsed, NodeId) {
    let p = parse(src);
    let b = p.proc_body(0);
    (p, b[0])
}

fn args_of(p: &Parsed, id: NodeId) -> Vec<NodeId> {
    match p.node(id) {
        ExprNode::ArgList { args } => args.clone(),
        _ => panic!("not an argument list"),
    }
}

#[test]
fn two_operand_form() {
    let (p, st) = first_statement("Sub S()\n Mid(target, 1) = \"x\"\nEnd Sub\n");
    assert_eq!(p.diags, 0);
    let ExprNode::MidAssign { byte_oriented: false, dollar: false, args, value } = p.node(st) else { panic!("not Mid") };
    let a = args_of(&p, *args);
    assert_eq!(a.len(), 2);
    assert!(matches!(p.node(a[0]), ExprNode::NameRef { sym, .. } if p.name(*sym) == "target"));
    assert!(matches!(p.node(a[1]), ExprNode::Literal { lit: AstLit::Int(1) }));
    assert!(matches!(p.node(*value), ExprNode::Literal { lit: AstLit::Str(s) } if &**s == "x"));
}

#[test]
fn three_operand_dollar_byte_form() {
    let (p, st) = first_statement("Sub S()\n MidB$(target, 2, 3) = \"x\"\nEnd Sub\n");
    assert_eq!(p.diags, 0);
    let ExprNode::MidAssign { byte_oriented: true, dollar: true, args, .. } = p.node(st) else { panic!("not MidB$") };
    let a = args_of(&p, *args);
    assert_eq!(a.len(), 3);
    assert!(matches!(p.node(a[1]), ExprNode::Literal { lit: AstLit::Int(2) }));
    assert!(matches!(p.node(a[2]), ExprNode::Literal { lit: AstLit::Int(3) }));
}

#[test]
fn target_may_be_an_element_or_a_member() {
    for target in ["target(1)", "obj.item", "obj!item", "Me.item"] {
        let (p, st) = first_statement(&format!("Sub S()\n Mid({target}, 1) = \"x\"\nEnd Sub\n"));
        assert_eq!(p.diags, 0, "{target}");
        let ExprNode::MidAssign { args, .. } = p.node(st) else { panic!("not Mid") };
        assert_eq!(args_of(&p, *args).len(), 2, "{target}");
    }
}

#[test]
fn first_operand_must_be_a_target() {
    // (operand, offset within it of the first rejected token)
    let cases: [(&str, u32); 5] = [
        ("\"abc\", 1", 0),
        ("t & \"a\", 1", 2),
        ("(t), 1", 0),
        ("t, start:=1", 8),
        ("New t, 1", 0),
    ];
    for (operand, rel) in cases {
        let src = format!("Sub S()\n Mid({operand}) = \"x\"\nEnd Sub\n");
        let p = parse(&src);
        assert_eq!(p.found[0].1, offset_of(&src, "Mid(", 0) + 4 + rel, "{operand}: {:?}", p.found);
    }
}

#[test]
fn operand_count_is_two_or_three() {
    // (operands, offset within them of the first rejected token)
    let cases: [(&str, u32); 4] = [("t", 1), ("t,", 2), ("t, , 2", 3), ("t, 1, 2, 3", 7)];
    for (operand, rel) in cases {
        let src = format!("Sub S()
 Mid({operand}) = \"x\"
End Sub
");
        let p = parse(&src);
        assert_eq!(p.found[0].1, offset_of(&src, "Mid(", 0) + 4 + rel, "{operand}: {:?}", p.found);
    }
}
