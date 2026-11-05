//! `TypeOf <target> Is <object type>`: the operand is a target expression and the
//! type is `Object` or a class name.

mod common;
use common::{offset_of, parse};
use vb6_syntax::frontend::ast::{ExprNode, NodeId};

fn in_if(cond: &str) -> String {
    format!("Sub S(x, y)\n    If {cond} Then y = 1\nEnd Sub\n")
}

#[test]
fn typeof_of_a_name_member_or_call_against_object_or_a_class_is_clean() {
    for cond in [
        "TypeOf x Is Object",
        "TypeOf x Is Foo",
        "TypeOf x Is Foo.Bar",
        "TypeOf x Is Foo.Bar.Baz",
        "TypeOf x.y Is Foo",
        "TypeOf f(1) Is Foo",
        "TypeOf Me Is Foo",
        "TypeOf .a Is Foo",
        "Not TypeOf x Is Foo",
        "TypeOf x Is Foo And y",
    ] {
        assert_eq!(parse(&in_if(cond)).found, vec![], "{cond}");
    }
}

#[test]
fn typeof_operand_must_be_a_target() {
    for cond in ["TypeOf (x) Is Foo", "TypeOf 1 Is Foo", "TypeOf \"a\" Is Foo", "TypeOf Nothing Is Foo"] {
        let src = in_if(cond);
        assert_eq!(parse(&src).found[0].1, offset_of(&src, "TypeOf", 0) + 7, "{cond}");
    }
}

#[test]
fn typeof_type_is_never_a_value_type_or_new() {
    for ty in ["Long", "Integer", "String", "Variant", "Date", "Currency", "Byte", "Boolean", "Single", "Double", "New Foo"] {
        let src = in_if(&format!("TypeOf x Is {ty}"));
        let found = parse(&src).found;
        assert_eq!(found[0].1, offset_of(&src, "Is ", 0) + 3, "{ty}");
    }
}

#[test]
fn typeof_node_keeps_operand_and_type() {
    let p = parse("Sub S(x)\n    y = TypeOf x Is Foo\nEnd Sub\n");
    assert_eq!(p.found, vec![]);
    let body = p.proc_body(0);
    let ExprNode::Assign { value, .. } = p.node(body[0]) else { panic!("assign") };
    let ExprNode::TypeOf { expr, type_spec } = p.node(*value) else { panic!("typeof") };
    let (e, t): (NodeId, NodeId) = (*expr, *type_spec);
    assert!(matches!(p.node(e), ExprNode::NameRef { .. }));
    assert!(matches!(p.node(t), ExprNode::UserType { child: None, .. }));
}
