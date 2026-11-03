//! `ReDim` of a member array inside a `With` block: `ReDim .Items(n)`.

mod common;
use common::parse;
use vb6_syntax::frontend::ast::ExprNode;

fn redim_items(p: &common::Parsed, stmt: vb6_syntax::frontend::ast::NodeId) -> Vec<vb6_syntax::frontend::ast::NodeId> {
    p.block(stmt)
}

#[test]
fn redim_and_redim_preserve_on_with_member() {
    let p = parse(
        "Sub Reset(ByVal n As Long)\n    With mHolder\n        ReDim .Items(n)\n        ReDim Preserve .Items(n), .Other(1, 2)\n    End With\nEnd Sub\n",
    );
    assert_eq!(p.diags, 0);
    let body = p.proc_body(0);
    let ExprNode::With { body: wb, .. } = p.node(body[0]) else { panic!("With") };
    let stmts = p.block(*wb);
    assert_eq!(stmts.len(), 2);

    let first = redim_items(&p, stmts[0]);
    assert_eq!(first.len(), 1);
    let ExprNode::ReDimItem { preserve, base, name, bounds, type_node } = p.node(first[0]) else { panic!() };
    assert!(!preserve);
    assert_eq!(p.name(*name), "Items");
    assert!(matches!(p.node(base.unwrap()), ExprNode::WithContext));
    let ExprNode::ArgList { args } = p.node(bounds.unwrap()) else { panic!() };
    assert_eq!(args.len(), 1);
    assert!(matches!(p.node(args[0]), ExprNode::NameRef { .. }));
    assert!(type_node.is_none());

    let second = redim_items(&p, stmts[1]);
    assert_eq!(second.len(), 2);
    let names: Vec<_> = second
        .iter()
        .map(|i| {
            let ExprNode::ReDimItem { preserve, base, name, bounds, .. } = p.node(*i) else { panic!() };
            assert!(*preserve);
            assert!(matches!(p.node(base.unwrap()), ExprNode::WithContext));
            let ExprNode::ArgList { args } = p.node(bounds.unwrap()) else { panic!() };
            (p.name(*name).to_string(), args.len())
        })
        .collect();
    assert_eq!(names, vec![("Items".to_string(), 1), ("Other".to_string(), 2)]);
}

#[test]
fn redim_with_member_chain_and_element() {
    let p = parse("Sub R()\n With o\n  ReDim .H.Items(3)\n  ReDim .Arr(2).Items(3)\n End With\nEnd Sub\n");
    assert_eq!(p.diags, 0);
    let ExprNode::With { body: wb, .. } = p.node(p.proc_body(0)[0]) else { panic!() };
    let stmts = p.block(*wb);
    let ExprNode::ReDimItem { base, name, .. } = p.node(p.block(stmts[0])[0]) else { panic!() };
    assert_eq!(p.name(*name), "Items");
    let ExprNode::MemberAccess { base: inner, member, .. } = p.node(base.unwrap()) else { panic!() };
    assert_eq!(p.name(*member), "H");
    assert!(matches!(p.node(*inner), ExprNode::WithContext));
    let ExprNode::ReDimItem { base, name, .. } = p.node(p.block(stmts[1])[0]) else { panic!() };
    assert_eq!(p.name(*name), "Items");
    let ExprNode::Call { func, .. } = p.node(base.unwrap()) else { panic!() };
    let ExprNode::MemberAccess { base: inner, member, .. } = p.node(*func) else { panic!() };
    assert_eq!(p.name(*member), "Arr");
    assert!(matches!(p.node(*inner), ExprNode::WithContext));
}
