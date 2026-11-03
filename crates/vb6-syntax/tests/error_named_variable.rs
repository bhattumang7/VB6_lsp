//! A variable named `Error` shadows the `Error <number>` statement.

mod common;
use common::parse;
use vb6_syntax::frontend::ast::ExprNode;

#[test]
fn local_named_error_assignment_and_comparison() {
    let p = parse(
        "Public Function GetCode() As Long\n    Dim Error As Long\n    Error = 5\n    If Error <> 0 Then GetCode = Error\nEnd Function\n",
    );
    assert_eq!(p.diags, 0);
    let body = p.proc_body(0);
    assert_eq!(body.len(), 3);
    let dims = p.block(body[0]);
    assert_eq!(dims.len(), 1);
    assert!(matches!(p.node(dims[0]), ExprNode::DimItem { .. }));
    let ExprNode::Assign { target, value } = p.node(body[1]) else { panic!("expected Assign") };
    assert!(matches!(p.node(*target), ExprNode::NameRef { .. }));
    assert!(matches!(p.node(*value), ExprNode::Literal { .. }));
    let ExprNode::If { cond, then_body, else_body } = p.node(body[2]) else { panic!() };
    assert!(else_body.is_none());
    assert!(matches!(p.node(*cond), ExprNode::BinOp { .. }));
    let then = p.block(*then_body);
    assert_eq!(then.len(), 1);
    let ExprNode::Assign { value, .. } = p.node(then[0]) else { panic!() };
    assert!(matches!(p.node(*value), ExprNode::NameRef { .. }));
}

#[test]
fn error_member_and_parameter_and_array() {
    let p = parse(
        "Sub A(Error As Long)\n Error = 1\n Error(2) = 3\nEnd Sub\n\
         Sub B()\n Dim Error(3) As Long\n Error(2) = 3\n Error.Foo = 1\nEnd Sub\n",
    );
    assert_eq!(p.diags, 0);
    for (proc_idx, count) in [(0, 2), (1, 3)] {
        let body = p.proc_body(proc_idx);
        assert_eq!(body.len(), count);
        for n in body.iter().skip(count - 2) {
            assert!(matches!(p.node(*n), ExprNode::Assign { .. }));
        }
    }
}

#[test]
fn error_statement_still_parses_when_not_shadowed() {
    let p = parse("Sub A()\n Error 5\n Error (7)\nEnd Sub\nSub B()\n Dim Error As Long\nEnd Sub\nSub C()\n Error (9)\nEnd Sub\n");
    assert_eq!(p.diags, 0);
    let a = p.proc_body(0);
    assert_eq!(a.len(), 2);
    assert!(a.iter().all(|n| matches!(p.node(*n), ExprNode::ErrorStmt { .. })));
    // `Error` declared in B() does not leak into C().
    let c = p.proc_body(2);
    assert!(matches!(p.node(c[0]), ExprNode::ErrorStmt { .. }));
}
