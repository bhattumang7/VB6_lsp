//! A bare `End` (terminate the program) is a statement, not a block
//! terminator: only `End If` / `End With` / `End Select` / … close a block.

mod common;
use common::parse;
use vb6_syntax::frontend::ast::ExprNode;

#[test]
fn bare_end_inside_block_if_with_trailing_comment() {
    let p = parse(
        "Public Sub QuitIfAsked(ByVal bQuit As Boolean)\n    If bQuit Then\n        End ' Terminate program\n    End If\nEnd Sub\n",
    );
    assert_eq!(p.diags, 0);
    let body = p.proc_body(0);
    assert_eq!(body.len(), 1);
    let ExprNode::If { then_body, else_body, .. } = p.node(body[0]) else { panic!() };
    assert!(else_body.is_none());
    let then = p.block(*then_body);
    assert_eq!(then.len(), 1);
    assert!(matches!(p.node(then[0]), ExprNode::EndStmt));
}

#[test]
fn bare_end_in_every_block_kind() {
    let p = parse(
        "Sub A()\n\
         For i = 1 To 3\n End\n Next\n\
         While x\n End\n Wend\n\
         Do\n End\n Loop\n\
         With y\n End\n End With\n\
         Select Case x\n Case 1: End\n Case 2\n End\n End Select\n\
         If a Then\n End\n Else\n End\n End If\n\
         End\n\
         End Sub\n",
    );
    assert_eq!(p.diags, 0);
    let body = p.proc_body(0);
    assert_eq!(body.len(), 7);
    let is_end = |id| matches!(p.node(id), ExprNode::EndStmt);
    let ExprNode::For { body: b, .. } = p.node(body[0]) else { panic!() };
    assert!(p.block(*b).iter().all(|n| is_end(*n)) && p.block(*b).len() == 1);
    let ExprNode::While { body: b, .. } = p.node(body[1]) else { panic!() };
    assert!(p.block(*b).len() == 1 && is_end(p.block(*b)[0]));
    let ExprNode::Do { body: b, .. } = p.node(body[2]) else { panic!() };
    assert!(p.block(*b).len() == 1 && is_end(p.block(*b)[0]));
    let ExprNode::With { body: b, .. } = p.node(body[3]) else { panic!() };
    assert!(p.block(*b).len() == 1 && is_end(p.block(*b)[0]));
    let ExprNode::SelectCase { cases, .. } = p.node(body[4]) else { panic!() };
    assert_eq!(cases.len(), 2);
    for c in cases {
        let ExprNode::CaseBlock { body: b, .. } = p.node(*c) else { panic!() };
        assert!(p.block(*b).len() == 1 && is_end(p.block(*b)[0]));
    }
    let ExprNode::If { then_body, else_body, .. } = p.node(body[5]) else { panic!() };
    assert!(is_end(p.block(*then_body)[0]));
    assert!(is_end(p.block(else_body.unwrap())[0]));
    assert!(is_end(body[6]));
}

#[test]
fn bare_end_before_colon_and_single_line_else() {
    let p = parse("Sub A()\n While x\n End: Wend\n If a Then End Else b = 1\nEnd Sub\n");
    assert_eq!(p.diags, 0);
    let body = p.proc_body(0);
    assert_eq!(body.len(), 2);
    let ExprNode::While { body: b, .. } = p.node(body[0]) else { panic!() };
    assert!(matches!(p.node(p.block(*b)[0]), ExprNode::EndStmt));
    let ExprNode::If { then_body, else_body, .. } = p.node(body[1]) else { panic!() };
    assert!(matches!(p.node(p.block(*then_body)[0]), ExprNode::EndStmt));
    assert!(matches!(p.node(p.block(else_body.unwrap())[0]), ExprNode::Assign { .. }));
}
