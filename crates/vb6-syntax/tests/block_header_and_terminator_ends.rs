//! A block header or terminator ends its statement: anything further on the
//! same line is reported at the first extra token.

mod common;
use common::{offset_of, parse};
use vb6_syntax::frontend::ast::ExprNode;

const EXPECTED: u32 = 0x9c6f;

fn extra_token(body: &str) {
    let src = format!("Sub S()\n{body}\nEnd Sub\n");
    let at = offset_of(&src, "JUNK", 0);
    let p = parse(&src);
    assert_eq!(p.found, vec![(EXPECTED, at, 4)], "{body:?}");
    assert_eq!(p.top.len(), 1, "{body:?}");
    assert!(matches!(p.node(p.top[0]), ExprNode::ProcDecl { .. }), "{body:?}");
}

#[test]
fn headers_end_their_statement() {
    extra_token("For i = 1 To 2 JUNK\nNext");
    extra_token("For i = 1 To 2 Step 1 JUNK\nNext");
    extra_token("For Each i In c JUNK\nNext");
    extra_token("Do While x JUNK\nLoop");
    extra_token("Do Until x JUNK\nLoop");
    extra_token("While x JUNK\nWend");
    extra_token("With x JUNK\nEnd With");
    extra_token("Select Case x JUNK\nEnd Select");
    extra_token("Select Case x\nCase 1 JUNK\nEnd Select");
}

#[test]
fn terminators_end_their_statement() {
    extra_token("If x Then\nEnd If JUNK");
    extra_token("While x\nWend JUNK");
    extra_token("Do\nLoop JUNK");
    extra_token("Do\nLoop While x JUNK");
    extra_token("With x\nEnd With JUNK");
    extra_token("Select Case x\nEnd Select JUNK");
}

#[test]
fn headers_and_terminators_accept_a_colon_and_a_comment() {
    for body in [
        "For i = 1 To 2: x = i: Next",
        "For i = 1 To 2 ' note\nNext ' note",
        "Do While x: Loop",
        "While x ' c\nWend ' c",
        "With x: End With",
        "Select Case x ' c\nCase 1 ' c\nEnd Select ' c",
        "If x Then ' c\nEnd If ' c",
    ] {
        let p = parse(&format!("Sub S()\n{body}\nEnd Sub\n"));
        assert_eq!(p.found, vec![], "{body:?}");
    }
}
