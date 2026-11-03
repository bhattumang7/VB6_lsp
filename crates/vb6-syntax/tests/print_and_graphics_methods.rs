//! `Print` with `;` separators, and the graphics methods `Line` / `Circle` /
//! `PSet` with `Step` points and the `B` / `BF` box flags, written as methods
//! of an object (`Frm.Print …`, `Frm.Line …`) or on the enclosing form.

mod common;
use common::{parse, Parsed};
use vb6_syntax::frontend::ast::{ExprNode, NodeId};

/// The statement's callee member name and its argument nodes.
fn call_stmt(p: &Parsed, id: NodeId) -> (String, Vec<NodeId>) {
    let ExprNode::CallStmt { callee, args } = p.node(id) else { panic!("expected CallStmt") };
    let name = match p.node(*callee) {
        ExprNode::MemberAccess { member, .. } => p.name(*member).to_string(),
        ExprNode::NameRef { sym, .. } => p.name(*sym).to_string(),
        ExprNode::Call { func, .. } => match p.node(*func) {
            ExprNode::MemberAccess { member, .. } => p.name(*member).to_string(),
            ExprNode::NameRef { sym, .. } => p.name(*sym).to_string(),
            _ => panic!(),
        },
        _ => panic!("callee"),
    };
    let ExprNode::ArgList { args } = p.node(*args) else { panic!() };
    (name, args.clone())
}

fn kinds(p: &Parsed, args: &[NodeId]) -> String {
    args.iter()
        .map(|a| match p.node(*a) {
            ExprNode::MissingArg => "_".to_string(),
            ExprNode::LineBox { filled: false } => "B".into(),
            ExprNode::LineBox { filled: true } => "BF".into(),
            ExprNode::GraphicsPoint { step: true, .. } => "step".into(),
            ExprNode::GraphicsPoint { step: false, .. } => "point".into(),
            ExprNode::NameRef { sym, .. } => p.name(*sym).to_string(),
            ExprNode::Literal { .. } => "lit".into(),
            ExprNode::BinOp { .. } => "binop".into(),
            ExprNode::Paren { .. } => "paren".into(),
            ExprNode::Call { .. } => "call".into(),
            _ => "other".into(),
        })
        .collect::<Vec<_>>()
        .join(" ")
}

#[test]
fn print_method_semicolons_on_every_receiver() {
    let p = parse(
        "Public Sub PrintTo(ByVal OutObj As Object, ByVal Frm As Form, ByVal word As String)\n\
         Printer.Print \" \" + word$;\n\
         Printer.Print \"A\"; word\n\
         Frm.Print \" \" + word$;\n\
         Frm.Print \"A\"; word\n\
         OutObj.Print \" \" + word$;\n\
         OutObj.Print \"A\"; word\n\
         OutObj.Print (\" \" + word$);\n\
         End Sub\n",
    );
    assert_eq!(p.diags, 0);
    let body = p.proc_body(0);
    assert_eq!(body.len(), 7);
    let shapes: Vec<_> = body
        .iter()
        .map(|s| {
            let (name, args) = call_stmt(&p, *s);
            assert_eq!(name, "Print");
            kinds(&p, &args)
        })
        .collect();
    assert_eq!(
        shapes,
        vec!["binop", "lit word", "binop", "lit word", "binop", "lit word", "paren"]
    );
}

#[test]
fn print_method_mixed_separators_and_parenthesised_first_item() {
    let p = parse(
        "Sub P(F As Form)\n F.Print (a), (b);\n F.Print (a) + 1; (b)\n F.Print , x; y,\n F.Print\n F.Print (a)\nEnd Sub\n",
    );
    assert_eq!(p.diags, 0);
    let body = p.proc_body(0);
    assert_eq!(body.len(), 5);
    assert_eq!(kinds(&p, &call_stmt(&p, body[0]).1), "paren paren");
    assert_eq!(kinds(&p, &call_stmt(&p, body[1]).1), "binop paren");
    assert_eq!(kinds(&p, &call_stmt(&p, body[2]).1), "x y");
    // A bare method name is just the member access; a lone parenthesised item
    // stays a call of it.
    assert!(matches!(p.node(body[3]), ExprNode::MemberAccess { .. }));
    assert!(matches!(p.node(body[4]), ExprNode::Call { .. }));
}

#[test]
fn line_method_box_flags_and_omitted_arguments() {
    let p = parse(
        "Sub D(Frm As Form, OutObj As Object)\n\
         Frm.Line (0, 0)-(100, 100), , B\n\
         Frm.Line (0, 0)-(100, 100), vbRed, BF\n\
         OutObj.Line (0, 0)-(100, 100), , B\n\
         End Sub\n",
    );
    assert_eq!(p.diags, 0);
    let body = p.proc_body(0);
    assert_eq!(body.len(), 3);
    // The first point is the call's own argument list; the rest follow.
    let ExprNode::CallStmt { callee, .. } = p.node(body[0]) else { panic!() };
    assert!(matches!(p.node(*callee), ExprNode::Call { .. }));
    assert_eq!(kinds(&p, &call_stmt(&p, body[0]).1), "lit lit _ B");
    assert_eq!(kinds(&p, &call_stmt(&p, body[1]).1), "lit lit vbRed BF");
    assert_eq!(kinds(&p, &call_stmt(&p, body[2]).1), "lit lit _ B");
    assert_eq!(call_stmt(&p, body[2]).0, "Line");
}

#[test]
fn step_points_in_every_position() {
    let p = parse(
        "Sub D(Frm As Form)\n\
         Frm.Line Step(1, 1)-Step(10, 10)\n\
         Frm.Line -Step(10, 10)\n\
         Frm.Line -(10, 10), vbRed\n\
         Frm.Line (0, 0)-Step(5, 5), , BF\n\
         Frm.Circle Step(5, 5), 10\n\
         Frm.PSet Step(1, 1)\n\
         Frm.PSet Step(1, 1), vbBlue\n\
         End Sub\n",
    );
    assert_eq!(p.diags, 0);
    let body = p.proc_body(0);
    assert_eq!(body.len(), 7);
    let shape = |i: usize| kinds(&p, &call_stmt(&p, body[i]).1);
    assert_eq!(shape(0), "step step");
    assert_eq!(shape(1), "_ step");
    assert_eq!(shape(2), "_ point vbRed");
    assert_eq!(shape(3), "step _ BF");
    assert_eq!(shape(4), "step lit");
    assert_eq!(shape(5), "step");
    assert_eq!(shape(6), "step vbBlue");
    // A stepped point keeps its coordinates.
    let (_, args) = call_stmt(&p, body[0]);
    let ExprNode::GraphicsPoint { x, y, .. } = p.node(args[1]) else { panic!() };
    let lit = |n: NodeId| match p.node(n) {
        ExprNode::Literal { lit } => format!("{lit:?}"),
        _ => panic!(),
    };
    assert_eq!((lit(*x), lit(*y)), ("Int(10)".to_string(), "Int(10)".to_string()));
}

#[test]
fn bare_line_statement_uses_the_same_forms() {
    let p = parse("Sub D()\n Line (0, 0)-(1, 1), vbRed, B\n Line Step(1, 1)-Step(2, 2)\n Line -(5, 5)\nEnd Sub\n");
    assert_eq!(p.diags, 0);
    let body = p.proc_body(0);
    assert_eq!(body.len(), 3);
    assert_eq!(call_stmt(&p, body[0]).0, "Line");
    assert_eq!(kinds(&p, &call_stmt(&p, body[0]).1), "lit lit vbRed B");
    assert_eq!(kinds(&p, &call_stmt(&p, body[1]).1), "step step");
    assert_eq!(kinds(&p, &call_stmt(&p, body[2]).1), "_ point");
}

#[test]
fn b_is_an_ordinary_name_outside_the_box_flag_position() {
    let p = parse("Sub D(Frm As Form)\n Frm.Line (0, 0)-(1, 1), B\n Frm.Circle (1, 1), B\nEnd Sub\n");
    assert_eq!(p.diags, 0);
    let body = p.proc_body(0);
    // `B` is the colour here (position 0) and the radius for Circle.
    assert_eq!(kinds(&p, &call_stmt(&p, body[0]).1), "lit lit B");
    assert_eq!(kinds(&p, &call_stmt(&p, body[1]).1), "B");
    let (_, args) = call_stmt(&p, body[0]);
    assert!(matches!(p.node(args[2]), ExprNode::NameRef { .. }));
}
