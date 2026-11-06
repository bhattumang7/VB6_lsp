//! The tree of the graphics method statements (`Print`, `Circle`, `Line`, `PSet`, `Scale`, and `?`)
//! as members of an object, after `Call`, and bare: which node is the callee, which nodes are the
//! points, the omitted operands and the box flag.

mod common;
use common::{parse, parse_kind, Parsed};
use vb6_syntax::frontend::ast::{AstLit, ExprNode, NodeId};
use vb6_syntax::frontend::parser::ModuleKind;

/// The expression `id` written out: names, member accesses (`!` for a bang, `[x]` for a
/// bracketed member, `WITH` for the object of a `With`), literals, points (`pt(x,y)`,
/// `step(x,y)`), omitted operands (`_`) and box flags.
fn show(p: &Parsed, id: NodeId) -> String {
    match p.node(id) {
        ExprNode::NameRef { sym, .. } => p.name(*sym).to_string(),
        ExprNode::Me => "Me".into(),
        ExprNode::WithContext => "WITH".into(),
        ExprNode::MemberAccess { base, member, bang, bracketed } => {
            let name = if *bracketed { format!("[{}]", p.name(*member)) } else { p.name(*member).to_string() };
            format!("{}{}{}", show(p, *base), if *bang { "!" } else { "." }, name)
        }
        ExprNode::Literal { lit: AstLit::Int(n) } => n.to_string(),
        ExprNode::Literal { lit: AstLit::Str(s) } => format!("{s:?}"),
        ExprNode::GraphicsPoint { step, x, y } => {
            format!("{}({},{})", if *step { "step" } else { "pt" }, show(p, *x), show(p, *y))
        }
        ExprNode::MissingArg => "_".into(),
        ExprNode::LineBox { filled } => if *filled { "BF" } else { "B" }.into(),
        ExprNode::Paren { inner } => format!("({})", show(p, *inner)),
        ExprNode::Call { func, args } => format!("{}{}", show(p, *func), show(p, *args)),
        ExprNode::ArgList { args } => {
            format!("[{}]", args.iter().map(|a| show(p, *a)).collect::<Vec<_>>().join(", "))
        }
        ExprNode::CallStmt { callee, args } => format!("stmt {} {}", show(p, *callee), show(p, *args)),
        other => panic!("not shown: {}", std::any::type_name_of_val(other)),
    }
}

fn statements(src: &str, kind: ModuleKind) -> (Parsed, Vec<NodeId>) {
    let p = parse_kind(src, kind);
    let body = p.proc_body(0);
    (p, body)
}

/// The shape of the only statement of `Sub S(...)` with `stmt` as its body; no diagnostics.
fn one(stmt: &str) -> String {
    let (p, body) = statements(&format!("Sub S(o As Object)\n    {stmt}\nEnd Sub\n"), ModuleKind::Class);
    assert_eq!(p.found, vec![], "{stmt}");
    assert_eq!(body.len(), 1, "{stmt}");
    show(&p, body[0])
}

#[test]
fn pset_takes_one_point_and_a_colour() {
    assert_eq!(one("o.PSet (1, 2)"), "stmt o.PSet [pt(1,2)]");
    assert_eq!(one("o.PSet Step(1, 2)"), "stmt o.PSet [step(1,2)]");
    assert_eq!(one("o.PSet Step (1, 2), 3"), "stmt o.PSet [step(1,2), 3]");
    assert_eq!(one("o.PSet (1, 2), (3)"), "stmt o.PSet [pt(1,2), (3)]");
}

#[test]
fn circle_takes_a_point_a_radius_and_four_optional_operands() {
    assert_eq!(one("o.Circle (1, 2), 3"), "stmt o.Circle [pt(1,2), 3]");
    assert_eq!(one("o.Circle Step(1, 2), 3, 4"), "stmt o.Circle [step(1,2), 3, 4]");
    assert_eq!(one("o.Circle (1, 2), 3, , 5"), "stmt o.Circle [pt(1,2), 3, _, 5]");
    assert_eq!(one("o.Circle (1, 2), 3, , , , 6"), "stmt o.Circle [pt(1,2), 3, _, _, _, 6]");
    assert_eq!(one("o.Circle (1, 2), 3, 4, 5, 6, 7".replace(", 7", "").as_str()), "stmt o.Circle [pt(1,2), 3, 4, 5, 6]");
}

#[test]
fn line_takes_two_points_a_colour_and_a_box_flag() {
    assert_eq!(one("o.Line (1, 2)-(3, 4)"), "stmt o.Line [pt(1,2), pt(3,4)]");
    assert_eq!(one("o.Line Step(1, 2)-Step(3, 4)"), "stmt o.Line [step(1,2), step(3,4)]");
    assert_eq!(one("o.Line (1, 2)-Step(3, 4), 5"), "stmt o.Line [pt(1,2), step(3,4), 5]");
    // Without the first point the statement starts at the current position.
    assert_eq!(one("o.Line -(3, 4)"), "stmt o.Line [_, pt(3,4)]");
    assert_eq!(one("o.Line -Step(3, 4), , B"), "stmt o.Line [_, step(3,4), _, B]");
    assert_eq!(one("o.Line (1, 2)-(3, 4), 5, B"), "stmt o.Line [pt(1,2), pt(3,4), 5, B]");
    assert_eq!(one("o.Line (1, 2)-(3, 4), 5, BF"), "stmt o.Line [pt(1,2), pt(3,4), 5, BF]");
    assert_eq!(one("o.Line (1, 2)-(3, 4), , b"), "stmt o.Line [pt(1,2), pt(3,4), _, B]");
}

#[test]
fn the_two_words_b_f_are_the_filled_box() {
    assert_eq!(one("o.Line (1, 2)-(3, 4), 5, B F"), "stmt o.Line [pt(1,2), pt(3,4), 5, BF]");
    assert_eq!(one("o.Line (1, 2)-(3, 4), 5, [B] [F]"), "stmt o.Line [pt(1,2), pt(3,4), 5, BF]");
    assert_eq!(one("o.Line (1, 2)-(3, 4), 5, [BF]"), "stmt o.Line [pt(1,2), pt(3,4), 5, BF]");
    assert_eq!(one("o.Line (1, 2)-(3, 4), 5, [B]"), "stmt o.Line [pt(1,2), pt(3,4), 5, B]");
}

#[test]
fn scale_is_bare_or_takes_two_points() {
    assert_eq!(one("o.Scale"), "o.Scale");
    assert_eq!(one("o.Scale (1, 2)-(3, 4)"), "stmt o.Scale [pt(1,2), pt(3,4)]");
    assert_eq!(one("o.Scale (a, o)-(1, 1)"), "stmt o.Scale [pt(a,o), pt(1,1)]");
}

#[test]
fn print_takes_a_list_of_items() {
    assert_eq!(one("o.Print"), "o.Print");
    assert_eq!(one("o.Print 1"), "stmt o.Print [1]");
    assert_eq!(one("o.Print 1; 2"), "stmt o.Print [1, 2]");
    assert_eq!(one("o.Print (1), 2"), "stmt o.Print [(1), 2]");
}

#[test]
fn the_question_mark_is_the_print_member() {
    assert_eq!(one("o.? 1; 2"), "stmt o.Print [1, 2]");
    assert_eq!(one("o.?"), "o.Print");
    assert_eq!(one("o.p.? 1"), "stmt o.p.Print [1]");
    assert_eq!(one("Call o.? 1"), "stmt o.Print [1]");
    // Inside a With block the leading dot gives the member the object of the block.
    let p = parse("Sub S(o As Object)\n    With o\n        .? \"a\"; 1\n        .?\n    End With\nEnd Sub\n");
    assert_eq!(p.found, vec![]);
    let ExprNode::With { body, .. } = p.node(p.proc_body(0)[0]) else { panic!("with block") };
    let stmts = p.block(*body);
    assert_eq!(show(&p, stmts[0]), "stmt WITH.Print [\"a\", 1]");
    assert_eq!(show(&p, stmts[1]), "WITH.Print");
}

#[test]
fn call_with_a_graphics_member_takes_the_same_operands() {
    assert_eq!(one("Call o.Line (1, 2)-(3, 4), 5, B"), "stmt o.Line [pt(1,2), pt(3,4), 5, B]");
    assert_eq!(one("Call o.Circle (1, 2), 3"), "stmt o.Circle [pt(1,2), 3]");
    assert_eq!(one("Call o.PSet Step(1, 2)"), "stmt o.PSet [step(1,2)]");
    assert_eq!(one("Call o.Scale"), "stmt o.Scale []");
    assert_eq!(one("Call o.Scale (1, 2)-(3, 4)"), "stmt o.Scale [pt(1,2), pt(3,4)]");
    assert_eq!(one("Call o.Print"), "stmt o.Print []");
    assert_eq!(one("Call o.Print 1; 2"), "stmt o.Print [1, 2]");
}

#[test]
fn the_object_may_be_any_chain() {
    assert_eq!(one("o.p.Line (1, 2)-(3, 4)"), "stmt o.p.Line [pt(1,2), pt(3,4)]");
    assert_eq!(one("o.p(1).Print 2"), "stmt o.p[1].Print [2]");
    assert_eq!(one("o!p.Circle (1, 2), 3"), "stmt o!p.Circle [pt(1,2), 3]");
    assert_eq!(one("Me.PSet (1, 2)"), "stmt Me.PSet [pt(1,2)]");
}

#[test]
fn a_graphics_word_that_is_not_a_method_call_is_an_ordinary_member() {
    // Bracketed, after a bang, or with the object of a With: an ordinary call.
    assert_eq!(one("o.[Print] 1, 2"), "stmt o.[Print] [1, 2]");
    assert_eq!(one("o.[Circle] 1"), "stmt o.[Circle] [1]");
    let p = parse_kind("Sub S(o As Object)\n    o!Circle = 1\n    With o\n        .Print 1, 2\n        .Line = 3\n    End With\nEnd Sub\n", ModuleKind::Class);
    assert_eq!(p.found, vec![]);
    let first = p.proc_body(0)[0];
    let ExprNode::Assign { target, .. } = p.node(first) else { panic!("assignment") };
    assert_eq!(show(&p, *target), "o!Circle");
    let ExprNode::With { body, .. } = p.node(p.proc_body(0)[1]) else { panic!("with block") };
    let stmts = p.block(*body);
    assert_eq!(show(&p, stmts[0]), "stmt WITH.Print [1, 2]");
    assert!(matches!(p.node(stmts[1]), ExprNode::Assign { .. }));
}

#[test]
fn a_member_chain_through_a_graphics_word_continues() {
    for (stmt, target) in [
        ("o.Circle.x = 1", "o.Circle.x"),
        ("o.Line.x = 1", "o.Line.x"),
        ("o.Scale.x = 1", "o.Scale.x"),
        ("o.PSet.x = 1", "o.PSet.x"),
        ("o.Line = 1", "o.Line"),
    ] {
        let (p, body) = statements(&format!("Sub S(o As Object)\n    {stmt}\nEnd Sub\n"), ModuleKind::Class);
        assert_eq!(p.found, vec![], "{stmt}");
        let ExprNode::Assign { target: lhs, .. } = p.node(body[0]) else { panic!("{stmt}: assignment") };
        assert_eq!(show(&p, *lhs), target, "{stmt}");
    }
}

#[test]
fn print_takes_what_follows_it_whatever_the_dot_that_follows() {
    // `o.Print.x = 1` prints the comparison `.x = 1` of the object of an enclosing With.
    let (p, body) = statements("Sub S(o As Object)\n    o.Print.x = 1\nEnd Sub\n", ModuleKind::Class);
    assert_eq!(p.found, vec![]);
    let ExprNode::CallStmt { callee, args } = p.node(body[0]) else { panic!("statement") };
    assert_eq!(show(&p, *callee), "o.Print");
    let ExprNode::ArgList { args } = p.node(*args) else { panic!() };
    assert_eq!(args.len(), 1);
    let ExprNode::BinOp { lhs, rhs, .. } = p.node(args[0]) else { panic!("comparison") };
    assert_eq!(show(&p, *lhs), "WITH.x");
    assert_eq!(show(&p, *rhs), "1");
}

#[test]
fn bare_statements_have_a_name_as_their_callee() {
    let (p, body) = statements(
        "Sub S()\n    Circle Step(1, 2), 3\n    Line -(3, 4), , BF\n    PSet (1, 2), 3\n    Scale (1, 2)-(3, 4)\n    Scale\nEnd Sub\n",
        ModuleKind::Class,
    );
    assert_eq!(p.found, vec![]);
    let shown: Vec<String> = body.iter().map(|s| show(&p, *s)).collect();
    assert_eq!(
        shown,
        vec![
            "stmt Circle [step(1,2), 3]",
            "stmt Line [_, pt(3,4), _, BF]",
            "stmt PSet [pt(1,2), 3]",
            "stmt Scale [pt(1,2), pt(3,4)]",
            "Scale",
        ]
    );
}

#[test]
fn a_bracketed_member_name_is_marked_in_the_tree() {
    let (p, body) = statements("Sub S(o As Object)\n    o.[Print] 1\n    o.Print 2\nEnd Sub\n", ModuleKind::Class);
    assert_eq!(p.found, vec![]);
    let bracketed = |stmt: NodeId| {
        let ExprNode::CallStmt { callee, .. } = p.node(stmt) else { panic!("statement") };
        let ExprNode::MemberAccess { bang, bracketed, .. } = p.node(*callee) else { panic!("member") };
        (*bang, *bracketed)
    };
    assert_eq!(bracketed(body[0]), (false, true));
    assert_eq!(bracketed(body[1]), (false, false));
}
