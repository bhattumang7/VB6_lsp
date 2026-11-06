//! What may stand between `Select Case` and the first `Case`: declarations, a bare `Close`, an
//! `Exit For`, and the graphics methods of an object are allowed; the other statements are
//! reported at their own line.

mod common;
use common::{offset_of, parse, Parsed};
use vb6_syntax::frontend::ast::{ExitKind, ExprNode, FileIoKind, NodeId};
use vb6_syntax::frontend::diagnostics::{
    ERR_EXIT_FOR_OUTSIDE_FOR, ERR_LOOP_WITHOUT_DO, ERR_NEXT_WITHOUT_FOR, ERR_STATEMENT_BEFORE_FIRST_CASE as PRE,
};

/// Parses `body` as the statements of a Sub and returns the parse with the preamble nodes and the
/// number of `Case` clauses of its first statement, which must be a `Select Case`.
fn select(src: &str) -> (Parsed, Vec<NodeId>, usize) {
    let p = parse(src);
    let body = p.proc_body(0);
    let select = body
        .iter()
        .copied()
        .find(|n| matches!(p.node(*n), ExprNode::SelectCase { .. }))
        .expect("a Select Case statement");
    let ExprNode::SelectCase { pre, cases, .. } = p.node(select) else { unreachable!() };
    let (pre, cases) = (pre.clone(), cases.len());
    (p, pre, cases)
}

fn sub(preamble: &str) -> String {
    format!("Sub S()\nSelect Case x\n{preamble}\nCase 1\nEnd Select\nEnd Sub\n")
}

/// The items of a declaration statement (a `Block` of items).
fn items(p: &Parsed, stmt: NodeId) -> Vec<NodeId> {
    let ExprNode::Block { stmts } = p.node(stmt) else { panic!("declaration block") };
    stmts.clone()
}

struct Item {
    name: String,
    is_const: bool,
    is_static: bool,
    is_new: bool,
    bounds: bool,
    ty: Option<u32>,
    init: bool,
}

fn dim_item(p: &Parsed, id: NodeId) -> Item {
    let ExprNode::DimItem { name, is_const, is_static, is_new, with_events, bounds, type_node, init } = p.node(id) else {
        panic!("DimItem")
    };
    assert!(!with_events);
    Item {
        name: p.name(*name).to_ascii_lowercase(),
        is_const: *is_const,
        is_static: *is_static,
        is_new: *is_new,
        bounds: bounds.is_some(),
        ty: type_node.map(|t| match p.node(t) {
            ExprNode::BuiltinType { kind } => *kind,
            ExprNode::UserType { .. } => USER_TYPE,
            _ => 0,
        }),
        init: init.is_some(),
    }
}

const LONG: u32 = 3;
const INTEGER: u32 = 2;
/// Stand-in kind for a named (non built-in) type in the checks above.
const USER_TYPE: u32 = 1000;

#[test]
fn a_dim_between_select_case_and_the_first_case_is_accepted() {
    let (p, pre, cases) = select(&sub("Dim y As Long"));
    assert_eq!(p.found, vec![]);
    assert_eq!(cases, 1);
    assert_eq!(pre.len(), 1);
    let it = items(&p, pre[0]);
    assert_eq!(it.len(), 1);
    let d = dim_item(&p, it[0]);
    assert_eq!((d.name.as_str(), d.is_const, d.is_static, d.is_new, d.bounds, d.ty, d.init), ("y", false, false, false, false, Some(LONG), false));
}

#[test]
fn a_dim_list_with_an_array_and_a_new_object_is_accepted() {
    let (p, pre, _) = select(&sub("Dim a As Long, b(3) As Integer\nDim o As New Collection"));
    assert_eq!(p.found, vec![]);
    assert_eq!(pre.len(), 2);
    let first = items(&p, pre[0]);
    assert_eq!(first.len(), 2);
    let a = dim_item(&p, first[0]);
    let b = dim_item(&p, first[1]);
    assert_eq!((a.name.to_ascii_lowercase().as_str(), a.bounds, a.ty), ("a", false, Some(LONG)));
    assert_eq!((b.name.to_ascii_lowercase().as_str(), b.bounds, b.ty), ("b", true, Some(INTEGER)));
    let second = items(&p, pre[1]);
    assert_eq!(second.len(), 1);
    let o = dim_item(&p, second[0]);
    assert_eq!((o.name.as_str(), o.is_new, o.is_static, o.ty), ("o", true, false, Some(USER_TYPE)));
}

#[test]
fn a_static_declaration_is_accepted() {
    let (p, pre, _) = select(&sub("Static s As Long, t(2) As Integer"));
    assert_eq!(p.found, vec![]);
    let it = items(&p, pre[0]);
    assert_eq!(it.len(), 2);
    let s = dim_item(&p, it[0]);
    let t = dim_item(&p, it[1]);
    assert_eq!((s.name.to_ascii_lowercase().as_str(), s.is_static, s.bounds, s.ty), ("s", true, false, Some(LONG)));
    assert_eq!((t.name.to_ascii_lowercase().as_str(), t.is_static, t.bounds, t.ty), ("t", true, true, Some(INTEGER)));
}

#[test]
fn a_const_declaration_with_several_items_is_accepted() {
    let (p, pre, _) = select(&sub("Const k = 1, j As Long = 2"));
    assert_eq!(p.found, vec![]);
    let it = items(&p, pre[0]);
    assert_eq!(it.len(), 2);
    let k = dim_item(&p, it[0]);
    let j = dim_item(&p, it[1]);
    assert_eq!((k.name.as_str(), k.is_const, k.init, k.ty), ("k", true, true, None));
    assert_eq!((j.name.as_str(), j.is_const, j.init, j.ty), ("j", true, true, Some(LONG)));
}

#[test]
fn a_redim_is_accepted_with_and_without_preserve() {
    let (p, pre, _) = select(&sub("ReDim a(3)\nReDim Preserve b(4) As Long"));
    assert_eq!(p.found, vec![]);
    assert_eq!(pre.len(), 2);
    let redim = |stmt: NodeId| -> Vec<(String, bool, bool, Option<u32>)> {
        let it = match p.node(stmt) {
            ExprNode::Block { stmts } => stmts.clone(),
            _ => vec![stmt],
        };
        it.iter()
            .map(|i| {
                let ExprNode::ReDimItem { preserve, base, name, bounds, type_node } = p.node(*i) else { panic!("ReDimItem") };
                assert!(base.is_none());
                (
                    p.name(*name).to_ascii_lowercase(),
                    *preserve,
                    bounds.is_some(),
                    type_node.map(|t| match p.node(t) {
                        ExprNode::BuiltinType { kind } => *kind,
                        _ => 0,
                    }),
                )
            })
            .collect()
    };
    assert_eq!(redim(pre[0]), vec![("a".to_string(), false, true, None)]);
    assert_eq!(redim(pre[1]), vec![("b".to_string(), true, true, Some(LONG))]);
}

#[test]
fn a_declaration_on_the_select_case_line_is_accepted() {
    let (p, pre, cases) = select("Sub S()\nSelect Case x: Dim y As Long\nCase 1\nEnd Select\nEnd Sub\n");
    assert_eq!(p.found, vec![]);
    assert_eq!((pre.len(), cases), (1, 1));
}

#[test]
fn a_statement_after_a_declaration_is_reported_at_that_statement() {
    let src = sub("Dim y As Long\ny = 1");
    let (p, pre, _) = select(&src);
    assert_eq!(p.found, vec![(PRE, offset_of(&src, "y = 1", 0), 1)]);
    assert_eq!(pre.len(), 2);
    assert!(matches!(p.node(pre[0]), ExprNode::Block { .. }));
    assert!(matches!(p.node(pre[1]), ExprNode::Assign { .. }));
}

#[test]
fn a_statement_after_a_declaration_on_one_line_is_reported_at_that_statement() {
    let src = sub("Dim y As Long: y = 1");
    let (p, pre, _) = select(&src);
    assert_eq!(p.found, vec![(PRE, offset_of(&src, "y = 1", 0), 1)]);
    assert_eq!(pre.len(), 2);
}

#[test]
fn a_label_in_front_of_a_declaration_is_reported_at_the_label() {
    let src = sub("L2: Dim y As Long");
    let (p, pre, _) = select(&src);
    assert_eq!(p.found, vec![(PRE, offset_of(&src, "L2", 0), 2)]);
    assert!(matches!(p.node(pre[0]), ExprNode::Label { .. }));
    assert_eq!(pre.len(), 2);
}

#[test]
fn a_bare_end_is_reported_at_its_own_line() {
    let src = sub("Dim y As Long\nEnd");
    let (p, pre, _) = select(&src);
    assert_eq!(p.found, vec![(PRE, offset_of(&src, "End\n", 0), 3)]);
    assert!(matches!(p.node(pre[1]), ExprNode::EndStmt));
    assert_eq!(src[..offset_of(&src, "End\n", 0) as usize].matches('\n').count() + 1, 4);
}

#[test]
fn a_bare_close_is_accepted_and_a_close_with_channels_is_not() {
    for text in ["Close", "Close ' note", "Close: Close"] {
        let (p, pre, _) = select(&sub(text));
        assert_eq!(p.found, vec![], "{text}");
        for stmt in &pre {
            let ExprNode::FileIoStmt { kind, channel, channel_marks, args, open_mode, open_flags } = p.node(*stmt) else { panic!("{text}") };
            assert!(channel_marks.is_empty(), "{text}");
            assert_eq!((*kind, channel.is_none(), args.len(), *open_mode, *open_flags), (FileIoKind::Close, true, 0, 0, 0), "{text}");
        }
        assert_eq!(pre.len(), if text == "Close: Close" { 2 } else { 1 }, "{text}");
    }
    for text in ["Close #1", "Close 1, 2", "Close #1, #2"] {
        let src = sub(text);
        let (p, pre, _) = select(&src);
        assert_eq!(p.found, vec![(PRE, offset_of(&src, "Close", 0), 5)], "{text}");
        assert!(matches!(p.node(pre[0]), ExprNode::FileIoStmt { kind: FileIoKind::Close, .. }), "{text}");
    }
}

#[test]
fn a_bare_close_followed_by_a_statement_reports_that_statement() {
    let src = sub("Close: x = 1");
    let (p, pre, _) = select(&src);
    assert_eq!(p.found, vec![(PRE, offset_of(&src, "x = 1", 0), 1)]);
    assert_eq!(pre.len(), 2);
}

/// Statement shape of a graphics method: the member name, the argument count of a call written
/// on the member itself (`None` when it is not called that way) and the argument count of the
/// statement's own list (`None` when the statement is not a `CallStmt`).
fn graphics(p: &Parsed, id: NodeId) -> (String, Option<usize>, Option<usize>) {
    let count = |args: NodeId| match p.node(args) {
        ExprNode::ArgList { args } => args.len(),
        _ => panic!("ArgList"),
    };
    let (callee, outer) = match p.node(id) {
        ExprNode::CallStmt { callee, args } => (*callee, Some(count(*args))),
        _ => (id, None),
    };
    let (target, inner) = match p.node(callee) {
        ExprNode::Call { func, args } => (*func, Some(count(*args))),
        _ => (callee, None),
    };
    let ExprNode::MemberAccess { member, bang, .. } = p.node(target) else { panic!("member access") };
    assert!(!*bang);
    (p.name(*member).to_string(), inner, outer)
}

#[test]
fn the_graphics_methods_of_an_object_are_accepted() {
    let cases: &[(&str, &str, Option<usize>, Option<usize>)] = &[
        ("f.Circle (1, 1), 2", "circle", None, Some(2)),
        ("f.Circle Step (1, 1), 2, 3, 4, 5, 6", "circle", None, Some(6)),
        ("f.Line (1, 1)-(2, 2)", "line", None, Some(2)),
        ("f.Line -(2, 2)", "line", None, Some(2)),
        ("f.Line (1, 1)-(2, 2), , B", "line", None, Some(4)),
        ("f.Line Step (1, 1)-Step (2, 2), 3, BF", "line", None, Some(4)),
        ("f.PSet (1, 1)", "pset", None, Some(1)),
        ("f.PSet Step (1, 1), 3", "pset", None, Some(2)),
        ("f.Scale", "scale", None, None),
        ("f.Scale (1, 1)-(2, 2)", "scale", None, Some(2)),
        ("f.Controls.Circle (1, 1), 2", "circle", None, Some(2)),
        ("f.Pset (1, 1)", "pset", None, Some(1)),
    ];
    for (text, name, inner, outer) in cases {
        let (p, pre, selects) = select(&sub(&format!("Dim f As Form\n{text}")));
        assert_eq!(p.found, vec![], "{text}");
        assert_eq!((pre.len(), selects), (2, 1), "{text}");
        let (got, i, o) = graphics(&p, pre[1]);
        assert_eq!((got.to_ascii_lowercase(), i, o), (name.to_string(), *inner, *outer), "{text}");
    }
}

#[test]
fn other_methods_assignments_and_calls_of_an_object_are_reported() {
    for text in ["f.Cls", "f.Print 1", "f.Line = 1", "f.Foo 1, 2", "f.Move 1, 2", "x = f.Circle", "Debug.Print 1"] {
        let src = sub(text);
        let (p, pre, _) = select(&src);
        let at = offset_of(&src, text, 0);
        assert_eq!(p.found.len(), 1, "{text}: {:?}", p.found);
        assert_eq!((p.found[0].0, p.found[0].1), (PRE, at), "{text}");
        assert_eq!(pre.len(), 1, "{text}");
    }
}

#[test]
fn a_graphics_method_followed_by_a_statement_reports_that_statement() {
    let src = sub("f.PSet (1, 1): x = 1");
    let (p, pre, _) = select(&src);
    assert_eq!(p.found, vec![(PRE, offset_of(&src, "x = 1", 0), 1)]);
    assert_eq!(pre.len(), 2);
}

#[test]
fn exit_for_inside_a_for_loop_is_accepted() {
    let src = "Sub S()\nFor i = 1 To 2\nSelect Case x\nExit For\nCase 1\nEnd Select\nNext\nEnd Sub\n";
    let p = parse(src);
    assert_eq!(p.found, vec![]);
    let body = p.proc_body(0);
    let ExprNode::For { body: for_body, .. } = p.node(body[0]) else { panic!("For") };
    let inner = p.block(*for_body);
    let ExprNode::SelectCase { pre, .. } = p.node(inner[0]) else { panic!("Select") };
    assert_eq!(pre.len(), 1);
    assert!(matches!(p.node(pre[0]), ExprNode::ExitStmt { kind: ExitKind::For }));
}

#[test]
fn exit_for_outside_a_for_loop_gets_only_its_own_error() {
    let src = sub("Exit For");
    let (p, pre, _) = select(&src);
    assert_eq!(p.found, vec![(ERR_EXIT_FOR_OUTSIDE_FOR, offset_of(&src, "Exit", 0), 4)]);
    assert!(matches!(p.node(pre[0]), ExprNode::ExitStmt { kind: ExitKind::For }));
}

#[test]
fn the_other_exits_are_reported_as_statements_before_the_first_case_even_when_misplaced() {
    // Each of these would otherwise be an error of its own (wrong kind of procedure, or no loop).
    for (text, kind) in [("Exit Do", ExitKind::Do), ("Exit Function", ExitKind::Function), ("Exit Property", ExitKind::Property)] {
        let src = sub(text);
        let (p, pre, _) = select(&src);
        assert_eq!(p.found, vec![(PRE, offset_of(&src, "Exit", 0), 4)], "{text}");
        let ExprNode::ExitStmt { kind: got } = p.node(pre[0]) else { panic!("{text}") };
        assert_eq!(*got, kind, "{text}");
    }
    let src = sub("Exit Sub");
    let (p, _, _) = select(&src);
    assert_eq!(p.found, vec![(PRE, offset_of(&src, "Exit", 0), 4)]);
}

#[test]
fn exit_sub_in_a_function_is_reported_as_a_statement_before_the_first_case_only() {
    let src = "Function F() As Long\nSelect Case x\nExit Sub\nCase 1\nEnd Select\nEnd Function\n";
    let p = parse(src);
    assert_eq!(p.found, vec![(PRE, offset_of(src, "Exit", 0), 4)]);
}

#[test]
fn exit_do_inside_a_do_loop_is_reported() {
    let src = "Sub S()\nDo\nSelect Case x\nExit Do\nCase 1\nEnd Select\nLoop\nEnd Sub\n";
    let p = parse(src);
    assert_eq!(p.found, vec![(PRE, offset_of(src, "Exit", 0), 4)]);
}

#[test]
fn loop_with_a_condition_is_reported_as_a_statement_before_the_first_case() {
    for text in ["Loop While x > 0", "Loop Until x > 0"] {
        let src = sub(text);
        let (p, pre, cases) = select(&src);
        assert_eq!(p.found, vec![(PRE, offset_of(&src, "Loop", 0), 4)], "{text}");
        assert_eq!((pre.len(), cases), (0, 1), "{text}");
    }
    let src = "Sub S()\nDo\nSelect Case x\nLoop While x\nCase 1\nEnd Select\nLoop\nEnd Sub\n";
    let p = parse(src);
    assert_eq!(p.found, vec![(PRE, offset_of(src, "Loop", 0), 4)]);
}

#[test]
fn a_bare_loop_and_a_next_keep_their_own_errors() {
    let src = sub("Loop");
    let (p, _, _) = select(&src);
    assert_eq!(p.found, vec![(ERR_LOOP_WITHOUT_DO, offset_of(&src, "Loop", 0), 4)]);
    let src = sub("Next");
    let (p, _, _) = select(&src);
    assert_eq!(p.found, vec![(ERR_NEXT_WITHOUT_FOR, offset_of(&src, "Next", 0), 4)]);
}

#[test]
fn blank_lines_comments_and_colons_are_still_fine() {
    let (p, pre, cases) = select("Sub S()\nSelect Case x\n\n' c\n:\nCase 1\nEnd Select\nEnd Sub\n");
    assert_eq!(p.found, vec![]);
    assert_eq!((pre.len(), cases), (0, 1));
}
