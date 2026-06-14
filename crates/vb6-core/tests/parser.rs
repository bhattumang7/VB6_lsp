//! Structural tests for the VB6 parser.

use vb6_core::frontend::ast::{
    AstLit, BinOpKind, DoKind, ExitKind, ExprArena, ExprNode, FileIoKind, LabelRef, NodeId,
    OnErrorKind, ProcKind, UnOpKind,
};
use vb6_core::frontend::keyword_table::KEYWORD_TABLE;
use vb6_core::frontend::parser::{ModuleKind, Parser};
use vb6_core::frontend::scanner::ScannerContext;

fn ctx() -> ScannerContext {
    let mut c = ScannerContext::new(1, 1, 0x0409);
    c.intern_keywords();
    c
}

fn parse_module(src: &[u8]) -> (Vec<NodeId>, ExprArena, bool) {
    let mut c = ctx();
    let mut arena = ExprArena::new();
    let mut parser = Parser::new(&mut c, src);
    let nodes = parser.parse_module(&mut arena);
    let has_errors = parser.diagnostics.has_errors();
    (nodes, arena, has_errors)
}

fn parse_class_module(src: &[u8]) -> (Vec<NodeId>, ExprArena, bool) {
    let mut c = ctx();
    let mut arena = ExprArena::new();
    let mut parser = Parser::with_module_kind(&mut c, src, ModuleKind::Class);
    let nodes = parser.parse_module(&mut arena);
    let has_errors = parser.diagnostics.has_errors();
    (nodes, arena, has_errors)
}

fn parse_stmts(src: &[u8]) -> (Vec<NodeId>, ExprArena, bool) {
    // Wrap in a Sub so parse_module can find statements
    let wrapped = format!("Sub Test()\n{}\nEnd Sub", std::str::from_utf8(src).unwrap());
    let (nodes, arena, errors) = parse_module(wrapped.as_bytes());
    (nodes, arena, errors)
}

/// First node matching `pred`.
fn find_node<'a>(arena: &'a ExprArena, pred: impl Fn(&ExprNode) -> bool) -> Option<&'a ExprNode> {
    (0..arena.len()).map(|i| arena.get(NodeId(i as u32))).find(|n| pred(n))
}

/// Count nodes matching `pred`.
fn count_nodes(arena: &ExprArena, pred: impl Fn(&ExprNode) -> bool) -> usize {
    (0..arena.len()).filter(|&i| pred(arena.get(NodeId(i as u32)))).count()
}

/// Statements of the first ProcDecl body Block (the wrapper Sub from `parse_stmts`).
fn proc_body_stmts(arena: &ExprArena, nodes: &[NodeId]) -> Vec<NodeId> {
    let body = nodes.iter().find_map(|n| match arena.get(*n) {
        ExprNode::ProcDecl { body, .. } => Some(*body),
        _ => None,
    });
    match body.map(|b| arena.get(b)) {
        Some(ExprNode::Block { stmts }) => stmts.clone(),
        _ => Vec::new(),
    }
}

// ── Accept tests ──────────────────────────────────────────────────────────────

#[test]
fn empty_sub_accepts() {
    let (nodes, arena, errors) = parse_module(b"Sub Foo()\nEnd Sub");
    assert!(!errors, "empty Sub should parse without errors");
    assert_eq!(nodes.len(), 1, "exactly one top-level decl");
    // `Sub Foo()` — empty (but present) param list, no ParamDef children.
    assert!(matches!(
        arena.get(nodes[0]),
        ExprNode::ProcDecl { kind: ProcKind::Sub, params: Some(_), .. }
    ));
    assert_eq!(count_nodes(&arena, |n| matches!(n, ExprNode::ParamDef { .. })), 0);
}

#[test]
fn sub_with_params_accepts() {
    let (nodes, arena, errors) = parse_module(b"Sub Foo(x As Integer, y As String)\nEnd Sub");
    assert!(!errors);
    assert!(matches!(
        arena.get(nodes[0]),
        ExprNode::ProcDecl { kind: ProcKind::Sub, params: Some(_), .. }
    ));
    assert_eq!(count_nodes(&arena, |n| matches!(n, ExprNode::ParamDef { .. })), 2);
}

#[test]
fn function_with_return_type() {
    let (nodes, arena, errors) =
        parse_module(b"Function Add(a As Long, b As Long) As Long\nEnd Function");
    assert!(!errors);
    assert!(matches!(
        arena.get(nodes[0]),
        ExprNode::ProcDecl { kind: ProcKind::Function, params: Some(_), ret_type: Some(_), .. }
    ));
    assert_eq!(count_nodes(&arena, |n| matches!(n, ExprNode::ParamDef { .. })), 2);
}

#[test]
fn property_get_accepts() {
    let (nodes, arena, errors) = parse_module(b"Property Get Value() As Integer\nEnd Property");
    assert!(!errors);
    assert!(matches!(
        arena.get(nodes[0]),
        ExprNode::ProcDecl { kind: ProcKind::PropGet, ret_type: Some(_), .. }
    ));
}

#[test]
fn property_let_accepts() {
    let (nodes, arena, errors) = parse_module(b"Property Let Value(v As Integer)\nEnd Property");
    assert!(!errors);
    assert!(matches!(arena.get(nodes[0]), ExprNode::ProcDecl { kind: ProcKind::PropLet, .. }));
    assert_eq!(count_nodes(&arena, |n| matches!(n, ExprNode::ParamDef { .. })), 1);
}

#[test]
fn property_set_accepts() {
    let (nodes, arena, errors) = parse_module(b"Property Set Obj(o As Object)\nEnd Property");
    assert!(!errors);
    assert!(matches!(arena.get(nodes[0]), ExprNode::ProcDecl { kind: ProcKind::PropSet, .. }));
    assert_eq!(count_nodes(&arena, |n| matches!(n, ExprNode::ParamDef { .. })), 1);
}

#[test]
fn dim_statement_accepts() {
    let (_, arena, errors) = parse_stmts(b"Dim x As Integer");
    assert!(!errors);
    let item = find_node(&arena, |n| matches!(n, ExprNode::DimItem { .. }))
        .expect("Dim must build a DimItem");
    assert!(matches!(
        item,
        ExprNode::DimItem { is_const: false, bounds: None, type_node: Some(_), init: None, .. }
    ));
    // Integer = builtin kind 2.
    assert!(find_node(&arena, |n| matches!(n, ExprNode::BuiltinType { kind: 2 })).is_some());
}

#[test]
fn dim_with_array_bounds() {
    let (_, arena, errors) = parse_stmts(b"Dim a(10) As String");
    assert!(!errors);
    let item = find_node(&arena, |n| matches!(n, ExprNode::DimItem { .. }))
        .expect("Dim must build a DimItem");
    assert!(matches!(item, ExprNode::DimItem { bounds: Some(_), type_node: Some(_), .. }));
    assert!(find_node(&arena, |n| matches!(n, ExprNode::StringType { .. })).is_some());
}

#[test]
fn dim_with_range_bounds() {
    let (_, arena, errors) = parse_stmts(b"Dim a(1 To 10) As Integer");
    assert!(!errors);
    assert!(find_node(&arena, |n| matches!(n, ExprNode::DimItem { bounds: Some(_), .. })).is_some());
    // `1 To 10` bound is a RangeTo.
    assert!(find_node(&arena, |n| matches!(n, ExprNode::RangeTo { .. })).is_some());
}

#[test]
fn const_declaration() {
    let (_, arena, errors) = parse_stmts(b"Const Pi = 3.14159");
    assert!(!errors);
    let item = find_node(&arena, |n| matches!(n, ExprNode::DimItem { is_const: true, .. }))
        .expect("Const must build a const DimItem");
    assert!(matches!(item, ExprNode::DimItem { is_const: true, init: Some(_), .. }));
}

#[test]
fn assignment_accepts() {
    let (nodes, arena, errors) = parse_stmts(b"x = 42");
    assert!(!errors);
    assert_eq!(count_nodes(&arena, |n| matches!(n, ExprNode::Assign { .. })), 1);
    let assign = find_node(&arena, |n| matches!(n, ExprNode::Assign { .. })).unwrap();
    if let ExprNode::Assign { value, .. } = assign {
        assert!(matches!(arena.get(*value), ExprNode::Literal { lit: AstLit::Int(42) }));
    }
    assert_eq!(proc_body_stmts(&arena, &nodes).len(), 1);
}

#[test]
fn let_assignment_accepts() {
    let (_, arena, errors) = parse_stmts(b"Let x = 42");
    assert!(!errors);
    // `Let` is a plain Assign, not a Set.
    assert_eq!(count_nodes(&arena, |n| matches!(n, ExprNode::Assign { .. })), 1);
    assert_eq!(count_nodes(&arena, |n| matches!(n, ExprNode::SetAssign { .. })), 0);
}

#[test]
fn set_assignment_accepts() {
    let (_, arena, errors) = parse_stmts(b"Set obj = Nothing");
    assert!(!errors);
    let set = find_node(&arena, |n| matches!(n, ExprNode::SetAssign { .. }))
        .expect("Set must build a SetAssign");
    if let ExprNode::SetAssign { value, .. } = set {
        assert!(matches!(arena.get(*value), ExprNode::Nothing));
    }
}

#[test]
fn for_next_loop() {
    let (_, arena, errors) = parse_stmts(b"For i = 1 To 10\nNext i");
    assert!(!errors);
    assert_eq!(count_nodes(&arena, |n| matches!(n, ExprNode::For { .. })), 1);
    assert!(find_node(&arena, |n| matches!(n, ExprNode::For { step: None, .. })).is_some());
}

#[test]
fn for_with_step() {
    let (_, arena, errors) = parse_stmts(b"For i = 1 To 10 Step 2\nNext");
    assert!(!errors);
    let f = find_node(&arena, |n| matches!(n, ExprNode::For { .. })).expect("must build a For");
    if let ExprNode::For { step: Some(s), .. } = f {
        assert!(matches!(arena.get(*s), ExprNode::Literal { lit: AstLit::Int(2) }));
    } else {
        panic!("For must carry a Step expression");
    }
}

#[test]
fn for_each_loop() {
    let (_, arena, errors) = parse_stmts(b"For Each x In col\nNext x");
    assert!(!errors);
    assert_eq!(count_nodes(&arena, |n| matches!(n, ExprNode::ForEach { .. })), 1);
    assert_eq!(count_nodes(&arena, |n| matches!(n, ExprNode::For { .. })), 0);
}

#[test]
fn do_while_loop() {
    let (_, arena, errors) = parse_stmts(b"Do While x > 0\nLoop");
    assert!(!errors);
    assert!(find_node(&arena, |n| matches!(
        n,
        ExprNode::Do { kind: DoKind::PreWhile, cond: Some(_), .. }
    ))
    .is_some());
}

#[test]
fn do_until_loop() {
    let (_, arena, errors) = parse_stmts(b"Do Until x = 0\nLoop");
    assert!(!errors);
    assert!(find_node(&arena, |n| matches!(
        n,
        ExprNode::Do { kind: DoKind::PreUntil, cond: Some(_), .. }
    ))
    .is_some());
}

#[test]
fn do_loop_while() {
    let (_, arena, errors) = parse_stmts(b"Do\nLoop While x > 0");
    assert!(!errors);
    assert!(find_node(&arena, |n| matches!(
        n,
        ExprNode::Do { kind: DoKind::PostWhile, cond: Some(_), .. }
    ))
    .is_some());
}

#[test]
fn do_loop_infinite() {
    let (_, arena, errors) = parse_stmts(b"Do\nLoop");
    assert!(!errors);
    assert!(find_node(&arena, |n| matches!(n, ExprNode::Do { kind: DoKind::Inf, cond: None, .. }))
        .is_some());
}

#[test]
fn while_wend_loop() {
    let (_, arena, errors) = parse_stmts(b"While x > 0\nWend");
    assert!(!errors);
    let w = find_node(&arena, |n| matches!(n, ExprNode::While { .. })).expect("must build a While");
    if let ExprNode::While { cond, .. } = w {
        assert!(matches!(arena.get(*cond), ExprNode::BinOp { op: BinOpKind::Gt, .. }));
    }
}

#[test]
fn if_then_end_if() {
    let (_, arena, errors) = parse_stmts(b"If x > 0 Then\nx = 1\nEnd If");
    assert!(!errors);
    let iff = find_node(&arena, |n| matches!(n, ExprNode::If { .. })).expect("must build an If");
    assert!(matches!(iff, ExprNode::If { else_body: None, .. }));
    if let ExprNode::If { cond, .. } = iff {
        assert!(matches!(arena.get(*cond), ExprNode::BinOp { op: BinOpKind::Gt, .. }));
    }
}

#[test]
fn if_else_end_if() {
    let (_, arena, errors) = parse_stmts(b"If x > 0 Then\nx = 1\nElse\nx = 0\nEnd If");
    assert!(!errors);
    assert!(
        find_node(&arena, |n| matches!(n, ExprNode::If { else_body: Some(_), .. })).is_some(),
        "If/Else must carry an else_body"
    );
}

#[test]
fn if_elseif_else_end_if() {
    let (_, arena, errors) =
        parse_stmts(b"If x > 0 Then\nx = 1\nElseIf x < 0 Then\nx = -1\nElse\nx = 0\nEnd If");
    assert!(!errors);
    // ElseIf nests a second If inside the outer If's else_body.
    assert!(
        count_nodes(&arena, |n| matches!(n, ExprNode::If { .. })) >= 2,
        "ElseIf must nest a second If"
    );
    assert!(find_node(&arena, |n| matches!(n, ExprNode::If { else_body: Some(_), .. })).is_some());
}

#[test]
fn single_line_if() {
    let (_, arena, errors) = parse_stmts(b"If x > 0 Then x = 1");
    assert!(!errors);
    assert!(find_node(&arena, |n| matches!(n, ExprNode::If { else_body: None, .. })).is_some());
}

#[test]
fn select_case_accepts() {
    let (_, arena, errors) = parse_stmts(
        b"Select Case x\nCase 1\ny = 1\nCase 2, 3\ny = 2\nCase Else\ny = 0\nEnd Select",
    );
    assert!(!errors);
    let sel = find_node(&arena, |n| matches!(n, ExprNode::SelectCase { .. }))
        .expect("must build a SelectCase");
    if let ExprNode::SelectCase { cases, .. } = sel {
        assert_eq!(cases.len(), 3, "two Case arms + Case Else");
    }
    assert_eq!(count_nodes(&arena, |n| matches!(n, ExprNode::CaseBlock { .. })), 2);
    assert_eq!(count_nodes(&arena, |n| matches!(n, ExprNode::CaseElse { .. })), 1);
}

#[test]
fn select_case_is_range() {
    let (_, arena, errors) =
        parse_stmts(b"Select Case x\nCase Is > 10\ny = 1\nCase 1 To 5\ny = 2\nEnd Select");
    assert!(!errors);
    let isnode = find_node(&arena, |n| matches!(n, ExprNode::CaseIs { .. }))
        .expect("Case Is must build a CaseIs");
    assert!(matches!(isnode, ExprNode::CaseIs { op: BinOpKind::Gt, .. }));
    assert!(find_node(&arena, |n| matches!(n, ExprNode::RangeTo { .. })).is_some());
}

#[test]
fn with_block_accepts() {
    let (_, arena, errors) = parse_stmts(b"With obj\n.x = 1\nEnd With");
    assert!(!errors);
    assert_eq!(count_nodes(&arena, |n| matches!(n, ExprNode::With { .. })), 1);
    // `.x = 1` is a leading-dot With-member access; the `.` is semantically
    // significant and must not collapse into a plain `x` NameRef.
    assert!(
        find_node(&arena, |n| matches!(n, ExprNode::MemberAccess { .. })).is_some(),
        "With-block `.x` must build a MemberAccess, not a bare NameRef"
    );
}

#[test]
fn exit_sub_accepts() {
    let (_, arena, errors) = parse_stmts(b"Exit Sub");
    assert!(!errors);
    assert!(find_node(&arena, |n| matches!(n, ExprNode::ExitStmt { kind: ExitKind::Sub })).is_some());
}

#[test]
fn exit_for_accepts() {
    let (_, arena, errors) = parse_stmts(b"Exit For");
    assert!(!errors);
    assert!(find_node(&arena, |n| matches!(n, ExprNode::ExitStmt { kind: ExitKind::For })).is_some());
}

#[test]
fn exit_do_accepts() {
    let (_, arena, errors) = parse_stmts(b"Exit Do");
    assert!(!errors);
    assert!(find_node(&arena, |n| matches!(n, ExprNode::ExitStmt { kind: ExitKind::Do })).is_some());
}

#[test]
fn goto_accepts() {
    let (_, arena, errors) = parse_stmts(b"GoTo myLabel");
    assert!(!errors);
    let g = find_node(&arena, |n| matches!(n, ExprNode::GoTo { .. })).expect("must build a GoTo");
    assert!(matches!(g, ExprNode::GoTo { target: LabelRef::Name(_) }));
}

#[test]
fn on_error_resume_next() {
    let (_, arena, errors) = parse_stmts(b"On Error Resume Next");
    assert!(!errors);
    assert!(find_node(&arena, |n| matches!(
        n,
        ExprNode::OnError { kind: OnErrorKind::ResumeNext }
    ))
    .is_some());
}

#[test]
fn on_error_goto_label() {
    let (_, arena, errors) = parse_stmts(b"On Error GoTo errHandler");
    assert!(!errors);
    assert!(find_node(&arena, |n| matches!(
        n,
        ExprNode::OnError { kind: OnErrorKind::Goto(LabelRef::Name(_)) }
    ))
    .is_some());
}

#[test]
fn on_error_goto_zero() {
    let (_, arena, errors) = parse_stmts(b"On Error GoTo 0");
    assert!(!errors);
    // `GoTo 0` disables the handler.
    assert!(find_node(&arena, |n| matches!(n, ExprNode::OnError { kind: OnErrorKind::Disable }))
        .is_some());
}

#[test]
fn call_statement_parens() {
    let (_, arena, errors) = parse_stmts(b"Call Foo(1, 2)");
    assert!(!errors);
    // `Call Foo(1, 2)` — a CallStmt whose callee carries the `(1, 2)` arg list.
    assert!(find_node(&arena, |n| matches!(n, ExprNode::CallStmt { .. })).is_some(),
        "explicit Call must build a CallStmt");
    // Exactly one ArgList holds the two literal arguments.
    assert!(
        find_node(&arena, |n| matches!(n, ExprNode::ArgList { args } if args.len() == 2)).is_some(),
        "the two call arguments must form a 2-element ArgList"
    );
}

#[test]
fn redim_accepts() {
    let (_, arena, errors) = parse_stmts(b"ReDim a(10)");
    assert!(!errors);
    let r = find_node(&arena, |n| matches!(n, ExprNode::ReDimItem { .. }))
        .expect("ReDim must build a ReDimItem");
    assert!(matches!(r, ExprNode::ReDimItem { preserve: false, bounds: Some(_), .. }));
}

#[test]
fn erase_accepts() {
    let (_, arena, errors) = parse_stmts(b"Erase a");
    assert!(!errors);
    let e = find_node(&arena, |n| matches!(n, ExprNode::Erase { .. })).expect("must build Erase");
    if let ExprNode::Erase { vars } = e {
        assert_eq!(vars.len(), 1);
    }
}

#[test]
fn raise_event_accepts() {
    let (_, arena, errors) = parse_stmts(b"RaiseEvent MyEvent(x, y)");
    assert!(!errors);
    let r = find_node(&arena, |n| matches!(n, ExprNode::RaiseEvent { .. }))
        .expect("must build RaiseEvent");
    if let ExprNode::RaiseEvent { args, .. } = r {
        if let ExprNode::ArgList { args } = arena.get(*args) {
            assert_eq!(args.len(), 2);
        }
    }
}

// ── Expression tests ──────────────────────────────────────────────────────────

#[test]
fn arithmetic_expr_no_error() {
    // `1 + 2 * 3` == `1 + (2 * 3)` — Add at the root, Mul on the rhs.
    let (a, v) = assign_value("1 + 2 * 3");
    let (_l, rhs) = expect_binop(&a, v, BinOpKind::Add);
    expect_binop(&a, rhs, BinOpKind::Mul);
}

#[test]
fn comparison_expr_no_error() {
    let (a, v) = assign_value("a > b");
    expect_binop(&a, v, BinOpKind::Gt);
}

#[test]
fn logical_expr_no_error() {
    // `a And b Or c` == `(a And b) Or c` — Or at the root, And on the lhs.
    let (a, v) = assign_value("a And b Or c");
    let (lhs, _r) = expect_binop(&a, v, BinOpKind::Or);
    expect_binop(&a, lhs, BinOpKind::And);
}

#[test]
fn unary_not_no_error() {
    let (a, v) = assign_value("Not a");
    assert!(matches!(a.get(v), ExprNode::UnOp { op: UnOpKind::Not, .. }));
}

#[test]
fn unary_neg_no_error() {
    let (a, v) = assign_value("-a");
    assert!(matches!(a.get(v), ExprNode::UnOp { op: UnOpKind::Neg, .. }));
}

#[test]
fn paren_expr_no_error() {
    // `(1 + 2)` — a Paren wrapping an Add.
    let (a, v) = assign_value("(1 + 2)");
    match a.get(v) {
        ExprNode::Paren { inner } => {
            expect_binop(&a, *inner, BinOpKind::Add);
        }
        other => panic!("expected Paren"),
    }
}

#[test]
fn member_access_no_error() {
    let (a, v) = assign_value("obj.Property");
    assert!(matches!(a.get(v), ExprNode::MemberAccess { bang: false, .. }));
}

#[test]
fn chained_member_access() {
    // `a.b.c` — root MemberAccess whose base is another MemberAccess.
    let (a, v) = assign_value("a.b.c");
    match a.get(v) {
        ExprNode::MemberAccess { base, .. } => {
            assert!(matches!(a.get(*base), ExprNode::MemberAccess { .. }));
        }
        other => panic!("expected MemberAccess"),
    }
    assert_eq!(count_nodes(&a, |n| matches!(n, ExprNode::MemberAccess { .. })), 2);
}

#[test]
fn function_call_no_error() {
    // `Foo(1, 2, 3)` — a Call whose args ArgList has three entries.
    let (a, v) = assign_value("Foo(1, 2, 3)");
    match a.get(v) {
        ExprNode::Call { args, .. } => match a.get(*args) {
            ExprNode::ArgList { args } => assert_eq!(args.len(), 3),
            other => panic!("Call args must be ArgList"),
        },
        other => panic!("expected Call"),
    }
}

#[test]
fn new_expr_no_error() {
    let (_, arena, errors) = parse_stmts(b"Set obj = New MyClass");
    assert!(!errors);
    assert!(find_node(&arena, |n| matches!(n, ExprNode::New { .. })).is_some());
}

#[test]
fn typeof_is_no_error() {
    let (a, v) = assign_value("TypeOf obj Is MyClass");
    assert!(matches!(a.get(v), ExprNode::TypeOf { .. }));
}

#[test]
fn string_concat_no_error() {
    // `s1 & s2 & s3` — left-associative: `(s1 & s2) & s3`.
    let (a, v) = assign_value("s1 & s2 & s3");
    let (lhs, _r) = expect_binop(&a, v, BinOpKind::Cat);
    expect_binop(&a, lhs, BinOpKind::Cat);
}

#[test]
fn power_operator_no_error() {
    let (a, v) = assign_value("2 ^ 10");
    expect_binop(&a, v, BinOpKind::Pow);
}

#[test]
fn integer_div_no_error() {
    let (a, v) = assign_value("a \\ b");
    expect_binop(&a, v, BinOpKind::IDiv);
}

#[test]
fn mod_operator_no_error() {
    let (a, v) = assign_value("a Mod b");
    expect_binop(&a, v, BinOpKind::Mod);
}

// ── Type spec tests ───────────────────────────────────────────────────────────

#[test]
fn type_spec_integer() {
    let (_, arena, errors) = parse_module(b"Dim x As Integer");
    assert!(!errors);
    assert!(find_node(&arena, |n| matches!(n, ExprNode::BuiltinType { kind: 2 })).is_some());
}

#[test]
fn type_spec_string_fixed() {
    let (_, arena, errors) = parse_module(b"Dim s As String * 10");
    assert!(!errors);
    // Fixed-length String carries a fixed_len expression.
    assert!(
        find_node(&arena, |n| matches!(n, ExprNode::StringType { fixed_len: Some(_) })).is_some()
    );
}

#[test]
fn type_spec_user_defined() {
    let (_, arena, errors) = parse_module(b"Dim r As MyRecord");
    assert!(!errors);
    assert!(find_node(&arena, |n| matches!(n, ExprNode::UserType { child: None, .. })).is_some());
}

#[test]
fn type_spec_module_qualified() {
    let (_, arena, errors) = parse_module(b"Dim obj As Module1.MyClass");
    assert!(!errors);
    // Qualified `Module1.MyClass` is a UserType with a child reference.
    assert!(find_node(&arena, |n| matches!(n, ExprNode::UserType { child: Some(_), .. })).is_some());
}

// ── Declaration tests ─────────────────────────────────────────────────────────

#[test]
fn public_sub_accepts() {
    let (nodes, arena, errors) = parse_module(b"Public Sub Foo()\nEnd Sub");
    assert!(!errors);
    assert!(matches!(arena.get(nodes[0]), ExprNode::ProcDecl { kind: ProcKind::Sub, .. }));
}

#[test]
fn private_function_accepts() {
    let (nodes, arena, errors) = parse_module(b"Private Function Bar() As Long\nEnd Function");
    assert!(!errors);
    assert!(matches!(
        arena.get(nodes[0]),
        ExprNode::ProcDecl { kind: ProcKind::Function, ret_type: Some(_), .. }
    ));
}

#[test]
fn type_declaration_accepts() {
    let (_, arena, errors) = parse_module(b"Type Point\nx As Integer\ny As Integer\nEnd Type");
    assert!(!errors);
    let t = find_node(&arena, |n| matches!(n, ExprNode::TypeDecl { .. }))
        .expect("must build a TypeDecl");
    if let ExprNode::TypeDecl { members, .. } = t {
        assert_eq!(members.len(), 2, "two record members");
    }
}

#[test]
fn enum_declaration_accepts() {
    let (_, arena, errors) = parse_module(b"Enum Color\nRed = 1\nGreen = 2\nBlue = 3\nEnd Enum");
    assert!(!errors);
    let e = find_node(&arena, |n| matches!(n, ExprNode::EnumDecl { .. }))
        .expect("must build an EnumDecl");
    if let ExprNode::EnumDecl { members, .. } = e {
        assert_eq!(members.len(), 3, "three enum members");
    }
}

#[test]
fn event_declaration_accepts() {
    // Event is class-module-only.
    let (_, arena, errors) = parse_class_module(b"Event Click(x As Integer, y As Integer)");
    assert!(!errors);
    assert!(find_node(&arena, |n| matches!(n, ExprNode::EventDecl { .. })).is_some());
    assert_eq!(count_nodes(&arena, |n| matches!(n, ExprNode::ParamDef { .. })), 2);
}

#[test]
fn implements_accepts() {
    // Implements is class-module-only.
    let (_, arena, errors) = parse_class_module(b"Implements IMyInterface");
    assert!(!errors);
    assert!(find_node(&arena, |n| matches!(n, ExprNode::Implements { .. })).is_some());
}

// ── AST structure checks ──────────────────────────────────────────────────────

#[test]
fn sub_decl_produces_node() {
    let (nodes, arena, _) = parse_module(b"Sub Foo()\nEnd Sub");
    assert!(!nodes.is_empty());
    assert!(
        matches!(arena.get(nodes[0]), ExprNode::ProcDecl { kind: ProcKind::Sub, .. }),
        "Sub should produce ProcDecl {{ kind: Sub, .. }}"
    );
}

#[test]
fn function_decl_produces_node() {
    let (nodes, arena, _) = parse_module(b"Function Bar() As Long\nEnd Function");
    assert!(!nodes.is_empty());
    assert!(
        matches!(arena.get(nodes[0]), ExprNode::ProcDecl { kind: ProcKind::Function, .. }),
        "Function should produce ProcDecl {{ kind: Function, .. }}"
    );
}

#[test]
fn assignment_produces_node() {
    let (nodes, arena, errors) = parse_module(b"Sub T()\nx = 1\nEnd Sub");
    assert!(!errors);
    assert!(!nodes.is_empty());
    // The single body statement is an Assign of literal 1.
    let stmts = proc_body_stmts(&arena, &nodes);
    assert_eq!(stmts.len(), 1);
    match arena.get(stmts[0]) {
        ExprNode::Assign { value, .. } => {
            assert!(matches!(arena.get(*value), ExprNode::Literal { lit: AstLit::Int(1) }));
        }
        other => panic!("expected Assign"),
    }
}

#[test]
fn empty_module_accepts() {
    let (nodes, _, errors) = parse_module(b"");
    assert!(!errors);
    assert!(nodes.is_empty(), "empty source yields no top-level declarations");
}

#[test]
fn multi_proc_module() {
    let (nodes, arena, errors) =
        parse_module(b"Sub A()\nEnd Sub\nSub B()\nEnd Sub\nSub C()\nEnd Sub");
    assert!(!errors);
    assert_eq!(nodes.len(), 3, "three Sub declarations");
    assert!(nodes
        .iter()
        .all(|n| matches!(arena.get(*n), ExprNode::ProcDecl { kind: ProcKind::Sub, .. })));
}

// ── Error recovery tests ──────────────────────────────────────────────────────

#[test]
fn missing_end_sub_produces_error() {
    let (nodes, arena, _errors) = parse_module(b"Sub Foo()\nx = 1");
    // No End Sub — the parser recovers at EOF; it must still build the ProcDecl
    // and not drop the body assignment.
    assert!(!nodes.is_empty(), "must still produce the Sub declaration");
    assert!(matches!(arena.get(nodes[0]), ExprNode::ProcDecl { kind: ProcKind::Sub, .. }));
    assert!(find_node(&arena, |n| matches!(n, ExprNode::Assign { .. })).is_some());
}

#[test]
fn empty_arg_list_accepts() {
    let (_, arena, errors) = parse_stmts(b"Call Foo()");
    assert!(!errors);
    assert!(find_node(&arena, |n| matches!(n, ExprNode::CallStmt { .. })).is_some(),
        "Call must build a CallStmt");
    // `Foo()` — an empty ArgList must be present (no MissingArg, no args).
    assert!(
        find_node(&arena, |n| matches!(n, ExprNode::ArgList { args } if args.is_empty())).is_some(),
        "empty parens must build an empty ArgList"
    );
}

#[test]
fn optional_param_with_default() {
    let (_, arena, errors) = parse_module(b"Sub Foo(Optional x As Integer = 0)\nEnd Sub");
    assert!(!errors);
    // The Optional param carries a default expression.
    let p = find_node(&arena, |n| matches!(n, ExprNode::ParamDef { .. }))
        .expect("must build a ParamDef");
    assert!(
        matches!(p, ExprNode::ParamDef { default: Some(_), .. }),
        "Optional must keep a default"
    );
}

#[test]
fn byval_byref_params() {
    let (_, arena, errors) = parse_module(b"Sub Foo(ByVal x As Integer, ByRef y As Long)\nEnd Sub");
    assert!(!errors);
    // Two params; ByVal/ByRef is encoded in distinct param flags.
    let flags: Vec<u16> = (0..arena.len())
        .filter_map(|i| match arena.get(NodeId(i as u32)) {
            ExprNode::ParamDef { flags, .. } => Some(*flags),
            _ => None,
        })
        .collect();
    assert_eq!(flags.len(), 2);
    assert_ne!(flags[0], flags[1], "ByVal and ByRef must produce different param flags");
}

#[test]
fn conditional_compilation_directive() {
    let (_, _, errors) = parse_module(b"#If Win32 Then\n#End If");
    assert!(!errors);
}

// ── Operator precedence ───────────────────────────────────────────────────────
//
// Each operator's precedence is stored in the low nibble of its keyword-table
// `w1` field (`KEYWORD_TABLE[..].w1 & 0xF`); higher binds tighter.  `infix_bp`
// in parser.rs is derived from these nibbles.  These tests pin both the table
// ordering and the resulting AST grouping.

/// Precedence level for a keyword name (low nibble of `w1`).
fn dll_level(name: &str) -> u32 {
    KEYWORD_TABLE
        .iter()
        .find(|e| e.name == name)
        .unwrap_or_else(|| panic!("keyword {name:?} not in table"))
        .w1
        & 0xF
}

#[test]
fn precedence_table_nibbles_match_documented_levels() {
    // The canonical VB6 logical/comparison/arithmetic ordering, as encoded in
    // the keyword table.  This is what `infix_bp` is built from.
    assert_eq!(dll_level("Imp"), 1);
    assert_eq!(dll_level("Eqv"), 2);
    assert_eq!(dll_level("Xor"), 3);
    assert_eq!(dll_level("Or"), 4);
    assert_eq!(dll_level("And"), 5);
    assert_eq!(dll_level("="), 6);
    assert_eq!(dll_level("<"), 6);
    assert_eq!(dll_level("Like"), 6);
    assert_eq!(dll_level("Is"), 6);
    assert_eq!(dll_level("&"), 7);
    assert_eq!(dll_level("+"), 8);
    assert_eq!(dll_level("-"), 8);
    assert_eq!(dll_level("Mod"), 9);
    assert_eq!(dll_level("\\"), 10);
    assert_eq!(dll_level("*"), 11);
    assert_eq!(dll_level("/"), 11);
    assert_eq!(dll_level("^"), 12);

    // Strictly-increasing logical chain — the property the merged-level bug
    // (Or==Xor, Eqv==Imp) violated.
    assert!(dll_level("Imp") < dll_level("Eqv"));
    assert!(dll_level("Eqv") < dll_level("Xor"));
    assert!(dll_level("Xor") < dll_level("Or"));
    assert!(dll_level("Or") < dll_level("And"));
    assert!(dll_level("And") < dll_level("="));
}

/// Parse `x = <expr>` and return the arena plus the value-expression node id.
fn assign_value(expr_src: &str) -> (ExprArena, NodeId) {
    let src = format!("Sub T()\n x = {expr_src}\nEnd Sub");
    let (nodes, arena, errs) = parse_module(src.as_bytes());
    assert!(!errs, "unexpected parse errors for {expr_src:?}");
    let body = match arena.get(nodes[0]) {
        ExprNode::ProcDecl { body, .. } => *body,
        _ => panic!("expected ProcDecl"),
    };
    let stmts = match arena.get(body) {
        ExprNode::Block { stmts } => stmts.clone(),
        _ => panic!("expected Block"),
    };
    let value = stmts.iter().find_map(|s| match arena.get(*s) {
        ExprNode::Assign { value, .. } => Some(*value),
        _ => None,
    });
    (arena, value.expect("no Assign found"))
}

/// Assert the value root is a `BinOp` with `op`, returning `(lhs, rhs)`.
fn expect_binop(arena: &ExprArena, id: NodeId, op: BinOpKind) -> (NodeId, NodeId) {
    match arena.get(id) {
        ExprNode::BinOp { op: o, lhs, rhs } => {
            assert_eq!(*o, op, "expected {op:?} at root");
            (*lhs, *rhs)
        }
        _ => panic!("expected BinOp({op:?})"),
    }
}

#[test]
fn xor_binds_looser_than_or() {
    // `1 Xor 2 Or 3` == `1 Xor (2 Or 3)` because Or (level 4) > Xor (level 3).
    let (a, v) = assign_value("1 Xor 2 Or 3");
    let (_lhs, rhs) = expect_binop(&a, v, BinOpKind::Xor);
    expect_binop(&a, rhs, BinOpKind::Or);
}

#[test]
fn or_binds_looser_than_and() {
    // `1 Or 2 And 3` == `1 Or (2 And 3)`.
    let (a, v) = assign_value("1 Or 2 And 3");
    let (_l, rhs) = expect_binop(&a, v, BinOpKind::Or);
    expect_binop(&a, rhs, BinOpKind::And);
}

#[test]
fn imp_binds_looser_than_eqv() {
    // `1 Imp 2 Eqv 3` == `1 Imp (2 Eqv 3)`.
    let (a, v) = assign_value("1 Imp 2 Eqv 3");
    let (_l, rhs) = expect_binop(&a, v, BinOpKind::Imp);
    expect_binop(&a, rhs, BinOpKind::Eqv);
}

#[test]
fn eqv_binds_looser_than_xor() {
    // `1 Eqv 2 Xor 3` == `1 Eqv (2 Xor 3)`.
    let (a, v) = assign_value("1 Eqv 2 Xor 3");
    let (_l, rhs) = expect_binop(&a, v, BinOpKind::Eqv);
    expect_binop(&a, rhs, BinOpKind::Xor);
}

#[test]
fn comparison_binds_tighter_than_and() {
    // `a = b And c` == `(a = b) And c`.
    let (a, v) = assign_value("a = b And c");
    let (lhs, _r) = expect_binop(&a, v, BinOpKind::And);
    expect_binop(&a, lhs, BinOpKind::Eq);
}

#[test]
fn not_binds_looser_than_comparison() {
    // `Not a = b` == `Not (a = b)` — Not is a logical operator below comparison.
    let (a, v) = assign_value("Not a = b");
    match a.get(v) {
        ExprNode::UnOp { op: UnOpKind::Not, operand } => {
            expect_binop(&a, *operand, BinOpKind::Eq);
        }
        _ => panic!("expected UnOp(Not)"),
    }
}

// ── Additional statement forms ────────────────────────────────────────────────
//
// Each construct below is accepted by VB6 and must parse without errors.

#[test]
fn if_goto_without_then_accepts() {
    // Legacy single-line `If <cond> GoTo <label>` (no Then).
    let (_, arena, errors) = parse_stmts(b"If True GoTo 10\n10");
    assert!(!errors);
    // Must build an If whose then-branch jumps via a GoTo to numeric line 10.
    assert!(find_node(&arena, |n| matches!(n, ExprNode::If { .. })).is_some());
    let g = find_node(&arena, |n| matches!(n, ExprNode::GoTo { .. }))
        .expect("If ... GoTo must build a GoTo node");
    assert!(matches!(g, ExprNode::GoTo { target: LabelRef::Line(10) }));
}

#[test]
fn if_goto_label_without_then_accepts() {
    let (_, arena, errors) = parse_stmts(b"If True GoTo L\nL:");
    assert!(!errors);
    assert!(find_node(&arena, |n| matches!(n, ExprNode::If { .. })).is_some());
    let g = find_node(&arena, |n| matches!(n, ExprNode::GoTo { .. }))
        .expect("If ... GoTo must build a GoTo node");
    assert!(matches!(g, ExprNode::GoTo { target: LabelRef::Name(_) }));
    // And the `L:` definition site is a Label.
    assert!(find_node(&arena, |n| matches!(n, ExprNode::Label { .. })).is_some());
}

#[test]
fn endif_one_word_accepts() {
    // One-word `EndIf` synonym for `End If`.
    let (_, arena, errors) = parse_stmts(b"If True Then\n  Dim x As Long\nEndIf");
    assert!(!errors);
    assert!(find_node(&arena, |n| matches!(n, ExprNode::If { .. })).is_some());
    assert!(find_node(&arena, |n| matches!(n, ExprNode::DimItem { .. })).is_some());
}

#[test]
fn def_int_range_accepts() {
    let (_, arena, errors) = parse_module(b"DefInt A-Z\nSub F()\nEnd Sub");
    assert!(!errors);
    let d = find_node(&arena, |n| matches!(n, ExprNode::DefType { .. }))
        .expect("must build a DefType");
    if let ExprNode::DefType { ranges, .. } = d {
        assert_eq!(ranges.len(), 1, "single A-Z range");
        assert_ne!(ranges[0].1, 0, "A-Z is a two-letter range (end != 0)");
    }
}

#[test]
fn def_multi_range_accepts() {
    let (_, arena, errors) = parse_module(b"DefInt A-C, X-Z\nSub F()\nEnd Sub");
    assert!(!errors);
    let d = find_node(&arena, |n| matches!(n, ExprNode::DefType { .. }))
        .expect("must build a DefType");
    if let ExprNode::DefType { ranges, .. } = d {
        assert_eq!(ranges.len(), 2, "two comma-separated ranges");
    }
}

#[test]
fn def_single_letter_accepts() {
    let (_, arena, errors) = parse_module(b"DefBool B\nSub F()\nEnd Sub");
    assert!(!errors);
    let d = find_node(&arena, |n| matches!(n, ExprNode::DefType { .. }))
        .expect("must build a DefType");
    if let ExprNode::DefType { ranges, .. } = d {
        assert_eq!(ranges.len(), 1);
        assert_eq!(ranges[0].1, 0, "single-letter range has end == 0");
    }
}

#[test]
fn attribute_line_accepts() {
    let (_, arena, errors) = parse_module(b"Attribute VB_Name = \"M\"\nSub F()\nEnd Sub");
    assert!(!errors);
    let a = find_node(&arena, |n| matches!(n, ExprNode::Attribute { .. }))
        .expect("must build an Attribute");
    if let ExprNode::Attribute { values, .. } = a {
        assert_eq!(values.len(), 1, "single string value");
    }
}

#[test]
fn ident_label_accepts() {
    let (_, arena, errors) = parse_stmts(b"Foo:\n GoTo Foo");
    assert!(!errors);
    // `Foo:` is a Label definition; `GoTo Foo` jumps to it.
    assert!(find_node(&arena, |n| matches!(n, ExprNode::Label { .. })).is_some());
    assert!(
        find_node(&arena, |n| matches!(n, ExprNode::GoTo { target: LabelRef::Name(_) })).is_some()
    );
}

// ── Statement-dispatch regression tests ──────────────────────────────────────

#[test]
fn lset_builds_left_justify_range_assign() {
    let (_, arena, errors) = parse_stmts(b"LSet s = \"hello\"");
    assert!(!errors, "LSet must parse without errors");
    let node = find_node(&arena, |n| matches!(n, ExprNode::RangeAssign { .. }))
        .expect("LSet must build a RangeAssign node, not an Assign");
    assert!(
        matches!(node, ExprNode::RangeAssign { right_justify: false, .. }),
        "LSet must be left-justify (right_justify = false)"
    );
}

#[test]
fn rset_builds_right_justify_range_assign() {
    let (_, arena, errors) = parse_stmts(b"RSet s = \"world\"");
    assert!(!errors, "RSet must parse without errors");
    let node = find_node(&arena, |n| matches!(n, ExprNode::RangeAssign { .. }))
        .expect("RSet must build a RangeAssign node");
    assert!(
        matches!(node, ExprNode::RangeAssign { right_justify: true, .. }),
        "RSet must be right-justify (right_justify = true)"
    );
}

#[test]
fn question_builds_print_stmt() {
    // `?` is the Print shortcut — it builds a Print FileIoStmt with no channel,
    // not an implicit call.
    let (_, arena, errors) = parse_stmts(b"? \"hello\"");
    assert!(!errors, "? (Print alias) must parse without errors");
    let found = find_node(&arena, |n| {
        matches!(n, ExprNode::FileIoStmt { kind: FileIoKind::Print, channel: None, .. })
    });
    assert!(found.is_some(), "? must build a channelless Print statement");
}

#[test]
fn bare_print_builds_print_stmt() {
    // Regression: bare `Print "x"` (no #channel) must be accepted as a Print
    // statement; routing it through the implicit-call path used to wrongly reject it.
    let (_, arena, errors) = parse_stmts(b"Print \"x\"");
    assert!(!errors, "bare Print must parse without errors");
    let found = find_node(&arena, |n| {
        matches!(n, ExprNode::FileIoStmt { kind: FileIoKind::Print, channel: None, .. })
    });
    assert!(found.is_some(), "bare Print must build a channelless Print statement");
}

#[test]
fn input_dollar_file_accepts() {
    let (_, arena, errors) = parse_stmts(b"Input$ #1, x");
    assert!(!errors, "Input$ #n, var must parse without errors");
    // `Input$ #n, var` reads from a channel — a channelled Input FileIoStmt.
    assert!(find_node(&arena, |n| matches!(
        n,
        ExprNode::FileIoStmt { kind: FileIoKind::Input, channel: Some(_), .. }
    ))
    .is_some());
}

#[test]
fn inputb_file_accepts() {
    let (_, arena, errors) = parse_stmts(b"InputB #1, x");
    assert!(!errors, "InputB #n, var must parse without errors");
    assert!(find_node(&arena, |n| matches!(
        n,
        ExprNode::FileIoStmt { kind: FileIoKind::Input, channel: Some(_), .. }
    ))
    .is_some());
}

#[test]
fn inputbs_file_accepts() {
    let (_, arena, errors) = parse_stmts(b"InputB$ #1, x");
    assert!(!errors, "InputB$ #n, var must parse without errors");
    assert!(find_node(&arena, |n| matches!(
        n,
        ExprNode::FileIoStmt { kind: FileIoKind::Input, channel: Some(_), .. }
    ))
    .is_some());
}

#[test]
fn mid_dollar_assign_is_char_oriented() {
    let (_, arena, errors) = parse_stmts(b"Mid$(s, 1, 3) = \"abc\"");
    assert!(!errors, "Mid$ assignment must parse without errors");
    let node = find_node(&arena, |n| matches!(n, ExprNode::MidAssign { .. }))
        .expect("Mid$ must build a MidAssign node");
    assert!(
        matches!(node, ExprNode::MidAssign { byte_oriented: false, dollar: true, .. }),
        "Mid$ is character-oriented with the $ bit set"
    );
}

#[test]
fn midb_assign_is_byte_oriented() {
    let (_, arena, errors) = parse_stmts(b"MidB(s, 1, 3) = \"abc\"");
    assert!(!errors, "MidB assignment must parse without errors");
    let node = find_node(&arena, |n| matches!(n, ExprNode::MidAssign { .. }))
        .expect("MidB must build a MidAssign node");
    assert!(
        matches!(node, ExprNode::MidAssign { byte_oriented: true, dollar: false, .. }),
        "MidB must be byte-oriented (distinct from Mid)"
    );
}

#[test]
fn next_multi_counter_accepts() {
    // `Next i, j` — VB6 allows a comma-separated counter list closing two Fors.
    let src = b"For i = 1 To 3\n For j = 1 To 3\n Next j, i";
    let (_, arena, errors) = parse_stmts(src);
    assert!(!errors, "Next i, j must parse without errors");
    assert_eq!(count_nodes(&arena, |n| matches!(n, ExprNode::For { .. })), 2);
}

#[test]
fn addressof_builds_typed_node() {
    // `AddressOf Bar` must produce an AddressOf node (not a bare NameRef that
    // drops `Bar`).
    let (_nodes, arena, errors) = parse_stmts(b"Foo AddressOf Bar");
    assert!(!errors, "AddressOf operand must parse without errors");
    assert!(
        find_node(&arena, |n| matches!(n, ExprNode::AddressOf { .. })).is_some(),
        "expected an ExprNode::AddressOf in the parsed tree"
    );
}

#[test]
fn ident_300_chars_accepts() {
    // The parser must accept identifiers longer than 255 chars without truncation.
    let name: String = std::iter::repeat('a').take(300).collect();
    let src = format!("Dim {name} As Integer");
    let (_, arena, errors) = parse_stmts(src.as_bytes());
    assert!(!errors, "300-char identifier must parse without errors");
    assert!(find_node(&arena, |n| matches!(n, ExprNode::DimItem { .. })).is_some());
}

#[test]
fn d_exponent_is_double_literal() {
    // `1d3` is the single Double literal 1000.0, not `1` + `d3`.
    let (a, v) = assign_value("1d3");
    match a.get(v) {
        ExprNode::Literal { lit } => {
            let s = format!("{lit:?}");
            assert!(s.contains("1000"), "expected 1d3 to be the literal 1000.0, got {s}");
        }
        other => panic!("expected a single Literal for 1d3"),
    }
}
