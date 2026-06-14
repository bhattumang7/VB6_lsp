//! Confirms the exact snippets used against the VB6 `/make` oracle also parse
//! through the Rust parser and produce the expected AST shapes.
//!
//! Every snippet here is ACCEPTED by the real VB6 SP6 compiler and must parse
//! through the Rust parser into the matching AST node.
//!
//! Run: cargo test -p vb6-core --test oracle_examples_ast -- --nocapture

use vb6_core::frontend::ast::{AstLit, ExprArena, ExprNode, LabelRef, NodeId, OnErrorKind, ResumeTarget};
use vb6_core::frontend::parser::Parser;
use vb6_core::frontend::scanner::ScannerContext;

const CONTROL: &str = "Attribute VB_Name = \"Module1\"\n\
Sub Main()\n    Dim x As Integer\n    x = 1\n    If x = 1 Then GoTo done\n    x = 2\ndone:\n    x = 3\nEnd Sub\n";

const ENDIF: &str = "Attribute VB_Name = \"Module1\"\n\
Sub Main()\n    Dim x As Integer\n    x = 1\n    If x = 1 Then\n        x = 2\n    EndIf\nEnd Sub\n";

const IFGOTO: &str = "Attribute VB_Name = \"Module1\"\n\
Sub Main()\n    Dim x As Integer\n    x = 1\n    If x = 1 GoTo done\n    x = 2\ndone:\n    x = 3\nEnd Sub\n";

const GOTO_SPACE: &str = "Attribute VB_Name = \"Module1\"\n\
Sub Main()\n    Dim x As Integer\n    x = 1\n    Go To done\n    x = 2\ndone:\n    x = 3\nEnd Sub\n";

const GOSUB_SPACE: &str = "Attribute VB_Name = \"Module1\"\n\
Sub Main()\n    Dim x As Integer\n    Go Sub handler\n    Exit Sub\nhandler:\n    x = 2\n    Return\nEnd Sub\n";

fn parse(src: &str) -> (ExprArena, bool) {
    let (_c, a, errs) = parse_ctx(src);
    (a, errs)
}

/// Like [`parse`], but also returns the [`ScannerContext`] so a captured symbol
/// id can be resolved back to its source identifier text.
fn parse_ctx(src: &str) -> (ScannerContext, ExprArena, bool) {
    let mut c = ScannerContext::new(1, 1, 0x0409);
    c.intern_keywords();
    let mut arena = ExprArena::new();
    let mut parser = Parser::new(&mut c, src.as_bytes());
    let _ = parser.parse_module(&mut arena);
    let has_errors = parser.diagnostics.has_errors();
    (c, arena, has_errors)
}

fn count(arena: &ExprArena, pred: impl Fn(&ExprNode) -> bool) -> usize {
    (0..arena.len()).filter(|&i| pred(arena.get(NodeId(i as u32)))).count()
}
fn has(a: &ExprArena, p: impl Fn(&ExprNode) -> bool) -> bool { count(a, p) > 0 }

/// Wrap a procedure body fragment in a minimal module + Sub.
fn in_sub(body: &str) -> String {
    format!("Attribute VB_Name = \"Module1\"\nSub Main()\n{body}\nEnd Sub\n")
}

fn find_resume(a: &ExprArena) -> ResumeTarget {
    for i in 0..a.len() {
        if let ExprNode::Resume { target } = a.get(NodeId(i as u32)) {
            return *target;
        }
    }
    panic!("no Resume node found");
}

fn find_on_error(a: &ExprArena) -> OnErrorKind {
    for i in 0..a.len() {
        if let ExprNode::OnError { kind } = a.get(NodeId(i as u32)) {
            return *kind;
        }
    }
    panic!("no OnError node found");
}

// ── Resume statement: all four VB6 forms (operand must be captured, never
//     dropped and re-parsed as a separate statement) ───────────────────────────

#[test]
fn resume_bare_is_retry() {
    let (a, errs) = parse(&in_sub("    On Error Resume Next\n    Resume"));
    assert!(!errs, "bare Resume must parse without errors");
    assert_eq!(find_resume(&a), ResumeTarget::Retry);
}

#[test]
fn resume_next() {
    // Two `Resume`-class nodes: the `On Error Resume Next` is an OnError node,
    // the standalone `Resume Next` is a Resume node with target Next.
    let (a, errs) = parse(&in_sub("    Resume Next"));
    assert!(!errs, "Resume Next must parse without errors");
    assert_eq!(find_resume(&a), ResumeTarget::Next);
}

#[test]
fn resume_label_operand_captured() {
    // R1: `Resume <label>` must record the label on the Resume node, not leave
    // it in the stream to be mis-parsed as a separate statement.
    let (c, a, errs) = parse_ctx(&in_sub("    Resume done\ndone:\n    Dim y As Integer"));
    assert!(!errs, "Resume <label> must parse without errors");
    let label_sym = match find_resume(&a) {
        ResumeTarget::At(LabelRef::Name(sym)) => {
            assert!(sym != 0, "label symbol must be interned");
            // Verify the captured symbol is the *correct* label, not just any.
            assert_eq!(c.sym_name(sym as usize).to_ascii_lowercase(), "done",
                "Resume must capture the `done` label, not some other symbol");
            sym
        }
        other => panic!("expected Resume At(Name), got {other:?}"),
    };
    // Exactly one Resume; the operand was consumed, not turned into a stray stmt.
    assert_eq!(count(&a, |n| matches!(n, ExprNode::Resume { .. })), 1);
    // Airtight leak check: under the old bug, `done` was left in the stream and
    // re-parsed as a standalone implicit-call/name-reference statement. The only
    // remaining `NameRef(done)` would be that leaked node — assert there is none.
    // (The label definition `done:` is a `Label` node, not a `NameRef`.)
    assert_eq!(
        count(&a, |n| matches!(n, ExprNode::NameRef { sym, .. } if *sym == label_sym)),
        0,
        "Resume operand must not leak as a separate name-reference statement",
    );
}

#[test]
fn resume_zero_is_line_zero() {
    let (a, errs) = parse(&in_sub("    Resume 0"));
    assert!(!errs, "Resume 0 must parse without errors");
    assert_eq!(find_resume(&a), ResumeTarget::At(LabelRef::Line(0)));
}

#[test]
fn resume_numeric_line_operand_captured() {
    let (a, errs) = parse(&in_sub("    Resume 100\n100:\n    Dim y As Integer"));
    assert!(!errs, "Resume <line#> must parse without errors");
    assert_eq!(find_resume(&a), ResumeTarget::At(LabelRef::Line(100)));
}

// ── On Error GoTo: 0 disables, nonzero line installs a handler ────────────────

#[test]
fn on_error_goto_zero_disables() {
    let (a, errs) = parse(&in_sub("    On Error GoTo 0"));
    assert!(!errs);
    assert_eq!(find_on_error(&a), OnErrorKind::Disable);
}

#[test]
fn on_error_goto_label() {
    let (c, a, errs) = parse_ctx(&in_sub("    On Error GoTo handler\nhandler:\n    Dim y As Integer"));
    assert!(!errs);
    match find_on_error(&a) {
        OnErrorKind::Goto(LabelRef::Name(sym)) => {
            assert!(sym != 0);
            assert_eq!(c.sym_name(sym as usize).to_ascii_lowercase(), "handler",
                "On Error GoTo must capture the `handler` label");
        }
        other => panic!("expected Goto(Name), got {other:?}"),
    }
}

#[test]
fn on_error_goto_numeric_line_is_not_disable() {
    // `On Error GoTo 100` installs a handler at line 100 — it is NOT `Disable`.
    let (a, errs) = parse(&in_sub("    On Error GoTo 100\n100:\n    Dim y As Integer"));
    assert!(!errs);
    assert_eq!(find_on_error(&a), OnErrorKind::Goto(LabelRef::Line(100)));
}

// ── Numeric line labels are preserved as jump targets ─────────────────────────

// ── With-block leading-dot / bang member references ───────────────────────────

/// Return the single `Assign { target, value }` node in the arena (panics if
/// there is not exactly one — keeps the structural assertions unambiguous).
fn the_assign(a: &ExprArena) -> (NodeId, NodeId) {
    let mut found = None;
    for i in 0..a.len() {
        if let ExprNode::Assign { target, value } = a.get(NodeId(i as u32)) {
            assert!(found.is_none(), "expected exactly one Assign node");
            found = Some((*target, *value));
        }
    }
    found.expect("no Assign node found")
}

/// Assert `id` is `MemberAccess { base: WithContext, member, bang }` with the
/// member symbol resolving (case-insensitively) to `name`.
fn assert_with_member(c: &ScannerContext, a: &ExprArena, id: NodeId, name: &str, bang: bool) {
    match a.get(id) {
        ExprNode::MemberAccess { base, member, bang: b } => {
            assert_eq!(*b, bang, "wrong bang flag for `{name}`");
            assert!(matches!(a.get(*base), ExprNode::WithContext),
                "base of `{name}` must be the implicit WithContext node");
            assert_eq!(c.sym_name(*member as usize).to_ascii_lowercase(), name,
                "member symbol must resolve to `{name}`");
        }
        _ => panic!("expected MemberAccess on WithContext for `{name}`"),
    }
}

/// Assert no `NameRef` in the arena resolves to `name` — proves the member was
/// not collapsed into a bare name reference (the dropped-dot bug).
fn assert_no_nameref(c: &ScannerContext, a: &ExprArena, name: &str) {
    for i in 0..a.len() {
        if let ExprNode::NameRef { sym, .. } = a.get(NodeId(i as u32)) {
            assert_ne!(c.sym_name(*sym as usize).to_ascii_lowercase(), name,
                "`{name}` leaked as a bare NameRef instead of a With member access");
        }
    }
}

#[test]
fn with_member_target_builds_member_access() {
    // `.Field = 1` — the assignment target is a MemberAccess on the implicit
    // With object; the value is the literal 1; and `Field` is NOT a bare name.
    let (c, a, errs) = parse_ctx(&in_sub("    With obj\n        .Field = 1\n    End With"));
    assert!(!errs, "With-block `.Field = 1` must parse without errors");
    let (target, value) = the_assign(&a);
    assert_with_member(&c, &a, target, "field", false);
    assert!(matches!(a.get(value), ExprNode::Literal { lit: AstLit::Int(1) }), "RHS must be literal 1");
    assert_no_nameref(&c, &a, "field");
    // The With subject `obj` is still a real name reference.
    assert!(has(&a, |n| matches!(n, ExprNode::With { .. })), "expected With node");
}

#[test]
fn with_member_expression_position_builds_member_access() {
    // `x = .Field` — the *value* side is the With member access; the target is
    // the plain variable `x`.
    let (c, a, errs) = parse_ctx(&in_sub("    Dim x\n    With obj\n        x = .Field\n    End With"));
    assert!(!errs, "With-block `.Field` in expression position must parse");
    let (target, value) = the_assign(&a);
    assert!(matches!(a.get(target), ExprNode::NameRef { sym, .. }
        if c.sym_name(*sym as usize).eq_ignore_ascii_case("x")), "target must be NameRef `x`");
    assert_with_member(&c, &a, value, "field", false);
    assert_no_nameref(&c, &a, "field");
}

#[test]
fn with_member_bang_builds_member_access() {
    // `!Key = 1` — leading bang (default-member access) in a With block: a
    // bang MemberAccess on WithContext, value literal 1, `Key` not a bare name.
    let (c, a, errs) = parse_ctx(&in_sub("    With obj\n        !Key = 1\n    End With"));
    assert!(!errs, "With-block `!Key = 1` must parse without errors");
    let (target, value) = the_assign(&a);
    assert_with_member(&c, &a, target, "key", true);
    assert!(matches!(a.get(value), ExprNode::Literal { lit: AstLit::Int(1) }), "RHS must be literal 1");
    assert_no_nameref(&c, &a, "key");
}

#[test]
fn named_goto_and_label_resolve_to_same_symbol() {
    // The `GoTo done` target and the `done:` definition must both carry the
    // interned symbol for `done` — and it must be the *same* symbol.
    let (c, a, errs) = parse_ctx(&in_sub("    GoTo done\ndone:\n    Dim y As Integer"));
    assert!(!errs);
    let goto_sym = (0..a.len()).find_map(|i| match a.get(NodeId(i as u32)) {
        ExprNode::GoTo { target: LabelRef::Name(s) } => Some(*s),
        _ => None,
    }).expect("expected GoTo(Name) node");
    let label_sym = (0..a.len()).find_map(|i| match a.get(NodeId(i as u32)) {
        ExprNode::Label { target: LabelRef::Name(s) } => Some(*s),
        _ => None,
    }).expect("expected Label(Name) node");
    assert_eq!(c.sym_name(goto_sym as usize).to_ascii_lowercase(), "done");
    assert_eq!(c.sym_name(label_sym as usize).to_ascii_lowercase(), "done");
    assert_eq!(goto_sym, label_sym, "GoTo target and label definition must share one symbol");
}

#[test]
fn numeric_line_label_definition_preserved() {
    let (a, errs) = parse(&in_sub("    GoTo 100\n100:\n    Dim y As Integer"));
    assert!(!errs, "numeric line label must parse without errors");
    // GoTo carries the numeric target.
    assert!(has(&a, |n| matches!(n, ExprNode::GoTo { target: LabelRef::Line(100) })),
        "GoTo 100 must carry Line(100)");
    // The `100:` definition is emitted as a Label node carrying the line number.
    assert!(has(&a, |n| matches!(n, ExprNode::Label { target: LabelRef::Line(100) })),
        "numeric label definition must be a Label(Line(100)) node");
}

#[test]
fn control_if_then_goto() {
    let (a, errs) = parse(CONTROL);
    assert!(!errs, "If…Then GoTo must parse without errors");
    assert!(has(&a, |n| matches!(n, ExprNode::If { .. })), "expected If node");
    assert!(has(&a, |n| matches!(n, ExprNode::GoTo { .. })), "expected GoTo node");
    assert!(has(&a, |n| matches!(n, ExprNode::Label { .. })), "expected Label node");
}

#[test]
fn endif_one_word() {
    let (a, errs) = parse(ENDIF);
    assert!(!errs, "one-word EndIf must parse without errors");
    assert!(has(&a, |n| matches!(n, ExprNode::If { .. })), "expected If node");
}

#[test]
fn if_goto_no_then() {
    let (a, errs) = parse(IFGOTO);
    assert!(!errs, "If cond GoTo (no Then) must parse without errors");
    assert!(has(&a, |n| matches!(n, ExprNode::If { .. })), "expected If node");
    assert!(has(&a, |n| matches!(n, ExprNode::GoTo { .. })), "expected GoTo node in If body");
    assert!(has(&a, |n| matches!(n, ExprNode::Label { .. })), "expected Label node");
}

#[test]
fn go_to_space() {
    let (a, errs) = parse(GOTO_SPACE);
    assert!(!errs, "Go To (space form) must parse without errors");
    assert_eq!(count(&a, |n| matches!(n, ExprNode::GoTo { .. })), 1,
        "expected exactly one GoTo node for `Go To done`");
    assert_eq!(count(&a, |n| matches!(n, ExprNode::GoSub { .. })), 0, "must not be a GoSub");
    assert!(has(&a, |n| matches!(n, ExprNode::Label { .. })), "expected Label node");
}

#[test]
fn go_sub_space() {
    let (a, errs) = parse(GOSUB_SPACE);
    assert!(!errs, "Go Sub (space form) must parse without errors");
    assert_eq!(count(&a, |n| matches!(n, ExprNode::GoSub { .. })), 1,
        "expected exactly one GoSub node for `Go Sub handler`");
    assert_eq!(count(&a, |n| matches!(n, ExprNode::GoTo { .. })), 0, "must not be a GoTo");
    assert!(has(&a, |n| matches!(n, ExprNode::Label { .. })), "expected Label node");
}
