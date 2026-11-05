//! `Exit` placement, `Next` control variables and `Do`/`Loop` conditions.

mod common;
use common::{offset_of, parse};
use vb6_syntax::frontend::ast::{DoKind, ExitKind, ExprNode};

const EXIT_FOR: u32 = 0x9c9b;
const EXIT_DO: u32 = 0x9c9c;
const EXIT_FUNCTION_IN_SUB_OR_PROPERTY: u32 = 0x9ca2;
const EXIT_SUB_IN_FUNCTION_OR_PROPERTY: u32 = 0x9ca3;
const EXIT_PROPERTY_IN_FUNCTION_OR_SUB: u32 = 0x9cbf;
const INVALID_NEXT_VARIABLE: u32 = 0x9ccd;
const FOR_VARIABLE_IN_USE: u32 = 0x9cce;
const NEXT_WITHOUT_FOR: u32 = 0x9c98;
const LOOP_WITHOUT_DO: u32 = 0x9c99;

fn exit_in(proc_src: &str, exit: &str, code: Option<u32>) {
    let src = proc_src.replace("EXIT", exit);
    let p = parse(&src);
    let expected: Vec<_> = code
        .map(|c| (c, offset_of(&src, exit, 0), 4))
        .into_iter()
        .collect();
    assert_eq!(p.found, expected, "for {src:?}");
}

#[test]
fn exit_for_and_do_need_an_enclosing_loop_of_that_kind() {
    exit_in("Sub S()\nEXIT\nEnd Sub\n", "Exit For", Some(EXIT_FOR));
    exit_in("Sub S()\nEXIT\nEnd Sub\n", "Exit Do", Some(EXIT_DO));
    exit_in("Sub S()\nDo\nEXIT\nLoop\nEnd Sub\n", "Exit For", Some(EXIT_FOR));
    exit_in("Sub S()\nFor i = 1 To 2\nEXIT\nNext\nEnd Sub\n", "Exit Do", Some(EXIT_DO));
    exit_in("Sub S()\nFor i = 1 To 2\nEXIT\nNext\nEnd Sub\n", "Exit For", None);
    exit_in("Sub S()\nFor Each v In c\nEXIT\nNext\nEnd Sub\n", "Exit For", None);
    exit_in("Sub S()\nDo While a\nEXIT\nLoop\nEnd Sub\n", "Exit Do", None);
    exit_in("Sub S()\nDo\nEXIT\nLoop Until a\nEnd Sub\n", "Exit Do", None);
}

#[test]
fn exit_finds_the_loop_through_other_blocks() {
    exit_in(
        "Sub S()\nFor i = 1 To 2\nWith o\nIf a Then\nSelect Case i\nCase 1\nEXIT\nEnd Select\nEnd If\nEnd With\nNext\nEnd Sub\n",
        "Exit For",
        None,
    );
    // A finished loop no longer counts.
    exit_in("Sub S()\nFor i = 1 To 2\nNext\nEXIT\nEnd Sub\n", "Exit For", Some(EXIT_FOR));
}

#[test]
fn exit_procedure_kind_must_match_the_procedure() {
    exit_in("Sub S()\nEXIT\nEnd Sub\n", "Exit Sub", None);
    exit_in("Sub S()\nEXIT\nEnd Sub\n", "Exit Function", Some(EXIT_FUNCTION_IN_SUB_OR_PROPERTY));
    exit_in("Sub S()\nEXIT\nEnd Sub\n", "Exit Property", Some(EXIT_PROPERTY_IN_FUNCTION_OR_SUB));
    exit_in("Function F()\nEXIT\nEnd Function\n", "Exit Function", None);
    exit_in("Function F()\nEXIT\nEnd Function\n", "Exit Sub", Some(EXIT_SUB_IN_FUNCTION_OR_PROPERTY));
    exit_in("Function F()\nEXIT\nEnd Function\n", "Exit Property", Some(EXIT_PROPERTY_IN_FUNCTION_OR_SUB));
    // A getter may be left with any Exit.
    let getter = "Property Get P()\nEXIT\nEnd Property\n";
    exit_in(getter, "Exit Property", None);
    exit_in(getter, "Exit Sub", None);
    exit_in(getter, "Exit Function", None);
    for head in ["Property Let P(v)", "Property Set P(v)"] {
        let src = format!("{head}\nEXIT\nEnd Property\n");
        exit_in(&src, "Exit Property", None);
        exit_in(&src, "Exit Sub", Some(EXIT_SUB_IN_FUNCTION_OR_PROPERTY));
        exit_in(&src, "Exit Function", Some(EXIT_FUNCTION_IN_SUB_OR_PROPERTY));
    }
}

#[test]
fn exit_inside_a_valid_loop_builds_the_exit_node() {
    let p = parse("Sub S()\nFor i = 1 To 2\nExit For\nNext\nEnd Sub\n");
    assert_eq!(p.found, vec![]);
    let body = p.proc_body(0);
    let ExprNode::For { body: for_body, .. } = p.node(body[0]) else { panic!("expected For") };
    let inner = p.block(*for_body);
    assert_eq!(inner.len(), 1);
    assert!(matches!(p.node(inner[0]), ExprNode::ExitStmt { kind: ExitKind::For }));
}

#[test]
fn next_variable_must_be_the_loops_control_variable() {
    let src = "Sub S()\nFor i = 1 To 2\nNext j\nEnd Sub\n";
    let p = parse(src);
    assert_eq!(p.found, vec![(INVALID_NEXT_VARIABLE, offset_of(src, "j", 0), 1)]);
    // The loop is still closed by that `Next`.
    assert_eq!(p.proc_body(0).len(), 1);

    let src = "Sub S()\nFor Each v In c\nNext w\nEnd Sub\n";
    assert_eq!(parse(src).found, vec![(INVALID_NEXT_VARIABLE, offset_of(src, "w", 0), 1)]);

    for ok in ["Next", "Next i", "Next I"] {
        let p = parse(&format!("Sub S()\nFor i = 1 To 2\n{ok}\nEnd Sub\n"));
        assert_eq!(p.found, vec![], "for {ok}");
    }
}

#[test]
fn one_next_can_close_several_nested_loops() {
    let p = parse("Sub S()\nFor i = 1 To 2\nFor j = 1 To 2\nx = 1\nNext j, i\ny = 2\nEnd Sub\n");
    assert_eq!(p.found, vec![]);
    let body = p.proc_body(0);
    assert_eq!(body.len(), 2, "the statement after `Next j, i` is outside both loops");
    let ExprNode::For { body: outer, .. } = p.node(body[0]) else { panic!("expected For") };
    let outer = p.block(*outer);
    assert_eq!(outer.len(), 1);
    let ExprNode::For { body: inner, .. } = p.node(outer[0]) else { panic!("expected inner For") };
    assert_eq!(p.block(*inner).len(), 1);
    assert!(matches!(p.node(body[1]), ExprNode::Assign { .. }));

    // Names are matched innermost first.
    let src = "Sub S()\nFor i = 1 To 2\nFor j = 1 To 2\nNext i, j\nEnd Sub\n";
    assert_eq!(
        parse(src).found,
        vec![
            (INVALID_NEXT_VARIABLE, offset_of(src, "i", 1), 1),
            (INVALID_NEXT_VARIABLE, offset_of(src, "j", 1), 1),
        ]
    );
}

#[test]
fn next_names_beyond_the_directly_nested_loops_are_stray() {
    let src = "Sub S()\nFor i = 1 To 2\nNext i, j\nEnd Sub\n";
    assert_eq!(parse(src).found, vec![(NEXT_WITHOUT_FOR, offset_of(src, "j", 0), 1)]);
    // A block between the loops breaks the nesting.
    let src = "Sub S()\nFor i = 1 To 2\nIf a Then\nFor j = 1 To 2\nNext j, i\nEnd If\nNext\nEnd Sub\n";
    assert_eq!(parse(src).found, vec![(NEXT_WITHOUT_FOR, offset_of(src, "i", 1), 1)]);
}

#[test]
fn nested_loops_cannot_reuse_a_control_variable() {
    let src = "Sub S()\nFor i = 1 To 2\nFor i = 1 To 3\nNext\nNext\nEnd Sub\n";
    assert_eq!(parse(src).found, vec![(FOR_VARIABLE_IN_USE, offset_of(src, "i", 1), 1)]);
    let src = "Sub S()\nFor i = 1 To 2\nFor Each i In c\nNext\nNext\nEnd Sub\n";
    assert_eq!(parse(src).found, vec![(FOR_VARIABLE_IN_USE, offset_of(src, "i", 1), 1)]);
    // Sequential loops may share one.
    let p = parse("Sub S()\nFor i = 1 To 2\nNext\nFor i = 1 To 2\nNext\nEnd Sub\n");
    assert_eq!(p.found, vec![]);
}

#[test]
fn a_loop_tests_its_condition_at_the_top_or_the_bottom_not_both() {
    let src = "Sub S()\nDo While a\nLoop While b\nEnd Sub\n";
    let p = parse(src);
    assert_eq!(p.found, vec![(LOOP_WITHOUT_DO, offset_of(src, "Loop", 0), 4)]);

    let src = "Sub S()\nDo Until a\nLoop Until b\nEnd Sub\n";
    assert_eq!(parse(src).found, vec![(LOOP_WITHOUT_DO, offset_of(src, "Loop", 0), 4)]);

    for (ok, kind) in [
        ("Do While a\nLoop", DoKind::PreWhile),
        ("Do Until a\nLoop", DoKind::PreUntil),
        ("Do\nLoop While a", DoKind::PostWhile),
        ("Do\nLoop Until a", DoKind::PostUntil),
        ("Do\nLoop", DoKind::Inf),
    ] {
        let p = parse(&format!("Sub S()\n{ok}\nEnd Sub\n"));
        assert_eq!(p.found, vec![], "for {ok:?}");
        let body = p.proc_body(0);
        let ExprNode::Do { kind: got, .. } = p.node(body[0]) else { panic!("expected Do") };
        assert_eq!(*got, kind, "for {ok:?}");
    }
}
