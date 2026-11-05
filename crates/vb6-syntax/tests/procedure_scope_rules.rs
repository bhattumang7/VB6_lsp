//! Executable statements are invalid outside a procedure; declarations of
//! procedures, types and options are invalid inside one.

mod common;
use common::{offset_of, parse};
use vb6_syntax::frontend::ast::ExprNode;

const INVALID_OUTSIDE_PROCEDURE: u32 = 0x9c73;
const INVALID_INSIDE_PROCEDURE: u32 = 0x9c74;

#[test]
fn statements_outside_a_procedure_are_reported_once_each() {
    for (stmt, len) in [
        ("x = 1", 1),
        ("Call Foo", 4),
        ("Foo 1, 2", 3),
        ("Debug.Print 1", 5),
        ("GoTo Done", 4),
        ("Exit Sub", 4),
        ("Set o = Nothing", 3),
        ("Print #1, x", 5),
        ("10 x = 1", 2),
        ("x = 1: y = 2", 1),
    ] {
        let src = format!("{stmt}\nSub S()\nEnd Sub\n");
        let p = parse(&src);
        assert_eq!(p.found, vec![(INVALID_OUTSIDE_PROCEDURE, 0, len)], "for {stmt:?}");
        assert_eq!(p.top.len(), 1, "only the procedure is kept for {stmt:?}");
    }
}

#[test]
fn a_block_statement_outside_a_procedure_is_one_error() {
    let src = "If a Then\nx = 1\nElse\nx = 2\nEnd If\nFor i = 1 To 3\nNext\nSub S()\nEnd Sub\n";
    let p = parse(src);
    assert_eq!(
        p.found,
        vec![
            (INVALID_OUTSIDE_PROCEDURE, 0, 2),
            (INVALID_OUTSIDE_PROCEDURE, offset_of(src, "For", 0), 3),
        ]
    );
    assert_eq!(p.top.len(), 1);
}

#[test]
fn declarations_outside_a_procedure_stay_clean() {
    let p = parse("Option Explicit\nPublic x As Long\nConst c = 1\nDeclare Sub D Lib \"k\" ()\nEnum E\nA\nEnd Enum\nType T\nm As Long\nEnd Type\n");
    assert_eq!(p.found, vec![]);
}

#[test]
fn procedure_declarations_inside_a_procedure_are_reported_and_skipped() {
    for (decl, len) in [
        ("Sub Inner()", 3),
        ("Function Inner()", 8),
        ("Property Get Inner()", 8),
        ("Declare Sub D Lib \"k\" ()", 7),
        ("Event Ev()", 5),
        ("Option Explicit", 6),
        ("DefInt A-Z", 6),
    ] {
        let src = format!("Sub S()\n{decl}\nx = 1\nEnd Sub\n");
        let p = parse(&src);
        let mut found = p.found.clone();
        // Class-only keywords in a standard module are reported as well.
        found.retain(|d| d.0 == INVALID_INSIDE_PROCEDURE);
        assert_eq!(
            found,
            vec![(INVALID_INSIDE_PROCEDURE, offset_of(&src, decl, 0), len)],
            "for {decl:?}"
        );
        let body = p.proc_body(0);
        assert_eq!(body.len(), 1, "for {decl:?}");
        assert!(matches!(p.node(body[0]), ExprNode::Assign { .. }), "for {decl:?}");
    }
}

#[test]
fn type_and_enum_blocks_inside_a_procedure_are_one_error() {
    let src = "Sub S()\nType T\nm As Long\nEnd Type\nx = 1\nEnum E\nA\nEnd Enum\nEnd Sub\n";
    let p = parse(src);
    assert_eq!(
        p.found,
        vec![
            (INVALID_INSIDE_PROCEDURE, offset_of(src, "Type", 0), 4),
            (INVALID_INSIDE_PROCEDURE, offset_of(src, "Enum", 0), 4),
        ]
    );
    let body = p.proc_body(0);
    assert_eq!(body.len(), 1);
    assert!(matches!(p.node(body[0]), ExprNode::Assign { .. }));
}
