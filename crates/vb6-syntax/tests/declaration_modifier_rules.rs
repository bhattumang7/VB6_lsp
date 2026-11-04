//! Declaration modifier combinations, and `On` jump statements.

mod common;
use common::{offset_of, parse};
use vb6_syntax::frontend::ast::{ExprArena, ExprNode, LabelRef, OnErrorKind};
use vb6_syntax::frontend::parser::{ModuleKind, Parser};
use vb6_syntax::frontend::scanner::ScannerContext;

const EXPECTED: u32 = 0x9c6f;
const OBJECT_MODULE_ONLY: u32 = 0xdee1;

fn found_in(kind: ModuleKind, src: &str) -> Vec<(u32, u32, u32)> {
    let mut ctx = ScannerContext::new(1, 1, 0x0409);
    ctx.intern_keywords();
    let mut arena = ExprArena::new();
    let mut parser = Parser::with_module_kind(&mut ctx, src.as_bytes(), kind);
    parser.parse_module(&mut arena);
    parser.diagnostics.items().iter().map(|d| (d.code, d.span.start, d.span.len)).collect()
}

#[test]
fn static_and_access_modifiers_combine_in_either_order() {
    for head in ["Static Private Sub", "Private Static Sub", "Static Public Sub", "Public Static Sub"] {
        let src = format!("{head} S()\nEnd Sub\n");
        let p = parse(&src);
        assert_eq!(p.found, vec![], "{head}");
        assert_eq!(p.top.len(), 1, "{head}");
        assert!(matches!(p.node(p.top[0]), ExprNode::ProcDecl { .. }), "{head}");
    }
}

#[test]
fn static_with_an_access_modifier_cannot_declare_a_variable() {
    for head in ["Public Static x As Long", "Private Static x As Long", "Static Private x As Long"] {
        let src = format!("{head}\n");
        let at = offset_of(&src, "x", 0);
        assert_eq!(
            parse(&src).found,
            vec![(EXPECTED, at, 1)],
            "{head}"
        );
    }
}

#[test]
fn friend_qualifies_only_a_procedure() {
    for head in ["Friend x As Long", "Friend Const x = 1", "Friend Dim x"] {
        let src = format!("{head}\n");
        let at = offset_of(&src, head.split(' ').nth(1).unwrap(), 0);
        let len = head.split(' ').nth(1).unwrap().len() as u32;
        assert_eq!(
            found_in(ModuleKind::Class, &src),
            vec![(EXPECTED, at, len)],
            "{head}"
        );
    }
    assert_eq!(found_in(ModuleKind::Class, "Friend Sub S()\nEnd Sub\n"), vec![]);
}

#[test]
fn friend_is_reported_in_a_standard_module() {
    let src = "Friend Sub S()\nEnd Sub\n";
    assert_eq!(found_in(ModuleKind::Standard, src), vec![(OBJECT_MODULE_ONLY, 0, 6)]);
}

#[test]
fn only_public_declares_an_event() {
    for head in ["Private", "Global", "Friend"] {
        let src = format!("{head} Event E()\n");
        let at = offset_of(&src, "Event", 0);
        assert_eq!(found_in(ModuleKind::Class, &src), vec![(EXPECTED, at, 5)], "{head}");
    }
    assert_eq!(found_in(ModuleKind::Class, "Public Event E()\n"), vec![]);
    assert_eq!(found_in(ModuleKind::Class, "Event E()\n"), vec![]);
}

#[test]
fn on_error_goto_minus_one_is_valid() {
    let p = parse("Sub S()\n On Error GoTo -1\nEnd Sub\n");
    assert_eq!(p.found, vec![]);
    let body = p.proc_body(0);
    assert_eq!(body.len(), 1);
    assert!(matches!(p.node(body[0]), ExprNode::OnError { kind: OnErrorKind::ClearError }));
}

#[test]
fn on_error_goto_other_negative_numbers_stay_invalid() {
    let src = "Sub S()\n On Error GoTo -2\nEnd Sub\n";
    assert!(!parse(src).found.is_empty());
}

#[test]
fn on_goto_targets_are_jump_targets_not_variable_reads() {
    let p = parse("Sub S()\n On x GoTo L1, 20\nL1:\n 20 x = 1\nEnd Sub\n");
    assert_eq!(p.found, vec![]);
    let body = p.proc_body(0);
    let ExprNode::OnGo { is_gosub, labels, .. } = p.node(body[0]) else { panic!("not OnGo") };
    assert!(!*is_gosub);
    assert_eq!(labels.len(), 2);
    let ExprNode::JumpTarget { target: LabelRef::Name(sym) } = p.node(labels[0]) else {
        panic!("first target")
    };
    assert_eq!(p.name(*sym), "L1");
    assert!(matches!(p.node(labels[1]), ExprNode::JumpTarget { target: LabelRef::Line(20) }));
}
