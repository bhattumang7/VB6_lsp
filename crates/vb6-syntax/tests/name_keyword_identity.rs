//! A name written in square brackets keeps the identity of the keyword it spells wherever the
//! grammar asks which word a name is, and nowhere else.

mod common;
use common::{offset_of, parse, parse_as};
use vb6_syntax::frontend::ast::{AstLit, ExprNode, FileIoKind, OnErrorKind};
use vb6_syntax::frontend::parser::ModuleKind;

const EXPECTED: u32 = 0x9c6f;
const SYNTAX_ERROR: u32 = 0x9c70;
const OPTION_PRIVATE_IN_OBJECT_MODULE: u32 = 0xdec9;

// ── Option ────────────────────────────────────────────────────────────────

#[test]
fn bracketed_option_words_are_the_option_words() {
    let p = parse("Option [Explicit]\n");
    assert_eq!(p.found, vec![]);
    assert_eq!(p.top.len(), 1);
    assert!(matches!(p.node(p.top[0]), ExprNode::OptionExplicit));

    let p = parse("Option [Base] 1\n");
    assert_eq!(p.found, vec![]);
    assert!(matches!(p.node(p.top[0]), ExprNode::OptionBase { value: 1 }));

    for (src, mode) in [
        ("Option [Compare] Text\n", 1),
        ("Option Compare [Text]\n", 1),
        ("Option Compare [Binary]\n", 0),
        ("Option [Compare] [Binary]\n", 0),
        ("Option compare [TEXT]\n", 1),
    ] {
        let p = parse(src);
        assert_eq!(p.found, vec![], "for {src:?}");
        assert_eq!(p.top.len(), 1, "for {src:?}");
        assert!(matches!(p.node(p.top[0]), ExprNode::OptionCompare { mode: m } if *m == mode), "for {src:?}");
    }
}

#[test]
fn bracketed_private_module_is_accepted_but_a_bracketed_private_is_not() {
    let p = parse("Option Private [Module]\n");
    assert_eq!(p.found, vec![]);

    // `Private` is a reserved word: written in brackets it is only a name.
    let src = "Option [Private] Module\n";
    let p = parse(src);
    assert_eq!(p.found, vec![(EXPECTED, offset_of(src, "[Private]", 0), 9)]);
}

#[test]
fn a_bracketed_word_that_is_not_an_option_word_is_rejected() {
    // `Database` is not accepted by `Option Compare`, bracketed or not.
    for src in ["Option Compare [Database]\n", "Option Compare Database\n"] {
        let p = parse(src);
        assert_eq!(p.messages, vec!["Expected: Text or Binary".to_string()], "for {src:?}");
        assert_eq!(p.found.len(), 1);
    }
    let p = parse("Option [Foo]\n");
    assert_eq!(p.found.len(), 1);
    assert_eq!(p.found[0].0, EXPECTED);
}

#[test]
fn option_private_is_checked_for_the_module_before_the_module_kind() {
    // In an object module `Option Private Module` is refused as a whole; anything else after
    // `Private` is first of all not the word `Module`.
    let head = "VERSION 1.0 CLASS\nBEGIN\nEND\n";
    let src = format!("{head}Option Private Foo\n");
    let p = parse_as(&src, ModuleKind::Class);
    assert_eq!(p.messages, vec!["Expected: Module".to_string()]);
    assert_eq!(p.found, vec![(EXPECTED, offset_of(&src, "Foo", 0), 3)]);

    let src = format!("{head}Option Private\n");
    let p = parse_as(&src, ModuleKind::Class);
    assert_eq!(p.messages, vec!["Expected: Module".to_string()]);

    let src = format!("{head}Option Private Module\n");
    let p = parse_as(&src, ModuleKind::Class);
    assert_eq!(p.found, vec![(OPTION_PRIVATE_IN_OBJECT_MODULE, offset_of(&src, "Option", 0), 6)]);

    let src = format!("{head}Option Private [Module]\n");
    let p = parse_as(&src, ModuleKind::Class);
    assert_eq!(p.found, vec![(OPTION_PRIVATE_IN_OBJECT_MODULE, offset_of(&src, "Option", 0), 6)]);
}

// ── On Error, For Step ─────────────────────────────────────────────────────

#[test]
fn bracketed_error_after_on_is_the_error_word() {
    let p = parse("Sub S()\nOn [Error] Resume Next\nOn [error] GoTo 0\nOn Local [Error] Resume Next\nEnd Sub\n");
    assert_eq!(p.found, vec![]);
    let body = p.proc_body(0);
    assert_eq!(body.len(), 3);
    assert!(matches!(p.node(body[0]), ExprNode::OnError { kind: OnErrorKind::ResumeNext }));
    assert!(matches!(p.node(body[1]), ExprNode::OnError { kind: OnErrorKind::Disable }));
    assert!(matches!(p.node(body[2]), ExprNode::OnError { kind: OnErrorKind::ResumeNext }));
}

#[test]
fn bracketed_step_in_a_for_header_is_the_step_word() {
    let p = parse("Sub S()\nDim i As Long\nFor i = 1 To 9 [Step] 2\nNext\nEnd Sub\n");
    assert_eq!(p.found, vec![]);
    let body = p.proc_body(0);
    let ExprNode::For { start, end, step, .. } = p.node(body[1]) else { panic!("expected For") };
    assert!(matches!(p.node(*start), ExprNode::Literal { lit: AstLit::Int(1) }));
    assert!(matches!(p.node(*end), ExprNode::Literal { lit: AstLit::Int(9) }));
    let step = step.expect("a step");
    assert!(matches!(p.node(step), ExprNode::Literal { lit: AstLit::Int(2) }));
}

// ── Open ───────────────────────────────────────────────────────────────────

fn open_of(clauses: &str) -> (common::Parsed, Option<(u8, u8)>) {
    let p = parse(&format!("Sub S()\nOpen \"f\" {clauses} As #1\nEnd Sub\n"));
    let found = if p.found.is_empty() {
        match p.node(p.proc_body(0)[0]) {
            ExprNode::FileIoStmt { kind: FileIoKind::Open, open_mode, open_flags, .. } => Some((*open_mode, *open_flags)),
            other => panic!("expected Open, got {:?}", std::mem::discriminant(other)),
        }
    } else {
        None
    };
    (p, found)
}

#[test]
fn bracketed_open_mode_words_are_the_mode_words() {
    for (word, mode) in [("Input", 1), ("Output", 2), ("Random", 4), ("Append", 8), ("Binary", 0x20)] {
        let (p, got) = open_of(&format!("For [{word}]"));
        assert_eq!(p.found, vec![], "for {word}");
        assert_eq!(got, Some((mode, 0)), "for {word}");
    }
}

#[test]
fn bracketed_access_and_lock_words_are_the_access_and_lock_words() {
    // `(clause, open_flags)`: the lock code in the high nibble, the access code in the low one.
    for (clause, flags) in [
        ("[Access] Read", 0x01),
        ("Access [Read]", 0x01),
        ("Access [Write]", 0x02),
        ("Access [Read] Write", 0x03),
        ("[Access] [Read] Write", 0x03),
        ("Lock [Read]", 0x30),
        ("Lock [Write]", 0x20),
        ("Lock [Read] Write", 0x10),
        ("Access [Read] Lock [Write]", 0x21),
    ] {
        let (p, got) = open_of(&format!("For Binary {clause}"));
        assert_eq!(p.found, vec![], "for {clause}");
        assert_eq!(got, Some((0x20, flags)), "for {clause}");
    }
}

#[test]
fn the_second_word_of_read_write_must_be_the_bare_keyword() {
    for clause in ["Access Read [Write]", "Lock Read [Write]"] {
        let src = format!("Sub S()\nOpen \"f\" For Binary {clause} As #1\nEnd Sub\n");
        let p = parse(&src);
        assert_eq!(p.found[0].0, SYNTAX_ERROR, "for {clause}");
        assert_eq!(p.found[0].1, offset_of(&src, "[Write]", 0), "for {clause}");
    }
    let (p, got) = open_of("For Binary Lock Read Write");
    assert_eq!(p.found, vec![]);
    assert_eq!(got, Some((0x20, 0x10)));
}

#[test]
fn bracketed_shared_lock_and_len_are_not_those_words() {
    for clause in ["[Shared]", "[Lock] Read"] {
        let src = format!("Sub S()\nOpen \"f\" For Binary {clause} As #1\nEnd Sub\n");
        let p = parse(&src);
        assert_eq!(p.found[0].0, SYNTAX_ERROR, "for {clause}");
        assert_eq!(p.found[0].1, offset_of(&src, clause.split(' ').next().unwrap(), 0), "for {clause}");
    }
    let src = "Sub S()\nOpen \"f\" For Binary As #1 [Len] = 10\nEnd Sub\n";
    let p = parse(src);
    assert_eq!(p.found[0].0, EXPECTED);
    assert_eq!(p.found[0].1, offset_of(src, "[Len]", 0));
}

#[test]
fn an_open_mode_that_is_no_mode_word_is_a_syntax_error() {
    let src = "Sub S()\nOpen \"f\" For [Foo] As #1\nEnd Sub\n";
    let p = parse(src);
    assert_eq!(p.found[0].0, SYNTAX_ERROR);
    assert_eq!(p.found[0].1, offset_of(src, "[Foo]", 0));
}
