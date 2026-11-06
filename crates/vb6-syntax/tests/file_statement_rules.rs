//! File statements: `Open` access against mode, `Len =`, `Lock`/`Unlock`
//! ranges, `Seek`/`Lock` without `#`, and a colon between module-level
//! declarations.

mod common;
use common::{offset_of, parse};
use vb6_syntax::frontend::ast::{AstLit, ExprNode, FileIoKind};

const SYNTAX_ERROR: u32 = 0x9c70;
const EXPECTED: u32 = 0x9c6f;
const UNEXPECTED_TOKEN: u32 = 40047;

fn in_sub(body: &str) -> (String, Vec<(u32, u32, u32)>) {
    let src = format!("Sub S()\n{body}\nEnd Sub\n");
    let found = parse(&src).found;
    (src, found)
}

#[test]
fn open_access_must_suit_the_mode() {
    for stmt in [
        "Open \"a\" For Input Access Write As #1",
        "Open \"a\" For Append Access Read As #1",
        "Open \"a\" For Output Access Read As #1",
        "Open \"a\" For Input Access Read Write As #1",
        "Open \"a\" For Output Access Read Write As #1",
    ] {
        let (src, found) = in_sub(stmt);
        assert_eq!(found, vec![(SYNTAX_ERROR, offset_of(&src, "Access", 0), 6)], "{stmt}");
    }
    for stmt in [
        "Open \"a\" For Input Access Read As #1",
        "Open \"a\" For Output Access Write As #1",
        "Open \"a\" For Append Access Write As #1",
        "Open \"a\" For Append Access Read Write As #1",
        "Open \"a\" For Random Access Read Write As #1",
        "Open \"a\" For Binary Access Read As #1",
        "Open \"a\" Access Read Write As #1",
    ] {
        assert_eq!(in_sub(stmt).1, vec![], "{stmt}");
    }
}

#[test]
fn open_access_and_lock_name_what_they_restrict() {
    // `Access` and `Lock` with neither `Read` nor `Write` after them: a syntax error at the token
    // that stands in their place.
    for (stmt, at) in [
        ("Open \"a\" For Input Access As #1", "As"),
        ("Open \"a\" For Binary Lock As #1", "As"),
        ("Open \"a\" For Input Access Foo As #1", "Foo"),
        ("Open \"a\" For Binary Lock Foo As #1", "Foo"),
        ("Open \"a\" For Binary Access", ""),
        ("Open \"a\" For Binary Lock", ""),
    ] {
        let (src, found) = in_sub(stmt);
        let offset = if at.is_empty() { src.find("\nEnd Sub").unwrap() as u32 } else { offset_of(&src, at, 0) };
        // At the end of the line the error has no text of its own.
        assert_eq!(found, vec![(SYNTAX_ERROR, offset, at.len() as u32)], "{stmt}");
    }
    for stmt in [
        "Open \"a\" For Binary Lock Read As #1",
        "Open \"a\" For Binary Lock Write As #1",
        "Open \"a\" For Binary Lock Read Write As #1",
        "Open \"a\" For Binary Access Read Lock Write As #1",
    ] {
        assert_eq!(in_sub(stmt).1, vec![], "{stmt}");
    }
}

#[test]
fn open_len_needs_an_equals_sign() {
    assert_eq!(in_sub("Open \"a\" For Random As #1 Len = 5").1, vec![]);
    let (src, found) = in_sub("Open \"a\" For Random As #1 Len 5");
    assert_eq!(found[0], (EXPECTED, offset_of(&src, "5", 0), 1));
}

#[test]
fn lock_range_may_start_with_to() {
    let p = parse("Sub S()\n Lock #1, To 10\nEnd Sub\n");
    assert_eq!(p.found, vec![]);
    let body = p.proc_body(0);
    let ExprNode::FileIoStmt { kind: FileIoKind::Lock, args, .. } = p.node(body[0]) else {
        panic!("not a Lock")
    };
    assert_eq!(args.len(), 2);
    // The first record is left out, not a record number of its own.
    assert!(matches!(p.node(args[0]), ExprNode::MissingArg));
    assert!(matches!(p.node(args[1]), ExprNode::Literal { lit: AstLit::Int(10) }));
}

#[test]
fn open_mode_access_and_the_first_lock_word_may_be_written_in_brackets() {
    let cases: &[(&str, u8, u8)] = &[
        ("Open \"a\" For [Input] As #1", 1, 0),
        ("Open \"a\" For [input] As #1", 1, 0),
        ("Open \"a\" For [Output] As #1", 2, 0),
        ("Open \"a\" For [Random] As #1", 4, 0),
        ("Open \"a\" For [Append] As #1", 8, 0),
        ("Open \"a\" For [Binary] As #1", 0x20, 0),
        ("Open \"a\" For Binary [Access] Read As #1", 0x20, 1),
        ("Open \"a\" For Binary Access [Read] As #1", 0x20, 1),
        ("Open \"a\" For Binary Access [Write] As #1", 0x20, 2),
        ("Open \"a\" For Binary Access [Read] Write As #1", 0x20, 3),
        ("Open \"a\" For [Input] Access Read As #1", 1, 1),
        ("Open \"a\" For Binary Lock [Read] As #1", 0x20, 0x30),
        ("Open \"a\" For Binary Lock [Write] As #1", 0x20, 0x20),
        ("Open \"a\" For Binary Lock [Read] Write As #1", 0x20, 0x10),
    ];
    for (stmt, mode, flags) in cases {
        let p = parse(&format!("Sub S()\n {stmt}\nEnd Sub\n"));
        assert_eq!(p.found, vec![], "{stmt}");
        let body = p.proc_body(0);
        let ExprNode::FileIoStmt { kind: FileIoKind::Open, open_mode, open_flags, .. } = p.node(body[0]) else {
            panic!("not an Open: {stmt}")
        };
        assert_eq!((*open_mode, *open_flags), (*mode, *flags), "{stmt}");
    }
    // The second of `Read Write` is the plain word, and the other words of the statement
    // are not words that can be bracketed.
    for stmt in [
        "Open \"a\" For Binary Access Read [Write] As #1",
        "Open \"a\" For Binary Access [Read] [Write] As #1",
        "Open \"a\" For Binary Lock Read [Write] As #1",
        "Open \"a\" [For] Input As #1",
        "Open \"a\" For Input [As] #1",
        "Open \"a\" For Binary [Lock] Read As #1",
        "Open \"a\" For Binary [Shared] As #1",
        "Open \"a\" For Output Access [Read] As #1",
    ] {
        let (_, found) = in_sub(stmt);
        assert_eq!(found.first().map(|d| d.0), Some(SYNTAX_ERROR), "{stmt}: {found:?}");
    }
    // `[Len]` is a name after the channel, where the statement should end.
    let (src, found) = in_sub("Open \"a\" For Random As #1 [Len] = 5");
    assert_eq!(found, vec![(UNEXPECTED_TOKEN, offset_of(&src, "[Len]", 0), 5)]);
}

/// The `#` marks of the channel operands, statement by statement.
#[test]
fn channel_marks_follow_the_written_hash() {
    let cases: &[(&str, FileIoKind, &[bool])] = &[
        ("Print #1, 2", FileIoKind::Print, &[true]),
        ("Write #1, 2", FileIoKind::Write, &[true]),
        ("Input #1, a", FileIoKind::Input, &[true]),
        ("Line Input #1, a", FileIoKind::LineInput, &[true]),
        ("Get #1, , a", FileIoKind::Get, &[true]),
        ("Get 1, , a", FileIoKind::Get, &[false]),
        ("Put #1, , a", FileIoKind::Put, &[true]),
        ("Put 1, , a", FileIoKind::Put, &[false]),
        ("Seek #1, 2", FileIoKind::Seek, &[true]),
        ("Seek 1, 2", FileIoKind::Seek, &[false]),
        ("Width #1, 2", FileIoKind::Width, &[true]),
        ("Lock #1", FileIoKind::Lock, &[true]),
        ("Lock 1, 2 To 3", FileIoKind::Lock, &[false]),
        ("Unlock #1, 2", FileIoKind::Unlock, &[true]),
        ("Open \"a\" For Input As #1", FileIoKind::Open, &[true]),
        ("Open \"a\" For Input As 1", FileIoKind::Open, &[false]),
        ("Close", FileIoKind::Close, &[]),
        ("Close #1", FileIoKind::Close, &[true]),
        ("Close 1", FileIoKind::Close, &[false]),
        ("Close #1, b, #3", FileIoKind::Close, &[true, false, true]),
    ];
    for (stmt, want_kind, want_marks) in cases {
        let p = parse(&format!("Sub S()\n {stmt}\nEnd Sub\n"));
        assert_eq!(p.found, vec![], "{stmt}");
        let body = p.proc_body(0);
        let ExprNode::FileIoStmt { kind, channel_marks, .. } = p.node(body[0]) else { panic!("{stmt}") };
        assert_eq!((*kind, channel_marks.as_slice()), (*want_kind, *want_marks), "{stmt}");
    }
}

#[test]
fn lock_unlock_and_seek_do_not_need_a_hash() {
    for (stmt, kind) in [
        ("Lock 1, 5 To 10", FileIoKind::Lock),
        ("Unlock 1, 5", FileIoKind::Unlock),
        ("Seek 1, 10", FileIoKind::Seek),
    ] {
        let p = parse(&format!("Sub S()\n {stmt}\nEnd Sub\n"));
        assert_eq!(p.found, vec![], "{stmt}");
        let body = p.proc_body(0);
        let ExprNode::FileIoStmt { kind: k, channel, .. } = p.node(body[0]) else {
            panic!("{stmt}: not a file statement")
        };
        assert_eq!(*k, kind, "{stmt}");
        assert!(matches!(p.node(channel.expect("channel")), ExprNode::Literal { lit: AstLit::Int(1) }), "{stmt}");
    }
}

#[test]
fn seek_stays_usable_as_a_name() {
    let p = parse("Sub S()\n x = Seek(1)\n Seek #1, 5\nEnd Sub\n");
    assert_eq!(p.found, vec![]);
}

#[test]
fn a_colon_may_separate_module_level_declarations() {
    let p = parse("Dim a As Long: Dim b As Long\nConst c = 1: Const d = 2\n");
    assert_eq!(p.found, vec![]);
    assert_eq!(p.top.len(), 4);
}

#[test]
fn open_for_needs_a_mode_word() {
    let (src, found) = in_sub("Open \"a\" For As #1");
    assert_eq!(found[0], (0x9c70, offset_of(&src, "As", 0), 2));
    assert_eq!(in_sub("Open \"a\" For Binary As #1").1, vec![]);
}

#[test]
fn open_needs_as_before_the_channel() {
    let (src, found) = in_sub("Open \"a\" 1");
    assert_eq!(found[0], (0x9c70, offset_of(&src, "1", 0), 1));
    let (src, found) = in_sub("Open \"a\" For Input #1");
    assert_eq!(found[0], (0x9c70, offset_of(&src, "#", 0), 1));
    assert_eq!(in_sub("Open \"a\" For Input As #1").1, vec![]);
}
