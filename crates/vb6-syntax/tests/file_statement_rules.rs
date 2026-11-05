//! File statements: `Open` access against mode, `Len =`, `Lock`/`Unlock`
//! ranges, `Seek`/`Lock` without `#`, and a colon between module-level
//! declarations.

mod common;
use common::{offset_of, parse};
use vb6_syntax::frontend::ast::{AstLit, ExprNode, FileIoKind};

const BAD_FILE_MODE: u32 = 54;
const EXPECTED: u32 = 0x9c6f;

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
        assert_eq!(found, vec![(BAD_FILE_MODE, offset_of(&src, "Access", 0), 6)], "{stmt}");
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
    assert!(matches!(p.node(args[0]), ExprNode::Literal { lit: AstLit::Int(0) }));
    assert!(matches!(p.node(args[1]), ExprNode::Literal { lit: AstLit::Int(10) }));
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
