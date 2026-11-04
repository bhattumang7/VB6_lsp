//! `Get`, `Put`, `Lock`, `Unlock` and `Seek` take their channel with or
//! without a `#`; `Write`, `Input`, `Print #` need the comma after it. A
//! reserved file keyword with no operands, or used as a plain name, is an error.

mod common;
use common::{offset_of, parse, Parsed};
use vb6_syntax::frontend::ast::{AstLit, ExprNode, FileIoKind, NodeId};

fn name_of(p: &Parsed, id: NodeId) -> String {
    let ExprNode::NameRef { sym, .. } = p.node(id) else { panic!("expected a name") };
    p.name(*sym).to_string()
}

fn only_stmt(src: &str) -> (Parsed, NodeId) {
    let p = parse(&format!("Sub S()\n    {src}\nEnd Sub\n"));
    let body = p.proc_body(0);
    assert_eq!(body.len(), 1, "{src}");
    let stmt = body[0];
    (p, stmt)
}

fn file_stmt(p: &Parsed, id: NodeId) -> (FileIoKind, NodeId, Vec<NodeId>) {
    let ExprNode::FileIoStmt { kind, channel, args, .. } = p.node(id) else { panic!("expected a file statement") };
    (*kind, channel.expect("a channel"), args.clone())
}

#[test]
fn get_and_put_without_a_hash() {
    for (src, kind) in [("Get F, offset, buf", FileIoKind::Get), ("Put F, offset, buf", FileIoKind::Put)] {
        let (p, stmt) = only_stmt(src);
        assert_eq!(p.found, vec![], "{src}");
        let (k, channel, args) = file_stmt(&p, stmt);
        assert_eq!(k, kind);
        assert_eq!(name_of(&p, channel), "F");
        assert_eq!(args.len(), 2);
        assert_eq!(name_of(&p, args[0]), "offset");
        assert_eq!(name_of(&p, args[1]), "buf");
    }
}

#[test]
fn get_with_a_hash_and_an_omitted_record() {
    let (p, stmt) = only_stmt("Get #1, , v");
    assert_eq!(p.found, vec![]);
    let (kind, channel, args) = file_stmt(&p, stmt);
    assert_eq!(kind, FileIoKind::Get);
    assert!(matches!(p.node(channel), ExprNode::Literal { lit: AstLit::Int(1) }));
    assert_eq!(args.len(), 1);
    assert_eq!(name_of(&p, args[0]), "v");
}

#[test]
fn get_with_an_arithmetic_record_and_a_following_statement() {
    let p = parse("Sub S()\n    Get gv, gv2 + f(p), l: x = 1\nEnd Sub\n");
    assert_eq!(p.found, vec![]);
    let body = p.proc_body(0);
    assert_eq!(body.len(), 2);
    let (kind, _, args) = file_stmt(&p, body[0]);
    assert_eq!(kind, FileIoKind::Get);
    assert_eq!(args.len(), 2);
    assert!(matches!(p.node(body[1]), ExprNode::Assign { .. }));
}

#[test]
fn lock_and_unlock_take_a_channel_in_any_form() {
    for src in ["Lock #1", "Lock 1, 2", "Unlock (1)", "Unlock .x = 1", "Unlock #1, 1 To 2"] {
        let (p, stmt) = only_stmt(src);
        assert_eq!(p.found, vec![], "{src}");
        let (kind, _, _) = file_stmt(&p, stmt);
        assert!(matches!(kind, FileIoKind::Lock | FileIoKind::Unlock), "{src}");
    }
}

#[test]
fn seek_takes_a_channel_and_a_position() {
    let (p, stmt) = only_stmt("Seek F, 5");
    assert_eq!(p.found, vec![]);
    let (kind, channel, args) = file_stmt(&p, stmt);
    assert_eq!(kind, FileIoKind::Seek);
    assert_eq!(name_of(&p, channel), "F");
    assert_eq!(args.len(), 1);
}

fn first_diag_offset(src: &str, at: &str) -> Vec<(u32, u32, u32)> {
    let full = format!("Sub S()\n    {src}\nEnd Sub\n");
    let _ = offset_of(&full, at, 0);
    parse(&full).found
}

#[test]
fn print_write_and_input_need_the_comma_after_the_channel() {
    for (src, kw) in [("Print #1", "Print"), ("Write #1", "Write"), ("Input #1", "Input")] {
        let found = first_diag_offset(src, kw);
        assert_eq!(found.len(), 1, "{src}");
        assert_eq!(found[0].0, 0x9c6f, "{src}");
    }
    for ok in ["Print #1,", "Write #1,", "Print #1, x", "Write #1, x", "Input #1, x"] {
        let (p, _) = only_stmt(ok);
        assert_eq!(p.found, vec![], "{ok}");
    }
}

#[test]
fn reserved_file_keywords_without_operands_are_syntax_errors() {
    for kw in ["Write", "Put", "Get", "Seek", "Lock", "Unlock", "Input"] {
        // Nothing that could be an operand follows: the keyword itself is reported.
        for tail in ["", ": x = 1", " = 1", ", 1"] {
            let src = format!("Sub S()\n    {kw}{tail}\nEnd Sub\n");
            let p = parse(&src);
            assert_eq!(p.found.len(), 1, "{kw}{tail}: {:?}", p.found);
            assert_eq!(p.found[0], (0x9c70, offset_of(&src, kw, 0), kw.len() as u32), "{kw}{tail}");
        }
        // A member access after the keyword is no name for it: rejected as well, here
        // or at the missing comma that follows an operand. (`Lock` and `Unlock` take
        // any expression as their channel, `.x = 1` included.)
        if !["Lock", "Unlock"].contains(&kw) {
            let src = format!("Sub S()\n    {kw}.x = 1\nEnd Sub\n");
            assert!(!parse(&src).found.is_empty(), "{kw}.x = 1");
        }
    }
}
