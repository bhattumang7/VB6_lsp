//! Control characters: a NUL ends the text of its line, 0x19 is a blank, every other
//! control character is invalid; a comment runs to the end of its line whatever it holds
//! and continues only when its last word is a blank followed by `_`.

mod common;
use common::{offset_of, parse};

fn in_sub(body: &str) -> (String, Vec<(u32, u32, u32)>) {
    let src = format!("Sub S()\n{body}\nEnd Sub\n");
    let found = parse(&src).found;
    (src, found)
}

#[test]
fn a_control_character_after_a_statement_is_invalid() {
    for c in ['\u{1}', '\u{8}', '\u{b}', '\u{c}', '\u{14}', '\u{1a}', '\u{1d}', '\u{7f}'] {
        let body = format!("n = 1{c}");
        let (src, found) = in_sub(&body);
        assert!(!found.is_empty(), "{:#x}", c as u32);
        assert_eq!(found[0].1, offset_of(&src, &c.to_string(), 0), "{:#x}", c as u32);
    }
}

#[test]
fn nineteen_is_a_blank_and_a_nul_ends_the_line_text() {
    for body in ["n = 1\u{19}+\u{19}2", "x = n \u{19}< \u{19}= 1", "x = 1\u{0} + 2", "x = 1\u{0}@@@ ("] {
        let (_, found) = in_sub(body);
        assert_eq!(found, vec![], "{body:?}");
    }
}

#[test]
fn a_comment_holds_any_control_character() {
    for body in ["n = 1 ' a\u{c}@@@ (", "Rem a\u{1}@@@ (", "' a\u{b}@@@ (", "' a\u{0}@@@ ("] {
        let (_, found) = in_sub(body);
        assert_eq!(found, vec![], "{body:?}");
    }
}

#[test]
fn a_comment_continues_only_after_a_blank_and_an_underscore() {
    for body in ["' x(_\n@@@", "'_\n@@@", "' x -_\n@@@", "Rem x(_\n@@@", "' x_\n@@@", "' x __\n@@@"] {
        let (src, found) = in_sub(body);
        assert!(!found.is_empty(), "{body:?}");
        assert_eq!(found[0].1, offset_of(&src, "@@@", 0), "{body:?}");
    }
    for body in ["' _\n@@@", "' x _  \n@@@", "' x\t_\n@@@", "' x _\t\n@@@", "Rem _\n@@@"] {
        let (_, found) = in_sub(body);
        assert_eq!(found, vec![], "{body:?}");
    }
}

#[test]
fn brackets_inside_a_bracketed_name_nest() {
    use vb6_syntax::frontend::ast::ExprNode;
    let p = parse("Sub S()\n o.[a[b]c] = 1\nEnd Sub\n");
    assert_eq!(p.found, vec![]);
    let body = p.proc_body(0);
    let ExprNode::Assign { target, .. } = p.node(body[0]) else { panic!("not an assignment") };
    let ExprNode::MemberAccess { member, .. } = p.node(*target) else { panic!("not a member") };
    assert_eq!(p.name(*member), "a[b]c");
    let (_, found) = in_sub("o.[a[b] = 1");
    assert!(!found.is_empty());
}

#[test]
fn identifier_length_counts_source_bytes_not_utf8_bytes() {
    use vb6_syntax::frontend::ast::ExprArena;
    use vb6_syntax::frontend::parser::Parser;
    use vb6_syntax::frontend::scanner::ScannerContext;
    // 130 accented letters are 130 bytes of source text.
    let mut src = b"Dim ".to_vec();
    src.extend(std::iter::repeat(0xE9u8).take(130));
    src.extend_from_slice(b" As Long
");
    let mut ctx = ScannerContext::new(1, 1, 0x0409);
    ctx.intern_keywords();
    let mut arena = ExprArena::new();
    let mut parser = Parser::new(&mut ctx, &src);
    parser.parse_module(&mut arena);
    assert_eq!(parser.diagnostics.items().len(), 0);
    let long: String = "a".repeat(256);
    assert_eq!(parse(&format!("Dim {long} As Long
")).found.first().map(|f| f.0), Some(0x9c6d));
}
