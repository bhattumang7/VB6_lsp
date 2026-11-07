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
    src.extend(std::iter::repeat_n(0xE9u8, 130));
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

#[test]
fn names_differing_in_the_case_of_an_accented_letter_are_one_symbol() {
    use vb6_syntax::frontend::ast::ExprArena;
    use vb6_syntax::frontend::parser::Parser;
    use vb6_syntax::frontend::scanner::ScannerContext;
    // 0xE9 and 0xC9 are the lower and upper case of the same letter.
    let src: Vec<u8> = b"Dim \xE9 As Long\n\xC9 = 1\n".to_vec();
    let mut ctx = ScannerContext::new(1, 1, 0x0409);
    ctx.intern_keywords();
    let mut arena = ExprArena::new();
    let mut parser = Parser::new(&mut ctx, &src);
    parser.parse_module(&mut arena);
    drop(parser);
    let a = ctx.lookup("\u{e9}").map(|s| s.name.clone());
    let b = ctx.lookup("\u{c9}").map(|s| s.name.clone());
    assert_eq!(a, b);
    assert!(a.is_some());
}

#[test]
fn string_bytes_are_read_in_the_ansi_code_page() {
    use vb6_syntax::frontend::ast::{AstLit, ExprArena, ExprNode};
    use vb6_syntax::frontend::parser::Parser;
    use vb6_syntax::frontend::scanner::ScannerContext;
    let mut src = b"Sub S()\n x = \"".to_vec();
    src.extend_from_slice(&[0x80, 0x81, 0x8E, 0x9F, 0xE9]);
    src.extend_from_slice(b"\"\nEnd Sub\n");
    let mut ctx = ScannerContext::new(1, 1, 0x0409);
    ctx.intern_keywords();
    let mut arena = ExprArena::new();
    let mut parser = Parser::new(&mut ctx, &src);
    let top = parser.parse_module(&mut arena);
    assert_eq!(parser.diagnostics.items().len(), 0);
    let ExprNode::ProcDecl { body, .. } = arena.get(top[0]) else { panic!("no procedure") };
    let ExprNode::Block { stmts } = arena.get(*body) else { panic!("no body") };
    let ExprNode::Assign { value, .. } = arena.get(stmts[0]) else { panic!("no assignment") };
    let ExprNode::Literal { lit: AstLit::Str(s) } = arena.get(*value) else { panic!("no string") };
    assert_eq!(s.chars().map(|c| c as u32).collect::<Vec<_>>(), vec![8364, 129, 381, 376, 233]);
}

#[test]
fn an_underscore_after_a_bang_does_not_start_a_member() {
    // `r!_x`: the `!` is the Single suffix of `r` and `_x` is not a name.
    let (_, found) = in_sub("n = r!_x");
    assert!(!found.is_empty());
    let (_, found) = in_sub("n = r!x");
    assert_eq!(found, vec![]);
}

#[test]
fn a_double_byte_character_is_one_unit_of_a_name_in_a_double_byte_code_page() {
    use vb6_syntax::frontend::ast::ExprArena;
    use vb6_syntax::frontend::parser::Parser;
    use vb6_syntax::frontend::scanner::ScannerContext;
    // Shift-JIS: 0x83 0x5C is one katakana letter whose trail byte is a backslash.
    let src: Vec<u8> = b"Dim a\x83\x5Cb As Long\n".to_vec();
    for (lcid, expect_errors) in [(0x0411u32, false), (0x0409u32, true)] {
        let mut ctx = ScannerContext::new(1, 1, lcid);
        ctx.intern_keywords();
        let mut arena = ExprArena::new();
        let mut parser = Parser::new(&mut ctx, &src);
        parser.parse_module(&mut arena);
        assert_eq!(!parser.diagnostics.items().is_empty(), expect_errors, "{lcid:#x}");
    }
}

fn parse_bytes_in(lcid: u32, src: &[u8]) -> (Vec<(u32, u32, u32)>, String) {
    use vb6_syntax::frontend::ast::{AstLit, BinOpKind, ExprArena, ExprNode};
    use vb6_syntax::frontend::parser::Parser;
    use vb6_syntax::frontend::scanner::ScannerContext;
    let mut ctx = ScannerContext::new(1, 1, lcid);
    ctx.intern_keywords();
    let mut arena = ExprArena::new();
    let mut parser = Parser::new(&mut ctx, src);
    let top = parser.parse_module(&mut arena);
    let found = parser.diagnostics.items().iter().map(|d| (d.code, d.span.start, d.span.len)).collect();
    drop(parser);
    let mut shape = String::new();
    if let Some(ExprNode::ProcDecl { body, .. }) = top.first().map(|n| arena.get(*n)) {
        if let ExprNode::Block { stmts } = arena.get(*body) {
            if let Some(ExprNode::Assign { target, value }) = stmts.first().map(|s| arena.get(*s)) {
                let name = |id| match arena.get(id) {
                    ExprNode::NameRef { sym, .. } => ctx.sym_name(*sym as usize).to_string(),
                    ExprNode::Literal { lit: AstLit::Int(n) } => n.to_string(),
                    ExprNode::Literal { lit: AstLit::Str(s) } => format!("{s:?}"),
                    _ => "?".to_string(),
                };
                shape = match arena.get(*value) {
                    ExprNode::BinOp { op: BinOpKind::Add, lhs, rhs } => format!("{} = {} + {}", name(*target), name(*lhs), name(*rhs)),
                    _ => format!("{} = {}", name(*target), name(*value)),
                };
            }
        }
    }
    (found, shape)
}

#[test]
fn full_width_forms_are_read_as_the_ascii_characters_in_a_double_byte_code_page() {
    // Shift-JIS: x, full-width = 1 + 2 (digits 0x8250.., plus 0x817B, equals 0x8181).
    let mut src = b"Sub S()\nx".to_vec();
    src.extend_from_slice(&[0x81, 0x81, 0x82, 0x50, 0x81, 0x7B, 0x82, 0x51]);
    src.extend_from_slice(b"\nEnd Sub\n");
    assert_eq!(parse_bytes_in(0x0411, &src), (vec![], "x = 1 + 2".to_string()));
    assert!(!parse_bytes_in(0x0409, &src).0.is_empty());
}

#[test]
fn full_width_letters_and_quotes_in_a_double_byte_code_page() {
    // Full-width A (0x8260) and a (0x8281) in a name are the ASCII letters; 0x8167 ... 0x8168
    // are the quotes of a string.
    let mut src = b"Sub S()\n".to_vec();
    src.extend_from_slice(&[0x82, 0x60, 0x82, 0x81]);
    src.extend_from_slice(b" = ");
    src.extend_from_slice(&[0x81, 0x67]);
    src.extend_from_slice(b"abc");
    src.extend_from_slice(&[0x81, 0x68]);
    src.extend_from_slice(b"\nEnd Sub\n");
    assert_eq!(parse_bytes_in(0x0411, &src), (vec![], "Aa = \"abc\"".to_string()));
    // Korean (949) and Chinese (936, 950) spell the digits too.
    let digit_one: [(u32, [u8; 2]); 3] = [(0x0412, [0xa3, 0xb1]), (0x0804, [0xa3, 0xb1]), (0x0404, [0xa2, 0xb0])];
    for (lcid, pair) in digit_one {
        let mut src = b"Sub S()\nx = ".to_vec();
        src.extend_from_slice(&pair);
        src.extend_from_slice(b"\nEnd Sub\n");
        assert_eq!(parse_bytes_in(lcid, &src), (vec![], "x = 1".to_string()), "{lcid:#x}");
    }
}
