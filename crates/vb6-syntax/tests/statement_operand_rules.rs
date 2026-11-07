//! Operand rules of individual statements: `Option`, `DefType`, `Dim`,
//! `Get`/`Put`, `End`, array bounds.

mod common;
use common::{offset_of, parse};
use vb6_syntax::frontend::ast::{AstLit, ExprNode};

const EXPECTED: u32 = 0x9c6f;

fn module(src: &str) -> Vec<(u32, u32, u32)> {
    parse(src).found
}

fn in_sub(body: &str) -> (String, Vec<(u32, u32, u32)>) {
    let src = format!("Sub S()\n{body}\nEnd Sub\n");
    let found = parse(&src).found;
    (src, found)
}

#[test]
fn option_base_takes_the_literal_zero_or_one() {
    for ok in ["Option Base 0\n", "Option Base 1\n"] {
        assert_eq!(module(ok), vec![], "{ok:?}");
    }
    for (src, bad) in [
        ("Option Base 2\n", "2"),
        ("Option Base x\n", "x"),
        ("Option Base 1.5\n", "1.5"),
    ] {
        assert_eq!(module(src), vec![(EXPECTED, offset_of(src, bad, 0), bad.len() as u32)], "{src:?}");
    }
    // Missing operand: reported at the end of the line.
    let src = "Option Base\n";
    assert_eq!(module(src), vec![(EXPECTED, 11, 0)]);
}

#[test]
fn option_base_value_is_kept() {
    let p = parse("Option Base 1\n");
    assert!(matches!(p.node(p.top[0]), ExprNode::OptionBase { value: 1 }));
}

#[test]
fn option_compare_takes_binary_or_text_only() {
    for (src, mode) in [("Option Compare Binary\n", 0u8), ("Option Compare Text\n", 1)] {
        let p = parse(src);
        assert_eq!(p.found, vec![], "{src:?}");
        assert!(matches!(p.node(p.top[0]), ExprNode::OptionCompare { mode: m } if *m == mode));
    }
    for (src, len) in [("Option Compare Foo\n", 3), ("Option Compare Database\n", 8)] {
        assert_eq!(module(src), vec![(EXPECTED, 15, len)], "{src:?}");
    }
}

#[test]
fn unknown_option_is_reported_at_its_operand() {
    for (src, at, len) in [("Option Strict On\n", 7, 6), ("Option Foo\n", 7, 3)] {
        assert_eq!(module(src), vec![(EXPECTED, at, len)], "{src:?}");
    }
}

#[test]
fn option_private_takes_module() {
    assert_eq!(module("Option Private Module\n"), vec![]);
    let src = "Option Private Foo\n";
    assert_eq!(module(src), vec![(EXPECTED, 15, 3)]);
}

#[test]
fn deftype_letters_are_single_characters() {
    assert_eq!(module("DefInt A-C, X\n"), vec![]);
    let src = "DefInt AB\n";
    assert_eq!(module(src), vec![(EXPECTED, 7, 2)]);
    let src = "DefInt A-BC\n";
    assert_eq!(module(src), vec![(EXPECTED, 9, 2)]);
}

#[test]
fn dim_has_no_initial_value() {
    for body in ["Dim x As Long = 5", "Dim x = 5"] {
        let (src, found) = in_sub(body);
        let at = offset_of(&src, "= 5", 0);
        assert_eq!(found, vec![(EXPECTED, at, 1)], "{body}");
    }
    let (_, found) = in_sub("Const c As Long = 5");
    assert_eq!(found, vec![]);
}

#[test]
fn array_bounds_have_no_trailing_comma() {
    let (src, found) = in_sub("Dim x(1 To 2,) As Long");
    let at = offset_of(&src, ")", 1);
    assert_eq!(found[0], (EXPECTED, at, 1));
}

#[test]
fn get_and_put_need_both_commas() {
    for stmt in ["Get", "Put"] {
        let (_, found) = in_sub(&format!("{stmt} #1, , x"));
        assert_eq!(found, vec![], "{stmt}");
        let (_, found) = in_sub(&format!("{stmt} #1, 5, x"));
        assert_eq!(found, vec![], "{stmt}");
        let (src, found) = in_sub(&format!("{stmt} #1, x"));
        let eol = offset_of(&src, "x\n", 0) + 1;
        assert_eq!(found, vec![(EXPECTED, eol, 0)], "{stmt} one comma");
        let (src, found) = in_sub(&format!("{stmt} #1"));
        let eol = offset_of(&src, "1\n", 0) + 1;
        assert_eq!(found, vec![(EXPECTED, eol, 0)], "{stmt} no comma");
    }
}

#[test]
fn end_takes_no_operand() {
    let (src, found) = in_sub("End x");
    assert_eq!(found, vec![(EXPECTED, offset_of(&src, "x\n", 0), 1)]);
    let (src, found) = in_sub("End 5");
    assert_eq!(found, vec![(EXPECTED, offset_of(&src, "5", 0), 1)]);
    let p = parse("Sub S()\n End\nEnd Sub\n");
    assert_eq!(p.found, vec![]);
    let body = p.proc_body(0);
    assert_eq!(body.len(), 1);
    assert!(matches!(p.node(body[0]), ExprNode::EndStmt));
    let _ = AstLit::Empty;
}

#[test]
fn const_declaration_requires_a_value() {
    let (src, found) = in_sub("Const c As Long");
    let eol = offset_of(&src, "Long\n", 0) + 4;
    assert_eq!(found[0], (EXPECTED, eol, 0));
    let (src, found) = in_sub("Const a = 1, b");
    let eol = offset_of(&src, "b\n", 0) + 1;
    assert_eq!(found[0], (EXPECTED, eol, 0));
    let (_, found) = in_sub("Const a = 1, b = 2");
    assert_eq!(found, vec![]);
}

#[test]
fn enum_members_keep_an_optional_value() {
    assert_eq!(module("Enum E\n A\n B = 2\nEnd Enum\n"), vec![]);
}

#[test]
fn assignment_and_loop_targets_name_something() {
    for body in ["For Each 5 In c\nNext", "Erase 5", "Erase \"a\"", "Set 5 = x"] {
        let (src, found) = in_sub(body);
        let at = body.find(['5', '"']).unwrap() as u32 + 8;
        assert_eq!(found[0].0, 0x9c6f, "{body}: {src:?}");
        assert_eq!(found[0].1, at, "{body}");
    }
}

#[test]
fn mid_statement_takes_two_or_three_arguments() {
    let (_, found) = in_sub("Mid(s, 1) = \"a\"");
    assert_eq!(found, vec![]);
    let (_, found) = in_sub("Mid$(s, 1, 2) = \"a\"");
    assert_eq!(found, vec![]);
    let (src, found) = in_sub("Mid(s) = \"a\"");
    assert_eq!(found, vec![(EXPECTED, offset_of(&src, ")", 1), 1)]);
    // A fourth operand is rejected where the closing bracket was due.
    let (src, found) = in_sub("Mid(s, 1, 2, 3) = \"a\"");
    assert_eq!(found, vec![(EXPECTED, offset_of(&src, ", 3", 0), 1)]);
}
