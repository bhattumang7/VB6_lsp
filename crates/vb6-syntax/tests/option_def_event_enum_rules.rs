//! Option Compare modes, Def statements, RaiseEvent names, Implements inside a
//! procedure, and the physical-line rules of Enum members.

mod common;
use common::{offset_of, parse};

const EXPECTED: u32 = 0x9c6f;
const INVALID_IN_ENUM: u32 = 0xdeff;

#[test]
fn option_compare_takes_text_or_binary_only() {
    for mode in ["Text", "Binary"] {
        assert_eq!(parse(&format!("Option Compare {mode}\n")).found, vec![], "{mode}");
    }
    for bad in ["Database", "Foo", ""] {
        let src = format!("Option Compare {bad}\n");
        let found = parse(&src).found;
        assert_eq!(found[0].0, EXPECTED, "{bad}");
        let at = if bad.is_empty() { src.find(char::from(10)).unwrap() as u32 } else { offset_of(&src, bad, 0) };
        assert_eq!(found[0].1, at, "{bad}");
    }
}

#[test]
fn def_dec_is_not_a_statement() {
    assert!(!parse("DefDec A\n").found.is_empty());
    assert_eq!(parse("DefInt A\nDefCur C\nDefDbl D\n").found, vec![]);
}

#[test]
fn raise_event_takes_a_name_not_a_reserved_word() {
    use vb6_syntax::frontend::ast::ExprArena;
    use vb6_syntax::frontend::parser::{ModuleKind, Parser};
    use vb6_syntax::frontend::scanner::ScannerContext;
    let class_found = |src: &str| {
        let mut ctx = ScannerContext::new(1, 1, 0x0409);
        ctx.intern_keywords();
        let mut arena = ExprArena::new();
        let mut parser = Parser::with_module_kind(&mut ctx, src.as_bytes(), ModuleKind::Class);
        parser.parse_module(&mut arena);
        parser.diagnostics.items().iter().map(|d| (d.code, d.span.start)).collect::<Vec<_>>()
    };
    for word in ["Event", "Implements", "Dim"] {
        let src = format!("Sub S()
    RaiseEvent {word}
End Sub
");
        let at = offset_of(&src, &format!(" {word}"), 0) + 1;
        assert!(class_found(&src).iter().any(|d| d.1 == at), "{word}");
    }
    assert_eq!(class_found("Sub S()
    RaiseEvent E(1)
End Sub
"), vec![]);
}

#[test]
fn implements_is_read_inside_a_procedure() {
    let src = "Sub S()\n    Implements Foo\nEnd Sub\n";
    let found = parse(src).found;
    // Only the standard-module rule applies: the class-only statement.
    assert_eq!(found, vec![(57057, offset_of(src, "Implements", 0), 10)]);
}

#[test]
fn an_enum_member_with_a_type_is_invalid_inside_enum() {
    let src = "Enum E\n    A As Long\nEnd Enum\n";
    assert_eq!(parse(src).found[0], (INVALID_IN_ENUM, offset_of(src, "A", 0), 1));
}

#[test]
fn a_bare_enum_member_cannot_be_followed_by_a_colon_but_a_valued_one_can() {
    let src = "Enum E\n    A: B\nEnd Enum\n";
    assert_eq!(parse(src).found[0], (INVALID_IN_ENUM, offset_of(src, "A", 0), 1));
    for text in [
        "Enum E\n    A = 1: B = 2\nEnd Enum\n",
        "Enum E\n    A = 1: B\nEnd Enum\n",
        "Enum E\n    A = 1:\n    B\nEnd Enum\n",
        "Enum E\n    :A\nEnd Enum\n",
    ] {
        assert_eq!(parse(text).found, vec![], "{text}");
    }
}

#[test]
fn an_enum_member_stays_on_one_physical_line() {
    let src = "Enum E\n    A = 1 _\n        + 2\nEnd Enum\n";
    assert_eq!(parse(src).found[0], (INVALID_IN_ENUM, offset_of(src, "A", 0), 1));
}

#[test]
fn an_enum_member_may_be_bracketed() {
    assert_eq!(parse("Enum E\n    [A b]\n    [C d] = 2\nEnd Enum\n").found, vec![]);
}
