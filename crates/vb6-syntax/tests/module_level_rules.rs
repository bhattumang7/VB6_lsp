//! Rules the compiler applies to a module as a whole: one `Option` of each
//! kind, `DefType` statements before declarations and without overlapping
//! letters, `Option Base` before arrays, `Option Private Module` only in a
//! standard module.

mod common;
use common::{offset_of, parse};
use vb6_syntax::frontend::ast::ExprNode;
use vb6_syntax::frontend::parser::{ModuleKind, Parser};
use vb6_syntax::frontend::ast::ExprArena;
use vb6_syntax::frontend::scanner::ScannerContext;

const ARRAY_ALREADY_DIMENSIONED: u32 = 0x9c77;
const DUPLICATE_OPTION: u32 = 0x9c78;
const DEFTYPE_AFTER_DECLARATION: u32 = 0x9c79;
const DUPLICATE_DEFTYPE: u32 = 0x9c7a;
const OPTION_PRIVATE_IN_OBJECT_MODULE: u32 = 0xdec9;

fn found_in(kind: ModuleKind, src: &str) -> Vec<(u32, u32, u32)> {
    let mut ctx = ScannerContext::new(1, 1, 0x0409);
    ctx.intern_keywords();
    let mut arena = ExprArena::new();
    let mut parser = Parser::with_module_kind(&mut ctx, src.as_bytes(), kind);
    parser.parse_module(&mut arena);
    parser
        .diagnostics
        .items()
        .iter()
        .map(|d| (d.code, d.span.start, d.span.len))
        .collect()
}

#[test]
fn a_second_option_of_the_same_kind_is_reported_at_the_option_keyword() {
    for (first, second) in [
        ("Option Explicit", "Option Explicit"),
        ("Option Base 0", "Option Base 1"),
        ("Option Compare Text", "Option Compare Binary"),
        ("Option Private Module", "Option Private Module"),
    ] {
        let src = format!("{first}\n{second}\n");
        let p = parse(&src);
        assert_eq!(p.found, vec![(DUPLICATE_OPTION, offset_of(&src, "Option", 1), 6)], "for {src:?}");
        assert_eq!(p.top.len(), 2, "both statements are kept for {src:?}");
    }
}

#[test]
fn different_option_kinds_do_not_conflict() {
    let p = parse("Option Explicit\nOption Base 1\nOption Compare Text\nOption Private Module\n");
    assert_eq!(p.found, vec![]);
    assert!(matches!(p.node(p.top[0]), ExprNode::OptionExplicit));
    assert!(matches!(p.node(p.top[1]), ExprNode::OptionBase { value: 1 }));
    assert!(matches!(p.node(p.top[2]), ExprNode::OptionCompare { mode: 1 }));
}

#[test]
fn option_base_after_a_module_level_array_is_reported() {
    let src = "Dim a(5) As Long\nOption Base 1\n";
    let p = parse(src);
    assert_eq!(p.found, vec![(ARRAY_ALREADY_DIMENSIONED, offset_of(src, "Option", 0), 6)]);
    let src = "Dim a() As Long\nOption Base 1\n";
    assert_eq!(parse(src).found, vec![(ARRAY_ALREADY_DIMENSIONED, offset_of(src, "Option", 0), 6)]);
    // A scalar, a constant, and a procedure-local array do not count.
    for ok in [
        "Dim n As Long\nOption Base 1\n",
        "Const n = 5\nOption Base 1\n",
        "Option Base 1\nSub S()\nDim a(5) As Long\nEnd Sub\n",
    ] {
        assert_eq!(parse(ok).found, vec![], "for {ok:?}");
    }
}

#[test]
fn option_private_module_is_for_standard_modules_only() {
    let src = "Option Private Module\n";
    assert_eq!(found_in(ModuleKind::Standard, src), vec![]);
    for kind in [ModuleKind::Class, ModuleKind::Form] {
        assert_eq!(
            found_in(kind, src),
            vec![(OPTION_PRIVATE_IN_OBJECT_MODULE, 0, 6)]
        );
    }
}

#[test]
fn deftype_after_a_declaration_is_reported_at_the_deftype_keyword() {
    for decl in [
        "Dim n As Long",
        "Public n As Long",
        "Private Declare Function F Lib \"k\" () As Long",
        "Const n = 1",
        "Enum E\n    A\nEnd Enum",
        "Type T\n    a As Long\nEnd Type",
    ] {
        let src = format!("{decl}\nDefInt A-Z\n");
        let p = parse(&src);
        assert_eq!(
            p.found,
            vec![(DEFTYPE_AFTER_DECLARATION, offset_of(&src, "DefInt", 0), 6)],
            "for {decl:?}"
        );
    }
}

#[test]
fn deftype_before_declarations_and_after_options_is_fine() {
    let p = parse("Option Explicit\nDefInt A-C\nDefStr S\nDefLng X-Z\nDim n As Long\n");
    assert_eq!(p.found, vec![]);
    assert!(matches!(p.node(p.top[1]), ExprNode::DefType { .. }));
}

#[test]
fn deftype_letters_may_not_overlap_an_earlier_deftype() {
    let src = "DefInt A-Z\nDefStr S\n";
    assert_eq!(parse(src).found, vec![(DUPLICATE_DEFTYPE, offset_of(src, "DefStr", 0), 6)]);
    let src = "DefInt A-C\nDefStr C-D\n";
    assert_eq!(parse(src).found, vec![(DUPLICATE_DEFTYPE, offset_of(src, "DefStr", 0), 6)]);
    let src = "DefInt a\nDefStr A\n";
    assert_eq!(parse(src).found, vec![(DUPLICATE_DEFTYPE, offset_of(src, "DefStr", 0), 6)]);
    let src = "DefInt S\nDefStr A-Z\n";
    assert_eq!(parse(src).found, vec![(DUPLICATE_DEFTYPE, offset_of(src, "DefStr", 0), 6)]);
    let src = "DefInt A-C, E, G-I\nDefStr F\nDefLng D\nDefDbl J-Z\n";
    assert_eq!(parse(src).found, vec![]);
    let src = "DefInt A-C, E, G-I\nDefStr I-J\n";
    assert_eq!(parse(src).found, vec![(DUPLICATE_DEFTYPE, offset_of(src, "DefStr", 0), 6)]);
}

#[test]
fn nothing_but_a_procedure_may_follow_a_procedure() {
    const AFTER_END_OF_PROCEDURE: u32 = 0x9c91;
    for decl in [
        "Option Explicit",
        "Option Base 1",
        "DefInt A-Z",
        "Dim n As Long",
        "Public n As Long",
        "Private n As Long",
        "Const n = 1",
        "Declare Sub D Lib \"k\" ()",
        "Type T\n    a As Long\nEnd Type",
        "Enum E\n    A\nEnd Enum",
    ] {
        for proc in ["Sub S()\nEnd Sub", "Function F()\nEnd Function", "Property Get P()\nEnd Property"] {
            let src = format!("{proc}\n{decl}\n");
            let p = parse(&src);
            let first = decl.split(' ').next().unwrap();
            assert_eq!(p.found, vec![(AFTER_END_OF_PROCEDURE, offset_of(&src, first, 0), first.len() as u32)], "for {src:?}");
        }
    }
    // Procedures, attribute lines and conditional compilation may follow.
    let src = "Sub S()\nEnd Sub\nAttribute VB_Name = \"m\"\n#Const k = 1\nPublic Sub T()\nEnd Sub\nProperty Get Q()\nEnd Property\n";
    assert_eq!(parse(src).found, vec![]);
}
