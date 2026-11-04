//! Rules for what a class module may declare publicly, modifier combinations
//! that are not declarations, WithEvents types and array-typed variables.

mod common;
use common::{offset_of, parse};
use vb6_syntax::frontend::ast::ExprArena;
use vb6_syntax::frontend::parser::{ModuleKind, Parser};
use vb6_syntax::frontend::scanner::ScannerContext;

const PRIVATE_HEADER: &str = "VERSION 1.0 CLASS\nBEGIN\n  MultiUse = -1  'True\nEND\nAttribute VB_Name = \"Class1\"\nAttribute VB_GlobalNameSpace = False\nAttribute VB_Creatable = True\nAttribute VB_PredeclaredId = False\nAttribute VB_Exposed = False\n";
const EXPOSED_HEADER: &str = "VERSION 1.0 CLASS\nBEGIN\n  MultiUse = -1  'True\nEND\nAttribute VB_Name = \"Class1\"\nAttribute VB_GlobalNameSpace = False\nAttribute VB_Creatable = True\nAttribute VB_PredeclaredId = False\nAttribute VB_Exposed = True\n";

fn class_found(header: &str, body: &str) -> (String, Vec<(u32, u32, u32)>) {
    let src = format!("{header}{body}\n");
    let mut ctx = ScannerContext::new(1, 1, 0x0409);
    ctx.intern_keywords();
    let mut arena = ExprArena::new();
    let mut parser = Parser::with_module_kind(&mut ctx, src.as_bytes(), ModuleKind::Class);
    parser.parse_module(&mut arena);
    let found = parser.diagnostics.items().iter().map(|d| (d.code, d.span.start, d.span.len)).collect();
    (src, found)
}

const NOT_ALLOWED: u32 = 0x9da8;
const PRIVATE_TYPE: u32 = 0x9da0;

#[test]
fn public_constants_declares_arrays_and_fixed_strings_are_not_allowed_in_a_class() {
    for (body, at) in [
        ("Public Const c = 1", "Public"),
        ("Public Const c As Long = 1", "Public"),
        ("Declare Function f Lib \"k\" () As Long", "Declare"),
        ("Public Declare Function f Lib \"k\" () As Long", "Public"),
        ("Public a(3) As Long", "Public"),
        ("Public a() As Long", "Public"),
        ("Public s As String * 3", "Public"),
    ] {
        let (src, found) = class_found(PRIVATE_HEADER, body);
        assert_eq!(found.len(), 1, "{body}: {found:?}");
        assert_eq!((found[0].0, found[0].1), (NOT_ALLOWED, offset_of(&src, at, 0)), "{body}");
    }
}

#[test]
fn private_and_default_members_of_a_class_are_fine() {
    for body in [
        "Const c = 1",
        "Private Const c = 1",
        "Private Declare Function f Lib \"k\" () As Long",
        "Dim a(3) As Long",
        "Private a(3) As Long",
        "Dim s As String * 3",
        "Public s As String",
        "Public x",
        "Public Event E()",
        "Public Enum E\n    A\nEnd Enum",
        "Private Type T\n    a As Long\nEnd Type",
        "Friend Sub P()\nEnd Sub",
        "Public Static Sub P()\nEnd Sub",
    ] {
        assert_eq!(class_found(PRIVATE_HEADER, body).1, vec![], "{body}");
    }
}

#[test]
fn global_is_not_allowed_in_a_class() {
    for body in ["Global x As Long", "Global Const c = 1", "Global Sub P()\nEnd Sub"] {
        let (src, found) = class_found(PRIVATE_HEADER, body);
        assert_eq!((found[0].0, found[0].1), (NOT_ALLOWED, PRIVATE_HEADER.len() as u32), "{body}");
        let _ = &src;
    }
}

#[test]
fn a_public_type_needs_an_exposed_class() {
    for body in ["Type T\n    a As Long\nEnd Type", "Public Type T\n    a As Long\nEnd Type"] {
        let (src, found) = class_found(PRIVATE_HEADER, body);
        assert_eq!(found.len(), 1, "{body}: {found:?}");
        let at = if body.starts_with("Public") { "Public" } else { "Type" };
        assert_eq!((found[0].0, found[0].1), (PRIVATE_TYPE, offset_of(&src, at, 0)), "{body}");
        assert_eq!(class_found(EXPOSED_HEADER, body).1, vec![], "{body}");
    }
}

#[test]
fn standard_modules_have_no_such_limits() {
    for text in [
        "Public Const c = 1\n",
        "Global Const c = 1\n",
        "Public a(3) As Long\n",
        "Public s As String * 3\n",
        "Public Type T\n    a As Long\nEnd Type\n",
        "Declare Function f Lib \"k\" () As Long\n",
    ] {
        assert_eq!(parse(text).found, vec![], "{text}");
    }
}

#[test]
fn a_modifier_cannot_be_followed_by_dim_or_implements() {
    for text in ["Public Dim x\n", "Private Dim x\n", "Global Dim x\n"] {
        let found = parse(text).found;
        assert_eq!(found[0].1, offset_of(text, "Dim", 0), "{text}: {found:?}");
    }
    let (src, found) = class_found(PRIVATE_HEADER, "Public Implements Class1");
    assert!(found.iter().any(|d| d.1 == offset_of(&src, "Implements", 0)), "{found:?}");
}

#[test]
fn a_with_events_variable_takes_a_named_class_only() {
    for ty in ["Long", "Object", "String", "Variant", "Integer", "Date", "String * 3"] {
        let (src, found) = class_found(PRIVATE_HEADER, &format!("Private WithEvents x As {ty}"));
        assert!(found.iter().any(|d| d.1 == offset_of(&src, ty.split(' ').next().unwrap(), 0) && d.0 == 0x9c6f), "{ty}: {found:?}");
    }
    let (src, found) = class_found(PRIVATE_HEADER, "Private WithEvents x As New Class1");
    assert_eq!(found[0].0, 0x9d9c, "{found:?}");
    assert_eq!(found[0].1, offset_of(&src, "New", 0));
    assert_eq!(class_found(PRIVATE_HEADER, "Private WithEvents x As Class1").1, vec![]);
}

#[test]
fn a_variable_type_takes_no_array_marker() {
    for text in ["Dim x As Long()\n", "Private x As Long()\n", "Public x As Long()\n", "Dim x() As Long()\n"] {
        assert!(!parse(text).found.is_empty(), "{text}");
    }
    assert_eq!(parse("Function F() As Long()\nEnd Function\n").found, vec![]);
}
