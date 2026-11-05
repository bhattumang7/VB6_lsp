//! A procedure, declare or event takes at most 60 parameters in a standard module and
//! 59 in an object module; the first parameter past the limit is reported.

mod common;
use common::parse;
use vb6_syntax::frontend::ast::ExprArena;
use vb6_syntax::frontend::parser::{ModuleKind, Parser};
use vb6_syntax::frontend::scanner::ScannerContext;

const TOO_MANY: u32 = 0x9c81;

fn params(n: usize) -> String {
    (0..n).map(|i| format!("a{i} As Long")).collect::<Vec<_>>().join(", ")
}

fn found(kind: ModuleKind, src: &str) -> Vec<(u32, u32, u32)> {
    let mut ctx = ScannerContext::new(1, 1, 0x0409);
    ctx.intern_keywords();
    let mut arena = ExprArena::new();
    let mut parser = Parser::with_module_kind(&mut ctx, src.as_bytes(), kind);
    parser.parse_module(&mut arena);
    parser.diagnostics.items().iter().map(|d| (d.code, d.span.start, d.span.len)).collect()
}

fn headers(n: usize) -> Vec<(String, &'static str)> {
    let p = params(n);
    vec![
        (format!("Sub Foo({p})\nEnd Sub\n"), "Sub"),
        (format!("Function Foo({p}) As Long\nEnd Function\n"), "Function"),
        (format!("Private Declare Function Foo Lib \"k\" ({p}) As Long\n"), "Declare"),
        (format!("Private Declare Sub Foo Lib \"k\" ({p})\n"), "DeclareSub"),
    ]
}

#[test]
fn a_standard_module_takes_sixty_parameters() {
    for (src, what) in headers(60) {
        assert_eq!(parse(&src).found, vec![], "{what}");
    }
    for (src, what) in headers(61) {
        let at = src.find("a60").unwrap() as u32;
        assert_eq!(parse(&src).found, vec![(TOO_MANY, at, 3)], "{what}");
    }
}

#[test]
fn an_object_module_takes_fifty_nine_parameters() {
    for (src, what) in headers(59) {
        assert_eq!(found(ModuleKind::Class, &src), vec![], "{what}");
    }
    for (src, what) in headers(60) {
        let at = src.find("a59").unwrap() as u32;
        assert_eq!(found(ModuleKind::Class, &src), vec![(TOO_MANY, at, 3)], "{what}");
    }
}

#[test]
fn an_event_and_a_property_follow_the_same_limit() {
    let src = format!("Event Ev({})\n", params(60));
    assert_eq!(found(ModuleKind::Class, &src), vec![(TOO_MANY, src.find("a59").unwrap() as u32, 3)]);
    let src = format!("Property Get P({}) As Long\nEnd Property\n", params(59));
    assert_eq!(found(ModuleKind::Class, &src), vec![]);
}
