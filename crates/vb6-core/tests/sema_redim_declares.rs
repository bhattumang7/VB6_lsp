//! `ReDim a(n)` of a name that is not declared at procedure or module level
//! declares a local dynamic array at that statement.

use vb6_core::sema::{bind, VbaType};
use vb6_core::frontend::ast::ExprArena;
use vb6_core::frontend::parser::{ModuleKind, Parser};
use vb6_core::frontend::scanner::ScannerContext;
use vb6_core::session::{DeclId, Session};

fn at(src: &str, ctx: &str, member: &str) -> u32 {
    (src.find(ctx).unwrap() + ctx.find(member).unwrap()) as u32
}

fn session(src: &str) -> Session {
    Session::from_sources(vec![("m.bas".to_string(), src.as_bytes().to_vec())])
}

#[test]
fn undeclared_redim_declares_a_local_array() {
    let src = "Option Explicit\nSub Run()\n    ReDim arr(5) As clsItem\n    arr(1).Save\n    n = UBound(arr)\nEnd Sub\n";
    let s = session(src);
    let local = Some(DeclId::Local { module: 0, proc: 0, idx: 0 });
    // The ReDim's name is the declaration; later uses resolve to it.
    assert_eq!(s.declaration_at(0, at(src, "ReDim arr", "arr")), local);
    assert_eq!(s.declaration_at(0, at(src, "arr(1).Save", "arr")), local);
    assert_eq!(s.declaration_at(0, at(src, "UBound(arr)", "arr")), local);
    // Declared at the ReDim: the only use sites are the two later ones.
    let refs = s.references(0, at(src, "ReDim arr", "arr"), false);
    let starts: Vec<u32> = refs.iter().map(|l| l.span.start).collect();
    assert_eq!(starts, vec![at(src, "arr(1).Save", "arr"), at(src, "UBound(arr)", "arr")]);
    // `n` is the only undeclared name.
    let diags = s.diagnostics(0);
    assert_eq!(diags.len(), 1);
    assert_eq!(diags[0].span.start, at(src, "n = UBound", "n"));
}

#[test]
fn declared_local_is_not_redeclared() {
    let src = "Sub Run()\n    Dim arr() As Long\n    ReDim arr(5)\n    ReDim arr(9)\nEnd Sub\n";
    let s = session(src);
    // One local: the Dim; both ReDims are uses of it.
    let decl = Some(DeclId::Local { module: 0, proc: 0, idx: 0 });
    assert_eq!(s.declaration_at(0, at(src, "ReDim arr(5)", "arr")), decl);
    assert_eq!(s.references(0, at(src, "Dim arr", "arr"), false).len(), 2);
}

#[test]
fn module_level_array_is_not_redeclared() {
    let src = "Private mArr() As Long\nSub Run()\n    ReDim mArr(3)\nEnd Sub\n";
    let s = session(src);
    assert_eq!(
        s.declaration_at(0, at(src, "ReDim mArr", "mArr")),
        Some(DeclId::ModuleVar { module: 0, idx: 0 })
    );
}

#[test]
fn public_array_of_another_module_is_not_redeclared() {
    let globals = "Public gArr() As Long\n";
    let main = "Sub Run()\n    ReDim gArr(3)\nEnd Sub\n";
    let s = Session::from_sources(vec![
        ("modGlobals.bas".to_string(), globals.as_bytes().to_vec()),
        ("modMain.bas".to_string(), main.as_bytes().to_vec()),
    ]);
    assert_eq!(
        s.declaration_at(1, at(main, "ReDim gArr", "gArr")),
        Some(DeclId::ModuleVar { module: 0, idx: 0 })
    );
}

#[test]
fn parameter_is_not_redeclared() {
    let src = "Sub Run(a() As Long)\n    ReDim a(3)\nEnd Sub\n";
    let s = session(src);
    assert_eq!(
        s.declaration_at(0, at(src, "ReDim a(3)", "a")),
        Some(DeclId::Param { module: 0, proc: 0, idx: 0 })
    );
}

#[test]
fn implicit_array_has_the_declared_element_type_and_dimensions() {
    let src = "Sub Run()\n    ReDim a(3, 4) As Long\n    ReDim b(2)\nEnd Sub\n";
    let mut ctx = ScannerContext::new(1, 1, 0x0409);
    ctx.intern_keywords();
    let mut arena = ExprArena::new();
    let mut parser = Parser::with_module_kind(&mut ctx, src.as_bytes(), ModuleKind::Standard);
    let top = parser.parse_module(&mut arena);
    let spans = std::mem::take(&mut parser.node_spans);
    let vis = std::mem::take(&mut parser.decl_public);
    drop(parser);
    let m = bind(&ctx, &arena, &top, &spans, &vis);
    let locals = &m.procs[0].locals;
    assert_eq!(locals.len(), 2);
    assert_eq!(locals[0].vba_type, VbaType::Array(Box::new(VbaType::Long)));
    assert_eq!(locals[0].array_dims, Some(2));
    assert_eq!(locals[1].vba_type, VbaType::Array(Box::new(VbaType::Variant)));
    assert_eq!(locals[1].array_dims, Some(1));
}
