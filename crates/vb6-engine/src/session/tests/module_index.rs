use super::*;
use crate::frontend::ast::ExprArena;
use crate::frontend::parser::Parser;
use crate::frontend::scanner::ScannerContext;
use crate::sema::bind;

/// Parse+bind one module, returning its bound form plus a name resolver.
fn bind_module(src: &str) -> (BoundModule, ScannerContext) {
    let mut ctx = ScannerContext::new(1, 1, 0x0409);
    ctx.intern_keywords();
    let mut arena = ExprArena::new();
    let mut parser = Parser::new(&mut ctx, src.as_bytes());
    let top = parser.parse_module(&mut arena);
    let spans = std::mem::take(&mut parser.node_spans);
    let vis = std::mem::take(&mut parser.decl_public);
    drop(parser);
    let m = bind(&ctx, &arena, &top, &spans, &vis);
    (m, ctx)
}

#[test]
fn indexes_public_procs_and_vars() {
    let (m, ctx) = bind_module(
        "Public gConfig As Long\nPublic Sub DoThing()\nEnd Sub\nFunction Calc() As Long\nEnd Function\n",
    );
    let mut idx = ModuleIndex::new();
    idx.add_module(0, &m, |s| ctx.symbol(s as usize).name.to_ascii_lowercase(), true);

    // Case-insensitive lookup.
    assert_eq!(idx.lookup("dothing").map(|r| r.decl), Some(ExternalDecl::Proc(0)));
    assert_eq!(idx.lookup("CALC").map(|r| r.decl), Some(ExternalDecl::Proc(1)));
    assert_eq!(idx.lookup("gconfig").map(|r| r.decl), Some(ExternalDecl::Var(0)));
    assert!(idx.lookup("missing").is_none());
}

#[test]
fn indexes_enum_members() {
    let (m, ctx) = bind_module(
        "Public Enum Color\n    Red\n    Green\n    Blue\nEnd Enum\n",
    );
    let mut idx = ModuleIndex::new();
    idx.add_module(2, &m, |s| ctx.symbol(s as usize).name.to_ascii_lowercase(), true);

    assert_eq!(idx.lookup("color").map(|r| r.decl), Some(ExternalDecl::Enum(0)));
    assert_eq!(
        idx.lookup("green").map(|r| r.decl),
        Some(ExternalDecl::EnumMember { enum_idx: 0, member_idx: 1 }),
    );
    // The referenced module index is propagated.
    assert_eq!(idx.lookup("red").map(|r| r.module), Some(2));
}

#[test]
fn first_module_wins_on_clash() {
    let (m0, c0) = bind_module("Public Sub Shared()\nEnd Sub\n");
    let (m1, c1) = bind_module("Public Sub Shared()\nEnd Sub\n");
    let mut idx = ModuleIndex::new();
    idx.add_module(0, &m0, |s| c0.symbol(s as usize).name.to_ascii_lowercase(), true);
    idx.add_module(1, &m1, |s| c1.symbol(s as usize).name.to_ascii_lowercase(), true);
    assert_eq!(idx.lookup("shared").map(|r| r.module), Some(0));
}
