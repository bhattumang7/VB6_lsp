use super::*;
use crate::frontend::ast::{ExprArena, NodeSpans, Span};
use crate::frontend::parser::Parser;
use crate::frontend::scanner::ScannerContext;
use crate::sema::{bind, BoundModule};

fn analyze(src: &str) -> (BoundModule, NodeSpans, String) {
    let mut ctx = ScannerContext::new(1, 1, 0x0409);
    ctx.intern_keywords();
    let mut arena = ExprArena::new();
    let mut parser = Parser::new(&mut ctx, src.as_bytes());
    let top = parser.parse_module(&mut arena);
    let spans = std::mem::take(&mut parser.node_spans);
    let vis = std::mem::take(&mut parser.decl_public);
    drop(parser);
    let m = bind(&ctx, &arena, &top, &spans, &vis);
    (m, spans, src.to_string())
}

fn text(src: &str, span: Span) -> &str {
    let lo = span.start as usize;
    &src[lo..lo + span.len as usize]
}

#[test]
fn collects_local_uses() {
    let src = "Sub Foo()\n    Dim y As Long\n    y = y + 1\n    y = y * 2\nEnd Sub\n";
    let (m, spans, src) = analyze(src);
    let mut idx = ReferenceIndex::new();
    idx.add_module(0, &m.resolutions);

    // `y` is local 0 of proc 0.
    let sites = idx.references(DeclId::Local { module: 0, proc: 0, idx: 0 });
    // Four uses: target+rhs on two assignment lines.
    assert_eq!(sites.len(), 4);
    for s in &sites {
        assert_eq!(s.module, 0);
        assert_eq!(text(&src, spans.get(crate::frontend::ast::NodeId(s.node))), "y");
    }
}

#[test]
fn collects_module_var_uses_and_skips_builtins() {
    let src = "Public gN As Long\nSub Foo()\n    gN = Len(\"hi\")\n    gN = gN + 1\nEnd Sub\n";
    let (m, _spans, _src) = analyze(src);
    let mut idx = ReferenceIndex::new();
    idx.add_module(0, &m.resolutions);

    // gN used three times (two targets + one rhs); Len is a builtin (no decl).
    let sites = idx.references(DeclId::ModuleVar { module: 0, idx: 0 });
    assert_eq!(sites.len(), 3);
}

#[test]
fn external_resolution_maps_to_target_module() {
    // Build a reverse index from a hand-made external resolution to confirm
    // decl_id_of routes the use to the *target* module, not the use's module.
    let mut resolutions = HashMap::new();
    resolutions.insert(
        7u32,
        NameResolution::External {
            module: 3,
            decl: ExternalDecl::Proc(2),
        },
    );
    let mut idx = ReferenceIndex::new();
    idx.add_module(0, &resolutions); // use lives in module 0

    let sites = idx.references(DeclId::Proc { module: 3, idx: 2 });
    assert_eq!(sites, vec![RefSite { module: 0, node: 7 }]);
}
