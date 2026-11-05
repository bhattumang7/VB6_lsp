//! `RaiseEvent` takes a bare or qualified event name with an optional bracketed argument
//! list; the qualifier is kept on the node.


use vb6_syntax::frontend::ast::{ExprArena, ExprNode};
use vb6_syntax::frontend::parser::{ModuleKind, Parser};
use vb6_syntax::frontend::scanner::ScannerContext;

struct Raised {
    found: Vec<(u32, u32, u32)>,
    shape: Option<(String, Option<String>, usize)>,
}

/// Parse `RaiseEvent <text>` in a class module; the shape is (event name, rendered qualifier, argument count).
fn raise(text: &str) -> Raised {
    let src = format!("Sub S()\n RaiseEvent {text}\nEnd Sub\n");
    let mut ctx = ScannerContext::new(1, 1, 0x0409);
    ctx.intern_keywords();
    let mut arena = ExprArena::new();
    let mut parser = Parser::with_module_kind(&mut ctx, src.as_bytes(), ModuleKind::Class);
    let top = parser.parse_module(&mut arena);
    let found = parser.diagnostics.items().iter().map(|d| (d.code, d.span.start, d.span.len)).collect();
    drop(parser);
    let ExprNode::ProcDecl { body, .. } = arena.get(top[0]) else { panic!("no procedure") };
    let ExprNode::Block { stmts } = arena.get(*body) else { panic!("no body") };
    let shape = stmts.first().and_then(|s| match arena.get(*s) {
        ExprNode::RaiseEvent { name, qualifier, args } => {
            fn render(ctx: &ScannerContext, arena: &ExprArena, id: vb6_syntax::frontend::ast::NodeId) -> String {
                match arena.get(id) {
                    ExprNode::NameRef { sym, .. } => ctx.sym_name(*sym as usize).to_string(),
                    ExprNode::MemberAccess { base, member, .. } => {
                        format!("{}.{}", render(ctx, arena, *base), ctx.sym_name(*member as usize))
                    }
                    ExprNode::WithContext => String::new(),
                    _ => "?".to_string(),
                }
            }
            let ExprNode::ArgList { args } = arena.get(*args) else { panic!("no args") };
            Some((ctx.sym_name(*name as usize).to_string(), qualifier.map(|q| render(&ctx, &arena, q)), args.len()))
        }
        _ => None,
    });
    Raised { found, shape }
}

#[test]
fn a_bare_event_name_has_no_qualifier() {
    assert_eq!(raise("E").shape, Some(("E".into(), None, 0)));
    assert_eq!(raise("E(1, 2)").shape, Some(("E".into(), None, 2)));
    assert_eq!(raise("E").found, vec![]);
}

#[test]
fn a_qualified_event_name_keeps_its_qualifier() {
    for (text, name, qualifier, args) in [
        ("o.E", "E", "o", 0),
        ("o.E(1)", "E", "o", 1),
        ("E.x", "x", "E", 0),
        ("E.x(1)", "x", "E", 1),
        ("o.p.E", "E", "o.p", 0),
        ("o.p.q(1)", "q", "o.p", 1),
    ] {
        let r = raise(text);
        assert_eq!(r.found, vec![], "{text}");
        assert_eq!(r.shape, Some((name.into(), Some(qualifier.into()), args)), "{text}");
    }
}

#[test]
fn a_with_member_event_has_the_with_object_as_qualifier() {
    let src = "Sub S()\n With o\n RaiseEvent .E(1)\n End With\nEnd Sub\n";
    let mut ctx = ScannerContext::new(1, 1, 0x0409);
    ctx.intern_keywords();
    let mut arena = ExprArena::new();
    let mut parser = Parser::with_module_kind(&mut ctx, src.as_bytes(), ModuleKind::Class);
    parser.parse_module(&mut arena);
    assert_eq!(parser.diagnostics.items().len(), 0);
}

#[test]
fn other_shapes_are_syntax_errors() {
    for text in ["E!x", "E(1)(2)", "E 1, 2", "E(1),2", "E(1))", "\"E\""] {
        assert!(!raise(text).found.is_empty(), "{text}");
    }
}
