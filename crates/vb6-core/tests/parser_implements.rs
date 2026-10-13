//! `Implements [Library.]Interface` records the interface and its optional
//! type-library qualifier.

use vb6_core::frontend::ast::{ExprArena, ExprNode};
use vb6_core::frontend::parser::{ModuleKind, Parser};
use vb6_core::frontend::scanner::ScannerContext;

/// (diagnostics, [(library, interface)]) of every top-level Implements.
fn implements(src: &str) -> (usize, Vec<(Option<String>, String)>) {
    let mut ctx = ScannerContext::new(1, 1, 0x0409);
    ctx.intern_keywords();
    let mut arena = ExprArena::new();
    let mut parser = Parser::with_module_kind(&mut ctx, src.as_bytes(), ModuleKind::Class);
    let nodes = parser.parse_module(&mut arena);
    let diags = parser.diagnostics.items().len();
    drop(parser);
    let name = |sym: u32| ctx.symbol(sym as usize).name.clone();
    let out = nodes
        .iter()
        .map(|n| match arena.get(*n) {
            ExprNode::Implements { library, name: i } => (library.map(name), name(*i)),
            _ => panic!("expected Implements"),
        })
        .collect();
    (diags, out)
}

#[test]
fn unqualified_interface() {
    assert_eq!(implements("Implements IDataItem\n"), (0, vec![(None, "IDataItem".into())]));
}

#[test]
fn library_qualified_interface() {
    assert_eq!(
        implements("Implements AFCustom.ICustomFill\n"),
        (0, vec![(Some("AFCustom".into()), "ICustomFill".into())])
    );
}

#[test]
fn several_implements_lines() {
    assert_eq!(
        implements("Implements IDataItem\nImplements AFCustom.ICustomMarker\n"),
        (0, vec![(None, "IDataItem".into()), (Some("AFCustom".into()), "ICustomMarker".into())])
    );
}

#[test]
fn missing_interface_after_the_dot_is_an_error() {
    let (diags, _) = implements("Implements AFCustom.\n");
    assert_eq!(diags, 1);
}
