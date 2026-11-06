//! A whole-number literal written in octal or hexadecimal says so in the tree: `&H10` and
//! `16` are the same Integer value but not the same literal.

use vb6_syntax::frontend::ast::{AstLit, ExprArena, ExprNode, NodeId, Radix};
use vb6_syntax::frontend::parser::Parser;
use vb6_syntax::frontend::scanner::ScannerContext;

/// Parse `x = <text>` and return the literal on the right with its radix.
fn literal(text: &str) -> (AstLit, Option<Radix>) {
    let mut ctx = ScannerContext::new(1, 1, 0x0409);
    ctx.intern_keywords();
    let mut arena = ExprArena::new();
    let src = format!("Sub S()\n x = {text}\nEnd Sub\n");
    let mut parser = Parser::new(&mut ctx, src.as_bytes());
    let top = parser.parse_module(&mut arena);
    assert!(parser.diagnostics.items().is_empty(), "{text}");
    let ExprNode::ProcDecl { body, .. } = arena.get(top[0]) else { panic!("no procedure") };
    let ExprNode::Block { stmts } = arena.get(*body) else { panic!("no block") };
    let ExprNode::Assign { value, .. } = arena.get(stmts[0]) else { panic!("no assignment") };
    let value: NodeId = *value;
    let ExprNode::Literal { lit } = arena.get(value) else { panic!("not a literal: {text}") };
    (lit.clone(), parser.literal_radix.get(&value.0).copied())
}

#[test]
fn hexadecimal_literals_carry_the_hexadecimal_marker() {
    assert_eq!(literal("&H10"), (AstLit::Int(16), Some(Radix::Hexadecimal)));
    assert_eq!(literal("&h10"), (AstLit::Int(16), Some(Radix::Hexadecimal)));
    assert_eq!(literal("&HFFFF"), (AstLit::Int(-1), Some(Radix::Hexadecimal)));
    assert_eq!(literal("&H10000"), (AstLit::Long(65536), Some(Radix::Hexadecimal)));
    assert_eq!(literal("&HFF&"), (AstLit::Long(255), Some(Radix::Hexadecimal)));
    assert_eq!(literal("&H7F%"), (AstLit::Int(127), Some(Radix::Hexadecimal)));
}

#[test]
fn octal_literals_carry_the_octal_marker_with_or_without_the_letter() {
    assert_eq!(literal("&O17"), (AstLit::Int(15), Some(Radix::Octal)));
    assert_eq!(literal("&o17"), (AstLit::Int(15), Some(Radix::Octal)));
    assert_eq!(literal("&17"), (AstLit::Int(15), Some(Radix::Octal)));
    assert_eq!(literal("&177777"), (AstLit::Int(-1), Some(Radix::Octal)));
    assert_eq!(literal("&O200000"), (AstLit::Long(65536), Some(Radix::Octal)));
    assert_eq!(literal("&17&"), (AstLit::Long(15), Some(Radix::Octal)));
}

#[test]
fn decimal_literals_carry_no_marker() {
    assert_eq!(literal("16"), (AstLit::Int(16), None));
    assert_eq!(literal("65535"), (AstLit::Long(65535), None));
    assert_eq!(literal("16&"), (AstLit::Long(16), None));
    assert_eq!(literal("1.5"), (AstLit::Double(1.5), None));
    assert_eq!(literal("\"&H10\""), (AstLit::Str("&H10".into()), None));
}

#[test]
fn the_same_value_written_two_ways_differs_only_in_the_marker() {
    let (hex, hex_radix) = literal("&H10");
    let (dec, dec_radix) = literal("16");
    assert_eq!(hex, dec);
    assert_ne!(hex_radix, dec_radix);
}
