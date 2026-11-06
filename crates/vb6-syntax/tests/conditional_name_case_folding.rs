//! Conditional-compilation names are compared ignoring letter case for every letter that has a
//! case partner, the accented ones included: in the project's argument text, in `#Const`
//! lines, and between the two.

use vb6_syntax::frontend::ast::{ExprArena, ExprNode};
use vb6_syntax::frontend::cond_compile::{parse_cond_comp_args, CcValue, ConstantDeclarationsError};
use vb6_syntax::frontend::diagnostics::ERR_CC_DUPLICATE_DEFINITION;
use vb6_syntax::frontend::parser::Parser;
use vb6_syntax::frontend::scanner::ScannerContext;

#[test]
fn project_arguments_differing_only_in_accented_letter_case_are_duplicates() {
    for text in ["\u{c9}x=1:\u{e9}x=2", "\u{e9}x=1:\u{c9}x=2", "\u{c0}a=1:\u{e0}A=2", "\u{de}=1:\u{fe}=2", "Ax=1:aX=2"] {
        assert_eq!(parse_cond_comp_args(text), Err(ConstantDeclarationsError::InvalidSyntax), "{text}");
    }
}

#[test]
fn project_argument_names_are_kept_folded() {
    assert_eq!(
        parse_cond_comp_args("\u{c9}x=1 : \u{d6}\u{c6}=2"),
        Ok(vec![
            ("\u{e9}x".to_string(), CcValue::Integer(1)),
            ("\u{f6}\u{e6}".to_string(), CcValue::Integer(2)),
        ])
    );
}

/// The name of the procedure kept by `#If <name> Then Sub Yes #Else Sub No` when the project
/// defines `args` and the condition is the raw bytes `name_bytes`.
fn taken(args: &str, name_bytes: &[u8]) -> String {
    let mut ctx = ScannerContext::new(1, 1, 0x0409);
    ctx.intern_keywords();
    let mut arena = ExprArena::new();
    let mut src = b"#If ".to_vec();
    src.extend_from_slice(name_bytes);
    src.extend_from_slice(b" Then\nSub Yes()\nEnd Sub\n#Else\nSub No()\nEnd Sub\n#End If\n");
    let mut parser = Parser::new(&mut ctx, &src);
    parser.set_conditional_constants(parse_cond_comp_args(args).unwrap());
    let top = parser.parse_module(&mut arena);
    assert!(parser.diagnostics.items().is_empty());
    drop(parser);
    let ExprNode::ProcDecl { name, .. } = arena.get(top[0]) else { panic!("no procedure") };
    ctx.sym_name(*name as usize).to_string()
}

#[test]
fn a_project_argument_is_found_by_a_name_written_in_the_other_case() {
    // (argument text, source bytes of the name in the condition)
    for (args, name) in [
        ("\u{c9}x=1", &[0xE9u8, b'X'][..]),
        ("\u{e9}x=1", &[0xC9, b'x'][..]),
    ] {
        assert_eq!(taken(args, name), "Yes", "{args}");
    }
}

#[test]
fn a_name_that_differs_by_more_than_case_is_not_found() {
    assert_eq!(taken("\u{c9}x=1", &[0xE8, b'x']), "No");
    assert_eq!(taken("\u{d6}=1", &[0x9C]), "No");
}

#[test]
fn a_module_constant_is_found_by_a_name_written_in_the_other_case() {
    let mut ctx = ScannerContext::new(1, 1, 0x0409);
    ctx.intern_keywords();
    let mut arena = ExprArena::new();
    let src: Vec<u8> = [
        &b"#Const "[..],
        &[0x8A, b'x'],
        &b" = 1\n#If "[..],
        &[0x9A, b'X'],
        &b" Then\nSub Yes()\nEnd Sub\n#Else\nSub No()\nEnd Sub\n#End If\n"[..],
    ]
    .concat();
    let mut parser = Parser::new(&mut ctx, &src);
    let top = parser.parse_module(&mut arena);
    assert!(parser.diagnostics.items().is_empty());
    drop(parser);
    let ExprNode::ProcDecl { name, .. } = arena.get(top[0]) else { panic!("no procedure") };
    assert_eq!(ctx.sym_name(*name as usize), "Yes");
}

#[test]
fn a_module_constant_defined_twice_in_different_case_is_a_duplicate() {
    let mut ctx = ScannerContext::new(1, 1, 0x0409);
    ctx.intern_keywords();
    let mut arena = ExprArena::new();
    let src: Vec<u8> = [&b"#Const "[..], &[0x8A, b'x'], &b" = 1\n#Const "[..], &[0x9A, b'X'], &b" = 2\n"[..]].concat();
    let mut parser = Parser::new(&mut ctx, &src);
    parser.parse_module(&mut arena);
    let codes: Vec<u32> = parser.diagnostics.items().iter().map(|d| d.code).collect();
    assert_eq!(codes, vec![ERR_CC_DUPLICATE_DEFINITION]);
}
