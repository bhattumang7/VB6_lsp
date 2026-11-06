//! A name written again in other letter case takes the new spelling, which every use of the
//! name then shows. Keywords keep theirs.

mod common;
use common::parse;
use vb6_syntax::frontend::ast::ExprNode;
use vb6_syntax::frontend::scanner::ScannerContext;

/// The spelling of the name on the left of each assignment in the first procedure, after the
/// whole module has been read.
fn assigned_spellings(src: &str) -> Vec<String> {
    let p = parse(src);
    assert_eq!(p.found, vec![], "{src}");
    p.proc_body(0)
        .iter()
        .filter_map(|id| match p.node(*id) {
            ExprNode::Assign { target, .. } => match p.node(*target) {
                ExprNode::NameRef { sym, .. } => Some(p.name(*sym).to_string()),
                _ => None,
            },
            _ => None,
        })
        .collect()
}

#[test]
fn every_use_shows_the_latest_spelling() {
    let names = assigned_spellings("Sub S()\n Dim abc As Long\n abc = 1\n ABC = 2\n Abc = 3\nEnd Sub\n");
    assert_eq!(names, vec!["Abc", "Abc", "Abc"]);
}

#[test]
fn different_names_keep_their_own_spelling() {
    let names = assigned_spellings("Sub S()\n Dim abc, abd As Long\n abc = 1\n abd = 2\n ABC = 3\nEnd Sub\n");
    assert_eq!(names, vec!["ABC", "abd", "ABC"]);
}

#[test]
fn accented_letters_take_the_latest_spelling_too() {
    let mut ctx = ScannerContext::new(1, 1, 0x0409);
    ctx.intern_keywords();
    let first = ctx.intern_ident("\u{e9}x").unwrap();
    assert_eq!(ctx.sym_name(first), "\u{e9}x");
    let again = ctx.intern_ident("\u{c9}X").unwrap();
    assert_eq!(first, again);
    assert_eq!(ctx.sym_name(first), "\u{c9}X");
}

#[test]
fn keywords_keep_their_own_spelling() {
    let mut ctx = ScannerContext::new(1, 1, 0x0409);
    ctx.intern_keywords();
    let id = ctx.intern_ident("print").unwrap();
    assert_eq!(ctx.sym_name(id), "Print");
    let id = ctx.intern_ident("PRINT").unwrap();
    assert_eq!(ctx.sym_name(id), "Print");
}
