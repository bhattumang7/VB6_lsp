//! Names that differ only in letter case are one name, for every letter the Western code page
//! has a case pair for: the accented Latin-1 letters and the four letters of the 0x80-0x9F
//! range (S-caron, OE, Z-caron, Y-diaeresis).

use vb6_syntax::frontend::ast::{ExprArena, ExprNode};
use vb6_syntax::frontend::parser::Parser;
use vb6_syntax::frontend::scanner::{fold_name, ScannerContext};

/// Parse raw bytes; returns the diagnostics' codes and the symbol ids and names of the
/// left-hand names of the assignments in the first procedure.
fn assigned_names(src: &[u8]) -> (Vec<u32>, Vec<(u32, String)>) {
    let mut ctx = ScannerContext::new(1, 1, 0x0409);
    ctx.intern_keywords();
    let mut arena = ExprArena::new();
    let (top, codes) = {
        let mut parser = Parser::new(&mut ctx, src);
        let top = parser.parse_module(&mut arena);
        let codes = parser.diagnostics.items().iter().map(|d| d.code).collect();
        (top, codes)
    };
    let mut names = Vec::new();
    for id in top {
        let ExprNode::ProcDecl { body, .. } = arena.get(id) else { continue };
        let ExprNode::Block { stmts } = arena.get(*body) else { panic!("body is not a block") };
        for stmt in stmts {
            let ExprNode::Assign { target, .. } = arena.get(*stmt) else { continue };
            if let ExprNode::NameRef { sym, .. } = arena.get(*target) {
                names.push((*sym, ctx.sym_name(*sym as usize).to_string()));
            }
        }
    }
    (codes, names)
}

fn module(first: u8, second: u8) -> Vec<u8> {
    let mut v = b"Sub S()\n".to_vec();
    v.extend_from_slice(b" ");
    v.push(first);
    v.extend_from_slice(b"x = 1\n ");
    v.push(second);
    v.extend_from_slice(b"X = 2\nEnd Sub\n");
    v
}

#[test]
fn windows_1252_letters_outside_latin_1_fold_with_their_case_partner() {
    // (upper, lower): S-caron, OE, Z-caron, Y-diaeresis.
    for (upper, lower) in [(0x8Au8, 0x9Au8), (0x8C, 0x9C), (0x8E, 0x9E), (0x9F, 0xFF)] {
        for (a, b) in [(upper, lower), (lower, upper)] {
            let (codes, names) = assigned_names(&module(a, b));
            assert_eq!(codes, Vec::<u32>::new(), "{a:#x} {b:#x}");
            assert_eq!(names.len(), 2, "{a:#x} {b:#x}");
            assert_eq!(names[0].0, names[1].0, "{a:#x} and {b:#x} are one name");
        }
    }
}

#[test]
fn latin_1_letters_still_fold() {
    // A-grave and a-grave, Thorn and thorn, O-slash and o-slash.
    for (upper, lower) in [(0xC0u8, 0xE0u8), (0xDE, 0xFE), (0xD8, 0xF8)] {
        let (_, names) = assigned_names(&module(upper, lower));
        assert_eq!(names[0].0, names[1].0, "{upper:#x}");
    }
}

#[test]
fn letters_without_a_case_partner_stay_apart() {
    // The multiplication sign and the division sign are not letters; the sharp s has no
    // upper-case partner in the code page; the micro sign likewise.
    for (a, b) in [(0xD7u8, 0xF7u8), (0xDF, 0xFF), (0x8A, 0x8C), (0x9A, 0x9E), (0xB5, 0xC0)] {
        let (_, names) = assigned_names(&module(a, b));
        assert_ne!(names[0].0, names[1].0, "{a:#x} {b:#x}");
    }
}

#[test]
fn a_jump_to_a_label_written_in_the_other_case_finds_it() {
    for (upper, lower) in [(0x8Au8, 0x9Au8), (0x8C, 0x9C), (0x8E, 0x9E), (0x9F, 0xFF)] {
        for (use_byte, label_byte) in [(upper, lower), (lower, upper)] {
            let mut src = b"Sub S()\n GoTo ".to_vec();
            src.push(use_byte);
            src.extend_from_slice(b"\n");
            src.push(label_byte);
            src.extend_from_slice(b":\nEnd Sub\n");
            let (codes, _) = assigned_names(&src);
            assert_eq!(codes, Vec::<u32>::new(), "{use_byte:#x} {label_byte:#x}");
        }
    }
}

#[test]
fn fold_name_maps_each_letter_to_its_lower_case_form() {
    assert_eq!(fold_name("\u{160}\u{152}\u{17d}\u{178}"), "\u{161}\u{153}\u{17e}\u{ff}");
    assert_eq!(fold_name("\u{161}\u{153}\u{17e}\u{ff}"), "\u{161}\u{153}\u{17e}\u{ff}");
    assert_eq!(fold_name("\u{c0}\u{d7}\u{de}Az"), "\u{e0}\u{d7}\u{fe}az");
}
