//! The exact shapes of bare `Scale`, `Line`, `Circle` and `PSet` statements, an
//! access word in a procedure, `Me` with a blank before its member dot, a
//! trailing comma in `Close`, and `RaiseEvent` in a standard module.

mod common;
use common::{offset_of, parse, parse_kind};
use vb6_syntax::frontend::parser::ModuleKind;

const SYNTAX_ERROR: u32 = 0x9c70;
const METHOD_NOT_VALID_WITHOUT_OBJECT: u32 = 0xdec8;
const INVALID_ATTRIBUTE_IN_PROCEDURE: u32 = 0x9cd3;

fn in_sub(stmt: &str) -> String {
    format!("Sub S(a, b)\n    {stmt}\nEnd Sub\n")
}

fn first_error(stmt: &str) -> Option<(u32, u32)> {
    parse(&in_sub(stmt)).found.first().map(|d| (d.0, d.1))
}

#[test]
fn bare_graphics_statements_with_a_complete_shape_are_clean() {
    for stmt in [
        "Scale",
        "Scale (1, 2)-(3, 4)",
        "Scale (a, b) - (3, 4)",
        "Line (1, 2)-(3, 4)",
        "Line -(3, 4)",
        "Line Step (1, 2)-Step (3, 4)",
        "Line -Step (3, 4)",
        "Line (1, 2)-(3, 4), 5",
        "Line (1, 2)-(3, 4), , B",
        "Line (1, 2)-(3, 4), 5, BF",
        "Circle (1, 2), 3",
        "Circle Step (1, 2), 3",
        "Circle (1, 2), 3, , 5, 6, 7",
        "PSet (1, 2)",
        "PSet Step (1, 2)",
        "PSet (1, 2), 3",
        "Line (1, 2)-(3, 4): a = 1",
    ] {
        // A class module and a form have a surface of their own to draw on.
        for kind in [ModuleKind::Class, ModuleKind::Form] {
            assert_eq!(parse_kind(&in_sub(stmt), kind).found, vec![], "{stmt}");
        }
    }
}

#[test]
fn bare_graphics_statements_of_a_standard_module_have_no_object_to_draw_on() {
    for (stmt, word) in [
        ("Scale", "Scale"),
        ("Scale (1, 2)-(3, 4)", "Scale"),
        ("Line (1, 2)-(3, 4)", "Line"),
        ("Line -(3, 4), 5, BF", "Line"),
        ("Circle (1, 2), 3", "Circle"),
        ("PSet Step (1, 2), 3", "PSet"),
        ("Print 1; 2", "Print"),
        ("Print", "Print"),
        ("? 1", "?"),
        ("?", "?"),
        ("Line (1, 2)-(3, 4): a = 1", "Line"),
        ("If a Then PSet (1, 2) Else b = 1", "PSet"),
    ] {
        let src = in_sub(stmt);
        let found = parse(&src).found;
        assert_eq!(
            found,
            vec![(METHOD_NOT_VALID_WITHOUT_OBJECT, offset_of(&src, word, 0), word.len() as u32)],
            "{stmt}"
        );
    }
    // Printing to a file needs no surface.
    for stmt in ["Print #1, 2", "? #1, 2", "Line Input #1, a"] {
        assert_eq!(parse(&in_sub(stmt)).found, vec![], "{stmt}");
    }
}

#[test]
fn a_bracketed_graphics_word_that_starts_a_statement_is_the_statement() {
    for stmt in ["[Print] 1", "[Circle] (1, 2), 3", "[Line] (1, 2)-(3, 4)", "[PSet] (1, 2)", "[Scale]"] {
        let src = in_sub(stmt);
        assert_eq!(
            parse(&src).found,
            vec![(METHOD_NOT_VALID_WITHOUT_OBJECT, offset_of(&src, "[", 0), stmt.find(']').unwrap() as u32 + 1)],
            "{stmt}"
        );
        assert_eq!(parse_kind(&src, ModuleKind::Class).found, vec![], "{stmt}");
    }
}

#[test]
fn bare_circle_pset_and_scale_name_nothing_but_the_statement() {
    for stmt in ["Circle = 1", "Circle.x = 1", "PSet = 1", "PSet!x = 1", "Scale = 1", "Scale.x", "Circle(1) = 2"] {
        let src = in_sub(stmt);
        let word = stmt.split([' ', '.', '!', '(']).next().unwrap();
        assert_eq!(
            parse_kind(&src, ModuleKind::Class).found,
            vec![(SYNTAX_ERROR, offset_of(&src, word, 0), word.len() as u32)],
            "{stmt}"
        );
    }
    // `Line` is not reserved: it can be a variable.
    for stmt in ["Line = 1", "Line.x = 1"] {
        assert_eq!(parse_kind(&in_sub(stmt), ModuleKind::Class).found, vec![], "{stmt}");
    }
}

#[test]
fn bare_graphics_statements_with_a_broken_shape_are_syntax_errors_at_the_word() {
    for stmt in [
        "Scale (1, 2)",
        "Scale (1, 2)-",
        "Scale -(3, 4)",
        "Scale (1)",
        "Scale (1, 2, 3)",
        "Scale Step (1, 2)-(3, 4)",
        "Line (1, 2)",
        "Line (1, 2)-",
        "Line -1",
        "Line (1, 2)-(3, 4), 5, B, 6",
        "Line (1, 2)-(3, 4),",
        "Line (1, 2)-(3, 4), ,",
        "Line (1, 2), (3, 4)",
        "Circle (1, 2)",
        "Circle (1, 2),",
        "Circle (1, 2), , 4",
        "Circle (1, 2), 3,",
        "Circle -(1, 2), 3",
        "PSet (1, 2),",
        "PSet (1, 2), 3, 4",
        "PSet (1, 2)-(3, 4)",
    ] {
        let src = in_sub(stmt);
        let word = stmt.split(' ').next().unwrap();
        assert_eq!(parse(&src).found[0], (SYNTAX_ERROR, offset_of(&src, word, 0), word.len() as u32), "{stmt}");
    }
}

#[test]
fn graphics_methods_of_an_object_are_not_held_to_the_bare_shapes() {
    for stmt in ["a.Line (1, 2)-(3, 4)", "a.Circle (1, 2), 3", "a.PSet (1, 2)", "a.Scale (1, 2)-(3, 4)", "a.Scale"] {
        assert_eq!(parse(&in_sub(stmt)).found, vec![], "{stmt}");
    }
}

#[test]
fn an_access_word_cannot_start_a_statement_in_a_procedure() {
    for word in ["Public", "Private", "Global"] {
        let src = in_sub(&format!("{word} x As Long"));
        assert_eq!(parse(&src).found[0], (INVALID_ATTRIBUTE_IN_PROCEDURE, offset_of(&src, word, 0), word.len() as u32), "{word}");
    }
    assert_eq!(parse(&in_sub("Dim x As Long")).found, vec![]);
    assert_eq!(parse(&in_sub("Static x As Long")).found, vec![]);
}

#[test]
fn me_and_its_member_dot_are_written_together() {
    for stmt in ["Me .a", "Me !a", "Me . a"] {
        let src = in_sub(stmt);
        let at = if stmt.contains('!') { "!a" } else { ".a" };
        let at = if stmt == "Me . a" { ". a" } else { at };
        assert_eq!(first_error(stmt), Some((SYNTAX_ERROR, offset_of(&src, at, 0))), "{stmt}");
    }
    for stmt in ["Me.a = 1", "Me. a = 1", "Me(1) = 2"] {
        assert_eq!(parse(&in_sub(stmt)).found, vec![], "{stmt}");
    }
}

#[test]
fn close_needs_a_channel_after_each_comma() {
    for stmt in ["Close a,", "Close #1,"] {
        let src = in_sub(stmt);
        let after = (src.find(stmt).unwrap() + stmt.len()) as u32;
        assert_eq!(first_error(stmt), Some((SYNTAX_ERROR, after)), "{stmt}");
    }
    for stmt in ["Close a, b", "Close #1, #2", "Close"] {
        assert_eq!(parse(&in_sub(stmt)).found, vec![], "{stmt}");
    }
}

#[test]
fn raise_event_is_not_a_statement_of_a_standard_module() {
    let src = in_sub("RaiseEvent E");
    assert_eq!(parse(&src).found[0], (SYNTAX_ERROR, offset_of(&src, "RaiseEvent", 0), 10));
}

#[test]
fn built_in_function_shapes_hold_inside_the_arguments_of_a_call_statement() {
    for stmt in ["Call f(UBound)", "Call f(UBound())", "Call f(String(1))", "Call f(StrComp(1))", "Call f(a, UBound.a)"] {
        assert!(!parse(&in_sub(stmt)).found.is_empty(), "{stmt}");
    }
    for stmt in ["Call f(UBound(a))", "Call f(String(3, \"x\"))", "Call f(StrComp(a, b))", "f UBound(a, 1)"] {
        assert_eq!(parse(&in_sub(stmt)).found, vec![], "{stmt}");
    }
}

#[test]
fn a_mid_assignment_leaves_no_argument_out() {
    for stmt in ["Mid(, 1) = \"a\"", "Mid(a,, 2) = \"a\"", "Mid(a, 1, ) = \"a\"", "MidB$(, 1, 2) = \"a\""] {
        let src = in_sub(stmt);
        assert!(!parse(&src).found.is_empty(), "{stmt}");
    }
    for stmt in ["Mid(a, 1) = \"a\"", "Mid$(a, 1, 2) = \"a\"", "MidB(a, 1) = \"a\""] {
        assert_eq!(parse(&in_sub(stmt)).found, vec![], "{stmt}");
    }
}

#[test]
fn rem_after_a_statement_in_a_single_line_if_depends_on_the_statement() {
    // These statements read their operand list up to the end of the line: Rem is no end.
    for stmt in ["Erase a", "Close #1", "Close #1, #2", "Input #1, b", "Call f", "Call f(1)"] {
        let src = in_sub(&format!("If a Then {stmt} Rem z"));
        let found = parse(&src).found;
        assert_eq!(found.len(), 1, "{stmt}: {found:?}");
        assert_eq!(found[0].1, offset_of(&src, "Rem", 0), "{stmt}");
    }
    // These stop at Rem like at a comment mark.
    for stmt in ["b = 1", "Exit Sub", "Stop", "End", "Print #1, 1", "Write #1, 1", "Debug.Print 1", "GoTo L"] {
        let src = format!("Sub S(a, b)\n    If a Then {stmt} Rem z\nL:\nEnd Sub\n");
        assert_eq!(parse(&src).found, vec![], "{stmt}");
    }
}
