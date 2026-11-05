//! The exact shapes of bare `Scale`, `Line`, `Circle` and `PSet` statements, an
//! access word in a procedure, `Me` with a blank before its member dot, a
//! trailing comma in `Close`, and `RaiseEvent` in a standard module.

mod common;
use common::{offset_of, parse};

const SYNTAX_ERROR: u32 = 0x9c70;
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
        assert_eq!(parse(&in_sub(stmt)).found, vec![], "{stmt}");
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
