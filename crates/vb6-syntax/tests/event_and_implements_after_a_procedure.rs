//! `Event` and `Implements` after the end of a procedure: a standard module refuses both as a
//! syntax error at the first word of the line; a class module takes `Implements`, and an
//! `Event` there is text after the procedure's end.

mod common;
use common::{offset_of, parse, parse_kind};
use vb6_syntax::frontend::parser::ModuleKind;

const SYNTAX_ERROR: u32 = 0x9c70;
const ONLY_COMMENTS_AFTER_END: u32 = 0x9c91;

fn after_procedure(line: &str) -> String {
    format!("Sub S()\nEnd Sub\n{line}\n")
}

#[test]
fn a_standard_module_refuses_event_and_implements_after_a_procedure() {
    for (line, len) in [("Event E()", 5), ("Public Event E()", 6), ("Implements Foo", 10)] {
        let src = after_procedure(line);
        assert_eq!(parse(&src).found, vec![(SYNTAX_ERROR, offset_of(&src, line, 0), len)], "{line}");
    }
}

#[test]
fn a_class_module_takes_implements_after_a_procedure() {
    let src = after_procedure("Implements Foo");
    assert_eq!(parse_kind(&src, ModuleKind::Class).found, vec![]);
}

#[test]
fn an_event_in_a_class_module_after_a_procedure_is_text_after_the_end() {
    for line in ["Event E()", "Public Event E()"] {
        let src = after_procedure(line);
        let found = parse_kind(&src, ModuleKind::Class).found;
        assert_eq!(found.len(), 1, "{line}: {found:?}");
        assert_eq!(found[0].0, ONLY_COMMENTS_AFTER_END, "{line}");
        assert_eq!(found[0].1, offset_of(&src, line, 0), "{line}");
    }
}
