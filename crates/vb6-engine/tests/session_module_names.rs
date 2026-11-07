//! The name an `Attribute VB_Name` line gives a module has to be one the project can take: the
//! checks that need the project (its name, the libraries it references, the other modules) and
//! those that need only the name run when the session links the project.

use vb6_syntax::frontend::diagnostics::{
    ERR_MAIN_NOT_ALLOWED_AS_NAME, ERR_NAME_CONFLICTS, ERR_NOT_LEGAL_OBJECT_NAME,
};
use vb6_engine::session::{Library, Session, SessionConfig};

const TAIL: &str = "Option Explicit\nPublic Sub Run()\nEnd Sub\n";

fn config() -> SessionConfig {
    let libraries = Library::bundled().into_iter().filter(|l| l.name != "stdole").collect();
    SessionConfig { project_name: Some("p".to_string()), libraries, conditional_compilation: None }
}

/// The module with `header` lines before an ordinary body, as a file named `path`.
fn file(path: &str, header: &str) -> (String, Vec<u8>) {
    (path.to_string(), format!("{header}{TAIL}").into_bytes())
}

/// `(code, message, start, len)` of every diagnostic of module `index`.
fn diagnostics(files: Vec<(String, Vec<u8>)>, index: usize) -> Vec<(u32, String, u32, u32)> {
    let session = Session::with_config(files, config());
    session
        .diagnostics(index)
        .iter()
        .map(|d| (d.code, d.message().expect("a known message"), d.span.start, d.span.len))
        .collect()
}

fn one(path: &str, header: &str) -> Vec<(u32, String, u32, u32)> {
    diagnostics(vec![file(path, header)], 0)
}

/// The offset of the value written in `header`: the quote that opens the first string.
fn value_at(header: &str) -> u32 {
    header.find('"').expect("a string value") as u32
}

#[test]
fn a_legal_unused_name_is_accepted() {
    for path in ["M.bas", "C.cls", "F.frm"] {
        assert_eq!(one(path, "Attribute VB_Name = \"Foo\"\n"), vec![], "{path}");
    }
}

#[test]
fn a_standard_module_name_that_is_not_legal_is_reported_with_the_name() {
    for name in ["a b", "1abc", "_a", "a.b", "a!b", "a$", "a-b", "ab ", " a"] {
        let header = format!("Attribute VB_Name = \"{name}\"\n");
        assert_eq!(
            one("M.bas", &header),
            vec![(
                ERR_NOT_LEGAL_OBJECT_NAME,
                format!("Not a legal object name: '{name}'"),
                value_at(&header),
                name.len() as u32 + 2,
            )],
            "{name:?}"
        );
    }
}

#[test]
fn a_standard_module_name_may_use_letters_digits_underscores_and_characters_above_ascii() {
    for name in ["a", "A1", "a_", "A__", "a1b2", "\u{e4}", "a\u{e9}b", "\u{a0}", "a\u{a0}b"] {
        let header = format!("Attribute VB_Name = \"{name}\"\n");
        assert_eq!(one("M.bas", &header), vec![], "{name:?}");
    }
}

#[test]
fn a_standard_module_name_has_at_most_127_characters() {
    let legal = "a".repeat(127);
    assert_eq!(one("M.bas", &format!("Attribute VB_Name = \"{legal}\"\n")), vec![]);
    let long = "a".repeat(128);
    let header = format!("Attribute VB_Name = \"{long}\"\n");
    assert_eq!(
        one("M.bas", &header),
        vec![(ERR_NOT_LEGAL_OBJECT_NAME, format!("Not a legal object name: '{long}'"), value_at(&header), 130)]
    );
}

#[test]
fn class_and_designer_modules_are_not_held_to_the_legal_name_rule() {
    for path in ["C.cls", "F.frm", "U.ctl", "P.pag"] {
        for name in ["a b", "1abc", "a.b", "ab "] {
            assert_eq!(one(path, &format!("Attribute VB_Name = \"{name}\"\n")), vec![], "{path} {name:?}");
        }
    }
}

#[test]
fn escapes_are_applied_before_the_name_is_judged() {
    // A line feed and a carriage return make the name illegal; a backslash before any other
    // character is dropped.
    for (written, shown) in [("a\\nb", "a\nb"), ("a\\rb", "a\rb"), ("a\\\\b", "a\\b")] {
        let header = format!("Attribute VB_Name = \"{written}\"\n");
        let found = one("M.bas", &header);
        assert_eq!(found.len(), 1, "{written}");
        assert_eq!(found[0].0, ERR_NOT_LEGAL_OBJECT_NAME, "{written}");
        assert_eq!(found[0].1, format!("Not a legal object name: '{shown}'"), "{written}");
    }
    for written in ["a\\qb", "a\\", "\\Foo", "Fo\\o", "a\\tb", "a\\0b", "a\\x41"] {
        assert_eq!(one("M.bas", &format!("Attribute VB_Name = \"{written}\"\n")), vec![], "{written}");
    }
    // The escaped name is the one compared with the project and the libraries.
    let header = "Attribute VB_Name = \"V\\BA\"\n";
    assert_eq!(one("M.bas", header)[0].0, ERR_NAME_CONFLICTS);
}

#[test]
fn an_empty_name_or_one_with_a_bracket_conflicts_in_every_kind_of_module() {
    for path in ["M.bas", "C.cls", "F.frm"] {
        for name in ["", "[", "a[b", "a]b", "[a]", "a b[c"] {
            let header = format!("Attribute VB_Name = \"{name}\"\n");
            assert_eq!(
                one(path, &header),
                vec![(
                    ERR_NAME_CONFLICTS,
                    "Name conflicts with existing module, project, or object library".to_string(),
                    value_at(&header),
                    name.len() as u32 + 2,
                )],
                "{path} {name:?}"
            );
        }
    }
    // A lone backslash is empty once its escape is applied.
    assert_eq!(one("M.bas", "Attribute VB_Name = \"\\\"\n")[0].0, ERR_NAME_CONFLICTS);
}

#[test]
fn the_project_name_and_the_library_names_conflict_in_every_kind_of_module() {
    for path in ["M.bas", "C.cls", "F.frm"] {
        for name in ["p", "P", "VBA", "vba", "VB", "Vb", "VBRUN", "vbrun"] {
            let found = one(path, &format!("Attribute VB_Name = \"{name}\"\n"));
            assert_eq!(found.len(), 1, "{path} {name}");
            assert_eq!(found[0].0, ERR_NAME_CONFLICTS, "{path} {name}");
        }
    }
    // A library the project does not reference does not.
    assert_eq!(one("M.bas", "Attribute VB_Name = \"stdole\"\n"), vec![]);
}

#[test]
fn main_is_refused_for_standard_and_class_modules_only() {
    for name in ["Main", "main", "MAIN"] {
        let header = format!("Attribute VB_Name = \"{name}\"\n");
        for path in ["M.bas", "C.cls"] {
            assert_eq!(
                one(path, &header),
                vec![(
                    ERR_MAIN_NOT_ALLOWED_AS_NAME,
                    "Main can't be module, type, project, or form name".to_string(),
                    value_at(&header),
                    name.len() as u32 + 2,
                )],
                "{path} {name}"
            );
        }
        assert_eq!(one("F.frm", &header), vec![], "{name}");
    }
    assert_eq!(one("M.bas", "Attribute VB_Name = \"Main1\"\n"), vec![]);
}

#[test]
fn a_name_that_is_not_legal_is_judged_before_it_is_compared() {
    // Trailing blank: not legal, although it spells a library, the project and `Main`.
    for name in ["VBA ", "p ", "Main ", " Main", "VBA!"] {
        let found = one("M.bas", &format!("Attribute VB_Name = \"{name}\"\n"));
        assert_eq!(found.len(), 1, "{name:?}");
        assert_eq!(found[0].0, ERR_NOT_LEGAL_OBJECT_NAME, "{name:?}");
    }
    // A class module is not judged for legality, so the same spellings are compared.
    assert_eq!(one("C.cls", "Attribute VB_Name = \"VBA \"\n"), vec![]);
}

#[test]
fn a_name_another_module_has_conflicts_in_standard_and_class_modules() {
    let files = vec![file("A.bas", "Attribute VB_Name = \"Alpha\"\n"), file("B.bas", "Attribute VB_Name = \"alpha\"\n")];
    for index in [0, 1] {
        let found = diagnostics(files.clone(), index);
        assert_eq!(found.len(), 1, "module {index}");
        assert_eq!(found[0].0, ERR_NAME_CONFLICTS, "module {index}");
    }
    // The other module's name can be its file name.
    let files = vec![file("Beta.bas", ""), file("B.cls", "Attribute VB_Name = \"BETA\"\n")];
    assert_eq!(diagnostics(files.clone(), 0), vec![]);
    assert_eq!(diagnostics(files, 1).len(), 1);
    // A designer takes its name from its designer section: another module's name is no conflict.
    let files = vec![file("Alpha.bas", ""), file("F.frm", "Attribute VB_Name = \"Alpha\"\n")];
    assert_eq!(diagnostics(files, 1), vec![]);
    // A module's own name is not another module's.
    assert_eq!(diagnostics(vec![file("Alpha.bas", "Attribute VB_Name = \"Alpha\"\n")], 0), vec![]);
}

#[test]
fn a_later_name_replaces_the_failure_of_an_earlier_one() {
    // `Foo` then `foo`: its own earlier name is no conflict.
    assert_eq!(one("M.bas", "Attribute VB_Name = \"Foo\"\nAttribute VB_Name = \"foo\"\n"), vec![]);
    // An illegal name, a name that is `Main`, and a name another module has are replaced.
    let ok = "Attribute VB_Name = \"Foo\"\n";
    for first in ["1a", "Main", "Other"] {
        let header = format!("Attribute VB_Name = \"{first}\"\n{ok}");
        let files = vec![file("M.bas", &header), file("O.bas", "Attribute VB_Name = \"Other\"\n")];
        assert_eq!(diagnostics(files, 0), vec![], "{first}");
    }
    // The last failure is the one left.
    let header = "Attribute VB_Name = \"1a\"\nAttribute VB_Name = \"1b\"\n";
    let found = one("M.bas", header);
    assert_eq!(found.len(), 1);
    assert_eq!(found[0].1, "Not a legal object name: '1b'");
    assert_eq!(found[0].2, header.rfind('"').unwrap() as u32 - 3);
    let header = "Attribute VB_Name = \"1a\"\nAttribute VB_Name = \"Main\"\n";
    assert_eq!(one("M.bas", header)[0].0, ERR_MAIN_NOT_ALLOWED_AS_NAME);
    let header = "Attribute VB_Name = \"Main\"\nAttribute VB_Name = \"1a\"\n";
    assert_eq!(one("M.bas", header)[0].0, ERR_NOT_LEGAL_OBJECT_NAME);
    // A value that is not a string gives no name: the failure before it stays.
    let header = "Attribute VB_Name = \"1a\"\nAttribute VB_Name = 5\n";
    assert_eq!(one("M.bas", header)[0].1, "Not a legal object name: '1a'");
    let header = "Attribute VB_Name = \"Main\"\nAttribute VB_Name = True\n";
    assert_eq!(one("M.bas", header)[0].0, ERR_MAIN_NOT_ALLOWED_AS_NAME);
}

#[test]
fn a_conflict_with_the_project_or_a_library_or_an_empty_name_stays() {
    for first in ["VBA", "p", "", "a[b"] {
        for second in ["Foo", "1b", "Main", "VB", ""] {
            let header = format!("Attribute VB_Name = \"{first}\"\nAttribute VB_Name = \"{second}\"\n");
            let found = one("M.bas", &header);
            assert_eq!(found.len(), 1, "{first:?} then {second:?}");
            assert_eq!(found[0].0, ERR_NAME_CONFLICTS, "{first:?} then {second:?}");
            assert_eq!(found[0].2, header.find('"').unwrap() as u32, "{first:?} then {second:?}");
        }
    }
}

#[test]
fn a_name_is_set_wherever_the_line_stands_in_a_standard_module_without_a_member() {
    // After a procedure, between the members of an `Enum`, a `Type` and in a procedure.
    for body in [
        "Public Sub Q()\nEnd Sub\nAttribute VB_Name = \"VBA\"\n",
        "Public Enum E\nA\nAttribute VB_Name = \"VBA\"\nEnd Enum\n",
        "Public Type T\nA As Long\nAttribute VB_Name = \"VBA\"\nEnd Type\n",
        "Public Sub Q()\nAttribute VB_Name = \"VBA\"\nEnd Sub\n",
    ] {
        let found = diagnostics(vec![("M.bas".to_string(), format!("Option Explicit\n{body}").into_bytes())], 0);
        assert_eq!(found.len(), 1, "{body}");
        assert_eq!(found[0].0, ERR_NAME_CONFLICTS, "{body}");
    }
    // With a member before the word, the name of the module is not set.
    let body = "Public Sub Q()\nAttribute Q.VB_Name = \"VBA\"\nEnd Sub\n";
    assert_eq!(diagnostics(vec![("M.bas".to_string(), format!("Option Explicit\n{body}").into_bytes())], 0), vec![]);
}

#[test]
fn a_value_that_is_not_a_string_sets_no_name() {
    for value in ["5", "-5", "True", "x", ""] {
        assert_eq!(one("M.bas", &format!("Attribute VB_Name = {value}\n")), vec![], "{value:?}");
    }
}
