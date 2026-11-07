//! The project's conditional-compilation arguments decide which code a
//! module contains.

use vb6_engine::session::{Session, SessionConfig};

const SRC: &[u8] = b"#If MyFlag Then\nSub Enabled()\nEnd Sub\n#Else\nSub Disabled()\nEnd Sub\n#End If\nSub Run()\n    Enabled\nEnd Sub\n";

fn codes(s: &Session) -> Vec<(u32, u32, u32)> {
    s.diagnostics(0).iter().map(|d| (d.code, d.span.start, d.span.len)).collect()
}

fn config(args: Option<&str>) -> SessionConfig {
    SessionConfig { conditional_compilation: args.map(String::from), ..SessionConfig::default() }
}

#[test]
fn a_project_argument_selects_the_branch() {
    // Without the argument `MyFlag` is Empty, so `Enabled` does not exist.
    let off = Session::with_config(vec![("m.bas".into(), SRC.to_vec())], config(None));
    assert_eq!(codes(&off).len(), 1);
    // With it, the first branch is compiled and `Enabled` resolves.
    let on = Session::with_config(vec![("m.bas".into(), SRC.to_vec())], config(Some("MyFlag = 1")));
    assert_eq!(codes(&on), vec![]);
}

#[test]
fn changing_the_arguments_rereads_every_module() {
    let mut s = Session::with_config(vec![("m.bas".into(), SRC.to_vec())], config(None));
    assert_eq!(codes(&s).len(), 1);
    s.set_config(config(Some("MyFlag = -1 : Other = 5")));
    assert_eq!(codes(&s), vec![]);
    s.set_config(config(Some("MyFlag = 0")));
    assert_eq!(codes(&s).len(), 1);
}

#[test]
fn an_edited_file_is_read_with_the_project_arguments() {
    let mut s = Session::with_config(vec![("m.bas".into(), b"Sub A()\nEnd Sub\n".to_vec())], config(Some("MyFlag = 1")));
    s.update_file("m.bas", SRC.to_vec());
    assert_eq!(codes(&s), vec![]);
}

#[test]
fn arguments_that_are_not_constant_declarations_define_nothing() {
    let s = Session::with_config(vec![("m.bas".into(), SRC.to_vec())], config(Some("MyFlag = 1 : Other = \"x\"")));
    assert!(!s.conditional_compilation_is_valid());
    assert_eq!(codes(&s).len(), 1);
    let s = Session::with_config(vec![("m.bas".into(), SRC.to_vec())], config(Some("MyFlag = 1")));
    assert!(s.conditional_compilation_is_valid());
}

#[test]
fn argument_text_follows_the_project_format() {
    let valid = |args: &str| {
        Session::with_config(vec![("m.bas".into(), SRC.to_vec())], config(Some(args))).conditional_compilation_is_valid()
    };
    // A sign alone is the value 0.
    assert!(valid("MyFlag = -"));
    assert!(valid("MyFlag = + : Other = 000001"));
    assert!(!valid("_MyFlag = 1"));
    assert!(!valid("MyFlag = 000001"));
    assert!(!valid("MyFlag = 1 : myflag = 2"));
    assert!(valid(&format!("{}=1", "A".repeat(1021))));
    assert!(!valid(&format!("{}=1 ", "A".repeat(1021))));
    // The sign alone gives 0, so the second branch is compiled.
    let s = Session::with_config(vec![("m.bas".into(), SRC.to_vec())], config(Some("MyFlag = -")));
    assert_eq!(codes(&s).len(), 1);
}
