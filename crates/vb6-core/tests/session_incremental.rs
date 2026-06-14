//! Incremental Session update tests: update_file / remove_file + relink.

use vb6_core::sema::binder::{ERR_SUB_OR_FUNCTION_NOT_DEFINED, ERR_VARIABLE_NOT_DEFINED};
use vb6_core::session::Session;

fn at(src: &str, needle: &str) -> u32 {
    (src.find(needle).expect("needle not found") + 1) as u32
}

#[test]
fn adding_a_module_resolves_a_previously_undefined_call() {
    let mod1 = "Sub Main()\n    Greet\nEnd Sub\n";
    let mut s = Session::from_sources(vec![("Mod1.bas".into(), mod1.as_bytes().to_vec())]);

    // Greet is undefined initially → flagged.
    assert!(s
        .diagnostics(0)
        .iter()
        .any(|d| d.code == ERR_SUB_OR_FUNCTION_NOT_DEFINED as u32));

    // Add a module that defines Greet publicly.
    s.update_file("Mod0.bas", b"Public Sub Greet()\nEnd Sub\n".to_vec());

    let m1 = s.module_of("Mod1.bas").unwrap();
    assert!(
        s.diagnostics(m1).iter().all(|d| d.code != ERR_SUB_OR_FUNCTION_NOT_DEFINED as u32),
        "call should resolve cross-module after the definition is added"
    );
    let def = s.definition(m1, at(mod1, "Greet")).expect("definition");
    assert_eq!(s.module_path(def.module), Some("Mod0.bas"));
}

#[test]
fn removing_a_module_breaks_cross_module_resolution() {
    let mut s = Session::from_sources(vec![
        ("Mod0.bas".into(), b"Public Sub Greet()\nEnd Sub\n".to_vec()),
        ("Mod1.bas".into(), b"Sub Main()\n    Greet\nEnd Sub\n".to_vec()),
    ]);
    let m1 = s.module_of("Mod1.bas").unwrap();
    assert!(s.diagnostics(m1).iter().all(|d| d.code != ERR_SUB_OR_FUNCTION_NOT_DEFINED as u32));

    assert!(s.remove_file("Mod0.bas"));
    assert!(!s.remove_file("Mod0.bas")); // already gone

    let m1 = s.module_of("Mod1.bas").unwrap(); // index may have shifted
    assert!(
        s.diagnostics(m1).iter().any(|d| d.code == ERR_SUB_OR_FUNCTION_NOT_DEFINED as u32),
        "call should be undefined again after the definition is removed"
    );
}

#[test]
fn repeated_identical_update_does_not_duplicate_diagnostics() {
    let src = b"Option Explicit\nSub F()\n    x = 1\nEnd Sub\n";
    let mut s = Session::from_sources(vec![("M.bas".into(), src.to_vec())]);
    let count = |s: &Session| {
        s.diagnostics(0).iter().filter(|d| d.code == ERR_VARIABLE_NOT_DEFINED as u32).count()
    };
    assert_eq!(count(&s), 1);
    s.update_file("M.bas", src.to_vec());
    s.update_file("M.bas", src.to_vec());
    assert_eq!(count(&s), 1, "relink resets from raw — no diagnostic duplication");
}

#[test]
fn update_reflects_new_content() {
    let mut s = Session::from_sources(vec![("M.bas".into(), b"Public gOld As Long\n".to_vec())]);
    assert_eq!(s.workspace_symbols("gOld").len(), 1);

    s.update_file("M.bas", b"Public gNew As String\n".to_vec());
    assert!(s.workspace_symbols("gOld").is_empty());
    assert_eq!(s.workspace_symbols("gNew").len(), 1);
}
