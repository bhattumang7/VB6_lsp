//! The full-engine side of the comparison reports name-resolution errors that the parser alone cannot.

use oracle_diff::oracle::ModuleKind;
use oracle_diff::ours::{diagnostics, full_diagnostics, OurDiagnostic};

const UNDEFINED: &[u8] = b"Option Explicit\r\nPublic Sub A()\r\n    x = 1\r\nEnd Sub\r\n";

#[test]
fn undefined_variable_is_reported_by_the_full_engine_only() {
    assert_eq!(diagnostics(UNDEFINED, ModuleKind::Bas), Vec::<OurDiagnostic>::new());
    assert_eq!(
        full_diagnostics(UNDEFINED, ModuleKind::Bas),
        vec![OurDiagnostic { code: 0x9caf, line: 3, message: "Variable not defined".to_string() }]
    );
}

#[test]
fn valid_module_has_no_diagnostics_in_either_mode() {
    let source = b"Public Sub A()\r\n    Dim n As Long\r\n    n = 1\r\nEnd Sub\r\n";
    assert_eq!(diagnostics(source, ModuleKind::Bas), Vec::<OurDiagnostic>::new());
    assert_eq!(full_diagnostics(source, ModuleKind::Bas), Vec::<OurDiagnostic>::new());
}

#[test]
fn parse_errors_are_kept_in_the_full_engine() {
    let source = b"Public Sub A()
    x = 1 +* 2
End Sub
";
    let expected = vec![
        OurDiagnostic { code: 40047, line: 2, message: "Expected: <token>".to_string() },
        OurDiagnostic { code: 40047, line: 2, message: "Expected: end of statement".to_string() },
    ];
    assert_eq!(diagnostics(source, ModuleKind::Bas), expected);
    assert_eq!(full_diagnostics(source, ModuleKind::Bas), expected);
}
