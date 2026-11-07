//! Tests for semantic diagnostics and binder errors.

use vb6_core::session::Session;

const DUPLICATE_DECLARATION: u32 = 40095;

#[test]
fn repeated_local_declaration_is_reported_at_the_second_name() {
    let src = "Sub Run()\n    Dim n As Long\n    Dim n As Long\nEnd Sub\n";
    let session = Session::from_sources(vec![("m.bas".to_string(), src.as_bytes().to_vec())]);
    let found: Vec<(u32, u32, u32)> =
        session.diagnostics(0).iter().map(|d| (d.code, d.span.start, d.span.len)).collect();
    let second = src.rfind("Dim n").unwrap() + "Dim ".len();
    assert_eq!(found, vec![(DUPLICATE_DECLARATION, second as u32, 1)]);
}
