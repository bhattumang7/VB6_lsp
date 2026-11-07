use super::*;

#[test]
fn known_codes_map_to_text() {
    assert_eq!(message_for_code(0x9caf), Some("Variable not defined"));
    assert_eq!(message_for_code(0x23), Some("Sub or Function not defined"));
    assert_eq!(message_for_code(0x9c9f), Some("Duplicate declaration in current scope"));
    assert_eq!(message_for_code(0xdee1), Some("Only valid in object module"));
    assert!(message_for_code(0x9c6f).is_some());
    assert_eq!(message_for_code(0xffff), None);
    assert_eq!(
        Diagnostic::new(0x9caf, Span::DUMMY).message(),
        Some("Variable not defined".to_string())
    );
    // Labeled 0x9c6f should produce "Expected: <label>"
    assert_eq!(
        Diagnostic::with_label(0x9c6f, Span::DUMMY, "Then").message(),
        Some("Expected: Then".to_string())
    );
    // Unlabeled 0x9c6f falls back to generic placeholder
    assert_eq!(
        Diagnostic::new(0x9c6f, Span::DUMMY).message(),
        Some("Expected: <token>".to_string())
    );
}
