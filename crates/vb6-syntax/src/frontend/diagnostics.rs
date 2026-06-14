//! Diagnostic (syntax error) collector for the VB6 parser.
//!
//! On a syntax error the parser records a `Diagnostic` and *continues* parsing,
//! reproducing VB6's "recover and continue" posture: each error pushes a
//! `Diagnostic` and parsing resumes.  No `longjmp`, no global flags.
//!
//! The `code` field matches VB6's 16-bit error codes (e.g. `0x9c6f` =
//! "Expected expression").  Unknown codes are used as-is.

use crate::frontend::token::Span;

/// A single recoverable parse diagnostic.
#[derive(Debug, Clone, PartialEq)]
pub struct Diagnostic {
    /// VB6 error code = VB6 string-resource id (e.g. `0x9c6f` = "Expected:
    /// <various>", `0xdee1` = "Only valid in object module"). See
    /// [`message_for_code`]. `0` = synthetic (no VB6 equivalent).
    pub code: u32,
    /// Source span where the error was detected.
    pub span: Span,
    /// For `0x9c6f` ("Expected: X") errors, the name of the expected token
    /// (e.g. `"Then"`, `")"`, `"Identifier"`). `None` for generic errors.
    pub label: Option<&'static str>,
}

impl Diagnostic {
    pub fn new(code: u32, span: Span) -> Self {
        Self { code, span, label: None }
    }

    pub fn with_label(code: u32, span: Span, label: &'static str) -> Self {
        Self { code, span, label: Some(label) }
    }

    /// Human-readable text for this diagnostic.
    ///
    /// For `0x9c6f` with a label, returns `"Expected: <label>"`.
    /// For other codes, returns the VB6 message text, or `None` if unknown.
    pub fn message(&self) -> Option<String> {
        if self.code == 0x9c6f {
            return Some(match self.label {
                Some(label) => format!("Expected: {label}"),
                None => "Expected: <token>".to_string(),
            });
        }
        message_for_code(self.code).map(|s| s.to_string())
    }
}

/// Map a diagnostic code to its VB6 message text (without label substitution).
///
/// Prefer [`Diagnostic::message`] when you have a `Diagnostic`, since it
/// incorporates the `label` field for `0x9c6f` errors.
pub fn message_for_code(code: u32) -> Option<&'static str> {
    Some(match code {
        0x9c6f => "Expected: <token>",                      // 40047
        0xdee1 => "Only valid in object module",            // 57057
        0x9caf => "Variable not defined",                   // 40111
        0x9c9f => "Duplicate declaration in current scope", // 40095
        0x23 => "Sub or Function not defined",              // 35
        _ => return None,
    })
}

/// Accumulated diagnostics for one parse run.
///
/// Implements the "record error + continue" pattern.
/// After parsing, callers check [`Diagnostics::has_errors`].
#[derive(Debug, Default, Clone)]
pub struct Diagnostics {
    items: Vec<Diagnostic>,
}

impl Diagnostics {
    pub fn new() -> Self {
        Self::default()
    }

    /// Record a diagnostic at `span` with the given VB6 error `code`.
    pub fn push(&mut self, code: u32, span: Span) {
        self.items.push(Diagnostic::new(code, span));
    }

    /// Record an `0x9c6f` ("Expected: X") diagnostic with a specific token label.
    pub fn push_labeled(&mut self, code: u32, span: Span, label: &'static str) {
        self.items.push(Diagnostic::with_label(code, span, label));
    }

    pub fn has_errors(&self) -> bool {
        !self.items.is_empty()
    }

    pub fn items(&self) -> &[Diagnostic] {
        &self.items
    }

    pub fn into_items(self) -> Vec<Diagnostic> {
        self.items
    }
}

#[cfg(test)]
mod tests {
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
}
