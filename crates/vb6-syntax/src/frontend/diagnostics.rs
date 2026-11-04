//! Diagnostic (syntax error) collector for the VB6 parser.
//!
//! On a syntax error the parser records a `Diagnostic` and *continues* parsing,
//! reproducing VB6's "recover and continue" posture: each error pushes a
//! `Diagnostic` and parsing resumes.  No `longjmp`, no global flags.
//!
//! The `code` field matches VB6's 16-bit error codes (e.g. `0x9c6f` =
//! "Expected expression").  Unknown codes are used as-is.

use crate::frontend::token::Span;

/// VB6 compile-error codes (code == the VB6 string-resource id, shown in decimal).
pub const ERR_EXPECTED_END_SUB: u32 = 0x9c8f; // 40079
pub const ERR_EXPECTED_END_FUNCTION: u32 = 0x9c90; // 40080
pub const ERR_AFTER_END_OF_PROCEDURE: u32 = 0x9c91; // 40081
pub const ERR_BLOCK_IF_WITHOUT_END_IF: u32 = 0x9c92; // 40082
pub const ERR_WHILE_WITHOUT_WEND: u32 = 0x9c93; // 40083
pub const ERR_DO_WITHOUT_LOOP: u32 = 0x9c94; // 40084
pub const ERR_FOR_WITHOUT_NEXT: u32 = 0x9c95; // 40085
pub const ERR_ELSE_WITHOUT_IF: u32 = 0x9c96; // 40086
pub const ERR_END_IF_WITHOUT_BLOCK_IF: u32 = 0x9c97; // 40087
pub const ERR_NEXT_WITHOUT_FOR: u32 = 0x9c98; // 40088
pub const ERR_LOOP_WITHOUT_DO: u32 = 0x9c99; // 40089
pub const ERR_WEND_WITHOUT_WHILE: u32 = 0x9c9a; // 40090
pub const ERR_EXIT_FOR_OUTSIDE_FOR: u32 = 0x9c9b; // 40091
pub const ERR_EXIT_DO_OUTSIDE_DO: u32 = 0x9c9c; // 40092
pub const ERR_EXIT_FUNCTION_NOT_ALLOWED: u32 = 0x9ca2; // 40098
pub const ERR_EXIT_SUB_NOT_ALLOWED: u32 = 0x9ca3; // 40099
pub const ERR_CASE_ELSE_OUTSIDE_SELECT: u32 = 0x9ca9; // 40105
pub const ERR_CASE_WITHOUT_SELECT: u32 = 0x9caa; // 40106
pub const ERR_END_SELECT_WITHOUT_SELECT: u32 = 0x9cab; // 40107
pub const ERR_SELECT_WITHOUT_END_SELECT: u32 = 0x9cac; // 40108
pub const ERR_EXPECTED_END_PROPERTY: u32 = 0x9cbe; // 40126
pub const ERR_EXIT_PROPERTY_NOT_ALLOWED: u32 = 0x9cbf; // 40127
pub const ERR_END_WITH_WITHOUT_WITH: u32 = 0x9cc6; // 40134
pub const ERR_EXPECTED_END_WITH: u32 = 0x9cc7; // 40135
pub const ERR_INVALID_NEXT_VARIABLE: u32 = 0x9ccd; // 40141
pub const ERR_FOR_VARIABLE_IN_USE: u32 = 0x9cce; // 40142
pub const ERR_ARRAY_ALREADY_DIMENSIONED: u32 = 0x9c77; // 40055
pub const ERR_DUPLICATE_OPTION: u32 = 0x9c78; // 40056
pub const ERR_DEFTYPE_AFTER_DECLARATION: u32 = 0x9c79; // 40057
pub const ERR_DUPLICATE_DEFTYPE: u32 = 0x9c7a; // 40058
pub const ERR_OPTION_PRIVATE_IN_OBJECT_MODULE: u32 = 0xdec9; // 57033
pub const ERR_CC_WITHOUT_IF: u32 = 0xdebd; // 57021
pub const ERR_CC_DUPLICATE_ELSE: u32 = 0xdebe; // 57022
pub const ERR_CC_ELSEIF_AFTER_ELSE: u32 = 0xdebf; // 57023
pub const ERR_CC_IF_NOT_TERMINATED: u32 = 0xdec0; // 57024
pub const ERR_END_TYPE_OUTSIDE_TYPE: u32 = 0x9c75; // 40053
pub const ERR_STATEMENT_IN_TYPE: u32 = 0x9c76; // 40054
pub const ERR_TYPE_WITHOUT_MEMBERS: u32 = 0x9cfc; // 40188
pub const ERR_ENUM_WITHOUT_MEMBERS: u32 = 0xdf43; // 57155
pub const ERR_IDENTIFIER_TOO_LONG: u32 = 0x9c6d; // 40045
pub const ERR_TOO_MANY_DIMENSIONS: u32 = 0x9c80; // 40064
pub const ERR_INVALID_FIXED_STRING_LENGTH: u32 = 0x9cb0; // 40112
pub const ERR_MUST_BE_FIRST_ON_LINE: u32 = 0x9c71; // 40049
pub const ERR_INVALID_OUTSIDE_PROCEDURE: u32 = 0x9c73; // 40051
pub const ERR_SYNTAX_ERROR: u32 = 0x9c70; // 40048
pub const ERR_ARRAY_ARGUMENT_MUST_BE_BYREF: u32 = 0x9ca0; // 40096
pub const ERR_PROPERTY_ARGUMENT_REQUIRED: u32 = 0x9cbd; // 40125
pub const ERR_OPTIONAL_ARRAY: u32 = 0x9ccf; // 40143
pub const ERR_PARAM_ARRAY_NOT_ARRAY: u32 = 0x9cd0; // 40144
pub const ERR_INVALID_EVENT_NAME: u32 = 0xdefb; // 57083
pub const ERR_PARAM_ARRAY_WITH_OPTIONAL: u32 = 0xdf00; // 57088
pub const ERR_LABEL_NOT_DEFINED: u32 = 0x9ca4; // 40100

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
        0x1cd => "Method or data member not found",         // 461
        0x9c74 => "Invalid inside procedure",               // 40052
        0xdeff => "Invalid inside Enum",                    // 57087
        0xdefa => "Assignment to constant not permitted",   // 57082
        54 => "Bad file mode",
        448 => "Named argument not found",
        449 => "Argument not optional",
        450 => "Wrong number of arguments or invalid property assignment",
        40074 => "Expected array",
        40109 => "Range has no values",
        40128 => "Named argument already specified",
        40136 => "Invalid or unqualified reference",
        40140 => "User-defined type may not be passed ByVal",
        40149 => "Invalid ReDim",
        57094 => "Can't assign to read-only property",
        57150 => "Event not found",
        ERR_LABEL_NOT_DEFINED => "Label not defined",
        ERR_SYNTAX_ERROR => "Syntax error",
        ERR_ARRAY_ARGUMENT_MUST_BE_BYREF => "Array argument must be ByRef",
        ERR_PROPERTY_ARGUMENT_REQUIRED => "Argument required for Property Let or Property Set",
        ERR_OPTIONAL_ARRAY => "Optional argument must be Variant or intrinsic type with a default value",
        ERR_PARAM_ARRAY_NOT_ARRAY => "ParamArray must be declared as an array of Variant",
        ERR_INVALID_EVENT_NAME => "Invalid event name",
        ERR_PARAM_ARRAY_WITH_OPTIONAL => "Can't have paramarrays with optional arguments",
        ERR_INVALID_OUTSIDE_PROCEDURE => "Invalid outside procedure",
        ERR_MUST_BE_FIRST_ON_LINE => "Must be first statement on the line",
        ERR_IDENTIFIER_TOO_LONG => "Identifier too long",
        ERR_TOO_MANY_DIMENSIONS => "Too many dimensions",
        ERR_INVALID_FIXED_STRING_LENGTH => "Invalid length for fixed-length string",
        ERR_END_TYPE_OUTSIDE_TYPE => "Statement invalid outside Type block",
        ERR_STATEMENT_IN_TYPE => "Statement invalid inside Type block",
        ERR_TYPE_WITHOUT_MEMBERS => "User-defined type without members not allowed",
        ERR_ENUM_WITHOUT_MEMBERS => "Enum without members not allowed",
        ERR_ARRAY_ALREADY_DIMENSIONED => "Array already dimensioned",
        ERR_DUPLICATE_OPTION => "Duplicate Option statement",
        ERR_DEFTYPE_AFTER_DECLARATION => "Deftype statements must precede declarations",
        ERR_DUPLICATE_DEFTYPE => "Duplicate Deftype statement",
        ERR_OPTION_PRIVATE_IN_OBJECT_MODULE => "Option Private Module not permitted in a object module",
        ERR_CC_WITHOUT_IF => "An # ElseIf, # Else, or # EndIf must be preceded by an # If clause",
        ERR_CC_DUPLICATE_ELSE => "An # Else clause must be preceded by an # If",
        ERR_CC_ELSEIF_AFTER_ELSE => "An # ElseIf must be preceded by an # If or # ElseIf and followed by an # Else or # EndIf",
        ERR_CC_IF_NOT_TERMINATED => "You must terminate the # If block with an # EndIf",
        ERR_EXPECTED_END_SUB => "Expected End Sub",
        ERR_EXPECTED_END_FUNCTION => "Expected End Function",
        ERR_AFTER_END_OF_PROCEDURE => "Only comments may appear after End Sub, End Function, or End Property",
        ERR_BLOCK_IF_WITHOUT_END_IF => "Block If without End If",
        ERR_WHILE_WITHOUT_WEND => "While without Wend",
        ERR_DO_WITHOUT_LOOP => "Do without Loop",
        ERR_FOR_WITHOUT_NEXT => "For without Next",
        ERR_ELSE_WITHOUT_IF => "Else without If",
        ERR_END_IF_WITHOUT_BLOCK_IF => "End If without block If",
        ERR_NEXT_WITHOUT_FOR => "Next without For",
        ERR_LOOP_WITHOUT_DO => "Loop without Do",
        ERR_WEND_WITHOUT_WHILE => "Wend without While",
        ERR_EXIT_FOR_OUTSIDE_FOR => "Exit For not within For...Next",
        ERR_EXIT_DO_OUTSIDE_DO => "Exit Do not within Do...Loop",
        ERR_EXIT_FUNCTION_NOT_ALLOWED => "Exit Function not allowed in Sub or Property",
        ERR_EXIT_SUB_NOT_ALLOWED => "Exit Sub not allowed in Function or Property",
        ERR_CASE_ELSE_OUTSIDE_SELECT => "Case Else outside Select Case",
        ERR_CASE_WITHOUT_SELECT => "Case without Select Case",
        ERR_END_SELECT_WITHOUT_SELECT => "End Select without Select Case",
        ERR_SELECT_WITHOUT_END_SELECT => "Select Case without End Select",
        ERR_EXPECTED_END_PROPERTY => "Expected End Property",
        ERR_EXIT_PROPERTY_NOT_ALLOWED => "Exit Property not allowed in Function or Sub",
        ERR_END_WITH_WITHOUT_WITH => "End With without With",
        ERR_EXPECTED_END_WITH => "Expected End With",
        ERR_INVALID_NEXT_VARIABLE => "Invalid Next control variable reference",
        ERR_FOR_VARIABLE_IN_USE => "For control variable already in use",
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

    /// Number of diagnostics recorded so far.
    pub fn len(&self) -> usize {
        self.items.len()
    }

    pub fn is_empty(&self) -> bool {
        self.items.is_empty()
    }

    /// Drop every diagnostic after the first `len`.
    pub fn truncate(&mut self, len: usize) {
        self.items.truncate(len);
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
