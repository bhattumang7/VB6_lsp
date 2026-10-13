//! Semantic-token classification for syntax highlighting.
//!
//! [`Session::semantic_tokens`](super::Session::semantic_tokens) re-scans a file
//! and emits a classified token per keyword / literal / comment / identifier.
//! Identifiers are classified from the bound model (variable vs function vs
//! parameter vs type vs enum member) — the part a static TextMate grammar cannot
//! do; unresolved identifiers and member names are left unstyled.

use crate::frontend::ast::Span;
use crate::sema::symbol::{ExternalDecl, NameResolution};

/// Semantic classification of a source token.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum SemTokenKind {
    Keyword,
    Function,
    Variable,
    Parameter,
    Type,
    EnumMember,
    String,
    Number,
    Comment,
}

/// A classified token: a byte [`Span`] and its kind. The host converts the span
/// to LSP positions (and delta-encodes) via `LineIndex`.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct SemToken {
    pub span: Span,
    pub kind: SemTokenKind,
}

/// Classify an identifier *use* from its name resolution. `None` leaves the
/// token unstyled (unresolved names).
pub(crate) fn kind_of_resolution(res: &NameResolution) -> Option<SemTokenKind> {
    Some(match res {
        NameResolution::Proc(_) | NameResolution::Builtin => SemTokenKind::Function,
        NameResolution::ModuleVar(_) | NameResolution::Local { .. } => SemTokenKind::Variable,
        NameResolution::Param { .. } => SemTokenKind::Parameter,
        NameResolution::EnumMember { .. } => SemTokenKind::EnumMember,
        NameResolution::External { decl, .. } => match decl {
            ExternalDecl::Proc(_) => SemTokenKind::Function,
            ExternalDecl::Var(_) => SemTokenKind::Variable,
            ExternalDecl::Type(_) | ExternalDecl::Enum(_) => SemTokenKind::Type,
            ExternalDecl::EnumMember { .. } => SemTokenKind::EnumMember,
            ExternalDecl::TypeMember { .. } | ExternalDecl::Control(_) => SemTokenKind::Variable,
        },
        NameResolution::Module(_)
        | NameResolution::Project
        | NameResolution::Library(_)
        | NameResolution::LibraryType { .. } => SemTokenKind::Type,
        // A library member may be a method or a property; without the library
        // model's kind the token is left to the client's base grammar.
        NameResolution::LibraryMember { .. } | NameResolution::Unresolved => return None,
    })
}
