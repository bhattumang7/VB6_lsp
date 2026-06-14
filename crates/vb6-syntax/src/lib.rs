//! VB6 language frontend: the lexer/scanner and the recursive-descent parser.
//!
//! This crate is self-contained (no dependency on semantic analysis, the
//! engine, or the runtime) so the lexer/parser can be consumed independently.

pub mod support;
pub mod frontend;
