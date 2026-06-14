//! VB6 analysis engine: the host-facing [`session::Session`] API plus the
//! designer-file (`frm`/`frx`) and project (`vbp`) parsers it builds on.
//!
//! This is the crate an editor/LSP shell depends on; it pulls in the lexer,
//! parser, and binder but **not** the runtime/compiler replica.

pub mod session;
pub mod frm;
pub mod frx;
pub mod vbp;

// Re-export dependency modules under their original crate-root paths so the
// moved code's `crate::frontend::…` / `crate::sema::…` references resolve.
pub use vb6_syntax::{frontend, support};
pub use vb6_sema::sema;
