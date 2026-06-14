//! VB6 semantic analysis: name resolution and type annotation (the binder).

pub mod sema;

// Re-export the frontend modules under their original crate-root paths so
// intra-crate `crate::frontend::…` / `crate::support::…` references in the
// moved code resolve unchanged.
pub use vb6_syntax::{frontend, support};
