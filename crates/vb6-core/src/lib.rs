//! VB6 runtime/compiler core, plus umbrella re-exports of the split-out
//! `vb6-syntax`, `vb6-sema`, and `vb6-engine` crates so existing
//! `vb6_core::{frontend,support,sema,session,frm,frx,vbp}` paths continue to
//! resolve.

pub mod builder;
pub mod context;
pub mod types;

// Re-exports keep this crate's own `crate::frontend::…` / `crate::sema::…`
// references resolving after the modules moved out, and preserve the
// `vb6_core::…` API for downstream crates.
pub use vb6_syntax::{frontend, support};
pub use vb6_sema::sema;
pub use vb6_engine::{frm, frx, session, vbp};
