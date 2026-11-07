//! VB6 Language Server Library
//!
//! Provides VB6 language support including:
//! - Opening a VB6 project or project group (`project`)
//! - VB6 form / FRX / RES binary companion file parsing
//! - LSP protocol implementation backed by vb6-engine

pub mod engine_glue;
pub mod lsp;
