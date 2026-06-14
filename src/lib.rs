//! VB6 Language Server Library
//!
//! Provides VB6 language support including:
//! - Project file (.vbp) parsing and workspace management
//! - VB6 form / FRX / RES binary companion file parsing
//! - LSP protocol implementation backed by vb6-engine

pub mod controls;
pub mod engine_glue;
pub mod lsp;
pub mod utils;
pub mod workspace;
