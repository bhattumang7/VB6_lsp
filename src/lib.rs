//! VB6 Language Server Library
//!
//! This library provides VB6 language support including:
//! - Tree-sitter based parsing
//! - Semantic analysis and symbol tables
//! - Project file (.vbp) parsing
//! - Workspace management
//! - LSP protocol implementation

pub mod analysis;
pub mod claude;
pub mod controls;
pub mod lsp;
pub mod parser;
pub mod utils;
pub mod workspace;
