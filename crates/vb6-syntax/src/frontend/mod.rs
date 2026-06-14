//! VB6 language frontend: the scanner/lexer and its symbol interner.
//!
//! These functions operate on a scanner context object and a shared symbol
//! table. The state is modeled natively in Rust (ownership replaces a manual
//! hash table + bump arena) and only the *observable behavior* is reproduced.

pub mod ast;
pub mod diagnostics;
pub mod keyword_table;
pub mod parser;
pub mod scanner;
pub mod token;
