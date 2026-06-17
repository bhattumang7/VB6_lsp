//! VB6 P-code back-end.
//!
//! Goal: given our parser's AST, emit the exact P-code byte stream the VB6
//! compiler produces. This crate implements the VB6 P-code emission rules,
//! driven by the opcode/operand tables in [`tables`].
//!
//! Fidelity rule: every opcode value, operand width, and write order matches the
//! VB6 P-code format. A not-yet-implemented path is a `todo!()`/`unimplemented!()`
//! describing what it must emit, never a guessed constant.
//!
//! ## Pipeline
//! ```text
//! our ExprArena (parser AST)
//!   -> lower to the raw 40-byte node graph
//!   -> bind + slot-alloc + rewrite
//!   -> emit
//!   -> p-code bytes (+ per-proc descriptor)
//! ```
//!
//! ## Buffer model
//! The main P-code output stream is a word cursor over a byte buffer: opcodes and
//! operands are little-endian 16-bit words, with literals and data blobs written
//! as raw bytes onto the same byte-addressed, 2-byte-aligned stream. That is
//! [`buffer::PcodeStream`].

pub mod bind;
pub mod buffer;
pub mod emit;
pub mod node;
pub mod tables;

pub use buffer::PcodeStream;
pub use emit::Emitter;
pub use node::{NodeArena, NodeRef, RawNode};
