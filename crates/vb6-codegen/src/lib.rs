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
//! The main P-code output stream is a dense byte buffer: each instruction is a
//! 1-byte opcode followed by type-specific operands at their natural width —
//! 2-byte signed frame offsets (i16 LE) for loads/stores, 4-byte or 8-byte
//! payloads for numeric literals.  No word-boundary alignment is enforced.
//! See [`buffer::PcodeStream`].

pub mod bind;
pub mod buffer;
pub mod emit;
pub mod node;
pub mod tables;

pub use buffer::PcodeStream;
pub use emit::Emitter;
pub use node::{NodeArena, NodeRef, RawNode};
