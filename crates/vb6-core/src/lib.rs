//! Compiler-side state and runtime value layouts: the per-compilation
//! [`context::CompilerContext`], the binder-phase node [`builder`], and the 32-bit COM
//! value types the generated code works with ([`types`]).

pub mod builder;
pub mod context;
pub mod types;
