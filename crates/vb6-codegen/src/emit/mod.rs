//! Expression / statement code generation: the runtime P-code byte stream.
//!
//! [`Emitter::emit_expr`] walks a bound expression/statement tree ([`NodeArena`])
//! and writes the dense little-endian P-code byte stream the runtime interpreter
//! executes. Dispatch is a single `match` on the node opcode
//! (`(short)word[0]`, [`RawNode::opcode`]); opcodes outside `1..=0x73` emit
//! nothing. Cases that finish emission `return`; cases that share the common tail
//! evaluate to the trailing opcode the tail emits.
//!
//! Opcode bytes are never hard-coded: every opcode is resolved through the
//! runtime opcode-byte table ([`RT_OPCODE_BYTE`]) by [`Emitter::emit_value2`] /
//! [`Emitter::emit_opcode2`], which apply the 1-or-2-byte escape encoding.
//!
//! ## Synthetic load/store nodes
//! Typed local / argument / global load and store byte sequences are driven by
//! synthetic node opcodes `0x74`..=`0x77` that the lowering pass builds; they are
//! handled ahead of the `1..=0x73` dispatch guard.
//!
//! ## Deferred cases
//! Cases that depend on machinery not yet built — the module symbol table, the
//! type/string pool, the type-conversion / instruction emitters — are
//! `unimplemented!()` with a description of what they need. They never emit a
//! guessed byte and never silently fall through.
//!
//! ## Module layout
//! The emitter's dispatch and helper methods are split by area:
//! [`expr`] (the main `emit_expr` dispatch), [`ops`] (binary-operation / type
//! validation / coercion support), [`reference`] (resolved-reference and
//! call-operand emission), [`calls`] (call-site emission), [`assign`] (typed
//! load/store and `=` assignment), [`stmt`] (statement-list walks, literals, and
//! output primitives), and [`intrinsics`] (explicit-conversion / unary-intrinsic
//! opcode tables).

use crate::buffer::PcodeStream;
use crate::node::{NodeArena, NodeRef, RawNode};
use crate::tables::{RT_BINOP_BASE, RT_CALL_TYPECODE, RT_DISPATCH_FLAG, RT_LOAD_BY_CTX, RT_OPCODE_BYTE, RT_RESULT_TYPE, RT_STORE_BY_CTX, RT_TYPE_OFFSET};
use crate::type_pool::TypePool;

mod assign;
mod calls;
mod expr;
mod intrinsics;
mod ops;
mod reference;
mod stmt;

/// Compute the call type-code for a call site: index [`RT_CALL_TYPECODE`] by the
/// reference-vs-value path.
///
/// * value path (`is_ref` false): index `(callee_type != 1) + (mask) * 2` → 0..3.
/// * reference path (`is_ref` true): index `(callee_type != 1) + 4` → 4..5.
///
/// `callee_type` and `mask` come from the callee's resolved descriptor (the
/// symbol-table model supplies them); this kernel is a pure function of its
/// inputs and the extracted table.
pub fn call_type_code(callee_type: i32, is_ref: bool, mask: bool) -> u16 {
    let idx = if is_ref {
        (callee_type != 1) as usize + 4
    } else {
        (callee_type != 1) as usize + (mask as usize) * 2
    };
    RT_CALL_TYPECODE[idx]
}

/// Map a call type-code to its runtime call opcode.
pub fn map_call_type_code(code: u16) -> u16 {
    match code {
        0x300 => 0x16a,
        0x310 => 0x169,
        0x320 => 0x34f,
        0x340 => 0x16b,
        0x350 => 0x16c,
        _ => 0x446,
    }
}

/// A resolved-reference descriptor — the input to [`Emitter::emit_reference`].
/// A reference resolver (the vb6-sema bridge) populates it; the emitter only
/// reads it to choose the typed load/store opcode.
///
/// * `kind` — storage class (1 = local, 2 = argument, 7 = indirect module-level);
///   selects the opcode base.
/// * `operand` — the 2-byte operand emitted after the opcode (a local's signed
///   frame offset, or a type size).
/// * `word6` — for argument / global kinds, bit 0 marks a ByRef slot (forces the
///   ByRef operation path). For the operator-reference kinds (8/9/0xb) it is the
///   opcode operand word (descriptor `+6`).
/// * `word8` — the descriptor's `+8` low word: the trailing word for kind 8, the
///   opcode operand for kind 0xb, and the extra word on the nested-type-expression
///   path (reached via the Variant-with-type chain).
/// * `flags1` — the descriptor's `+4` flag byte; bit `0x04` gates the finalize
///   tail for the operator-reference kinds.
#[derive(Clone, Copy, Debug, Default, PartialEq, Eq)]
pub struct RefDescriptor {
    pub kind: i32,
    pub operand: u16,
    pub word6: u16,
    pub word8: u16,
    pub flags1: u8,
}

/// A resolved call site — the input to [`Emitter::emit_call`]. The binder
/// (symbol-table model) populates it from the resolved callee; the emitter only
/// reads it. Fields mirror the bound call node the runtime emitter consumes:
///
/// * `kind` — the callee's call-convention kind (4..8); selects the dispatch
///   record and the call type-code.
/// * `byref` — 0 = by-value / value-returning, 1 = by-reference (method/Sub).
/// * `flags` — the call node's flag word (`word[1]`): bits `0x8000` / `0x800` /
///   `0x2000` / `0x1000` / `0x200` select the emission path; byte 1 carries the
///   `0x20` / `0x40` sub-flags.
/// * `node_word0` — the call node's `word[0]` (type tag in the high half, type
///   region in the high 16 bits).
/// * `callee` — the callee reference sub-node (`word[6]`), emitted with context 6.
/// * `arg_list` — the argument-list node (`word[5]`), or null.
/// * `member_id` — the callee's member-dispatch id (`word[7]`), emitted as a
///   trailing word on the dispatch path.
/// * `size` — the callee's resolved type size (from `word[8]`), emitted only
///   when the dispatch record requests a size operand.
/// * `node9` — the call node's `word[9]`, interned through the type pool to
///   produce the finalize step's trailing word on the dispatch path (emit
///   mode 1).
#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub struct CallDescriptor {
    pub kind: i32,
    pub byref: i32,
    pub flags: u32,
    pub node_word0: u32,
    pub callee: NodeRef,
    pub arg_list: NodeRef,
    pub member_id: u16,
    pub size: u16,
    pub node9: u32,
}

/// Drives [`Emitter::emit_expr`] over a [`NodeArena`], writing the runtime
/// P-code byte stream.
/// The module symbol context a member-reference (`0x60`) emission needs: the
/// compiled records heap the resolver reads, the member's byte offset within it
/// (`EbGetExprContext`), the compiler-context flag byte (`in_ECX + 0xc`), and the
/// binder-resolved convention `(kind, byref)` for the category-4 path.
#[derive(Clone, Debug, Default)]
pub struct SymbolContext {
    pub heap: Vec<u8>,
    pub member_off: usize,
    pub ctx_flag_c: u8,
    pub binding: Option<(i32, i32)>,
}

pub struct Emitter<'a> {
    arena: &'a NodeArena,
    stream: PcodeStream,
    type_pool: TypePool,
    sym: Option<SymbolContext>,
}

impl<'a> Emitter<'a> {
    pub fn new(arena: &'a NodeArena) -> Self {
        Self {
            arena,
            stream: PcodeStream::new(),
            type_pool: TypePool::new(),
            sym: None,
        }
    }

    /// Attach the module symbol context used by member-reference (`0x60`)
    /// emission.
    pub fn with_symbol_context(mut self, sym: SymbolContext) -> Self {
        self.sym = Some(sym);
        self
    }

    /// The type-intern pool accumulated during emission (for inspection/tests).
    pub fn type_pool(&self) -> &TypePool {
        &self.type_pool
    }

    /// The runtime P-code bytes emitted so far.
    pub fn bytes(&self) -> &[u8] {
        self.stream.bytes()
    }

    /// Consume the emitter, yielding the full runtime P-code byte stream.
    pub fn into_bytes(self) -> Vec<u8> {
        self.stream.into_bytes()
    }
}

#[cfg(test)]
#[path = "../tests/emit_tests.rs"]
mod tests;
