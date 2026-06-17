//! Expression code generation — runtime P-code byte-stream form.
//!
//! The emitter walks bound expression nodes and writes the **runtime** P-code
//! byte stream: the same type-specific, byte-packed encoding stored in the VB6
//! `.exe`.  Contrast this with the intermediate compile-time word stream (which
//! uses generic 16-bit opcodes); the runtime form uses 1-byte type-specific
//! opcodes and 2-byte signed frame offsets for every load and store.
//!
//! ## Emit format for loads and stores
//! Each typed local-variable load or store is 3 bytes:
//! ```text
//! [opcode:u8] [frame_offset_lo:u8] [frame_offset_hi:u8]
//! ```
//! where `opcode` comes from [`RT_LOAD_BY_CTX`] or [`RT_STORE_BY_CTX`]
//! indexed by the expression's type context, and the 2-byte field is a
//! little-endian signed i16 (locals have negative offsets from the proc frame
//! pointer, e.g. the first `Double` local in a standard Sub is at −140 =
//! `0xff74`).
//!
//! ## Branches not yet implemented
//! Branches that depend on the binder, slot allocator, or opcode-survey data
//! not yet gathered are `todo!()` or `unimplemented!()` — never a guessed
//! constant.  A `todo!()` marks a path we know the exact behaviour of but
//! haven't yet ported (usually a later phase); `unimplemented!()` marks a path
//! that requires additional empirical opcode-survey work before the values can
//! be filled in.

use crate::buffer::PcodeStream;
use crate::node::{NodeArena, NodeRef, RawNode};
use crate::tables::{
    RT_BINOP_BASE, RT_DISPATCH_FLAG, RT_LOAD_BY_CTX, RT_OPCODE_BYTE, RT_STORE_BY_CTX,
    RT_TYPE_OFFSET,
};

/// Drives [`Emitter::emit_expr`] over a [`NodeArena`], writing the runtime
/// P-code byte stream.
pub struct Emitter<'a> {
    arena: &'a NodeArena,
    stream: PcodeStream,
}

impl<'a> Emitter<'a> {
    pub fn new(arena: &'a NodeArena) -> Self {
        Self {
            arena,
            stream: PcodeStream::new(),
        }
    }

    /// The runtime P-code bytes emitted so far.
    pub fn bytes(&self) -> &[u8] {
        self.stream.bytes()
    }

    /// Consume the emitter, yielding the full runtime P-code byte stream.
    pub fn into_bytes(self) -> Vec<u8> {
        self.stream.into_bytes()
    }

    /// Emit the runtime P-code for one expression node.
    ///
    /// `call_ctx` is the calling-convention context (0 for a normal value
    /// read).  Non-zero call contexts change the load encoding and are not
    /// yet fully mapped.
    pub fn emit_expr(&mut self, node: NodeRef, call_ctx: u32) {
        let n = *self.arena.get(node);
        let op = n.opcode();

        // ── Short-opcode family (op < 0x6) ───────────────────────────────────
        if op < 0x6 {
            match op {
                1 => self.emit_int_literal(&n),
                2 => self.emit_currency_literal(&n),
                3 => self.emit_float_literal(&n, call_ctx),
                4 => self.emit_string_literal(&n),
                5 => unimplemented!(
                    "emit_expr: type-reference node (op 5)"
                ),
                _ => {} // op 0: no-op
            }
            return;
        }

        // ── Comparison binary operators (op 0x6–0xd) ─────────────────────────
        // EQ/NE/LT/GT/LE/GE/Like/Is comparisons.  The binder sets the node's
        // type_tag to the operand comparison type (not the Boolean result), so
        // the standard arithmetic dispatch path in emit_binop_value is correct.
        if op <= 0xd {
            self.emit_expr(n.lhs(), 0);
            self.emit_expr(n.rhs(), 0);
            self.emit_binop_value(&n);
            return;
        }

        // ── Short-opcode family (op 0xe–0x15) ────────────────────────────────
        // op is >= 0xe here; 0x0..=0xd are already handled above.
        if op == 0xe {
            todo!(
                "emit_expr: typed-load 0x0e path — runtime opcode from \
                 TYPE_SHIFT table; Phase 3/4"
            );
        }
        // op 0xf..=0x15: no-op pass-through nodes.
        if op < 0x16 {
            return;
        }

        // ── Arithmetic / logical binary operators (op 0x16–0x2b) ─────────────
        // All ops with a non-sentinel RT_BINOP_BASE entry use the EbEmitBinaryOperation2
        // arithmetic dispatch path; RT_DISPATCH_FLAG & 0x10 is 0 for every op in
        // this range.  Ops 0x1b and 0x1c have sentinel base (0x0446) and are no-ops.
        // Op 0x24 (Is) has its own explicit case in EbEmitStatement — it does NOT
        // route through EbEmitBinaryOperation2.
        if op < 0x2c {
            match op {
                0x16..=0x1a | 0x1d..=0x23 | 0x25..=0x2b => {
                    self.emit_expr(n.lhs(), 0);
                    self.emit_expr(n.rhs(), 0);
                    self.emit_binop_value(&n);
                }
                0x24 => self.emit_is_op(&n, call_ctx),
                _ => {} // 0x1b, 0x1c have sentinel base — no runtime mapping
            }
            return;
        }

        // ── Name / call / typed family (op 0x2c–0x60) ────────────────────────
        if op < 0x62 {
            if op == 0x61 {
                todo!(
                    "emit_expr: name/call node 0x61 (argument list + dispatch); \
                     Phase 3"
                );
            }
            if op == 0x36 {
                // EbEmitStatement case 0x36: emit LHS(ctx=1), RHS(ctx=1), then
                // n_opc=0xd2.  RT_OPCODE_BYTE[0xd2]=0xfb → extended form [0xfb, 0xd2].
                self.emit_expr(n.lhs(), 1);
                self.emit_expr(n.rhs(), 1);
                self.emit_value2(0xd2);
                return;
            }
            // Other [0x2c, 0x62): no-op.
            return;
        }

        // ── Argument / call / variable-load family (op 0x62–0x74) ────────────
        if op < 0x75 {
            match op {
                0x74 => self.emit_var_load(&n, call_ctx),
                0x62 => todo!(
                    "emit_expr: argument node 0x62 (arg emission); Phase 3"
                ),
                0x63 => todo!(
                    "emit_expr: overload node 0x63 (resolve + 0x9d); Phase 3"
                ),
                0x66 => todo!(
                    "emit_expr: overload node 0x66 (resolve + 0xc6); Phase 3"
                ),
                _ => {}
            }
            return;
        }

        // ── Name / property / typed-load switch (op 0x75–0x91) ───────────────
        if op > 0x87 {
            if op == 0x91 {
                unimplemented!(
                    "emit_expr: node 0x91 — runtime form of typed 0x104 \
                     emit not yet mapped"
                );
            }
            return;
        }
        if op == 0x87 {
            unimplemented!(
                "emit_expr: node 0x87 — runtime opcode not yet confirmed"
            );
        }
        match op {
            // 0x76 routes to the same variable-load body as 0x74.
            0x76 => self.emit_var_load(&n, call_ctx),
            0x75 => todo!(
                "emit_expr: node 0x75 (typed emit variant); Phase 3/4"
            ),
            0x77 => unimplemented!(
                "emit_expr: node 0x77 — runtime opcode not yet mapped"
            ),
            0x78 => unimplemented!(
                "emit_expr: node 0x78 — runtime opcode not yet mapped"
            ),
            0x79 => unimplemented!(
                "emit_expr: node 0x79 — runtime opcode not yet mapped"
            ),
            0x7a => todo!(
                "emit_expr: node 0x7a (name/call path); Phase 3"
            ),
            0x7d => todo!(
                "emit_expr: node 0x7d (property/array typed emit); Phase 3/4"
            ),
            0x7e => todo!(
                "emit_expr: node 0x7e (property typed emit); Phase 3/4"
            ),
            0x7f => todo!(
                "emit_expr: node 0x7f (chained comparison); Phase 3"
            ),
            _ => {}
        }
    }

    /// Emit the runtime byte(s) for `n_opc`, mirroring EbEmitValue2.
    ///
    /// Looks up `RT_OPCODE_BYTE[n_opc]`; if the byte is < 0xfb, emits it as a
    /// single-byte instruction; otherwise emits that byte followed by `n_opc as u8`
    /// (extended form, used when the opcode space overflows one byte).
    fn emit_value2(&mut self, n_opc: usize) {
        let rt_byte = RT_OPCODE_BYTE[n_opc];
        if rt_byte < 0xfb {
            self.stream.emit_byte(rt_byte);
        } else {
            self.stream.emit_byte(rt_byte);
            self.stream.emit_byte(n_opc as u8);
        }
    }

    /// Emit a typed local-variable load instruction.
    ///
    /// The type context (typeCtx) lives in `word[5]` of the variable-load node
    /// and selects both the 1-byte runtime opcode from [`RT_LOAD_BY_CTX`] and
    /// the operand interpretation.  The bound symbol child carries the signed
    /// frame offset in its `type_info()` field (high 16 bits of its `word[4]`);
    /// that offset is emitted as a 2-byte little-endian i16.
    ///
    /// Node types 0x74 and 0x76 both route here.  `call_ctx != 0` is not yet
    /// mapped to a confirmed runtime encoding.
    fn emit_var_load(&mut self, n: &RawNode, call_ctx: u32) {
        if call_ctx != 0 {
            unimplemented!(
                "emit_var_load: call_ctx {} — runtime encoding not yet confirmed",
                call_ctx
            );
        }
        let type_ctx = n.word(5) as usize;
        let opcode = RT_LOAD_BY_CTX
            .get(type_ctx)
            .copied()
            .unwrap_or(0);
        if opcode == 0 {
            unimplemented!(
                "emit_var_load: no confirmed runtime opcode for typeCtx {}",
                type_ctx
            );
        }
        let sym = self.arena.get(n.lhs());
        let frame_offset = sym.type_info() as i16;
        self.stream.emit_byte(opcode);
        self.stream.emit_i16(frame_offset);
    }

    /// Emit the Is-operator node (op 0x24).
    ///
    /// EbEmitStatement case 0x24 is an explicit case — it does NOT route through
    /// EbEmitBinaryOperation2.  Both operands are emitted with call_ctx=1, then
    /// the opcode depends on the node's type_tag:
    ///
    /// - 0xf (UDT/object with explicit size): EbEmitOpcode2(0xef, type_size) then
    ///   EbValidateTypeOperation(0xf, 0x17, ctx).  Requires EbGetTypeSize3; Phase 4.
    /// - 0x10: EbEmitValue2(0xf0) → RT_OPCODE_BYTE[0xf0]=0x2a; no further emission
    ///   (type_tag 0x10 has no matching case in EbValidateTypeOperation).
    /// - 2 (String): EbValidateTypeOperation(2, 0x17, _) → always n_opc=0x18b →
    ///   RT_OPCODE_BYTE[0x18b]=0xfc → [0xfc, 0x8b].
    /// - 10 (Single): n_opc=0x189 → RT_OPCODE_BYTE[0x189]=0x37 (single byte),
    ///   emitted only when call_ctx is 1 or 3.
    /// - 0xb/0xc (Double/Currency): n_opc=0x18a → RT_OPCODE_BYTE[0x18a]=0x39
    ///   (single byte), emitted only when call_ctx is 1 or 3.
    /// - other type_tag: no opcode emitted (EbValidateTypeOperation returns 0).
    fn emit_is_op(&mut self, n: &RawNode, call_ctx: u32) {
        self.emit_expr(n.lhs(), 1);
        self.emit_expr(n.rhs(), 1);
        match n.type_tag() {
            0xf => todo!(
                "emit_is_op: type_tag 0xf (UDT/object) requires EbGetTypeSize3(node.word(6)) \
                 for EbEmitOpcode2(0xef, size); Phase 4"
            ),
            0x10 => {
                self.emit_value2(0xf0);
            }
            2 => {
                self.emit_value2(0x18b);
            }
            10 => {
                if call_ctx == 1 || call_ctx == 3 {
                    self.emit_value2(0x189);
                }
            }
            0xb | 0xc => {
                if call_ctx == 1 || call_ctx == 3 {
                    self.emit_value2(0x18a);
                }
            }
            _ => {}
        }
    }

    /// Emit a runtime integer literal (op 1).
    ///
    /// type_tag 3/5/6 (Integer): if the value fits in a signed byte use
    /// n_opc=0x41a (rt_byte=0xf4) + 1 byte; otherwise n_opc=0x3b5
    /// (rt_byte=0xf3) + 2-byte LE i16.  type_tag 8/0x10 (Long/Byte):
    /// n_opc=0x3b8 (rt_byte=0xf5) + 4-byte LE i32.
    fn emit_int_literal(&mut self, n: &RawNode) {
        let tag = n.type_tag();
        match tag {
            3 | 5 | 6 => {
                let val = n.word(4) as i16;
                if val >= -128 && val < 128 {
                    self.stream.emit_byte(RT_OPCODE_BYTE[0x41a]);
                    self.stream.emit_byte(val as u8);
                } else {
                    self.emit_value2(0x3b5);
                    self.stream.emit_i16(val);
                }
            }
            8 | 0x10 => {
                self.stream.emit_byte(RT_OPCODE_BYTE[0x3b8]);
                self.stream.emit_bytes(&n.word(4).to_le_bytes());
            }
            _ => unimplemented!(
                "emit_int_literal: integer literal type_tag {tag} — \
                 not yet mapped to a literal opcode"
            ),
        }
    }

    /// Emit a runtime Currency literal (op 2).
    ///
    /// Emits n_opc=0x3bb (rt_byte=0xf6) followed by the 8-byte LE
    /// i64×10000 payload from `word[4]`/`word[5]`.
    fn emit_currency_literal(&mut self, n: &RawNode) {
        self.stream.emit_byte(RT_OPCODE_BYTE[0x3bb]);
        self.stream.emit_bytes(&n.literal8());
    }

    /// Emit a runtime floating-point literal (op 3).
    ///
    /// `call_ctx == 2` selects the "assign context" n_opc variants (0x3ba for
    /// Single, 0x3bd for Double/Date); all other call contexts use the
    /// non-assign variants.  Single literals (type_tag 10) are converted from
    /// the f64 stored in the node to f32 before emission; Double and Date
    /// (type_tag 11/12) emit the raw 8-byte f64.
    fn emit_float_literal(&mut self, n: &RawNode, call_ctx: u32) {
        let tag = n.type_tag();
        if tag == 10 {
            let n_opc = if call_ctx == 2 { 0x3ba } else { 0x3b9 };
            self.stream.emit_byte(RT_OPCODE_BYTE[n_opc]);
            let f32_bits = (n.literal_f64() as f32).to_bits();
            self.stream.emit_bytes(&f32_bits.to_le_bytes());
        } else if tag > 10 && tag < 13 {
            let n_opc = if call_ctx == 2 { 0x3bd } else { 0x3bc };
            self.stream.emit_byte(RT_OPCODE_BYTE[n_opc]);
            self.stream.emit_bytes(&n.literal8());
        } else {
            unimplemented!(
                "emit_float_literal: op 3 type_tag {tag} not in Single/Double/Date range"
            );
        }
    }

    /// Emit a runtime String literal (op 4).
    ///
    /// Two sub-cases driven by bit 15 of `word[1]` (`node+5` byte 0x80):
    /// * **Null/zero string** (bit set): emit n_opc=0x3b8 (rt_byte=0xf5) + 4
    ///   zero bytes — equivalent to emitting `Long 0` as a null BSTR pointer.
    /// * **Non-empty string** (bit clear): requires resolving a type descriptor
    ///   from `word[4]` via the pool type system (EbExtractTypeValue2 /
    ///   EbParseExpression2 / EbRegisterTypeInfo2), which is Phase 4 work.
    fn emit_string_literal(&mut self, n: &RawNode) {
        if (n.w[1] >> 15) & 1 != 0 {
            // Null/zero string: emit as Long-zero literal.
            self.stream.emit_byte(RT_OPCODE_BYTE[0x3b8]);
            self.stream.emit_bytes(&[0u8; 4]);
        } else {
            todo!(
                "emit_string_literal: non-null string requires pool type resolution \
                 (EbExtractTypeValue2 / EbParseExpression2); Phase 4"
            );
        }
    }

    /// Emit the runtime binary-operation byte(s) for `n`, using the three-level
    /// table dispatch from `EbEmitBinaryOperation2`.
    ///
    /// Precondition: the LHS and RHS of `n` have already been emitted onto the
    /// virtual stack by the caller.
    fn emit_binop_value(&mut self, n: &RawNode) {
        let op = n.opcode() as usize;
        let base = RT_BINOP_BASE[op] as i32;
        debug_assert!(
            base != 0x0446,
            "emit_binop_value: op {op:#x} has no runtime mapping (base=0x0446)"
        );

        // All known ops have RT_DISPATCH_FLAG[op] & 0x10 == 0 → arithmetic path:
        // use the node's own type_tag (set by the binder to the operation type).
        // The comparison path (flag & 0x10 != 0) is kept as a guard in case an
        // op outside the known table ever arrives.
        let flag = RT_DISPATCH_FLAG[op];
        if flag & 0x10 != 0 {
            todo!(
                "emit_binop_value: comparison-path dispatch (flag {flag:#x}) \
                 for op {op:#x} — requires LHS type_tag lookup; Phase 3/4"
            );
        }

        let type_tag = n.type_tag();
        let raw_offset = RT_TYPE_OFFSET[type_tag as usize];
        let offset = match raw_offset {
            10 => 4,
            9 => 1,
            x => x,
        };
        let n_opc = (base + offset) as usize;

        // When the node's type_tag is 0xf (UDT/object with explicit size), the
        // instruction takes a 2-byte type-size operand; all primitive types take
        // a plain value byte via EbEmitValue2.
        if type_tag == 0xf {
            todo!(
                "emit_binop_value: UDT type_tag=0xf path — \
                 requires EbGetTypeSize3(node.word(6)) operand; Phase 4"
            );
        }

        self.emit_value2(n_opc);
    }

    /// Emit a typed local-variable store instruction.
    ///
    /// Mirror of [`Self::emit_var_load`] using [`RT_STORE_BY_CTX`].  Call site
    /// is responsible for having already emitted the value to store onto the
    /// virtual stack before calling this.
    pub fn emit_var_store(&mut self, type_ctx: usize, frame_offset: i16) {
        let opcode = RT_STORE_BY_CTX
            .get(type_ctx)
            .copied()
            .unwrap_or(0);
        if opcode == 0 {
            unimplemented!(
                "emit_var_store: no confirmed runtime opcode for typeCtx {}",
                type_ctx
            );
        }
        self.stream.emit_byte(opcode);
        self.stream.emit_i16(frame_offset);
    }
}

#[cfg(test)]
#[path = "tests/emit_tests.rs"]
mod tests;
