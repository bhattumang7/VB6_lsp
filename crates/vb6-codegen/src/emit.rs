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
                1 => todo!(
                    "emit_expr: type-1 Dim/type-spec path — \
                     opcodes selected by type tag; Phase 4/6"
                ),
                2 => unimplemented!(
                    "emit_expr: Currency literal (op 2) — runtime literal opcode"
                ),
                3 => {
                    let tag = n.type_tag();
                    unimplemented!(
                        "emit_expr: numeric literal (op 3, type_tag {tag}) — \
                         runtime literal opcode"
                    );
                }
                4 => unimplemented!(
                    "emit_expr: String literal (op 4) — runtime-helper call sequence"
                ),
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
        if op < 0x2c {
            match op {
                // String concatenation (&): EbEmitStatement case 0x1a emits
                // 0xce (with size operand) or 0xcf directly.
                0x1a => {
                    self.emit_expr(n.lhs(), 0);
                    self.emit_expr(n.rhs(), 0);
                    todo!(
                        "emit_expr: string concat 0x1a — \
                         needs result type_tag to choose 0xce (sized) vs 0xcf"
                    );
                }
                // MUL Currency: EbEmitStatement case 0x18 handles this specially,
                // emitting 0xd1 (Integer RHS) or 0xb3 (other RHS) via EbEmitValue2.
                0x18 if n.type_tag() == 13 => {
                    self.emit_expr(n.lhs(), 0);
                    self.emit_expr(n.rhs(), 0);
                    todo!(
                        "emit_expr: MUL Currency — \
                         needs RHS type_tag to choose EbEmitValue2(0xb3) vs (0xd1)"
                    );
                }
                // Is / TypeOf operator: EbEmitStatement case 0x24 emits
                // 0xef (with size) for object type or 0xf0 for string type.
                0x24 => {
                    self.emit_expr(n.lhs(), 0);
                    self.emit_expr(n.rhs(), 0);
                    todo!(
                        "emit_expr: Is operator 0x24 — \
                         needs result type_tag to choose 0xef (sized) vs 0xf0"
                    );
                }
                // All other binary ops: standard EbEmitBinaryOperation2 path.
                0x16 | 0x17 | 0x18 | 0x19
                | 0x1d ..= 0x23 | 0x25 ..= 0x2b => {
                    self.emit_expr(n.lhs(), 0);
                    self.emit_expr(n.rhs(), 0);
                    self.emit_binop_value(&n);
                }
                _ => {} // other ops in this range are no-ops
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
                // `Is` operator: emit both operands then the Is opcode.  The
                // runtime opcode for Is is not yet mapped.
                unimplemented!(
                    "emit_expr: Is operator 0x36 — \
                     runtime opcode not yet mapped"
                );
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

        let rt_byte = RT_OPCODE_BYTE[n_opc];
        if rt_byte < 0xfb {
            self.stream.emit_byte(rt_byte);
        } else {
            self.stream.emit_byte(0xfb);
            self.stream.emit_byte(n_opc as u8);
        }
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
mod tests {
    use super::*;
    use crate::node::NodeArena;
    use crate::tables::{RT_LOAD_BY_CTX, RT_STORE_BY_CTX};

    /// Build a bound-symbol node whose frame offset (stored in type_info =
    /// high 16 bits of word[4]) is `offset`.
    fn sym(arena: &mut NodeArena, offset: i16) -> NodeRef {
        arena.alloc(NodeArena::node(0, 0, (offset as u16 as u32) << 16, 0, 0, 0))
    }

    /// Build a variable-load node (type 0x74) with the given type context and
    /// a bound symbol at `offset`.
    fn var_load(arena: &mut NodeArena, type_ctx: u16, offset: i16) -> NodeRef {
        let s = sym(arena, offset);
        arena.alloc(NodeArena::node(0x74, 0, s.0, type_ctx as u32, 0, 0))
    }

    fn emit(arena: &NodeArena, root: NodeRef) -> Vec<u8> {
        let mut e = Emitter::new(arena);
        e.emit_expr(root, 0);
        e.into_bytes()
    }

    // ── Table sanity ──────────────────────────────────────────────────────────

    #[test]
    fn rt_load_table_confirmed_entries() {
        assert_eq!(RT_LOAD_BY_CTX[1], 0x6b, "Integer load");
        assert_eq!(RT_LOAD_BY_CTX[2], 0x6c, "Long load");
        assert_eq!(RT_LOAD_BY_CTX[3], 0x6e, "Single load");
        assert_eq!(RT_LOAD_BY_CTX[4], 0x6f, "Double load");
        assert_eq!(RT_LOAD_BY_CTX[6], 0x6d, "Currency load");
    }

    #[test]
    fn rt_store_table_confirmed_entries() {
        assert_eq!(RT_STORE_BY_CTX[1], 0x70, "Integer store");
        assert_eq!(RT_STORE_BY_CTX[2], 0x71, "Long store");
        assert_eq!(RT_STORE_BY_CTX[3], 0x73, "Single store");
        assert_eq!(RT_STORE_BY_CTX[4], 0x74, "Double store");
        assert_eq!(RT_STORE_BY_CTX[6], 0x72, "Currency store");
    }

    // ── Variable loads ────────────────────────────────────────────────────────

    #[test]
    fn double_load_emits_6f_and_frame_offset() {
        // From the empirical probe: Double local `a` is at frame offset -140
        // (0xff74 as i16). Runtime: opcode 0x6f + LE i16 0xff74.
        let mut a = NodeArena::new();
        let v = var_load(&mut a, 4, 0xff74u16 as i16); // typeCtx 4 = Double
        assert_eq!(emit(&a, v), &[0x6f, 0x74, 0xff]);
    }

    #[test]
    fn double_load_second_local_probe_offset() {
        // Double local `b` at frame offset -148 (0xff6c).
        let mut a = NodeArena::new();
        let v = var_load(&mut a, 4, 0xff6cu16 as i16);
        assert_eq!(emit(&a, v), &[0x6f, 0x6c, 0xff]);
    }

    #[test]
    fn long_load_emits_6c_and_frame_offset() {
        let mut a = NodeArena::new();
        let v = var_load(&mut a, 2, -8i16); // typeCtx 2 = Long, offset -8
        assert_eq!(emit(&a, v), &[0x6c, 0xf8, 0xff]);
    }

    #[test]
    fn integer_load_emits_6b_and_frame_offset() {
        let mut a = NodeArena::new();
        let v = var_load(&mut a, 1, -4i16); // typeCtx 1 = Integer, offset -4
        assert_eq!(emit(&a, v), &[0x6b, 0xfc, 0xff]);
    }

    #[test]
    fn single_load_emits_6e_and_frame_offset() {
        let mut a = NodeArena::new();
        let v = var_load(&mut a, 3, -4i16); // typeCtx 3 = Single, offset -4
        assert_eq!(emit(&a, v), &[0x6e, 0xfc, 0xff]);
    }

    #[test]
    fn currency_load_emits_6d_and_frame_offset() {
        let mut a = NodeArena::new();
        let v = var_load(&mut a, 6, -8i16); // typeCtx 6 = Currency, offset -8
        assert_eq!(emit(&a, v), &[0x6d, 0xf8, 0xff]);
    }

    #[test]
    fn node_0x76_routes_to_same_var_load() {
        // Node type 0x76 uses the same body as 0x74.
        let mut a = NodeArena::new();
        let s = sym(&mut a, 0xff74u16 as i16);
        let v = a.alloc(NodeArena::node(0x76, 0, s.0, 4, 0, 0)); // typeCtx 4
        assert_eq!(emit(&a, v), &[0x6f, 0x74, 0xff]);
    }

    // ── Variable stores ───────────────────────────────────────────────────────

    #[test]
    fn emit_var_store_double_emits_74_and_offset() {
        let arena = NodeArena::new();
        let mut e = Emitter::new(&arena);
        e.emit_var_store(4, 0xff5cu16 as i16); // Double r at 0xff5c (-164)
        assert_eq!(e.into_bytes(), &[0x74, 0x5c, 0xff]);
    }

    #[test]
    fn emit_var_store_long_emits_71_and_offset() {
        let arena = NodeArena::new();
        let mut e = Emitter::new(&arena);
        e.emit_var_store(2, -4i16); // Long
        assert_eq!(e.into_bytes(), &[0x71, 0xfc, 0xff]);
    }

    // ── Positive frame offsets (unusual but valid) ────────────────────────────

    #[test]
    fn positive_frame_offset_encodes_correctly() {
        // Positive offsets are unusual (locals are negative) but must round-trip.
        let mut a = NodeArena::new();
        let v = var_load(&mut a, 4, 8i16);
        assert_eq!(emit(&a, v), &[0x6f, 0x08, 0x00]);
    }

    // ── No-op branches ────────────────────────────────────────────────────────

    #[test]
    fn op_zero_emits_nothing() {
        let mut a = NodeArena::new();
        let n = a.alloc(NodeArena::node(0x0, 0, 0, 0, 0, 0));
        assert_eq!(emit(&a, n), &[]);
    }

    #[test]
    fn op_in_0x2c_to_0x35_emits_nothing() {
        let mut a = NodeArena::new();
        let n = a.alloc(NodeArena::node(0x30, 0, 0, 0, 0, 0));
        assert_eq!(emit(&a, n), &[]);
    }

    // ── Binary-op opcode dispatch ─────────────────────────────────────────────
    //
    // Each test builds two Long-typed variable-load nodes (as stand-ins for the
    // two operands), wraps them in a binary-op node, and verifies the exact
    // byte sequence: [lhs-load:3 bytes] [rhs-load:3 bytes] [op-byte(s)].
    //
    // Opcode derivation (verified against RT_BINOP_BASE / RT_TYPE_OFFSET /
    // RT_OPCODE_BYTE tables extracted from the DLL binary):
    //
    //   AND (op 0x23=35): base=RT_BINOP_BASE[35]=0x0021=33, Long offset=2,
    //     n_opc=35, RT_OPCODE_BYTE[35]=0xc4  → 1 byte 0xc4
    //
    //   OR  (op 0x21=33): base=RT_BINOP_BASE[33]=0x0019=25, Long offset=2,
    //     n_opc=27, RT_OPCODE_BYTE[27]=0xc5  → 1 byte 0xc5
    //
    //   XOR (op 0x22=34): base=RT_BINOP_BASE[34]=0x0011=17, Long offset=2,
    //     n_opc=19, RT_OPCODE_BYTE[19]=0xfb  → 2 bytes [0xfb, 0x13]
    //
    //   ADD (op 0x16=22): base=RT_BINOP_BASE[22]=0x008e=142, Long offset=2,
    //     n_opc=144, RT_OPCODE_BYTE[144]=0xaa  → 1 byte 0xaa
    //
    //   ADD (op 0x16=22): Integer offset=1,
    //     n_opc=143, RT_OPCODE_BYTE[143]=0xa9  → 1 byte 0xa9
    //
    //   EQ  (op 0x6):  base=RT_BINOP_BASE[6]=0x00be=190, Long offset=2,
    //     n_opc=192, RT_OPCODE_BYTE[192]=0xc3  → 1 byte 0xc3

    /// Helper: allocate a minimal Long-load node at the given frame offset.
    fn long_load(a: &mut NodeArena, offset: i16) -> NodeRef {
        var_load(a, 2, offset) // typeCtx 2 = Long
    }

    /// Helper: allocate a binary-op node with opcode `op`, result type_tag
    /// `type_tag`, and the given LHS / RHS children.
    fn binop(a: &mut NodeArena, op: u16, type_tag: u16, lhs: NodeRef, rhs: NodeRef) -> NodeRef {
        a.alloc(NodeArena::node(op, type_tag, lhs.0, rhs.0, 0, 0))
    }

    #[test]
    fn and_long_emits_c4() {
        let mut a = NodeArena::new();
        let lhs = long_load(&mut a, -8);
        let rhs = long_load(&mut a, -12);
        let n = binop(&mut a, 0x23, 8, lhs, rhs); // AND, Long result
        assert_eq!(
            emit(&a, n),
            &[0x6c, 0xf8, 0xff, 0x6c, 0xf4, 0xff, 0xc4]
        );
    }

    #[test]
    fn or_long_emits_c5() {
        let mut a = NodeArena::new();
        let lhs = long_load(&mut a, -8);
        let rhs = long_load(&mut a, -12);
        let n = binop(&mut a, 0x21, 8, lhs, rhs); // OR, Long result
        assert_eq!(
            emit(&a, n),
            &[0x6c, 0xf8, 0xff, 0x6c, 0xf4, 0xff, 0xc5]
        );
    }

    #[test]
    fn xor_long_emits_fb_13() {
        let mut a = NodeArena::new();
        let lhs = long_load(&mut a, -8);
        let rhs = long_load(&mut a, -12);
        let n = binop(&mut a, 0x22, 8, lhs, rhs); // XOR, Long result
        // n_opc=19=0x13 → rt_byte=0xfb → extended form [0xfb, 0x13]
        assert_eq!(
            emit(&a, n),
            &[0x6c, 0xf8, 0xff, 0x6c, 0xf4, 0xff, 0xfb, 0x13]
        );
    }

    #[test]
    fn add_long_emits_aa() {
        let mut a = NodeArena::new();
        let lhs = long_load(&mut a, -8);
        let rhs = long_load(&mut a, -12);
        let n = binop(&mut a, 0x16, 8, lhs, rhs); // ADD, Long result
        assert_eq!(
            emit(&a, n),
            &[0x6c, 0xf8, 0xff, 0x6c, 0xf4, 0xff, 0xaa]
        );
    }

    #[test]
    fn add_integer_emits_a9() {
        let mut a = NodeArena::new();
        let lhs = var_load(&mut a, 1, -4); // Integer
        let rhs = var_load(&mut a, 1, -6);
        let n = binop(&mut a, 0x16, 6, lhs, rhs); // ADD, Integer result
        assert_eq!(
            emit(&a, n),
            &[0x6b, 0xfc, 0xff, 0x6b, 0xfa, 0xff, 0xa9]
        );
    }

    #[test]
    fn eq_long_emits_c3() {
        let mut a = NodeArena::new();
        let lhs = long_load(&mut a, -8);
        let rhs = long_load(&mut a, -12);
        // Comparison node: op=6 (EQ), type_tag=8 (Long comparison type,
        // set by binder to operand type, not Boolean result).
        let n = binop(&mut a, 0x6, 8, lhs, rhs);
        assert_eq!(
            emit(&a, n),
            &[0x6c, 0xf8, 0xff, 0x6c, 0xf4, 0xff, 0xc3]
        );
    }
}
