//! Expression code generation.
//!
//! The expression emitter recursively walks an expression node and writes its
//! P-code to the main output stream ([`PcodeStream`]). It is a postfix
//! (reverse-Polish) emitter: for a binary operator it emits the left subtree,
//! then the right subtree, then the operator opcode looked up in
//! [`OP_BY_NODE_TYPE`] by node type.
//!
//! Branches that depend on the bind/slot-allocation and type-coercion phases
//! (name/call/property/typed-load, the type-1 Dim path, the Single conversion,
//! and string-pool sourcing) are `todo!()` describing what they must emit, never
//! a guessed encoding — they are filled in as those phases land.

use crate::buffer::PcodeStream;
use crate::node::{NodeArena, NodeRef};
use crate::tables::{OP_BY_NODE_TYPE, TYPE_CTX_BITS, TYPE_SHIFT, VARLOAD_OP_BY_CALLCTX};

/// Drives [`Emitter::emit_expr`] over a [`NodeArena`], accumulating P-code bytes.
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

    /// The P-code bytes emitted so far.
    pub fn bytes(&self) -> &[u8] {
        self.stream.bytes()
    }

    /// Consume the emitter, yielding the P-code byte stream.
    pub fn into_bytes(self) -> Vec<u8> {
        self.stream.into_bytes()
    }

    /// Emit the 2-byte type-info operand stored at `node + 0x12`.
    fn emit_type_info3(&mut self, node: NodeRef) {
        let info = self.arena.get(node).type_info();
        self.stream.emit_word(info);
    }

    /// Fold the type-context bits into a typed opcode:
    /// `TYPE_CTX_BITS[typeCtx] << 10 | opcode`. When `type_ctx == 0` and the node
    /// carries the indirect-type flag, the node is redirected to its pooled type
    /// and the context becomes 5; otherwise the opcode is returned unchanged.
    fn resolve_typed_opcode(&self, opcode: u16, node: &mut NodeRef, type_ctx: i32) -> u16 {
        if type_ctx == 0 {
            if self.arena.get(*node).has_indirect_type() {
                // node -> type-pool[node+0x10]; type context becomes 5.
                todo!(
                    "resolve_typed_opcode: type-pool indirection (type-info pool) \
                     - Phase 4"
                );
            }
            // No indirection: the opcode is returned unchanged.
            return opcode;
        }
        (TYPE_CTX_BITS[type_ctx as usize] as u16) << 10 | opcode
    }

    /// Emit a typed opcode word (resolved from `type_ctx` unless `f_raw`), then
    /// the node's type-info operand via [`Self::emit_type_info3`].
    pub fn emit_typed_pcode(
        &mut self,
        type_or_opcode: u16,
        node: NodeRef,
        type_ctx: i32,
        f_raw: bool,
    ) {
        let mut node = node;
        let word = if f_raw {
            type_or_opcode
        } else {
            self.resolve_typed_opcode(type_or_opcode, &mut node, type_ctx)
        };
        self.stream.emit_word(word);
        self.emit_type_info3(node);
    }

    /// Emit the P-code for an expression node.
    pub fn emit_expr(&mut self, node: NodeRef, call_ctx: u32) {
        let n = *self.arena.get(node);
        let op = n.opcode();

        if op < 0xf {
            if op == 0xe {
                // Typed-load / unary-load path. Emit the operand, then a single
                // typed-load opcode formed from the node's type tag:
                // word = (TYPE_SHIFT[tag] << 10) | 0x58, with the two special-
                // cased tags 0xf and 0x1f.
                self.emit_expr(n.lhs(), 0);
                let tag = n.type_tag();
                let mut shift: i32 = 0;
                let mut word: u32 = 0x58;
                if tag != 0xf {
                    if tag == 0x1f {
                        word = 0x59;
                        shift = 10;
                    } else {
                        shift = (TYPE_SHIFT[tag as usize] as i16) as i32;
                    }
                }
                word = ((shift << 10) as u32) | word;
                self.stream.emit_word(word as u16);
                return;
            }

            // op != 0xe, op < 0xf
            let tag = n.type_tag();
            if op != 1 {
                match op {
                    // Currency literal: opcode 0xa9 + 8-byte value.
                    2 => {
                        self.stream.emit_literal8(0xa9, n.literal8());
                    }
                    // 8-byte float family by type tag.
                    3 => {
                        if tag == 10 {
                            // Single: convert the 8-byte double payload to f32,
                            // then emit opcode 0xb3 + the 4-byte single.
                            let single = n.literal_f64() as f32;
                            self.stream.emit_opcode4(0xb3, single.to_le_bytes());
                            return;
                        }
                        let opcode = if tag == 0xb {
                            0xb4 // Date
                        } else if tag == 0xc {
                            0xaa // Variant
                        } else {
                            return; // other tags emit nothing
                        };
                        self.stream.emit_literal8(opcode, n.literal8());
                    }
                    // String literal: opcode 0xb6 + length + data. The data
                    // pointer is a byte offset into the arena blob store, and the
                    // even-rounded byte count is copied verbatim.
                    4 => {
                        let len = n.word(5) as u16;
                        let copy = ((len as usize) + 1) & !1;
                        let src = self.arena.blob(n.word(6), copy).to_vec();
                        self.stream.emit_word_and_data(0xb6, len, &src);
                    }
                    // Unary ops (types 6,7): emit the single operand (word[4]),
                    // then the operator opcode.
                    6 | 7 => {
                        self.emit_expr(n.lhs(), 0);
                        self.stream.emit_word(OP_BY_NODE_TYPE[op as usize]);
                    }
                    // Types 0, 5, 8..13: emit nothing.
                    _ => {}
                }
                return;
            }

            // op == 1: type-spec / Dim emission (opcodes 0xb7 / 0xac-0xb2)
            // selected by the type tag.
            todo!(
                "emit_expr: type-1 Dim/type-spec path (opcodes 0xb7/0xac-0xb2, by \
                 type tag {tag}) - Phase 4/6"
            );
        }

        if op < 0x2c {
            // Binary operator: node types 0x16-0x1a and 0x1d-0x2b. Postfix:
            // lhs, rhs, then the operator opcode.
            if op < 0x1d {
                if op < 0x16 {
                    return;
                }
                if op > 0x1a {
                    return;
                }
            }
            self.emit_expr(n.lhs(), 0);
            self.emit_expr(n.rhs(), 0);
            self.stream.emit_word(OP_BY_NODE_TYPE[op as usize]);
            return;
        }

        // op >= 0x2c: the name/call/property/typed-load family.
        if op < 0x62 {
            if op == 0x61 {
                todo!("emit_expr: name/call node 0x61 (argument list + call) - Phase 3");
            }
            if op == 0x36 {
                // `Is` operator: emit both operands, then opcode 0x14.
                self.emit_expr(n.lhs(), 0);
                self.emit_expr(n.rhs(), 0);
                self.stream.emit_word(0x14);
                return;
            }
            // Other [0x2c, 0x62) node types: emit nothing.
            return;
        }

        // op >= 0x62: argument/call/overload/variable-load/property family.
        if op < 0x75 {
            match op {
                // Variable load.
                0x74 => self.emit_var_load(&n, call_ctx),
                0x62 => todo!("emit_expr: argument node 0x62 (arg emission) - Phase 3"),
                0x63 => todo!("emit_expr: node 0x63 (overload resolve + 0x9d) - Phase 3"),
                0x66 => todo!("emit_expr: node 0x66 (overload resolve + 0xc6) - Phase 3"),
                // Other [0x62, 0x75): emit nothing.
                _ => {}
            }
            return;
        }

        // op >= 0x75: name/property/typed-load switch.
        if op > 0x87 {
            if op != 0x91 {
                return;
            }
            // 0x91: typed emit of opcode 0x104 over child=word[4], typeCtx=word[5].
            self.emit_typed_pcode(0x104, n.lhs(), n.word(5) as i32, false);
            return;
        }
        if op == 0x87 {
            // Emit the single operand (word[4]), then the operator opcode.
            self.emit_expr(n.lhs(), 0);
            self.stream.emit_word(OP_BY_NODE_TYPE[op as usize]);
            return;
        }
        match op {
            // 0x76 routes to the same body as 0x74.
            0x76 => self.emit_var_load(&n, call_ctx),
            // Direct opcode emits.
            0x77 => {
                self.stream.emit_word(0xb0);
            }
            0x78 => {
                self.stream.emit_word(0x8b7);
            }
            0x79 => {
                self.stream.emit_word(0xcb7);
            }
            0x75 => todo!("emit_expr: node 0x75 (typed emit variant) - Phase 3/4"),
            0x7a => todo!("emit_expr: node 0x7a (name/call path) - Phase 3"),
            0x7d => todo!("emit_expr: node 0x7d (property/array typed emit) - Phase 3/4"),
            0x7e => todo!("emit_expr: node 0x7e (property typed emit) - Phase 3/4"),
            0x7f => todo!("emit_expr: node 0x7f (chained comparison) - Phase 3"),
            // Other [0x75, 0x88): emit nothing.
            _ => {}
        }
    }

    /// The variable-load body shared by node types 0x74 and 0x76. Base opcode by
    /// call context; the byref flag (0x8000) sets bit 0x8000 and emits the opcode
    /// raw. The bound symbol child (word[4]) carries the type-info operand at
    /// +0x12; word[5] is the type context.
    fn emit_var_load(&mut self, n: &crate::node::RawNode, call_ctx: u32) {
        let mut opcode = VARLOAD_OP_BY_CALLCTX[call_ctx as usize];
        let f_raw = (n.flags() & 0x8000) != 0;
        if f_raw {
            opcode |= 0x8000;
        }
        self.emit_typed_pcode(opcode, n.lhs(), n.word(5) as i32, f_raw);
    }
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::node::NodeArena;

    /// A Currency literal node (type 2) carrying `value` in its 8-byte payload.
    fn currency(arena: &mut NodeArena, value: i64) -> NodeRef {
        let b = value.to_le_bytes();
        let w4 = u32::from_le_bytes([b[0], b[1], b[2], b[3]]);
        let w5 = u32::from_le_bytes([b[4], b[5], b[6], b[7]]);
        arena.alloc(NodeArena::node(2, 0, w4, w5, 0, 0))
    }

    fn emit(arena: &NodeArena, root: NodeRef) -> Vec<u8> {
        let mut e = Emitter::new(arena);
        e.emit_expr(root, 0);
        e.into_bytes()
    }

    fn lit_bytes(value: i64) -> Vec<u8> {
        let mut v = vec![0xa9, 0x00];
        v.extend_from_slice(&value.to_le_bytes());
        v
    }

    #[test]
    fn currency_literal_emits_opcode_and_eight_bytes() {
        let mut a = NodeArena::new();
        let n = currency(&mut a, 12345);
        assert_eq!(emit(&a, n), lit_bytes(12345));
    }

    #[test]
    fn binop_is_postfix_lhs_rhs_op() {
        // Node type 0x1f -> OP_BY_NODE_TYPE[0x1f] = 0x0000. Operands are literals.
        let mut a = NodeArena::new();
        let l = currency(&mut a, 1);
        let r = currency(&mut a, 2);
        let op = a.alloc(NodeArena::node(0x1f, 0, l.0, r.0, 0, 0));
        let mut expect = lit_bytes(1);
        expect.extend(lit_bytes(2));
        expect.extend_from_slice(&OP_BY_NODE_TYPE[0x1f].to_le_bytes()); // 0x0000
        assert_eq!(emit(&a, op), expect);
    }

    #[test]
    fn binop_right_associative_nesting() {
        // a ^ (b ^ c): outer.lhs = a, outer.rhs = (inner: b,c). Right-assoc
        // grouping shows up as: load a; load b; load c; op_inner; op_outer.
        let mut a = NodeArena::new();
        let na = currency(&mut a, 1);
        let nb = currency(&mut a, 2);
        let nc = currency(&mut a, 3);
        let inner = a.alloc(NodeArena::node(0x1f, 0, nb.0, nc.0, 0, 0));
        let outer = a.alloc(NodeArena::node(0x1f, 0, na.0, inner.0, 0, 0));

        let opw = OP_BY_NODE_TYPE[0x1f].to_le_bytes();
        let mut expect = lit_bytes(1);
        expect.extend(lit_bytes(2));
        expect.extend(lit_bytes(3));
        expect.extend_from_slice(&opw); // inner op
        expect.extend_from_slice(&opw); // outer op
        assert_eq!(emit(&a, outer), expect);
    }

    #[test]
    fn unary_type6_emits_operand_then_op() {
        // Type 6 -> OP_BY_NODE_TYPE[6] = 0x0015.
        let mut a = NodeArena::new();
        let operand = currency(&mut a, 7);
        let u = a.alloc(NodeArena::node(6, 0, operand.0, 0, 0, 0));
        let mut expect = lit_bytes(7);
        expect.extend_from_slice(&OP_BY_NODE_TYPE[6].to_le_bytes());
        assert_eq!(OP_BY_NODE_TYPE[6], 0x0015);
        assert_eq!(emit(&a, u), expect);
    }

    #[test]
    fn is_operator_emits_operands_then_0x14() {
        let mut a = NodeArena::new();
        let l = currency(&mut a, 1);
        let r = currency(&mut a, 2);
        let is = a.alloc(NodeArena::node(0x36, 0, l.0, r.0, 0, 0));
        let mut expect = lit_bytes(1);
        expect.extend(lit_bytes(2));
        expect.extend_from_slice(&0x14u16.to_le_bytes());
        assert_eq!(emit(&a, is), expect);
    }

    /// A bound-symbol node whose type-info operand (node+0x12, high half of
    /// word[4]) is `info`.
    fn sym_with_type_info(arena: &mut NodeArena, info: u16) -> NodeRef {
        arena.alloc(NodeArena::node(0, 0, (info as u32) << 16, 0, 0, 0))
    }

    #[test]
    fn emit_type_info3_emits_node_plus_0x12() {
        let mut a = NodeArena::new();
        let n = sym_with_type_info(&mut a, 0x0042);
        let mut e = Emitter::new(&a);
        e.emit_type_info3(n);
        assert_eq!(e.into_bytes(), &[0x42, 0x00]);
    }

    #[test]
    fn typed_pcode_raw_emits_opcode_then_operand() {
        let mut a = NodeArena::new();
        let n = sym_with_type_info(&mut a, 0x1234);
        let mut e = Emitter::new(&a);
        e.emit_typed_pcode(0x0020, n, 0, true); // f_raw: opcode used verbatim
        assert_eq!(e.into_bytes(), &[0x20, 0x00, 0x34, 0x12]);
    }

    #[test]
    fn typed_pcode_resolves_type_ctx_bits() {
        // type ctx 5 -> TYPE_CTX_BITS[5] = 0x08; opcode 0x0020 -> 0x08<<10 | 0x20
        // = 0x2020. Then the type-info operand.
        let mut a = NodeArena::new();
        let n = sym_with_type_info(&mut a, 0x0001);
        let mut e = Emitter::new(&a);
        e.emit_typed_pcode(0x0020, n, 5, false);
        assert_eq!(TYPE_CTX_BITS[5], 0x08);
        assert_eq!(e.into_bytes(), &[0x20, 0x20, 0x01, 0x00]);
    }

    #[test]
    fn typed_pcode_type_ctx_zero_no_indirection_leaves_opcode() {
        // type ctx 0 with no indirect-type flag returns the opcode unchanged.
        let mut a = NodeArena::new();
        let n = sym_with_type_info(&mut a, 0x00aa);
        let mut e = Emitter::new(&a);
        e.emit_typed_pcode(0x0030, n, 0, false);
        assert_eq!(e.into_bytes(), &[0x30, 0x00, 0xaa, 0x00]);
    }

    #[test]
    fn single_literal_converts_double_and_emits_opcode4() {
        // Type 3, tag 10: 0.5 as f64 -> f32 (exact), opcode 0xb3 + 4-byte single.
        let mut a = NodeArena::new();
        let b = 0.5f64.to_le_bytes();
        let w4 = u32::from_le_bytes([b[0], b[1], b[2], b[3]]);
        let w5 = u32::from_le_bytes([b[4], b[5], b[6], b[7]]);
        let n = a.alloc(NodeArena::node(3, 10, w4, w5, 0, 0));
        let mut e = Emitter::new(&a);
        e.emit_expr(n, 0);
        let mut expect = vec![0xb3, 0x00];
        expect.extend_from_slice(&0.5f32.to_le_bytes());
        assert_eq!(e.into_bytes(), expect);
    }

    #[test]
    fn string_literal_emits_opcode_len_and_data() {
        // Type 4: opcode 0xb6, length word, then the (even-rounded) bytes.
        let mut a = NodeArena::new();
        let off = a.alloc_blob(b"Hi"); // even length, no padding
        let n = a.alloc(NodeArena::node(4, 0, 0, 2, off, 0));
        let mut e = Emitter::new(&a);
        e.emit_expr(n, 0);
        assert_eq!(e.into_bytes(), &[0xb6, 0x00, 0x02, 0x00, b'H', b'i']);
    }

    #[test]
    fn string_literal_odd_length_pads_from_blob() {
        // Odd length 3: count word holds 3, 4 bytes copied (incl. the trailing
        // source byte).
        let mut a = NodeArena::new();
        let off = a.alloc_blob(b"abc\0");
        let n = a.alloc(NodeArena::node(4, 0, 0, 3, off, 0));
        let mut e = Emitter::new(&a);
        e.emit_expr(n, 0);
        assert_eq!(e.into_bytes(), &[0xb6, 0x00, 0x03, 0x00, b'a', b'b', b'c', 0x00]);
    }

    /// A bound local-variable reference: a 0x74 node whose child (word[4]) is a
    /// bound symbol carrying `slot` at +0x12, with type context `type_ctx`.
    fn var_ref(arena: &mut NodeArena, slot: u16, type_ctx: u16) -> NodeRef {
        let sym = sym_with_type_info(arena, slot);
        arena.alloc(NodeArena::node(0x74, 0, sym.0, type_ctx as u32, 0, 0))
    }

    #[test]
    fn variable_load_emits_typed_opcode_and_slot() {
        // callCtx 0 -> VARLOAD_OP_BY_CALLCTX[0] = 0x20; type context 1 ->
        // TYPE_CTX_BITS[1] = 0x02 -> resolved 0x02<<10 | 0x20 = 0x820; operand =
        // slot 0x10.
        let mut a = NodeArena::new();
        let v = var_ref(&mut a, 0x0010, 1);
        let mut e = Emitter::new(&a);
        e.emit_expr(v, 0);
        assert_eq!(VARLOAD_OP_BY_CALLCTX[0], 0x0020);
        assert_eq!(e.into_bytes(), &[0x20, 0x08, 0x10, 0x00]);
    }

    #[test]
    fn variable_load_byref_sets_high_bit_and_emits_raw() {
        // Byref flag 0x8000: opcode |= 0x8000 and is emitted raw (no type-ctx
        // folding). callCtx 0 -> 0x20 | 0x8000 = 0x8020; operand = slot.
        let mut a = NodeArena::new();
        let sym = sym_with_type_info(&mut a, 0x0007);
        let mut node = NodeArena::node(0x74, 0, sym.0, 1, 0, 0);
        node.w[1] = 0x8000; // byref flag in word[1]
        let v = a.alloc(node);
        let mut e = Emitter::new(&a);
        e.emit_expr(v, 0);
        assert_eq!(e.into_bytes(), &[0x20, 0x80, 0x07, 0x00]);
    }

    #[test]
    fn variables_a_pow_b_pow_c_right_associative() {
        // r = a ^ b ^ c with variable operands. Right-assoc tree a ^ (b ^ c)
        // emits load a; load b; load c; pow; pow. Each variable load is two words
        // [0x820][slot]; pow is OP_BY_NODE_TYPE[0x1f].
        let mut a = NodeArena::new();
        let va = var_ref(&mut a, 0x10, 1);
        let vb = var_ref(&mut a, 0x14, 1);
        let vc = var_ref(&mut a, 0x18, 1);
        let inner = a.alloc(NodeArena::node(0x1f, 0, vb.0, vc.0, 0, 0));
        let outer = a.alloc(NodeArena::node(0x1f, 0, va.0, inner.0, 0, 0));

        let load = |slot: u8| vec![0x20, 0x08, slot, 0x00];
        let pow = OP_BY_NODE_TYPE[0x1f].to_le_bytes();
        let mut expect = Vec::new();
        expect.extend(load(0x10)); // a
        expect.extend(load(0x14)); // b
        expect.extend(load(0x18)); // c
        expect.extend_from_slice(&pow); // b ^ c
        expect.extend_from_slice(&pow); // a ^ (b ^ c)

        let mut e = Emitter::new(&a);
        e.emit_expr(outer, 0);
        assert_eq!(e.into_bytes(), expect);
    }

    #[test]
    fn node_0x76_routes_to_variable_load() {
        // 0x76 shares the 0x74 var-load body. callCtx 0, type ctx 1 -> 0x820 + slot.
        let mut a = NodeArena::new();
        let sym = sym_with_type_info(&mut a, 0x0012);
        let v = a.alloc(NodeArena::node(0x76, 0, sym.0, 1, 0, 0));
        let mut e = Emitter::new(&a);
        e.emit_expr(v, 0);
        assert_eq!(e.into_bytes(), &[0x20, 0x08, 0x12, 0x00]);
    }

    #[test]
    fn direct_opcode_nodes_0x77_0x78_0x79() {
        for (ty, word) in [(0x77u16, 0x00b0u16), (0x78, 0x08b7), (0x79, 0x0cb7)] {
            let mut a = NodeArena::new();
            let n = a.alloc(NodeArena::node(ty, 0, 0, 0, 0, 0));
            let mut e = Emitter::new(&a);
            e.emit_expr(n, 0);
            assert_eq!(e.into_bytes(), word.to_le_bytes(), "node 0x{ty:x}");
        }
    }

    #[test]
    fn node_0x87_emits_operand_then_op_0x1d() {
        // OP_BY_NODE_TYPE[0x87] = 0x001d.
        let mut a = NodeArena::new();
        let operand = currency(&mut a, 5);
        let n = a.alloc(NodeArena::node(0x87, 0, operand.0, 0, 0, 0));
        assert_eq!(OP_BY_NODE_TYPE[0x87], 0x001d);
        let mut expect = lit_bytes(5);
        expect.extend_from_slice(&OP_BY_NODE_TYPE[0x87].to_le_bytes());
        let mut e = Emitter::new(&a);
        e.emit_expr(n, 0);
        assert_eq!(e.into_bytes(), expect);
    }

    #[test]
    fn node_0x91_emits_typed_0x104_then_operand() {
        // 0x91 -> typed emit of opcode 0x104 over child, type ctx 0, no
        // indirection -> opcode 0x104 unchanged, then child+0x12 operand.
        let mut a = NodeArena::new();
        let sym = sym_with_type_info(&mut a, 0x0003);
        let n = a.alloc(NodeArena::node(0x91, 0, sym.0, 0, 0, 0));
        let mut e = Emitter::new(&a);
        e.emit_expr(n, 0);
        assert_eq!(e.into_bytes(), &[0x04, 0x01, 0x03, 0x00]);
    }

    #[test]
    fn typed_load_0xe_path_uses_type_shift_table() {
        // Node 0xe with type tag 2: shift = TYPE_SHIFT[2] = 0x18,
        // word = (0x18 << 10) | 0x58 = 0x6058.
        let mut a = NodeArena::new();
        let operand = currency(&mut a, 9);
        let load = a.alloc(NodeArena::node(0x0e, 2, operand.0, 0, 0, 0));
        let expected_word = ((0x18u32 << 10) | 0x58) as u16;
        assert_eq!(expected_word, 0x6058);
        let mut expect = lit_bytes(9);
        expect.extend_from_slice(&expected_word.to_le_bytes());
        assert_eq!(emit(&a, load), expect);
    }
}
