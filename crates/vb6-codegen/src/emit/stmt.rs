use super::*;

impl<'a> Emitter<'a> {
    // ── Output primitives ────────────────────────────────────────────────────

    /// Resolve `n_opc` through [`RT_OPCODE_BYTE`] and emit it: a single byte when
    /// `< 0xfb`, otherwise the escape byte followed by `n_opc as u8`.
    pub(super) fn emit_value2(&mut self, n_opc: usize) {
        let rt_byte = RT_OPCODE_BYTE[n_opc];
        if rt_byte < 0xfb {
            self.stream.emit_byte(rt_byte);
        } else {
            self.stream.emit_byte(rt_byte);
            self.stream.emit_byte(n_opc as u8);
        }
    }

    /// Emit the opcode byte(s) for `n_opc`, then a 2-byte little-endian operand.
    pub(super) fn emit_opcode2(&mut self, n_opc: usize, operand: u16) {
        self.emit_value2(n_opc);
        self.stream.emit_word(operand);
    }

    /// Emit a 4-byte little-endian value.
    pub(super) fn emit_dword(&mut self, value: u32) {
        self.stream.emit_bytes(&value.to_le_bytes());
    }

    /// Emit a bare 2-byte little-endian value (no opcode byte).
    pub(super) fn emit_word2(&mut self, value: u16) {
        self.stream.emit_word(value);
    }

    // ── Statement-list walks ─────────────────────────────────────────────────

    /// Walk a forward-linked statement list (opcodes `0x37` / `0x33`), emitting
    /// each list element's child (`word[4]`) in order with `mode`, then emit the
    /// trailing non-list node, if any.
    pub(super) fn process_linked_list(&mut self, list: NodeRef, mode: u32) {
        let mut cur = list;
        while cur.0 != 0 {
            let n = *self.arena.get(cur);
            let opcode = n.w[0] & 0xffff;
            if opcode != 0x37 && opcode != 0x33 {
                break;
            }
            self.emit_expr(NodeRef(n.w[4]), mode);
            cur = NodeRef(n.w[5]);
        }
        if cur.0 != 0 {
            self.emit_expr(cur, mode);
        }
    }

    /// Walk a statement list (opcodes `0x37` / `0x33`) and emit each element.
    /// List structure: `word[4]` = child statement, `word[5]` = next list node.
    /// Recurses on the sibling first, then emits the child — yielding
    /// right-to-left emission order for a forward list.
    pub fn traverse_node_tree(&mut self, node: NodeRef, n_mode: u32) {
        if node.0 == 0 {
            return;
        }
        let n = *self.arena.get(node);
        let mut active = node;
        if (n.w[0] & 0xffff) == 0x37 || (n.w[0] & 0xffff) == 0x33 {
            if n.w[5] != 0 {
                self.traverse_node_tree(NodeRef(n.w[5]), n_mode);
            }
            active = NodeRef(n.w[4]);
        }
        if active.0 != 0 {
            self.emit_expr(active, n_mode);
        }
    }

    // ── Literal emitters ─────────────────────────────────────────────────────

    /// Floating-point literal. `context == 2` selects the assign-context opcode
    /// variants. Single (type_tag 10) is converted to f32 and emitted as 4 bytes;
    /// Double/Date (11/12) emit the raw 8-byte f64. Returns 1 in assign context
    /// for the typed variants, else 0.
    pub(super) fn emit_float_literal(&mut self, n: &RawNode, context: u32) -> u32 {
        let type_tag = n.type_tag();
        let mut assign_ctx = 0u32;
        let mut emit_eight = true;
        let n_opc: usize;
        if type_tag == 10 {
            emit_eight = false;
            if context == 2 {
                assign_ctx = 1;
                n_opc = 0x3ba;
            } else {
                n_opc = 0x3b9;
            }
        } else if type_tag > 10 && type_tag < 0xd {
            if context == 2 {
                assign_ctx = 1;
                n_opc = 0x3bd;
            } else {
                n_opc = 0x3bc;
            }
        } else {
            n_opc = context as usize;
        }
        self.emit_value2(n_opc);
        if emit_eight {
            self.stream.emit_bytes(&n.literal8());
        } else {
            let f = n.literal_f64() as f32;
            self.stream.emit_bytes(&f.to_bits().to_le_bytes());
        }
        assign_ctx
    }

    /// String literal. Flag-byte bit `0x80` set → null string, emitted as a
    /// Long-zero (`0x3b8` + 4 zero bytes). Clear → a pooled string literal, which
    /// needs the type/string pool.
    pub(super) fn emit_string_literal(&mut self, n: &RawNode) {
        if (n.w[1] >> 8) & 0x80 == 0 {
            unimplemented!(
                "pooled string literal: needs the type/string pool; Phase 4"
            );
        }
        self.emit_value2(0x3b8);
        self.emit_dword(0);
    }
}
