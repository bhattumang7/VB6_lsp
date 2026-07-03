use super::*;

impl<'a> Emitter<'a> {
    // ── Binary operations ────────────────────────────────────────────────────

    /// Emit a binary operation: both operands (each in context 2) then the
    /// type-class-selected opcode.
    ///
    /// The opcode index is the operation's base ([`RT_BINOP_BASE`]) plus a type
    /// offset. Two dispatch modes, selected by `RT_DISPATCH_FLAG[op] & 0x10`:
    /// * clear → arithmetic: offset from the **node's own** type tag.
    /// * set → comparison / string: offset from the **left operand's** type tag,
    ///   with special cases for the `0x72` / type-3 / object(`0xf`) / `0xd` forms.
    pub(super) fn emit_binary_operation(&mut self, n: &RawNode, context: u32) -> u32 {
        self.emit_expr(n.lhs(), 2);
        if n.w[5] != 0 {
            self.emit_expr(n.rhs(), 2);
        }
        let op = n.opcode() as usize;
        let mut opcode_index = RT_BINOP_BASE[op] as i32;
        let type_tag = n.type_tag();

        if RT_DISPATCH_FLAG[op] & 0x10 == 0 {
            // Arithmetic: use the node's own type tag.
            let mut offset = RT_TYPE_OFFSET[type_tag as usize];
            if offset == 10 {
                offset = 4;
            } else if offset == 9 {
                offset = 1;
            }
            opcode_index += offset;
        } else {
            // Comparison / string: use the left operand's type tag.
            let lhs = *self.arena.get(n.lhs());
            let lhs_word0 = lhs.w[0];
            if (n.w[0] & 0xffff) == 0x72
                || (n.w[0] & 0xffff_0000) != 0x0003_0000
                || (lhs_word0 & 0xffff_0000) != 0x000f_0000
            {
                let rhs = *self.arena.get(n.rhs());
                let rhs_tag = (rhs.w[0] as i32) >> 16;
                if (lhs_word0 & 0xffff_0000) == 0x000d_0000
                    && (rhs_tag == 10 || rhs_tag == 0xb || rhs_tag == 0xc)
                {
                    opcode_index += 0xc;
                } else {
                    let lhs_tag = (lhs_word0 as i32) >> 16;
                    let mut offset = RT_TYPE_OFFSET[lhs_tag as usize];
                    if offset == 10 {
                        offset = 4;
                    } else if offset == 9 {
                        offset = 1;
                    }
                    opcode_index += offset;
                    if (n.w[1] >> 8) & 0x80 != 0 {
                        opcode_index += 2;
                    }
                }
            } else if (n.w[1] >> 8) & 0x80 == 0 {
                opcode_index += 10;
            } else {
                opcode_index += 0xb;
            }
        }

        if (n.w[0] & 0xffff_0000) == 0x000f_0000 {
            let size = self.emit_get_type_size3(n.w[6]);
            self.emit_opcode2(opcode_index as usize, size as u16);
        } else {
            self.emit_value2(opcode_index as usize);
        }
        self.emit_validate_type_operation(type_tag, 0, context)
    }

    // ── Per-type validation ──────────────────────────────────────────────────

    /// Emit the per-type validation / conversion opcode after an operation.
    pub(super) fn emit_validate_type_operation(&mut self, op_type: i32, variant: i32, type_flags: u32) -> u32 {
        match op_type {
            2 => {
                if variant == 0x17 {
                    self.emit_value2(0x18b);
                }
            }
            10 | 0xb | 0xc => {
                if type_flags != 3 && type_flags != 1 {
                    return 1;
                }
                if op_type == 10 {
                    self.emit_value2(0x189);
                    return 0;
                }
                if op_type > 10 && op_type < 0xd {
                    self.emit_value2(0x18a);
                    return 0;
                }
                self.emit_value2(type_flags as usize);
                return 0;
            }
            0xf => {
                if type_flags == 3 {
                    self.emit_value2(0x18c);
                    return 0;
                }
            }
            _ => {}
        }
        0
    }

    // ── Type-descriptor size lookup ──────────────────────────────────────────

    /// Resolve the byte size of a UDT / object type from its type descriptor.
    ///
    /// `type_desc` is an arena reference (a node index) to a type-descriptor
    /// record. The record's `word[0]` holds the descriptor kind; for a
    /// fixed-size type (kind `4`) the resolved byte size sits in the low half of
    /// `word[4]`. Any other kind — or a null reference — has no fixed size and
    /// yields the `0xffff_ffff` sentinel (emitted as a `0xffff` operand).
    ///
    /// The size value is the type's packed byte size, resolved upstream (by the
    /// type-resolution pass) and carried in the descriptor; this routine only
    /// reads it back, matching the runtime's behaviour exactly.
    pub(super) fn emit_get_type_size3(&self, type_desc: u32) -> u32 {
        if type_desc == 0 {
            return 0xffff_ffff;
        }
        let desc = self.arena.get(NodeRef(type_desc));
        if desc.w[0] == 4 {
            desc.w[4] & 0xffff
        } else {
            0xffff_ffff
        }
    }

    /// Emit the trailing type-validation for an object / UDT operation: a `0x202`
    /// guard when the node is object-typed with its flag-byte bit `0x80` set,
    /// then the per-type validation opcode.
    pub(super) fn emit_object_type(&mut self, n: &RawNode, context: u32) -> u32 {
        if (n.w[1] >> 8) & 0x80 != 0 && (n.w[0] & 0xffff_0000) == 0xf0000 {
            self.emit_value2(0x202);
        }
        self.emit_validate_type_operation(n.type_tag(), 0, context)
    }

    // ── Type coercion / conversion ───────────────────────────────────────────

    /// Sum the byte sizes contributed by an argument list, walking the `0x37` /
    /// `0x33` list nodes (child = `word[4]`, next = `word[5]`). Each element's
    /// size is `RT_RESULT_TYPE[element type tag]`, except an `0xf0000`-region
    /// element contributes 4 when `flag` is set. (Port of `EbCalculateStructSize`.)
    pub(super) fn emit_calculate_struct_size(&self, list: NodeRef, flag: bool) -> i32 {
        let mut total = 0i32;
        let mut cur = list;
        while cur.0 != 0 {
            let n = *self.arena.get(cur);
            let opc = (n.w[0] & 0xffff) as u16;
            let elem = if opc == 0x37 || opc == 0x33 {
                let e = NodeRef(n.w[4]);
                cur = NodeRef(n.w[5]);
                e
            } else {
                let e = cur;
                cur = NodeRef(0);
                e
            };
            if elem.0 != 0 {
                let e = *self.arena.get(elem);
                if e.w[0] & 0xffff_0000 == 0xf_0000 && flag {
                    total += 4;
                } else {
                    let tag = (e.w[0] as i32) >> 16;
                    total += RT_RESULT_TYPE[tag as usize] as i32;
                }
            }
        }
        total
    }

    /// Emit a structured type coercion (`EbEmitTypeCoercion4`): traverse the
    /// source list, optionally emit the `0x411` prefix, then the coercion opcode
    /// (operand = the interned type value of the source descriptor) followed by
    /// the accumulated struct size.
    pub(super) fn emit_type_coercion4(&mut self, target: i32, src: NodeRef) {
        let n = *self.arena.get(src);
        let child = NodeRef(n.w[5]);
        let mut size = self.emit_calculate_struct_size(child, true) + 4;
        if child.0 != 0 {
            self.traverse_node_tree(child, 1);
        }
        if target == 0x40d && (n.w[1] >> 8) & 0x40 != 0 {
            self.emit_value2(0x411);
            size += 4;
        }
        let descriptor = *self.arena.get(NodeRef(n.w[4]));
        let v = self.type_pool.extract_type_value2(descriptor.w[4]);
        self.emit_opcode2(target as usize, v);
        self.emit_word2(size as u16);
    }

    /// Emit a type conversion (`EbEmitTypeConversion2`): emit the operand list
    /// (linked-list walk when implicit, tree traversal when explicit), then the
    /// conversion opcode with an operand taken from the source type node — either
    /// the literal type code (`0x01`) or the interned type value (`0x6f`).
    pub(super) fn emit_type_conversion2(&mut self, target: i32, src: NodeRef, explicit: bool) {
        let n = *self.arena.get(src);
        let inner = *self.arena.get(NodeRef(n.w[5]));
        let list = NodeRef(inner.w[5]);
        if explicit {
            self.traverse_node_tree(list, 1);
        } else {
            self.process_linked_list(list, 1);
        }
        let mut p = *self.arena.get(NodeRef(inner.w[4]));
        if p.w[0] & 0xffff == 0x11 {
            p = *self.arena.get(NodeRef(p.w[4]));
        }
        let opc = (p.w[0] as u16 as i16) as i32;
        if opc == 1 {
            self.emit_opcode2(target as usize, p.w[4] as u16);
        } else if opc == 0x6f {
            let v = self.type_pool.extract_type_value2(p.w[4]);
            self.emit_opcode2(target as usize, v);
        }
    }

    /// Emit expression code (`EbEmitExpressionCode2`): emit the child (context 1),
    /// then for an `0xf`-type child a sized opcode (`0xeb`/`0xed`), and always the
    /// per-type validation. `store` selects the store variant (+2). The
    /// `0x10`-type child branch is gated (its decompiled opcode is unverified).
    pub(super) fn emit_expression_code2(&mut self, store: bool, node: NodeRef, type_info: u32) {
        let n = *self.arena.get(node);
        self.emit_expr(NodeRef(n.w[4]), 1);
        let child_tag = (self.arena.get(NodeRef(n.w[4])).w[0] as i32) >> 16;
        let mut variant = 0;
        if child_tag == 0xf {
            let sz = self.emit_get_type_size3(n.w[6]);
            self.emit_opcode2(0xeb + if store { 2 } else { 0 }, sz as u16);
            variant = 0x17;
        } else if child_tag == 0x10 {
            // Opcode 0xec (load) / 0xee (store); the variant stays 0 here.
            self.emit_value2(0xec + if store { 2 } else { 0 });
        }
        self.emit_validate_type_operation((n.w[0] as i32) >> 16, variant, type_info);
    }

    /// Emit a typed node (`FUN_0fabd27e`): dispatch on the node's opcode.
    /// `0x12` re-emits its child (context 5 promoted to 6); a plain node emits
    /// with context 1 (object region) or the given mode; an opcode-`5` wrapper
    /// emits the object guard + per-type validation when its child is at/above
    /// the `0x12` boundary. The `0x60`/`0x69` resolved-reference forms need the
    /// reference resolver and remain gated.
    pub(super) fn emit_typed_node(&mut self, node: NodeRef, mode: u32) {
        let n = *self.arena.get(node);
        let kind = n.w[0] & 0xffff;
        if kind == 0x12 {
            let m = if mode == 5 { 6 } else { mode };
            self.emit_expr(NodeRef(n.w[4]), m);
            return;
        }
        if kind == 0x60 || kind == 0x69 {
            unimplemented!(
                "typed node 0x60/0x69 reference resolution (FUN_0fab33e9/FUN_0fab397a); Phase 5"
            );
        }
        if kind != 5 {
            let m = if n.w[0] & 0xffff_0000 == 0xf_0000 { 1 } else { mode };
            self.emit_expr(node, m);
            return;
        }
        // opcode 5 wrapper (the EbGetTypeSize3 read here is discarded).
        let child = *self.arena.get(NodeRef(n.w[4]));
        if child.w[0] & 0xffff >= 0x12 {
            if (n.w[1] >> 8) & 0x40 != 0 {
                self.emit_value2(0x202);
            }
            self.emit_validate_type_operation(0xf, 0, mode);
        }
    }

    /// Read the type-word operand a complex-binop sub-node contributes: the
    /// source type node (`p`) gives a literal type code (`0x01`), a pooled value
    /// (`0x11` indirect / `0x6f`), or the fallback value.
    pub(super) fn complex_binop_type_word(&mut self, p: NodeRef, fallback: u32) -> u16 {
        let pn = *self.arena.get(p);
        let s9 = (pn.w[0] as u16 as i16) as i32;
        if s9 == 1 {
            pn.w[4] as u16
        } else if s9 == 0x11 {
            self.type_pool.extract_type_value2(self.arena.get(NodeRef(pn.w[4])).w[4])
        } else if s9 == 0x6f {
            self.type_pool.extract_type_value2(pn.w[4])
        } else {
            fallback as u16
        }
    }

    /// Emit a complex binary operation (`EbEmitComplexBinaryOp`): three flag-driven
    /// shapes (`0x4000`/`0x2000`/`0x8000`) over a nested operand tree, each
    /// emitting the operand, an opcode, and the type/operand words.
    pub(super) fn emit_complex_binary_op(&mut self, node: NodeRef) {
        let n = *self.arena.get(node);
        let flags = n.w[1] & 0xffff;
        if flags & 0x4000 == 0 {
            let n5 = *self.arena.get(NodeRef(n.w[5]));
            let uvar2 = self.arena.get(NodeRef(n5.w[4])).w[4] as u16;
            let an = *self.arena.get(NodeRef(n5.w[5]));
            let uvar3 = self.arena.get(NodeRef(an.w[4])).w[4] as u16;
            let i10n = *self.arena.get(NodeRef(an.w[5]));
            let uvar6 = i10n.w[4];
            let bn = *self.arena.get(NodeRef(i10n.w[5]));
            let uvar11 = self.complex_binop_type_word(NodeRef(bn.w[4]), uvar6);
            let bw5 = *self.arena.get(NodeRef(bn.w[5]));
            let uvar4 = self.arena.get(NodeRef(bw5.w[4])).w[4] as u16;
            self.traverse_node_tree(NodeRef(bw5.w[5]), 1);
            self.emit_expr(NodeRef(uvar6), 1);
            let opcode = ((flags & 0x8000) | 0x1c7_0000) >> 0xf;
            self.emit_opcode2(opcode as usize, uvar4);
            self.emit_word2(uvar11);
            self.emit_word2(uvar3);
            self.emit_word2(uvar2);
            return;
        }
        let n5 = *self.arena.get(NodeRef(n.w[5]));
        let b0 = *self.arena.get(NodeRef(n5.w[5]));
        let uvar8 = b0.w[4];
        let uvar2 = self.arena.get(NodeRef(n5.w[4])).w[4] as u16;
        let b1 = *self.arena.get(NodeRef(b0.w[5]));
        let uvar3 = self.arena.get(NodeRef(b1.w[4])).w[4] as u16;
        if flags & 0x2000 == 0 {
            self.traverse_node_tree(NodeRef(b1.w[5]), 1);
            self.emit_expr(NodeRef(uvar8), 1);
            let opcode = ((flags & 0x8000) | 0x1c8_0000) >> 0xf;
            self.emit_opcode2(opcode as usize, uvar3);
            self.emit_word2(uvar2);
            return;
        }
        let b2 = *self.arena.get(NodeRef(b1.w[5]));
        let word = self.complex_binop_type_word(NodeRef(b2.w[4]), b1.w[5]);
        self.traverse_node_tree(NodeRef(b2.w[5]), 1);
        self.emit_expr(NodeRef(uvar8), 1);
        let opcode = if flags & 0x8000 != 0 { 0x43c } else { 0x43b };
        self.emit_opcode2(opcode, uvar3);
        self.emit_word2(uvar2);
        self.emit_word2(word);
    }

    /// Post-operation operand dispatch (`EbDispatchOpcodeToEmitter`): walk `depth`
    /// `word[5]` links, unwrap a `0x37` list node (at depth 1) and a `0x11`/`0x2d`
    /// wrapper, then emit the operand with context 2 (or 5 when `type_tag` is
    /// `0x17`).
    pub(super) fn emit_dispatch_opcode(&mut self, start: NodeRef, depth: i32, type_tag: i32) {
        let mut node = start;
        let mut d = depth;
        while d > 1 {
            node = NodeRef(self.arena.get(node).w[5]);
            d -= 1;
        }
        if d == 1 && (self.arena.get(node).w[0] & 0xffff) == 0x37 {
            node = NodeRef(self.arena.get(node).w[4]);
        }
        let op0 = self.arena.get(node).w[0] & 0xffff;
        if op0 == 0x11 || op0 == 0x2d {
            node = NodeRef(self.arena.get(node).w[4]);
        }
        self.emit_expr(node, if type_tag != 0x17 { 2 } else { 5 });
    }

    /// Emit a method/instruction site (`EbEmitInstruction2`): for `0x6c`/`0x6d`
    /// nodes first emit the resolved argument operand, then the target reference
    /// (context 5 for an object, else 6), the instruction opcode, an optional
    /// argument-size word (flag `0x4000`), the call result-size word (`is_call`),
    /// the type-pool index or 4-byte payload, and the trailing member word.
    /// Walk a `0x37` wrapper chain emitting each child (and the final node) with
    /// `flags`, up to `depth` links; return the unconsumed tail. (`EbFindActualNode`.)
    pub(super) fn emit_find_actual_node(&mut self, mut node: NodeRef, flags: u32, mut depth: i32) -> NodeRef {
        while node.0 != 0 && (self.arena.get(node).w[0] & 0xffff) == 0x37 && depth != 0 {
            let nn = *self.arena.get(node);
            self.emit_expr(NodeRef(nn.w[4]), flags);
            node = NodeRef(nn.w[5]);
            depth -= 1;
        }
        if node.0 != 0 && depth != 0 {
            self.emit_expr(node, flags);
            node = NodeRef(0);
        }
        node
    }

    /// Walk a `0x37` chain emitting each element's trailing payload — a 4-byte
    /// value (`0x01` node) or a pooled type word (`0x6f`). (`EbTraverseExprTree3`.)
    pub(super) fn emit_traverse_expr_tree3(&mut self, mut node: NodeRef, mut depth: i32) {
        while node.0 != 0 && depth != 0 {
            let n = *self.arena.get(node);
            let (mut elem, next) = if n.w[0] & 0xffff == 0x37 {
                (NodeRef(n.w[4]), NodeRef(n.w[5]))
            } else {
                (node, NodeRef(0))
            };
            if self.arena.get(elem).w[0] & 0xffff == 0x11 {
                elem = NodeRef(self.arena.get(elem).w[4]);
            }
            let en = *self.arena.get(elem);
            let s = en.w[0] as u16 as i16;
            if s == 1 {
                self.emit_dword(en.w[4]);
            } else if s == 0x6f {
                let v = self.type_pool.extract_type_value2(en.w[4]);
                self.emit_word2(v);
            }
            node = next;
            depth -= 1;
        }
    }

    pub(super) fn emit_instruction2(&mut self, node: NodeRef, opcode: usize, has_arg: bool, is_call: bool) {
        let n = *self.arena.get(node);
        let op = n.w[0] & 0xffff;
        // Top-level argument emission (returns the unconsumed tail for the
        // trailing expr-tree pass).
        let uvar6 = if n.w[5] != 0 {
            self.emit_find_actual_node(NodeRef(n.w[5]), 3, (n.w[8] as i16) as i32)
        } else {
            NodeRef(0)
        };
        if op == 0x6c || op == 0x6d {
            let mut cur = NodeRef(n.w[5]);
            loop {
                let c = *self.arena.get(cur);
                if c.w[0] & 0xffff != 0x37 {
                    break;
                }
                cur = NodeRef(if c.w[5] != 0 { c.w[5] } else { c.w[4] });
            }
            self.emit_expr(cur, 3);
        }
        let child_region = self.arena.get(NodeRef(n.w[4])).w[0] & 0xffff_0000;
        let ctx = if child_region == 0xf_0000 { 5 } else { 6 };
        self.emit_expr(NodeRef(n.w[4]), ctx);
        self.emit_value2(opcode);
        if n.w[1] & 0x4000 != 0 {
            // word[8] high half (`node + 0x22`) is an argument count.
            let cnt = (n.w[8] >> 16) as i16 as i32;
            let mut v = if n.w[1] & 0x8000 == 0 { (cnt + 2) << 1 } else { (cnt << 2) + 6 };
            if is_call {
                v += 2;
            }
            self.emit_word2(v as u16);
        }
        if is_call {
            let sz = self.emit_get_type_size3(n.w[6]);
            self.emit_word2(sz as u16);
        }
        let byte5 = (n.w[1] >> 8) & 0xff;
        if byte5 & 0x80 == 0 {
            let v = self.type_pool.extract_type_value2(n.w[7]);
            self.emit_word2(v);
        } else {
            self.emit_dword(n.w[7]);
        }
        if (n.w[8] as i16) != 0 || has_arg {
            self.emit_word2(n.w[8] as u16);
        }
        if byte5 & 0x40 != 0 {
            self.emit_traverse_expr_tree3(uvar6, (n.w[8] >> 16) as i16 as i32);
        }
    }
}
