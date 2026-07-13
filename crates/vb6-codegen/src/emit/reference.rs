use super::*;

impl<'a> Emitter<'a> {
    // ── Resolved-reference emission ──────────────────────────────────────────

    /// Port of `EbSetupBinaryOperation` (`EbResolveReference2`'s 0x69 path):
    /// traverse + emit the two operands, then build the operator descriptor
    /// (kind 9/0xa/0xb) the value-emitter consumes.
    ///
    /// * flag `0x2000` set → kind `0xb` (typed operator): `word6` = `node[8]`,
    ///   `word8` = the interned type value of `node[6]`.
    /// * else `node[8]` low word == 1 → kind `0xa` (with flag `0x8000` and a
    ///   `node[6]` type: `word6` bit 0 set, `word8` = its size).
    /// * else → kind `9` (operator): `word6` = `node[8]`.
    /// A `0x160000`-region node additionally sets `flags1` bits `0x05`.
    pub(super) fn emit_setup_binary_operation(&mut self, node: NodeRef) -> RefDescriptor {
        let n = *self.arena.get(node);
        self.traverse_node_tree(NodeRef(n.w[5]), 1);
        self.emit_expr(NodeRef(n.w[4]), 1);

        let flags = n.w[1];
        let w8 = (n.w[8] & 0xffff) as u16;
        let mut desc = RefDescriptor::default();
        if flags & 0x2000 == 0 {
            if w8 as i16 == 1 {
                desc.kind = 0xa;
                if flags & 0x8000 != 0 && n.w[6] != 0 {
                    desc.word6 |= 1;
                    desc.word8 = crate::resolver::get_type_size3(self.arena, n.w[6]) as u16;
                }
            } else {
                desc.kind = 9;
                desc.word6 = w8;
            }
        } else {
            desc.kind = 0xb;
            desc.word6 = w8;
            desc.word8 = self.type_pool.extract_type_value2(n.w[6]) as u16;
        }
        if n.w[0] & 0xffff_0000 == 0x16_0000 {
            desc.flags1 |= 5;
        }
        desc
    }

    /// Emit a resolved reference (typed local / argument load or store) from a
    /// [`RefDescriptor`] — the routine that turns a resolved reference into the
    /// typed load/store opcode.
    ///
    /// This is a **pure** function of the descriptor and the emit parameters; it
    /// needs no symbol table. The descriptor's `kind`, `operand`, and the
    /// `f_flags` / `n_type` parameters are produced upstream by the reference
    /// resolver (the vb6-sema bridge). `n_op` selects the operation: 1 = value
    /// load, 4 = store; others as below.
    ///
    /// Only descriptor kinds and sub-paths that do not reach not-yet-built
    /// machinery are implemented; the rest are `unimplemented!()`. The opcode
    /// base by kind: local = `0x1e0`, argument = `0x210`, kind-7 = `0x240`.
    pub fn emit_reference(&mut self, desc: &RefDescriptor, n_op_in: i32, f_flags: u32, n_type_in: i32) {
        let mut n_op = n_op_in;
        let n_type = n_type_in;
        // Final-emit selector: 1 = opcode + 2-byte operand, 2 = opcode + operand
        // + trailing word, 0 = opcode only. The trailing-word path is only
        // reached via the deferred nType==0x17 chain, so it is always 1 here.
        let emit_mode: i16 = 1;
        let operand_word: u16 = desc.operand;

        // Operator-reference descriptor kinds (8/9/0xb) emit their opcode
        // directly from the descriptor (no resolver-supplied base) and finish
        // through the shared finalize tail.
        match desc.kind {
            // kind 8: typed coercion reference. Opcode `0x3ca`/`0x3cb` by nOp,
            // operand = descriptor `+6`; then the `+8` low word and the `+10`
            // operand word.
            8 => {
                self.emit_opcode2(((n_op != 5) as i32 + 0x3ca) as usize, desc.word6);
                self.emit_word2(desc.word8);
                self.emit_word2(desc.operand);
                self.expr2_finalize_tail(n_op, f_flags, n_type, desc.flags1);
                return;
            }
            // kind 9: operator reference. nOp is normalized first; opcode
            // `0x18d`/`0x18e` by nOp, operand = descriptor `+6`.
            9 => {
                Self::map_operator_type(n_type, &mut n_op);
                self.emit_opcode2(((n_op == 5) as i32 + 0x18d) as usize, desc.word6);
                self.expr2_finalize_tail(n_op, f_flags, n_type, desc.flags1);
                return;
            }
            // kind 0xb: typed operator reference. nOp normalized; opcode
            // `0x406`/`0x407` by nOp, operand = descriptor `+8` low word; then
            // the `+6` word.
            0xb => {
                Self::map_operator_type(n_type, &mut n_op);
                self.emit_opcode2(((n_op == 5) as i32 + 0x406) as usize, desc.word8);
                self.emit_word2(desc.word6);
                self.expr2_finalize_tail(n_op, f_flags, n_type, desc.flags1);
                return;
            }
            _ => {}
        }

        let opcode_base: i32;

        match desc.kind {
            // kind 1: local variable.
            1 => opcode_base = 0x1e0,
            // kind 2: argument/parameter. A ByRef slot (word6 bit 0) forces the
            // ByRef operation before the operation switch.
            2 => {
                opcode_base = 0x210;
                if desc.word6 & 1 != 0 {
                    n_op = 2;
                }
            }
            // kind 7: indirect module-level variable. Same ByRef promotion.
            7 => {
                opcode_base = 0x240;
                if desc.word6 & 1 != 0 {
                    n_op = 2;
                }
            }
            // kind 6: indirect module-level variable, ByRef form (the
            // by-reference counterpart of kind 7). Unconditional — no ByRef
            // promotion branch.
            6 => opcode_base = 0x3d0,
            // kinds 3/4/5/0xa: the opcode base is supplied by the resolver's
            // call chain, not available without the full module context.
            3 | 4 | 5 | 0xa => unimplemented!(
                "reference kind {}: needs the resolver-supplied opcode base from \
                 the full module compilation context",
                desc.kind
            ),
            // kinds 8/9/0xb handled above (operator-reference direct emission).
            8 | 9 | 0xb => unreachable!(),
            // default: needs the resolver-supplied opcode base.
            _ => unimplemented!(
                "reference kind {} (default): needs the resolver-supplied opcode \
                 base from the full module compilation context",
                desc.kind
            ),
        }

        // When nType==0x12 the usage-context flag may promote it to 0x17; that
        // flag is not threaded through this entry point.
        if n_type == 0x12 {
            unimplemented!(
                "reference nType 0x12: needs the usage-context flag (not threaded \
                 through emit_reference)"
            );
        }
        // Map nType 0x11/0x12 with nOp in {1,2,3} → nOp 5.
        if (n_type == 0x12 || n_type == 0x11) && (n_op == 1 || n_op == 2 || n_op == 3) {
            n_op = 5;
        }

        let opcode_index: i32 = match n_op {
            1 => {
                if f_flags & 0x4000 != 0 {
                    // Object / type-expression path.
                    if n_type == 0x10 {
                        0x23f
                    } else if n_type == 0x17 {
                        unimplemented!(
                            "Variant-with-type expression: needs the type-pool emit \
                             path"
                        );
                    } else {
                        0x262
                    }
                } else if f_flags & 0x1000 != 0 {
                    0x1b2
                } else {
                    let off = RT_TYPE_OFFSET[n_type as usize];
                    if off == 10 {
                        opcode_base | 4
                    } else if off == 9 {
                        opcode_base | 1
                    } else {
                        opcode_base | off
                    }
                }
            }
            // nOp 2 shares the 0x4000 path with nOp 1.
            2 => {
                if f_flags & 0x4000 != 0 {
                    if n_type == 0x10 {
                        0x23f
                    } else if n_type == 0x17 {
                        unimplemented!(
                            "Variant-with-type expression: needs the type-pool emit \
                             path"
                        );
                    } else {
                        0x262
                    }
                } else if f_flags & 0x1000 != 0 {
                    0x1b2
                } else {
                    let mut off = RT_TYPE_OFFSET[n_type as usize];
                    if off == 10 {
                        off = 4;
                    } else if off == 9 {
                        off = 1;
                    }
                    let mut v = off | opcode_base;
                    if off == 3 || off == 4 {
                        v += 6;
                    }
                    v
                }
            }
            3 => {
                if f_flags & 0x1000 != 0 {
                    (if f_flags & 0x2000 != 0 { 1 } else { 0 }) + 0x1b3
                } else if n_type == 0xf {
                    opcode_base + 0xd
                } else {
                    let off = RT_TYPE_OFFSET[n_type as usize];
                    if off == 10 {
                        opcode_base | 4
                    } else if off == 9 {
                        opcode_base | 1
                    } else {
                        opcode_base | off
                    }
                }
            }
            4 => {
                let off3 = RT_TYPE_OFFSET[n_type as usize];
                let base_off = if off3 == 10 {
                    4
                } else if off3 == 9 {
                    1
                } else {
                    off3
                };
                let mut v = (opcode_base + 0x10) | base_off;
                if f_flags & 0x8000 == 0 {
                    if f_flags & 0x20 != 0 {
                        v = (if f_flags & 0x1000 != 0 { 1 } else { 0 }) + 0x1b5;
                    } else if f_flags & 0x80 != 0 {
                        v += 6;
                    } else if f_flags & 0x200 != 0 {
                        v = (if v != 0x1f6 { 0x11 } else { 0 }) + 0x3f8;
                    } else if f_flags & 0x400 != 0 {
                        v = 0x3f7;
                    } else if f_flags & 0x1000 != 0 {
                        if n_type == 0x10 || n_type == 0xf {
                            v += 10;
                        }
                    } else if f_flags & 0x800 != 0 {
                        v = 0x439;
                    } else if off3 == 3 || off3 == 4 {
                        // A field whose RAW (pre-remap) RT_TYPE_OFFSET class
                        // is directly 3 or 4 needs +6 on the plain scalar-
                        // store index — the same adjustment the sibling
                        // nOp-2 branch above already applies. Gated on the
                        // RAW `off3`, not the post-remap `base_off`: a field
                        // whose raw class is 10 (remapped to base_off == 4)
                        // must NOT get +6 — `ref_store_currency` oracle-
                        // confirms that case stays at the unadjusted index
                        // (byte `0x72`). A UDT Double field store (`t.C =
                        // 3`), whose raw class is directly 4, oracle-
                        // confirms it DOES need +6 (index `0x1fa`, byte
                        // `0x74`, `RT_STORE_BY_CTX`'s own Double entry) — no
                        // capture had exercised a raw-off3-4 field through
                        // this branch until an 8-byte UDT field reached the
                        // real pipeline (only Integer/Long UDT fields were
                        // tested before).
                        v += 6;
                    }
                } else if f_flags & 0x20 == 0 {
                    // Store with conversion: opcode = base + 0x10 +
                    // EXPR_STORE_CONV[type offset][sub], where sub combines the
                    // inverted flag bits 0x1000 and 0x800. Reached only for type
                    // offsets 2..=9 (the table's valid domain).
                    use crate::tables::EXPR_STORE_CONV;
                    if !(2..=9).contains(&off3) {
                        unimplemented!(
                            "store conversion (0x8000 path) for type offset {off3}: \
                             outside the conversion table domain"
                        );
                    }
                    let inv12 = 1 - ((f_flags >> 0xc) & 1) as i32;
                    let inv11 = 1 - ((f_flags >> 0xb) & 1) as i32;
                    let sub = (2 * inv12 + inv11) as usize;
                    let conv = EXPR_STORE_CONV[(off3 - 2) as usize][sub] as i32;
                    v = conv + 0x10 + opcode_base;
                } else {
                    v = (if f_flags & 0x800 != 0 { 2 } else { 0 }) + 0x1b7;
                }
                // (f_flags & 0x40) with a local descriptor: opcode remap for
                // some non-scalar-store variants. `0x1f2` (the plain scalar
                // Long/Integer/small-type store) does NOT remap — oracle-
                // confirmed (e2e_udt_field_scalar_access: `t.X = 1` emits a
                // bare `71 <offset>` store, no trailing reload; the earlier
                // `0x1f2 => { emit 0x1f2; 0x1e2 }` arm produced a spurious
                // extra `6c <offset>` load that no capture had caught until
                // this fixture reached the real pipeline). The other arms
                // (String/Variant-shaped classes) are unverified by any
                // oracle fixture yet — left as ported, not re-derived here.
                if f_flags & 0x40 != 0 && desc.kind == 1 {
                    v = match v {
                        0x1f6 => 0x263,
                        0x1f7 => 0x264,
                        0x1f8 => 0x29c,
                        0x1ff => 0x29d,
                        0x200 => {
                            self.emit_opcode2(v as usize, operand_word);
                            0x1ec
                        }
                        0x201 => {
                            self.emit_opcode2(v as usize, operand_word);
                            0x1e7
                        }
                        other => other,
                    };
                }
                v
            }
            5 => {
                if f_flags & 0x1000 == 0 {
                    opcode_base + 0xc
                } else {
                    0x1b2
                }
            }
            6 => opcode_base + 0xb,
            _ => unimplemented!("reference operation {} not handled", n_op),
        };

        // The trailing-word emit mode (opcode + operand + extra word) is only
        // reached via the deferred nType==0x17 chain; emit_mode is always 1 here.
        if emit_mode == 0 {
            self.emit_value2(opcode_index as usize);
        } else {
            self.emit_opcode2(opcode_index as usize, operand_word);
            if emit_mode == 2 {
                self.stream.emit_word(desc.word8);
            }
        }
    }

    /// Normalize an operator nOp: for a Variant / fixed-string operand type
    /// (`0x12` / `0x11`) the load/store/coerce operations (1/2/3) collapse to the
    /// Variant operation (5).
    pub(super) fn map_operator_type(n_type: i32, n_op: &mut i32) {
        if (n_type == 0x12 || n_type == 0x11) && (*n_op == 1 || *n_op == 2 || *n_op == 3) {
            *n_op = 5;
        }
    }

    /// The shared finalize tail of the operator-reference value-emitter kinds
    /// (8/9/0xb). For a value/store operation (nOp not 5/6, or nOp 6 with the
    /// descriptor's `+4` bit `0x04` set) it re-enters the value emitter with a
    /// freshly-built coercion descriptor — that recursion's opcode base is
    /// supplied by the resolver's register state and is not reproducible without
    /// the full module context, so it stays gated. nOp 5 (and nOp 6 without the
    /// flag) finishes cleanly here.
    pub(super) fn expr2_finalize_tail(&mut self, n_op: i32, _f_flags: u32, _n_type: i32, flags1: u8) {
        if n_op == 5 || n_op == 6 {
            if n_op != 6 {
                return;
            }
            if flags1 & 4 == 0 {
                return;
            }
        }
        unimplemented!(
            "value-emitter finalize tail (EbBuildExprDescriptor): the re-entry \
             coercion descriptor's opcode base is resolver-supplied; reached for \
             nOp {n_op}"
        );
    }
}
