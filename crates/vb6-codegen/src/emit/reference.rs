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
    /// base by kind: local = `0x1e0`, argument = `0x210`, kind-7 = `0x240`,
    /// kind-6 = `0x3d0`, kind-3 = `0x270`, kind-5 = `0x2a0`, kind-0xa = `0x190`.
    pub fn emit_reference(&mut self, desc: &RefDescriptor, n_op_in: i32, f_flags: u32, n_type_in: i32) {
        let mut n_op = n_op_in;
        let n_type = n_type_in;
        // Final-emit selector: 1 = opcode + 2-byte operand, 2 = opcode + operand
        // + trailing word, 0 = opcode only. The trailing-word path is only
        // reached via the deferred nType==0x17 chain, so it is always 1 here.
        let emit_mode: i16 = 1;
        // The operand word fed to the final emit. Kinds 1/2/6/7 (and the
        // default fallback) use the descriptor's canonical `operand`; a few
        // kinds source a different descriptor field instead — overridden
        // per-kind in the `opcode_base` match below.
        let mut operand_word: u16 = desc.operand;

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
            // kind 3: opcode base 0x270; the operand is the descriptor's `+6`
            // field (not the canonical `+0xa` operand), unconditional.
            3 => {
                opcode_base = 0x270;
                operand_word = desc.word6;
            }
            // kind 5: opcode base 0x2a0; same shape as kind 3 (operand from
            // `+6`, unconditional), different template constant.
            5 => {
                opcode_base = 0x2a0;
                operand_word = desc.word6;
            }
            // kind 0xa: opcode base 0x190. The `+6` field's low bit selects
            // the operand source: set -> the descriptor's `+8` field. Clear
            // -> a distinct fallback shape (the raw nType argument doubles as
            // both the opcode base and the operand) whose downstream byte
            // consumption was not captured in the available trace; gated
            // rather than guessed.
            0xa => {
                opcode_base = 0x190;
                if desc.word6 & 1 != 0 {
                    operand_word = desc.word8;
                } else {
                    unimplemented!(
                        "reference kind 0xa, +6 field bit 0 clear: the \
                         nType-as-operand fallback shape is not captured on disk"
                    );
                }
            }
            // kind 4: needs a second stashed operand word the descriptor has
            // no field for, plus a non-default finalize-emit mode whose
            // downstream byte consumption was not captured in the available
            // trace; gated rather than guessed.
            4 => unimplemented!(
                "reference kind 4: second operand word / finalize-emit mode \
                 not captured on disk"
            ),
            // kinds 8/9/0xb handled above (operator-reference direct emission).
            8 | 9 | 0xb => unreachable!(),
            // default (kind 0 / out of range): both the opcode base and the
            // operand are the raw nType argument itself.
            _ => {
                opcode_base = n_type;
                operand_word = n_type as u16;
            }
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
                    // inverted flag bits 0x1000 and 0x800. The table's confirmed
                    // exact extent is offsets 2..=9 (RT_TYPE_OFFSET's raw range
                    // also includes 0/1/10/12/14, which real destination/source
                    // type combinations CAN produce here — this is a genuine,
                    // unresolved gap, not a dead default; needs its own
                    // extraction, not guessed).
                    use crate::tables::EXPR_STORE_CONV;
                    if !(2..=9).contains(&off3) {
                        unimplemented!(
                            "store conversion (0x8000 path) for type offset {off3}: \
                             outside the conversion table's confirmed 2..=9 extent"
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
    /// freshly-built descriptor: kind 3, every other field zeroed, the same
    /// `nOp`/`fFlags`/`nType` forwarded unchanged. nOp 5 (and nOp 6 without the
    /// flag) finishes cleanly here without recursing.
    pub(super) fn expr2_finalize_tail(&mut self, n_op: i32, f_flags: u32, n_type: i32, flags1: u8) {
        if n_op == 5 || n_op == 6 {
            if n_op != 6 {
                return;
            }
            if flags1 & 4 == 0 {
                return;
            }
        }
        let coercion_desc = RefDescriptor {
            kind: 3,
            ..RefDescriptor::default()
        };
        self.emit_reference(&coercion_desc, n_op, f_flags, n_type);
    }

    /// Port of `ConvertExpressionType`: propagates coercion flags/offsets
    /// between a type-descriptor node and an expression-result node (both a
    /// 4-word shape: `word0`=kind 1..=0xb, `word1` low16=flags, `word1`
    /// high16=an accumulator field, `word2` high16=another accumulator
    /// field, `word3`=an opaque payload word), re-entering the value emitter
    /// (`emit_reference`, n_op 6) on several arms. `ctx_flag_c` is byte
    /// `+0xc` of the compiler context — the same field [`resolve_ident_ref`]
    /// already threads through; only bit 1 (`&2`) is read here.
    pub(super) fn convert_expression_type(
        &mut self,
        type_desc: &mut RawNode,
        expr_desc: &mut RawNode,
        ctx_flag_c: u8,
    ) {
        match type_desc.w[0] as i32 {
            // kind 1: bit 2 clear -> accumulate into word2-high, merge flag
            // bits 0/2 from expr_desc, copy type_desc into expr_desc wholesale.
            // Bit 2 set falls into the shared "settype4" block (kind 2/4's arm).
            1 => {
                if flags_lo(type_desc) as u8 >> 2 & 1 == 0 {
                    let b = flags1_byte(expr_desc);
                    let new_hi2 = hi2(type_desc).wrapping_add(hi1(expr_desc));
                    set_hi2(type_desc, new_hi2);
                    set_flags_lo(type_desc, merge_flag_bits(flags_lo(type_desc), b));
                    copy_wholesale(expr_desc, type_desc);
                    return;
                }
                self.convert_expr_type_settype4(type_desc, expr_desc, ctx_flag_c);
            }
            // kind 2/4: bit 2 set -> the shared direct-emit block (kind 6's
            // "bit 2 set" arm); bit 2 clear -> the shared "settype4" block.
            2 | 4 => {
                if flags_lo(type_desc) as u8 >> 2 & 1 != 0 {
                    self.convert_expr_type_direct_emit(type_desc, ctx_flag_c);
                    return;
                }
                self.convert_expr_type_settype4(type_desc, expr_desc, ctx_flag_c);
            }
            // kind 3: its own accumulator field (word1-high, not word2-high).
            3 => {
                if flags_lo(type_desc) as u8 >> 2 & 1 == 0 {
                    let b = flags1_byte(expr_desc);
                    let new_hi1 = hi1(type_desc).wrapping_add(hi1(expr_desc));
                    set_hi1(type_desc, new_hi1);
                    set_flags_lo(type_desc, merge_flag_bits(flags_lo(type_desc), b));
                    copy_wholesale(expr_desc, type_desc);
                    return;
                }
                let desc = ref_descriptor_from_node(type_desc);
                self.emit_reference(&desc, 6, 0, 0);
            }
            // kind 5/7: direct re-emit, then a triple-gated 0x408 marker.
            5 | 7 => {
                let desc = ref_descriptor_from_node(type_desc);
                self.emit_reference(&desc, 6, 0, 0);
                if flags_lo(type_desc) as u8 >> 2 & 1 == 0 {
                    return;
                }
                if ctx_flag_c & 2 == 0 {
                    return;
                }
                if flags_lo(type_desc) & 1 == 0 {
                    return;
                }
                self.emit_value2(0x408);
            }
            // kind 6: same accumulator field as kind 1; bit 2 set falls to
            // the shared direct-emit block.
            6 => {
                if flags_lo(type_desc) as u8 >> 2 & 1 == 0 {
                    let b = flags1_byte(expr_desc);
                    let new_hi2 = hi2(type_desc).wrapping_add(hi1(expr_desc));
                    set_hi2(type_desc, new_hi2);
                    set_flags_lo(type_desc, merge_flag_bits(flags_lo(type_desc), b));
                    copy_wholesale(expr_desc, type_desc);
                    return;
                }
                self.convert_expr_type_direct_emit(type_desc, ctx_flag_c);
            }
            // kinds 8/9/10/0xb: bare re-emit, no flag propagation.
            8 | 9 | 10 | 0xb => {
                let desc = ref_descriptor_from_node(type_desc);
                self.emit_reference(&desc, 6, 0, 0);
            }
            // default (any other kind): bare return, no emit at all.
            _ => {}
        }
    }

    /// The shared "settype4" block (kind 1's bit-2-set arm and kind 2/4's
    /// bit-2-clear arm): when `expr_desc` is itself kind 3 or has its `+4`
    /// bit `0x08` set, recast `type_desc` to kind 4, fold `expr_desc`'s
    /// word6 into `type_desc`'s word3 (low 16 bits), merge flag bits 0/2,
    /// then copy `type_desc` into `expr_desc` wholesale. Otherwise falls
    /// through to a bare re-emit (the switch's own `break` case).
    fn convert_expr_type_settype4(&mut self, type_desc: &mut RawNode, expr_desc: &mut RawNode, _ctx_flag_c: u8) {
        if expr_desc.w[0] as i32 == 3 || flags1_byte(expr_desc) & 8 != 0 {
            let s = hi1(expr_desc);
            type_desc.w[0] = 4;
            let new_w3_lo = (type_desc.w[3] as u16 as i16).wrapping_add(s);
            type_desc.w[3] = (type_desc.w[3] & 0xffff_0000) | (new_w3_lo as u16 as u32);
            let b = flags1_byte(expr_desc);
            set_flags_lo(type_desc, merge_flag_bits(flags_lo(type_desc), b));
            copy_wholesale(expr_desc, type_desc);
            return;
        }
        let desc = ref_descriptor_from_node(type_desc);
        self.emit_reference(&desc, 6, 0, 0);
    }

    /// The shared direct-emit block (kind 2/4's bit-2-set arm and kind 6's
    /// bit-2-set arm): re-emit `type_desc`, then emit `0x408` when both
    /// `ctx_flag_c` bit 1 and `type_desc`'s own flags-byte bit 0 are set.
    fn convert_expr_type_direct_emit(&mut self, type_desc: &RawNode, ctx_flag_c: u8) {
        let desc = ref_descriptor_from_node(type_desc);
        self.emit_reference(&desc, 6, 0, 0);
        if ctx_flag_c & 2 != 0 && flags1_byte(type_desc) & 1 != 0 {
            self.emit_value2(0x408);
        }
    }

    /// Port of `EbEmitBinaryOpCode`'s scratch-descriptor construction and
    /// its re-entry into the value emitter (`emit_reference`, n_op 6):
    /// builds a kind-1 descriptor with the constant operand `8` when either
    /// of `type_desc_word4`'s bits `0x100`/`0x200` is clear, else a kind-5
    /// descriptor whose `word6` is either an interned type value (the
    /// common case) or the constant sentinel `0xffff` (`ctx_flag_c` bit 0
    /// set — a COM-bypass edge case, same convention as
    /// [`crate::resolver::fill_binding_desc`]'s own bypass).
    ///
    /// This ports only the function's FIRST half. Its second half — building
    /// the final kind-8 descriptor the caller consumes — reads a side-table
    /// located via a pointer stored in the module heap (byte offset `0x1c`)
    /// that this pipeline has no model for; that part is not ported (see
    /// the private note), never guessed.
    pub(super) fn emit_binary_op_code_temp_descriptor(
        &mut self,
        type_desc_word4: u16,
        ctx_flag_c: u8,
        type_value: u32,
    ) {
        if type_desc_word4 & 0x100 == 0 || type_desc_word4 & 0x200 == 0 {
            let desc = RefDescriptor {
                kind: 1,
                operand: 8,
                ..RefDescriptor::default()
            };
            self.emit_reference(&desc, 6, 0, 0);
        } else {
            let word6 = if ctx_flag_c & 1 == 0 {
                self.type_pool.extract_type_value2(type_value)
            } else {
                0xffff
            };
            let desc = RefDescriptor {
                kind: 5,
                word6,
                ..RefDescriptor::default()
            };
            self.emit_reference(&desc, 6, 0, 0);
        }
    }
}

// ── `ConvertExpressionType`'s 4-word node accessors ──────────────────────────
// word0 = kind; word1 low16 = flags (byte `+4` = `flags1_byte`, the same
// convention `RefDescriptor::flags1` uses); word1 high16 = the `+6`
// accumulator field; word2 high16 = the `+0xa` accumulator field (also
// `RefDescriptor::operand`'s convention); word3 = an opaque payload word.

fn flags_lo(n: &RawNode) -> u16 {
    (n.w[1] & 0xffff) as u16
}

fn flags1_byte(n: &RawNode) -> u8 {
    (n.w[1] & 0xff) as u8
}

fn set_flags_lo(n: &mut RawNode, v: u16) {
    n.w[1] = (n.w[1] & 0xffff_0000) | v as u32;
}

fn hi1(n: &RawNode) -> i16 {
    (n.w[1] >> 16) as u16 as i16
}

fn hi2(n: &RawNode) -> i16 {
    (n.w[2] >> 16) as u16 as i16
}

fn set_hi1(n: &mut RawNode, v: i16) {
    n.w[1] = (n.w[1] & 0xffff) | ((v as u16 as u32) << 16);
}

fn set_hi2(n: &mut RawNode, v: i16) {
    n.w[2] = (n.w[2] & 0xffff) | ((v as u16 as u32) << 16);
}

fn copy_wholesale(dst: &mut RawNode, src: &RawNode) {
    dst.w[0] = src.w[0];
    dst.w[1] = src.w[1];
    dst.w[2] = src.w[2];
    dst.w[3] = src.w[3];
}

/// Copy bits 0 and 2 from `b` into `a` (the repeated `(b^a)&mask ^ a` idiom).
fn merge_flag_bits(a: u16, b: u8) -> u16 {
    let a1 = (((b ^ (a as u8)) & 4) as u16) ^ a;
    (((a1 as u8 ^ b) & 1) as u16) ^ a1
}

/// Build the [`RefDescriptor`] `emit_reference` expects from a 4-word node
/// (the same field mapping `flags1_byte`/`hi1`/`hi2` document above).
fn ref_descriptor_from_node(n: &RawNode) -> RefDescriptor {
    RefDescriptor {
        kind: n.w[0] as i32,
        operand: (n.w[2] >> 16) as u16,
        word6: (n.w[1] >> 16) as u16,
        word8: (n.w[2] & 0xffff) as u16,
        flags1: (n.w[1] & 0xff) as u8,
    }
}
