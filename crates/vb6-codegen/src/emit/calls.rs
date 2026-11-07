use super::*;

impl<'a> Emitter<'a> {
    // ── Call sites ───────────────────────────────────────────────────────────

    /// Coercion type-code map: `10 -> 4`, `9 -> 1`, everything else unchanged.
    pub(super) fn map_type_code3(t: i32) -> i32 {
        match t {
            10 => 4,
            9 => 1,
            x => x,
        }
    }

    /// The behaviour-flag byte (`+0x1d`) of the dispatch record selected for a
    /// call of the given convention kind / by-reference mode.
    pub(super) fn call_record_flag_1d(kind: i32, byref: i32) -> u8 {
        use crate::tables::{RT_CALL_CONV_RECORDS, RT_CALL_KIND_CLASS, RT_CALL_SPECIAL_RECORD};
        if byref == 1 && kind == 4 {
            RT_CALL_SPECIAL_RECORD[0x1d]
        } else {
            let class = (RT_CALL_KIND_CLASS[kind as usize] & 3) as usize;
            RT_CALL_CONV_RECORDS[class * 0x1e + 0x1d]
        }
    }

    /// Emit a call site from a resolved [`CallDescriptor`].
    ///
    /// Direct emission of the by-reference (method/Sub) call path: compute the
    /// call opcode from the convention kind and node flags, emit the argument
    /// list and the callee reference, then the call opcode (and the member-id
    /// word on the dispatch path). Paths that need machinery not yet built — the
    /// early-bound dispatch lookup, the type-expression argument path, the
    /// value-returning (ByVal) result-type path, and the coercion sequence — are
    /// `unimplemented!()` with what they require.
    pub fn emit_call(&mut self, desc: &CallDescriptor, context: u32) -> u32 {
        let kind = desc.kind;
        let byref = desc.byref;
        let flags = desc.flags;
        let byte5 = (flags >> 8) & 0xff;
        let rec_1d = Self::call_record_flag_1d(kind, byref);

        // Early-bound dispatch (byte 1 bit 0x20): needs the dispatch-type lookup.
        if byref == 1 && flags & 0x2000 != 0 {
            unimplemented!(
                "call early-bound dispatch (flag 0x2000): needs the dispatch-type \
                 lookup; Phase 5"
            );
        }

        let (call_op, emit_mode) = Self::call_opcode_and_mode(desc, rec_1d);

        // Sized prefix (dispatch record requests a size operand).
        if rec_1d & 0x10 != 0 {
            self.emit_opcode2(0x265, desc.size);
        }
        self.emit_call_arguments(desc, byte5);
        // Value-returning (ByVal) calls take the result-type table path.
        if byref != 1 {
            unimplemented!(
                "value-returning (ByVal) call: needs the result-type table; Phase 5"
            );
        }
        // By-reference (method / Sub) call: emit the callee reference, the call
        // opcode, then (dispatch path) the member-id word.
        if byte5 & 0x20 == 0 {
            self.emit_expr(desc.callee, 6);
        }
        self.emit_value2(call_op as usize);
        if rec_1d & 0x10 != 0 {
            self.emit_word2(desc.size);
        }
        self.finish_call_by_mode(desc, rec_1d, emit_mode, context);
        0
    }

    /// Emit the argument list of a call site, if any.
    fn emit_call_arguments(&mut self, desc: &CallDescriptor, byte5: u32) {
        if desc.arg_list.0 == 0 {
            return;
        }
        if byte5 & 0x40 == 0 {
            self.process_linked_list(desc.arg_list, 3);
        } else {
            self.traverse_node_tree(desc.arg_list, 3);
        }
    }

    /// The call opcode and finalize emit mode for a call descriptor.
    fn call_opcode_and_mode(desc: &CallDescriptor, rec_1d: u8) -> (i32, i32) {
        let flags = desc.flags;
        let call_op = call_type_code(desc.kind, desc.byref == 1, flags & 0x2000 != 0) as i32;
        if flags & 0x8000 == 0 {
            return (Self::call_op_by_region(desc, call_op, rec_1d), 0);
        }
        if flags & 0x1000 == 0 || call_op != 800 {
            (map_call_type_code(call_op as u16) as i32, 1)
        } else {
            (0x41e, 1)
        }
    }

    /// Refine the call opcode by the call node's type region (non-dispatch path).
    fn call_op_by_region(desc: &CallDescriptor, call_op: i32, rec_1d: u8) -> i32 {
        let flags = desc.flags;
        if flags & 0x800 != 0 {
            unimplemented!(
                "call type-expression argument (flag 0x800): needs the type-pool \
                 expression path; Phase 5"
            );
        }
        match desc.node_word0 & 0xffff_0000 {
            0x20000 => {
                if flags & 0x200 != 0 && desc.byref == 1 {
                    unimplemented!(
                        "call coercion sequence (0x20000 + flag 0x200, ByRef): needs \
                         the coercion path; Phase 5"
                    );
                }
                call_op + 0xc
            }
            0x120000 => call_op | Self::map_type_code3(2),
            _ => Self::call_op_by_type_tag(desc, call_op, rec_1d),
        }
    }

    /// Merge the result-type class of the call node's type tag into the opcode.
    fn call_op_by_type_tag(desc: &CallDescriptor, call_op: i32, rec_1d: u8) -> i32 {
        let type_tag = (desc.node_word0 as i32) >> 16;
        let m = Self::map_type_code3(RT_TYPE_OFFSET[type_tag as usize]);
        let merged = call_op | m;
        let wide_scalar = matches!(type_tag, 10..=0xc) && rec_1d & 8 != 0;
        if wide_scalar && (m == 3 || m == 4) {
            merged + 6
        } else {
            merged
        }
    }

    /// Complete the call site according to the finalize emit mode.
    fn finish_call_by_mode(&mut self, desc: &CallDescriptor, rec_1d: u8, emit_mode: i32, context: u32) {
        match emit_mode {
            // Common path: the finalize step emits the member-id word, then the
            // per-type tail.
            0 => self.finalize_call(desc, rec_1d, desc.member_id, context),
            // Dispatch path: the member-id word is emitted here; the finalize
            // step's trailing word is the type-pool index of node[9].
            1 => {
                let trailing = self.type_pool.extract_type_value2(desc.node9);
                self.emit_word2(desc.member_id);
                self.finalize_call(desc, rec_1d, trailing, context);
            }
            // Result-descriptor path: unreachable while the type-expression
            // argument gate (flag 0x800) is the only site that selects this
            // mode — it throws before the mode is ever set to 2.
            _ => unimplemented!(
                "call result-descriptor path (emit mode 2): needs the struct-size / \
                 member-type model; Phase 5"
            ),
        }
    }

    /// The call-site finalize step: emit the caller-supplied `trailing` word,
    /// then — driven by the call node's type region — either re-enter the
    /// emitter on a synthetic dispatch-type node (gated; needs the type-pool
    /// allocator) or emit the per-type validation opcode. The `0xf0000` /
    /// dispatch-record `0x08` cases validate; everything else terminates the
    /// call site after the trailing word.
    pub(super) fn finalize_call(&mut self, desc: &CallDescriptor, rec_1d: u8, trailing: u16, context: u32) {
        self.emit_word2(trailing);
        let region = desc.node_word0 & 0xffff_0000;
        if region == 0x140000 {
            unimplemented!(
                "call finalize type-node path (region 0x140000): builds a 0x60 \
                 dispatch-type node and re-enters the emitter; needs the type-pool \
                 allocator; Phase 5"
            );
        }
        if rec_1d & 8 != 0 || region == 0xf_0000 {
            let type_tag = (desc.node_word0 as i32) >> 16;
            self.emit_validate_type_operation(type_tag, 0, context);
        }
    }
}
