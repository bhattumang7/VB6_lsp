use super::*;
use super::intrinsics::{explicit_conversion_bytes, unary_intrinsic_bytes};

/// Outcome of emitting one node kind: either the node is complete and yields a
/// sub-result, or it falls through to the common tail that emits one trailing
/// opcode.
enum Step {
    Done(u32),
    Tail(i32),
}

/// Completed emission with no sub-result.
const DONE: Step = Step::Done(0);

impl<'a> Emitter<'a> {
    // ── Node emission ────────────────────────────────────────────────────────

    /// Emit the runtime P-code for one node. `context` is the usage context
    /// (0 = normal value; 1/2/3/5/6 select specialized emission). The return
    /// value propagates a sub-result for the few cases that compute one (0
    /// otherwise).
    pub fn emit_expr(&mut self, node: NodeRef, context: u32) -> u32 {
        let n = *self.arena.get(node);
        let op = n.opcode();

        if let Some(result) = self.emit_synthetic_node(&n, op, context) {
            return result;
        }

        // Guard: opcodes outside `1..=0x73` emit nothing (opcode 0 wraps to
        // 0xffffffff and is also rejected).
        if (op - 1) as u32 > 0x72 {
            return 0;
        }

        match self.emit_node_kind(node, &n, op, context) {
            Step::Done(result) => result,
            // Common tail: emit the trailing opcode, then return 0.
            Step::Tail(tail_opcode) => {
                self.emit_value2(tail_opcode as usize);
                0
            }
        }
    }

    /// The raw node stored at a reference.
    fn node_at(&self, r: NodeRef) -> RawNode {
        *self.arena.get(r)
    }

    // ── Synthetic IR nodes ───────────────────────────────────────────────────

    /// Synthetic IR nodes (opcodes `>= 0x74`, which the dispatch guard rejects)
    /// whose byte form is produced directly. `None` for every other opcode.
    fn emit_synthetic_node(&mut self, n: &RawNode, op: i32, context: u32) -> Option<u32> {
        match op {
            // Synthetic typed-load IR nodes.
            0x74 | 0x76 => self.emit_var_load(n, context),
            // Synthetic ByRef parameter load: same layout as a local load node,
            // routed through the ByRef opcode (base + 0x14).
            0x75 => self.emit_byref_param_node(n),
            // Synthetic module-global load: module descriptor in the low u16 of
            // word[4], field offset in the high u16, type context word[5].
            0x77 => self.emit_global_node_load(n),
            // Synthetic operand coercion: widen a binary-operation operand to the
            // operation type. Node type tag = target type; word[4] = the operand.
            0x78 => self.emit_coerce_node(n),
            // Synthetic string literal: push a pooled string constant by index.
            0x79 => self.emit_opcode2(0x3bf, n.w[4] as u16),
            // Synthetic Static-local load: read from the procedure's static block.
            0x7b => self.emit_static_load(n),
            // Synthetic explicit type-conversion intrinsic.
            0x7c => self.emit_explicit_conversion_node(n),
            // Synthetic dedicated-opcode unary intrinsic.
            0x7d => self.emit_unary_intrinsic_node(n),
            // Synthetic function-call-in-expression: a pre-emitted call byte
            // blob (result left on the stack) at blob offset word[4], length
            // word[5]. Emit the blob verbatim.
            0x7e => {
                let bytes = self.arena.blob(n.w[4], n.w[5] as usize).to_vec();
                self.stream.emit_bytes(&bytes);
            }
            _ => return None,
        }
        Some(0)
    }

    /// Explicit conversion intrinsic (CInt/CLng/CSng/CDbl/CCur/CStr): emit the
    /// child operand, then the explicit-conversion opcode for (dest = node type
    /// tag, src = child type tag). Differs from the implicit assignment coercion.
    fn emit_explicit_conversion_node(&mut self, n: &RawNode) {
        self.emit_expr(NodeRef(n.w[4]), 2);
        let dest = n.type_tag();
        let src = self.arena.get(NodeRef(n.w[4])).type_tag();
        for &b in explicit_conversion_bytes(dest, src) {
            self.stream.emit_byte(b);
        }
    }

    /// Unary intrinsic (Len/Abs/Sgn/Int/Fix): emit the argument, then the
    /// intrinsic opcode selected by the argument's type (kind in word[5]).
    fn emit_unary_intrinsic_node(&mut self, n: &RawNode) {
        self.emit_expr(NodeRef(n.w[4]), 2);
        let arg = self.arena.get(NodeRef(n.w[4])).type_tag();
        for &b in unary_intrinsic_bytes(n.w[5], arg) {
            self.stream.emit_byte(b);
        }
    }

    // ── Opcode dispatch ──────────────────────────────────────────────────────

    /// Emit one regular node kind (opcodes `1..=0x73`).
    fn emit_node_kind(&mut self, node: NodeRef, n: &RawNode, op: i32, context: u32) -> Step {
        match op {
            1 => self.op_int_literal(n),
            2 => self.op_wide_literal(n),
            3 => Step::Done(self.emit_float_literal(n, context)),
            4 => {
                self.emit_string_literal(n);
                DONE
            }
            5 => self.op_typed_node(n, context),
            0xb => self.op_negate(n),
            0xc => {
                self.emit_expression_code2(false, node, context);
                DONE
            }
            0xd => {
                self.emit_expression_code2(true, node, context);
                DONE
            }
            0xe => self.op_load_assign(n, context),
            0xf => self.op_name_coerce(n),
            0x10 => {
                self.emit_expr(n.lhs(), 1);
                Step::Tail(0x135)
            }
            0x11 => {
                self.emit_typed_node(NodeRef(n.w[4]), 5);
                DONE
            }
            0x12 => self.op_member_deref(n, context),
            0x13 => {
                self.emit_expr(n.rhs(), 1);
                Step::Tail(0x397)
            }
            0x14 => self.op_object_cast(n),
            0x15 => self.op_select_by_child_tag(n),
            0x18 => self.op_concat(n, context),
            0x1a => self.op_power(n, context),
            // no-op group → emit nothing.
            0x1b | 0x1c | 0x2e | 0x35 | 0x3c | 0x3d | 0x40 | 0x5b | 0x62 | 0x64 | 0x6f
            | 0x70 | 0x71 => DONE,
            0x24 => self.op_is(n, context),
            0x2c => self.op_assign_stmt(n, context),
            0x2d => self.op_typed_assign(n),
            0x2f | 0x30 | 0x31 => self.op_sequence(n, context),
            // type coercion (opcode by flag byte bit 0x80).
            0x32 => {
                let target = if (n.w[1] >> 8) & 0x80 == 0 { 0x40d } else { 0x40e };
                self.emit_type_coercion4(target, node);
                DONE
            }
            // traverse the child list, emitting each element.
            0x33 => {
                self.traverse_node_tree(node, 1);
                DONE
            }
            0x34 => {
                self.emit_type_coercion4(0x40f, node);
                DONE
            }
            // `Like` operator.
            0x36 => {
                self.emit_expr(n.lhs(), 1);
                self.emit_expr(n.rhs(), 1);
                Step::Tail(0xd2)
            }
            // process the forward statement list.
            0x37 => {
                self.process_linked_list(node, context);
                DONE
            }
            0x38 => self.op_nested_member_size(n),
            0x39 => self.op_end_with_size(n),
            0x3a => self.traverse_then_tail(n, 0x1ca),
            0x3b => self.traverse_then_tail(n, 0x1cb),
            0x3e => self.op_inner_member_value(n),
            0x3f => self.op_member_size_list(n),
            0x41 => self.op_arg_list(n),
            0x42 | 0x43 => self.op_dispatch_type(n, op),
            0x44..=0x47 => self.op_conversion_dispatch(node, n, op),
            0x48..=0x4b => self.op_traverse_dispatch(n, op),
            0x4c => {
                self.emit_type_conversion2(0x35d, node, true);
                DONE
            }
            0x4d => {
                self.emit_expr(NodeRef(n.w[5]), 1);
                Step::Tail(0x23d)
            }
            0x4e => Step::Tail(0x23e),
            0x4f => self.traverse_then_tail(n, 0xfb),
            0x50 => self.traverse_then_tail(n, 0xfa),
            0x51 | 0x52 => self.op_operator_class(n, op),
            0x53 => {
                let tail = if (n.w[1] >> 8) & 0x40 == 0 { 0x1c0 } else { 0x1bf };
                self.traverse_then_tail(n, tail)
            }
            0x54 => {
                self.emit_type_coercion4(0x410, node);
                DONE
            }
            0x55 => {
                self.emit_type_conversion2(0x35e, node, true);
                DONE
            }
            0x56 => self.op_conversion_or_traverse(node, n),
            0x57 => self.op_traverse_validate(n, context),
            0x58 => self.op_traverse_sized(n, context),
            0x59 => self.traverse_then_tail(n, 0x1c1),
            0x5a => {
                self.emit_complex_binary_op(node);
                DONE
            }
            0x5c => Step::Tail(0x162),
            0x5d => unimplemented!("type-library-driven cast; Phase 6"),
            0x5e => {
                self.emit_expr(NodeRef(n.w[5]), 1);
                Step::Tail(0x40a)
            }
            0x5f => {
                self.emit_expr(NodeRef(n.w[5]), context);
                Step::Tail(0x40b)
            }
            0x60 | 0x69 => self.op_member_reference(node, n, op, context),
            0x61 => self.op_call_site(n, context),
            0x63 => self.op_pooled_member(n),
            0x65 => Step::Done(self.emit_expr(n.lhs(), context)),
            0x66 => {
                let v = self.type_pool.extract_type_value2(n.w[5]);
                self.emit_opcode2(0x2f4, v);
                DONE
            }
            0x67 => {
                self.emit_expr(NodeRef(n.w[4]), 5);
                let v = self.type_pool.extract_type_value2(n.w[5]);
                self.emit_opcode2(0x2f5, v);
                DONE
            }
            0x68 => self.op_member_ref_value(n, context),
            0x6a => self.op_member_call(n, context),
            0x6b => self.op_store_let(node, n, context),
            0x6c => self.op_copy_instruction(node, n),
            0x6d => self.op_instruction(node, n),
            0x6e => self.op_flag_instruction(node, n),
            0x72 => unimplemented!(
                "member type-node coercion: needs type-node construction; Phase 4"
            ),
            // emit opcode 0x266 with the child's word[4] low half.
            0x73 => {
                let cn = self.node_at(n.lhs());
                self.emit_opcode2(0x266, (cn.w[4] & 0xffff) as u16);
                DONE
            }
            // every other binary op (arithmetic, logical, comparison).
            _ => Step::Done(self.emit_binary_operation(n, context)),
        }
    }

    /// Traverse the child list in word[5], then fall through to `tail`.
    fn traverse_then_tail(&mut self, n: &RawNode, tail: i32) -> Step {
        self.traverse_node_tree(NodeRef(n.w[5]), 1);
        Step::Tail(tail)
    }

    // ── Literals ─────────────────────────────────────────────────────────────

    /// Integer literal.
    fn op_int_literal(&mut self, n: &RawNode) -> Step {
        match n.type_tag() {
            3 | 5 | 6 => {
                self.emit_small_int_literal(n);
                DONE
            }
            8 | 0x10 => {
                self.emit_value2(0x3b8);
                self.emit_dword(n.w[4]);
                DONE
            }
            0x16 => Step::Tail(0x163),
            _ => DONE,
        }
    }

    /// 16-bit integer literal: a one-byte form when it fits, else a word form.
    fn emit_small_int_literal(&mut self, n: &RawNode) {
        let v = n.w[4] as i32;
        if -0x81 < v && v < 0x80 {
            self.emit_value2(0x41a);
            self.stream.emit_byte(v as u8);
        } else {
            self.emit_opcode2(0x3b5, n.w[4] as u16);
        }
    }

    /// 8-byte literal — Currency (0x3bb → 0xf6) or, when tagged Date (0xc), an
    /// OLE date serial (0x3bd → 0xfa).
    fn op_wide_literal(&mut self, n: &RawNode) -> Step {
        let push = if n.type_tag() == 0xc { 0x3bd } else { 0x3bb };
        self.emit_value2(push);
        self.stream.emit_bytes(&n.literal8());
        DONE
    }

    // ── Typed / load / assign nodes ──────────────────────────────────────────

    /// Typed node. Below the 0x12 opcode boundary the child opcode is propagated
    /// as the sub-result; otherwise emit the object guard (flag byte 0x40) and
    /// the per-type validation.
    fn op_typed_node(&mut self, n: &RawNode, context: u32) -> Step {
        let child = self.node_at(NodeRef(n.w[4]));
        let child_op = child.w[0] & 0xffff;
        if child_op < 0x12 {
            return Step::Done(child_op);
        }
        if (n.w[1] >> 8) & 0x40 != 0 {
            self.emit_value2(0x202);
        }
        Step::Done(self.emit_validate_type_operation(0xf, 0, context))
    }

    /// Unary minus / negate (type-class-dependent opcode).
    fn op_negate(&mut self, n: &RawNode) -> Step {
        self.emit_expr(n.lhs(), 2);
        let lhs = self.node_at(n.lhs());
        let lhs_tag = (lhs.w[0] as i32) >> 16;
        let tail = match RT_TYPE_OFFSET[lhs_tag as usize] {
            10 => 0xf6,
            9 => 1 + 0xf2,
            other => other + 0xf2,
        };
        Step::Tail(tail)
    }

    /// Load / assign path. A non-object target always runs the body; an object
    /// target needs the object-reference emit path (deferred).
    fn op_load_assign(&mut self, n: &RawNode, context: u32) -> Step {
        let lhs_kind = (n.w[0] as i32) >> 16;
        // Object target: needs the object-reference emit path.
        if n.w[0] & 0xffff_0000 == 0xf0000 {
            unimplemented!(
                "assignment to an object-typed target: needs the \
                 object-reference emit path; Phase 4"
            );
        }
        // Emit the value being assigned (context 2).
        self.emit_expr(NodeRef(n.w[4]), 2);
        // 0x4000 flag: object Set assignment.
        if n.w[1] & 0x4000 != 0 {
            unimplemented!(
                "object Set assignment (0x4000): needs the object-reference \
                 emit path; Phase 4"
            );
        }
        let op_class = (n.w[1] >> 8) & 7;
        if op_class != 0 {
            return Step::Done(self.emit_assign_class(n, op_class, context));
        }
        // Plain assignment: emit the store opcode, then the per-type validation.
        self.emit_assign_op(n);
        Step::Done(self.emit_validate_type_operation(lhs_kind, 0, context))
    }

    /// Non-plain assignment classes: each emits its opcode, then the
    /// object-type validation.
    fn emit_assign_class(&mut self, n: &RawNode, op_class: u32, context: u32) -> u32 {
        match op_class {
            // UDT copy: sized copy opcode.
            1 => self.emit_sized_object_assign(n, 0x2fe, context),
            // Set (assign reference): fixed opcode.
            2 => {
                self.emit_value2(0x2fd);
                self.emit_object_type(n, context)
            }
            // Set with addref: sized opcode.
            3 => self.emit_sized_object_assign(n, 0x2f9, context),
            // Set with release: sized opcode.
            4 => self.emit_sized_object_assign(n, 0x2fa, context),
            // UDT move: sized opcode.
            5 => self.emit_sized_object_assign(n, 0x2fc, context),
            // Array Set/copy: builds and emits a synthetic type node — needs the
            // type-node construction path.
            6 => unimplemented!(
                "assignment op-class 6 (array Set/copy): needs type-node \
                 construction; Phase 4"
            ),
            // Me assignment: fixed opcode.
            7 => {
                self.emit_value2(0x41b);
                self.emit_object_type(n, context)
            }
            _ => unreachable!(),
        }
    }

    /// Sized assignment opcode followed by the object-type validation.
    fn emit_sized_object_assign(&mut self, n: &RawNode, opcode: usize, context: u32) -> u32 {
        let size = self.emit_get_type_size3(n.w[6]);
        self.emit_opcode2(opcode, size as u16);
        self.emit_object_type(n, context)
    }

    /// Name / coerce path — sub-dispatch on the 3-bit op class.
    fn op_name_coerce(&mut self, n: &RawNode) -> Step {
        match (n.w[1] >> 8) & 7 {
            0 => self.op_name_plain(n),
            1 => unimplemented!(
                "typed name reference: needs symbol / type resolution; Phase 3/4"
            ),
            2 => {
                self.emit_expr(n.lhs(), 3);
                DONE
            }
            // sized name reference: emit the child (context 3), then the sized
            // coercion opcode 0x2c6 with the node's resolved type size.
            3 => {
                self.emit_expr(n.lhs(), 3);
                let size = self.emit_get_type_size3(n.w[6]);
                self.emit_opcode2(0x2c6, size as u16);
                DONE
            }
            // in-place name reference: sets the child's flag bit 0x1000 before
            // emitting it (context 3). The emit-time arena is immutable, so the
            // in-place flag mutation is not yet supported.
            4 => unimplemented!(
                "in-place name reference: needs in-place child-node flag \
                 mutation, unsupported by the immutable emit arena"
            ),
            _ => DONE,
        }
    }

    /// Plain name reference (op class 0).
    fn op_name_plain(&mut self, n: &RawNode) -> Step {
        // Flag 0x8000 clear: emit the child directly with context 1.
        if n.w[1] & 0x8000 == 0 {
            self.emit_expr(n.lhs(), 1);
            return DONE;
        }
        // Flag set: inspect the child's opcode / flag byte.
        let child = n.lhs();
        let cn = self.node_at(child);
        if Self::name_child_passes_through(cn.w[0] & 0xffff, (cn.w[1] >> 8) & 0xff) {
            self.emit_expr(child, 3);
            return DONE;
        }
        self.emit_expr(child, 1);
        Step::Tail(0x18c)
    }

    /// Whether a flagged name reference's child is emitted as-is (context 3).
    fn name_child_passes_through(c_op: u32, c_flag: u32) -> bool {
        let member = c_op == 0x60 && c_flag & 0x20 != 0;
        let indexed = c_op == 0x69 && c_flag & 0x80 != 0;
        !member && !indexed && c_op != 0x5e
    }

    /// Member dereference.
    fn op_member_deref(&mut self, n: &RawNode, context: u32) -> Step {
        let child = n.lhs();
        let cn = self.node_at(child);
        let node_hi = n.w[0] & 0xffff_0000;
        let child_hi = cn.w[0] & 0xffff_0000;
        let child_op = cn.w[0] & 0xffff;
        if node_hi == 0x160000 && child_op == 0x67 {
            self.emit_expr(NodeRef(cn.w[4]), 5);
            let opcode = if context == 6 { 0x2f7 } else { 0x2f6 };
            let v = self.type_pool.extract_type_value2(cn.w[5]);
            self.emit_opcode2(opcode, v);
            return DONE;
        }
        if node_hi == 0xf0000 && child_hi == 0x170000 {
            self.emit_expr(child, 2);
            return DONE;
        }
        if node_hi != 0x170000 || child_hi != 0x170000 {
            if child_op == 0x3f {
                self.emit_expr(child, context);
            }
            return DONE;
        }
        self.emit_expr(child, context);
        Step::Tail(0x15c)
    }

    /// Object cast (self-contained) or type-library item.
    fn op_object_cast(&mut self, n: &RawNode) -> Step {
        let node_hi = n.w[0] & 0xffff_0000;
        if node_hi == 0xf0000 {
            self.emit_expr(n.lhs(), 3);
            return Step::Tail(0x3f9);
        }
        if node_hi != 0x170000 {
            return DONE;
        }
        unimplemented!("type-library item load; Phase 6");
    }

    /// Select the opcode by the child's type tag.
    fn op_select_by_child_tag(&mut self, n: &RawNode) -> Step {
        let lhs = self.node_at(n.lhs());
        if (lhs.w[0] & 0xffff_0000) == 0x000f_0000 {
            Step::Tail(0x42e)
        } else {
            Step::Tail(0x42d)
        }
    }

    // ── Binary-style operators ───────────────────────────────────────────────

    /// Concatenation-style binary op (only the 0xd-type form; every other form
    /// is a plain binary operation).
    fn op_concat(&mut self, n: &RawNode, context: u32) -> Step {
        if n.w[0] & 0xffff_0000 != 0x000d_0000 {
            return Step::Done(self.emit_binary_operation(n, context));
        }
        self.emit_expr(n.lhs(), 2);
        self.emit_expr(n.rhs(), 2);
        let rhs = self.node_at(n.rhs());
        if (rhs.w[0] & 0xffff_0000) == 0x0006_0000 {
            Step::Tail(0xd1)
        } else {
            Step::Tail(0xb3)
        }
    }

    /// Power operator.
    fn op_power(&mut self, n: &RawNode, context: u32) -> Step {
        self.emit_expr(n.lhs(), 2);
        self.emit_expr(n.rhs(), 2);
        if n.w[0] & 0xffff_0000 == 0x000f_0000 {
            let size = self.emit_get_type_size3(n.w[6]);
            self.emit_opcode2(0xce, size as u16);
        } else {
            self.emit_value2(0xcf);
        }
        Step::Done(self.emit_validate_type_operation(n.type_tag(), 0, context))
    }

    /// `Is` operator.
    fn op_is(&mut self, n: &RawNode, context: u32) -> Step {
        let type_tag = n.type_tag();
        self.emit_expr(n.lhs(), 1);
        self.emit_expr(n.rhs(), 1);
        if type_tag == 0xf {
            let size = self.emit_get_type_size3(n.w[6]);
            self.emit_opcode2(0xef, size as u16);
            self.emit_validate_type_operation(0xf, 0x17, context);
            return DONE;
        }
        if type_tag == 0x10 {
            self.emit_value2(0xf0);
        }
        self.emit_validate_type_operation(type_tag, 0x17, context);
        DONE
    }

    // ── Assignment statements ────────────────────────────────────────────────

    /// Assignment statement: emit the RHS value, resolve the LHS reference, then
    /// emit its store (nOp 4).
    fn op_assign_stmt(&mut self, n: &RawNode, context: u32) -> Step {
        let lhs = NodeRef(n.w[4]);
        let rhs = NodeRef(n.w[5]);
        let flags = n.w[1] & 0xffff;
        let ln = self.node_at(lhs);
        let rn = self.node_at(rhs);
        let lhs_op = ln.w[0] & 0xffff;

        // Dispatch-binding LHS (node+5 bit 1 with a 0x60 LHS).
        if flags & 0x200 != 0 && lhs_op == 0x60 {
            unimplemented!(
                "assignment dispatch-binding LHS: needs the dispatch-binding \
                 check; Phase 6"
            );
        }

        let rhs_result = self.emit_expr(rhs, 2);

        // 0x400 flag with a 0x69 LHS (compound-op store).
        if flags & 0x400 != 0 && lhs_op == 0x69 {
            self.emit_compound_op_store(lhs);
            return DONE;
        }
        Self::gate_assign_stmt_shapes(flags, &ln, &rn);

        let desc = self.resolve_assign_lhs(lhs, &ln);
        let (f_flags, also_emit_lhs) = Self::assign_store_flags(n.w[0], flags, rhs_result, context);
        let n_type = ((ln.w[0] & 0xffff_0000) >> 16) as i32;
        self.emit_reference(&desc, 4, f_flags, n_type);
        if also_emit_lhs {
            self.emit_expr(lhs, context);
        }
        DONE
    }

    /// Assignment shapes that are not yet supported.
    fn gate_assign_stmt_shapes(flags: u32, ln: &RawNode, rn: &RawNode) {
        let lhs_region = ln.w[0] & 0xffff_0000;
        let rhs_region = rn.w[0] & 0xffff_0000;
        // Array / special LHS class.
        if flags & 0x6000 != 0 {
            unimplemented!("assignment array/special LHS (flags & 0x6000); Phase 6");
        }
        // Stack-arg pass-through.
        if flags & 0x8000 != 0 && lhs_region == 0xf_0000 && rhs_region == 0xf_0000 {
            unimplemented!("assignment stack-arg pass-through; Phase 6");
        }
        // Object-assignment region.
        if lhs_region == 0x12_0000 {
            unimplemented!("assignment object-assignment (LHS region 0x120000); Phase 6");
        }
        // ByRef-init store.
        if ln.w[0] & 0xffff == 0x69 && ln.w[1] & 0x8000 != 0 {
            unimplemented!("assignment ByRef-init store; Phase 6");
        }
    }

    /// Resolve the LHS reference of an assignment statement.
    fn resolve_assign_lhs(&self, lhs: NodeRef, ln: &RawNode) -> RefDescriptor {
        let sym = self.sym.as_ref().expect(
            "0x2c assignment needs the LHS symbol context \
             (Emitter::with_symbol_context)",
        );
        if ln.w[4] != 0 {
            unimplemented!("assignment LHS member sub-expression; Phase 6");
        }
        crate::resolver::resolve_reference2(
            self.arena,
            lhs,
            &sym.heap,
            sym.member_off,
            sym.ctx_flag_c,
            sym.binding,
        )
    }

    /// Store flags for an assignment statement and whether the LHS is also
    /// emitted as a value afterwards.
    fn assign_store_flags(node_word0: u32, flags: u32, rhs_result: u32, context: u32) -> (u32, bool) {
        let mut f_flags = flags & 0xff00;
        if rhs_result & 1 != 0 {
            f_flags |= 0x80;
        }
        if (node_word0 & 0xffff_0000) == 0x2_0000 {
            return (f_flags, false);
        }
        if context == 6 {
            (f_flags, true)
        } else {
            (f_flags | 0x40, false)
        }
    }

    /// Typed assignment / error recovery. When the flag byte is clear and the
    /// assigned child is not object-typed, a child of type `0x10` emits its
    /// source (context 2) then the sized coercion opcode `0x2c7`; every other
    /// child type is a source-level type mismatch that the compiler reports as a
    /// diagnostic (no p-code). The remaining shapes re-emit through the
    /// assignment statement and need the resolver.
    fn op_typed_assign(&mut self, n: &RawNode) -> Step {
        let child = self.node_at(NodeRef(n.w[4]));
        let byte5 = (n.w[1] >> 8) & 0xff;
        if byte5 != 0 || (child.w[0] & 0xffff_0000) == 0xf0000 {
            unimplemented!(
                "typed-assignment fallthrough (object child / flag set): re-emits \
                 through the assignment statement (resolver); Phase 5"
            );
        }
        let child_tag = (child.w[0] as i32) >> 16;
        if child_tag != 0x10 {
            unimplemented!(
                "typed-assignment type mismatch (child type {child_tag}): a \
                 source-level type error, not valid bound input"
            );
        }
        self.emit_expr(NodeRef(n.w[5]), 2);
        let size = self.emit_get_type_size3(child.w[5]);
        self.emit_opcode2(0x2c7, size as u16);
        DONE
    }

    /// Sequence — emit both children.
    fn op_sequence(&mut self, n: &RawNode, context: u32) -> Step {
        let mut result = 0;
        if n.w[4] != 0 {
            result = self.emit_expr(n.lhs(), context);
        }
        if n.w[5] == 0 {
            return Step::Done(result);
        }
        Step::Done(self.emit_expr(n.rhs(), context))
    }

    // ── Member / list nodes ──────────────────────────────────────────────────

    /// Nested member-access size. Emit the inner list element, then the member
    /// size opcode 0x20d and per-type validation, both taken from the
    /// doubly-nested type node.
    fn op_nested_member_size(&mut self, n: &RawNode) -> Step {
        let inner = self.node_at(NodeRef(n.w[5]));
        self.emit_expr(NodeRef(inner.w[5]), 1);
        let mid = self.node_at(NodeRef(inner.w[4]));
        let piv = self.node_at(NodeRef(mid.w[4]));
        let sz = self.emit_get_type_size3(piv.w[5]);
        self.emit_opcode2(0x20d, sz as u16);
        Step::Done(self.emit_validate_type_operation((piv.w[0] as i32) >> 16, 0x17, 0))
    }

    /// Sized end-with opcode 0x20e.
    fn op_end_with_size(&mut self, n: &RawNode) -> Step {
        let inner = self.node_at(NodeRef(n.w[5]));
        let size = self.emit_get_type_size3(inner.w[5]);
        self.emit_opcode2(0x20e, size as u16);
        DONE
    }

    /// Process the inner list, then a fixed "no member" opcode or a pooled
    /// member-value opcode.
    fn op_inner_member_value(&mut self, n: &RawNode) -> Step {
        let inner = self.node_at(NodeRef(n.w[5]));
        self.process_linked_list(NodeRef(inner.w[5]), 1);
        if inner.w[4] == 0 {
            self.emit_value2(0x1c3);
            return DONE;
        }
        let mut p = self.node_at(NodeRef(inner.w[4]));
        if p.w[0] & 0xffff == 0x11 {
            p = self.node_at(NodeRef(p.w[4]));
        }
        let opcode = if n.w[1] & 0x8000 != 0 { 0x360 + 0xd1 } else { 0x360 };
        let v = self.type_pool.extract_type_value2(p.w[4]);
        self.emit_opcode2(opcode, v);
        DONE
    }

    /// Process the list, then the member size opcode 0x20f.
    fn op_member_size_list(&mut self, n: &RawNode) -> Step {
        self.process_linked_list(NodeRef(n.w[5]), 1);
        let child = self.node_at(NodeRef(n.w[4]));
        let size = self.emit_get_type_size3(child.w[5]);
        self.emit_opcode2(0x20f, size as u16);
        DONE
    }

    /// Argument-list emission. Emit the list opcode (selected by op-class and
    /// flag 0x8000), with the argument count when > 1, then each argument's
    /// resolved type size as a trailing word.
    fn op_arg_list(&mut self, n: &RawNode) -> Step {
        let flags = n.w[1] & 0xffff;
        let opcode = Self::arg_list_opcode(flags);
        let mut count: u16 = 1;
        let mut cur = NodeRef(n.w[5]);
        if flags & 0x8000 == 0 {
            let n5 = self.node_at(cur);
            count = self.arena.get(NodeRef(n5.w[4])).w[4] as u16;
            cur = NodeRef(n5.w[5]);
        }
        if count < 2 {
            self.emit_value2(opcode);
        } else {
            self.emit_opcode2(opcode, count * 2);
        }
        self.emit_arg_list_sizes(cur);
        DONE
    }

    /// The argument-list opcode for the node flags.
    fn arg_list_opcode(flags: u32) -> usize {
        let adj = if flags & 0x8000 != 0 { 0x119 } else { 0 };
        match (flags & 0x300) >> 8 {
            0 => 0x3b2 - adj,
            1 => 0x3b3 - adj,
            2 => 0x3b4 - adj,
            _ => unimplemented!(
                "argument-list op-class 3: opcode undefined in the decompile"
            ),
        }
    }

    /// Emit each argument's resolved type size as a trailing word.
    fn emit_arg_list_sizes(&mut self, first: NodeRef) {
        let mut cur = first;
        while cur.0 != 0 {
            let c = self.node_at(cur);
            let (elem, next) = if c.w[0] & 0xffff == 0x37 {
                (NodeRef(c.w[4]), NodeRef(c.w[5]))
            } else {
                (cur, NodeRef(0))
            };
            let elem = self.skip_type_wrappers(elem);
            let sz = self.emit_get_type_size3(self.arena.get(elem).w[5]);
            self.emit_word2(sz as u16);
            cur = next;
        }
    }

    /// Follow `0x11` wrapper nodes to the wrapped node.
    fn skip_type_wrappers(&self, mut elem: NodeRef) -> NodeRef {
        while self.arena.get(elem).w[0] & 0xffff == 0x11 {
            elem = NodeRef(self.arena.get(elem).w[4]);
        }
        elem
    }

    /// Dispatch-type resolution. The type value comes from the first child
    /// (`word[5].word[4]`, unwrapping an `0x11` wrapper); the emitted reference
    /// is the second child (`word[5].word[5]`, taken as `.word[4]`). An object
    /// member (`0x60` child after unwrapping `0x11`/`0x12`, with node type tag
    /// 2) takes a dispatch-binding path that stays gated unless the binding is
    /// recognised; every other shape uses the common typed path.
    fn op_dispatch_type(&mut self, n: &RawNode, op: i32) -> Step {
        let w5 = self.node_at(NodeRef(n.w[5]));
        let mut a = self.node_at(NodeRef(w5.w[4]));
        if a.w[0] & 0xffff == 0x11 {
            a = self.node_at(NodeRef(a.w[4]));
        }
        let type_value = self.type_pool.extract_type_value2(a.w[4]);
        let p6 = self.skip_dispatch_wrappers(NodeRef(w5.w[5]));
        if self.try_emit_dispatch_binding(n, op, p6, type_value) {
            return DONE;
        }
        let second = self.node_at(NodeRef(w5.w[5]));
        self.emit_typed_node(NodeRef(second.w[4]), 5);
        self.emit_opcode2(0x42a, type_value);
        Step::Done(self.emit_validate_type_operation((n.w[0] as i32) >> 16, 0x17, 1))
    }

    /// Follow `0x11` / `0x12` wrapper nodes.
    fn skip_dispatch_wrappers(&self, mut p: NodeRef) -> NodeRef {
        loop {
            let k = self.arena.get(p).w[0] & 0xffff;
            if k != 0x11 && k != 0x12 {
                return p;
            }
            p = NodeRef(self.arena.get(p).w[4]);
        }
    }

    /// Emit the dispatch-binding form when the unwrapped child is an object
    /// member whose binding is recognised; returns whether it was emitted.
    fn try_emit_dispatch_binding(&mut self, n: &RawNode, op: i32, p6: NodeRef, type_value: u16) -> bool {
        let p6n = self.node_at(p6);
        if (p6n.w[0] & 0xffff) != 0x60 || ((n.w[0] as i32) >> 16) != 2 {
            return false;
        }
        let ctx_kind = self.arena.get(NodeRef(p6n.w[5])).w[0] as i32;
        let sym = self.sym.as_ref().expect(
            "dispatch-binding path needs the module symbol context \
             (Emitter::with_symbol_context)",
        );
        if !crate::resolver::is_dispatch_binding(ctx_kind, &sym.heap, sym.member_off) {
            return false;
        }
        let size = self.emit_get_type_size3(p6n.w[5]);
        // case 0x42 -> base 0x2ff; case 0x43 -> base 0x3f6.
        let opcode = if op == 0x43 { 0x3f6 } else { 0x2ff };
        self.emit_opcode2(opcode, size as u16);
        self.emit_word2(type_value);
        true
    }

    // ── Conversion / traversal groups ────────────────────────────────────────

    /// Cases 0x44..=0x47: type conversion, then operand dispatch (depth 2) for
    /// results that are not the 0x20000 form.
    fn op_conversion_dispatch(&mut self, node: NodeRef, n: &RawNode, op: i32) -> Step {
        let target = match op {
            0x44 => 0x3f5,
            0x45 if n.w[1] & 0x8000 != 0 => 0x437,
            0x45 => 0x2ca,
            0x46 => 0x15d,
            _ => 0x15e,
        };
        self.emit_type_conversion2(target, node, true);
        if n.w[0] & 0xffff_0000 != 0x20000 {
            self.emit_dispatch_opcode(NodeRef(n.w[5]), 2, n.type_tag());
        }
        DONE
    }

    /// Cases 0x48..=0x4b: traverse list, emit a fixed opcode; results that are
    /// not the 0x20000 form additionally dispatch the operand (depth 1).
    fn op_traverse_dispatch(&mut self, n: &RawNode, op: i32) -> Step {
        self.traverse_node_tree(NodeRef(n.w[5]), 1);
        let value = match op {
            0x48 => 0x158,
            0x49 => 0x15a,
            0x4a => 0x159,
            _ => 0x15b,
        };
        self.emit_value2(value);
        if n.w[0] & 0xffff_0000 != 0x20000 {
            self.emit_dispatch_opcode(NodeRef(n.w[5]), 1, n.type_tag());
        }
        DONE
    }

    /// Operator classification (cases 0x51 / 0x52). The op-class (word[1] bits
    /// 8..10) selects the opcode `K` and the emission mode: a plain value emit
    /// (`typed` false) or a typed emit with a pooled type-value operand drawn
    /// from the nested operand. For `0x51` the opcode is `K - 2`, for `0x52` it
    /// is `K`.
    fn op_operator_class(&mut self, n: &RawNode, op: i32) -> Step {
        let op_class = (n.w[1] >> 8) & 7;
        let (k, typed): (i32, bool) = match op_class {
            0 => (0x177, false),
            1 => (0x178, false),
            2 => (0x418, true),
            3 => (0x419, true),
            4 => (0x414, false),
            5 => (0x415, false),
            // op-class 6/7 take their opcode from a caller-passed value that is
            // the node pointer itself — never a real operator.
            _ => unimplemented!(
                "operator classification op-class {op_class}: degenerate \
                 caller-supplied opcode path, not reached for real operators"
            ),
        };
        let opcode = if op == 0x51 { k - 2 } else { k };
        if !typed {
            self.traverse_node_tree(NodeRef(n.w[5]), 1);
            self.emit_value2(opcode as usize);
        } else {
            let inner = self.node_at(NodeRef(n.w[5]));
            self.traverse_node_tree(NodeRef(inner.w[5]), 1);
            let p = self.node_at(NodeRef(inner.w[4]));
            let typeval = self.type_pool.extract_type_value2(p.w[4]);
            self.emit_opcode2(opcode as usize, typeval);
        }
        DONE
    }

    /// Type conversion (flag 0x4000 clear) or traverse + flag-selected opcode
    /// (set).
    fn op_conversion_or_traverse(&mut self, node: NodeRef, n: &RawNode) -> Step {
        if n.w[1] & 0x4000 == 0 {
            let target = ((n.w[1] & 0x8000) | 0x6f_0000) >> 0xe;
            self.emit_type_conversion2(target as i32, node, true);
            return DONE;
        }
        let value = (if n.w[1] & 0x8000 != 0 { 2 } else { 0 }) + 0x1bb;
        self.traverse_then_tail(n, value)
    }

    /// Traverse list, then a fixed opcode + validation (flag 0x4000 clear) or a
    /// sized opcode + object-type validation (set).
    fn op_traverse_validate(&mut self, n: &RawNode, context: u32) -> Step {
        self.traverse_node_tree(NodeRef(n.w[5]), 1);
        let flags16 = n.w[1] & 0xffff;
        let type_tag = n.type_tag();
        if flags16 & 0x4000 == 0 {
            let value = if flags16 & 0x8000 != 0 { 0x3fb } else { 0x3fd };
            self.emit_value2(value);
            self.emit_validate_type_operation(type_tag, 0, context);
        } else {
            let size = self.emit_get_type_size3(n.w[6]);
            let opcode = (((!flags16) & 0x8000) | 0xff_0000) >> 0xe;
            self.emit_opcode2(opcode as usize, size as u16);
            self.emit_validate_type_operation(type_tag, 0x17, context);
        }
        DONE
    }

    /// Traverse list, then a sized opcode + object-type validation (flag bit
    /// 0x40 set) or the fixed opcode 0x3ff (clear).
    fn op_traverse_sized(&mut self, n: &RawNode, context: u32) -> Step {
        self.traverse_node_tree(NodeRef(n.w[5]), 1);
        if (n.w[1] >> 8) & 0x40 == 0 {
            return Step::Tail(0x3ff);
        }
        let size = self.emit_get_type_size3(n.w[6]);
        self.emit_opcode2(0x400, size as u16);
        self.emit_validate_type_operation(n.type_tag(), 0x17, context);
        DONE
    }

    // ── References / calls ───────────────────────────────────────────────────

    /// Member-reference coercion (cases 0x60 / 0x69): resolve the reference,
    /// then emit it via the value-emitter with nOp = context, fFlags = 0,
    /// nType = node[0] >> 16. For 0x69 the operator setup first traverses and
    /// emits the two operands; for 0x60 the reference resolver reads the symbol
    /// heap. The resolver itself gates the member-sub-expression and
    /// method-binding sub-paths.
    fn op_member_reference(&mut self, node: NodeRef, n: &RawNode, op: i32, context: u32) -> Step {
        let n_type = (n.w[0] >> 16) as i32;
        let desc = if op == 0x69 {
            self.operator_reference_descriptor(node, n)
        } else {
            self.member_reference_descriptor(node)
        };
        self.emit_reference(&desc, context as i32, 0, n_type);
        DONE
    }

    /// Descriptor for the operator-reference (0x69) form.
    fn operator_reference_descriptor(&mut self, node: NodeRef, n: &RawNode) -> RefDescriptor {
        if n.w[1] & 0x8000 != 0 {
            unimplemented!(
                "binary-operation ByRef stack-init (0x69, node+5 bit 0x80); Phase 6"
            );
        }
        self.emit_setup_binary_operation(node)
    }

    /// Descriptor for the member-reference (0x60) form.
    fn member_reference_descriptor(&self, node: NodeRef) -> RefDescriptor {
        let sym = self.sym.as_ref().expect(
            "0x60 member reference needs the module symbol context \
             (Emitter::with_symbol_context)",
        );
        crate::resolver::resolve_reference2(
            self.arena,
            node,
            &sym.heap,
            sym.member_off,
            sym.ctx_flag_c,
            sym.binding,
        )
    }

    /// Call site. Assemble a `CallDescriptor` from the bound call node and
    /// dispatch to `emit_call`. The convention kind and by-reference mode are
    /// resolved by the binder and carried on the node: `word[2]` = kind,
    /// `word[3]` = byref. The remaining fields mirror the call node — `word[0]`
    /// type/region, `word[1]` flags, `word[5]` arg list, `word[6]` callee,
    /// `word[7]` low half the member-dispatch id, `word[8]` the callee's type
    /// descriptor (resolved to a size only when the dispatch record asks for
    /// one), `word[9]` the finalize dispatch-word type value (interned on the
    /// emit-mode-1 path).
    fn op_call_site(&mut self, n: &RawNode, context: u32) -> Step {
        let desc = CallDescriptor {
            kind: n.w[2] as i32,
            byref: n.w[3] as i32,
            flags: n.w[1],
            node_word0: n.w[0],
            callee: NodeRef(n.w[6]),
            arg_list: NodeRef(n.w[5]),
            member_id: n.w[7] as u16,
            size: self.emit_get_type_size3(n.w[8]) as u16,
            node9: n.w[9],
        };
        Step::Done(self.emit_call(&desc, context))
    }

    /// Emit the child reference, then a pooled member opcode (0x38d for a
    /// non-object child, 0x16f for an object one).
    fn op_pooled_member(&mut self, n: &RawNode) -> Step {
        self.emit_expr(NodeRef(n.w[4]), 1);
        let child = self.node_at(NodeRef(n.w[4]));
        let opcode = if child.w[0] & 0xffff_0000 != 0xf_0000 { 0x38d } else { 0x16f };
        let v = self.type_pool.extract_type_value2(n.w[5]);
        self.emit_opcode2(opcode, v);
        DONE
    }

    /// Emit child, then (context 6) opcode 0x29f, else a type-class-selected
    /// opcode. For a `0x160000`-region node with a `0x160000`-region child the
    /// opcode is `0x2f2` followed by the pooled type value (`word[5]`). An
    /// object-typed child selects its opcode from the type descriptor's
    /// attribute / optional flags, which need the type-descriptor attribute
    /// model and stay gated.
    fn op_member_ref_value(&mut self, n: &RawNode, context: u32) -> Step {
        let child = n.lhs();
        self.emit_expr(child, 1);
        if context == 6 {
            return Step::Tail(0x29f);
        }
        let cn = self.node_at(child);
        let (value, needs_word) =
            Self::member_ref_opcode(n.w[0] & 0xffff_0000, cn.w[0] & 0xffff_0000, context);
        self.emit_value2(value as usize);
        if needs_word {
            let w = self.type_pool.extract_type_value2(n.w[5]);
            self.emit_word2(w);
        }
        DONE
    }

    /// The opcode for a member reference and whether it carries the pooled type
    /// value word.
    fn member_ref_opcode(node_hi: u32, child_hi: u32, context: u32) -> (u32, bool) {
        if node_hi != 0x160000 {
            return (context, false);
        }
        match child_hi {
            0xf0000 => unimplemented!(
                "member reference (object child): opcode from the \
                 type-descriptor attribute / optional flags; needs the \
                 type-descriptor model; Phase 5"
            ),
            0x160000 => (0x2f2, true),
            _ => (context, false),
        }
    }

    /// Member-call instruction. Process the argument list, emit the callee
    /// reference (context 1), then an op-class-selected instruction opcode
    /// (word[1] bits 8..9): classes 1/2 emit a sized opcode (0x3b0/0x3b1) with
    /// the word[8] operand and finish; classes 0/3 emit 0x3ae/0x3af with the
    /// word[6] type size (or 0x3f2 + the word[7] size when the flag byte bit
    /// 0x80 is set), the word[8] trailing word, and the per-type validation.
    fn op_member_call(&mut self, n: &RawNode, context: u32) -> Step {
        if n.w[5] != 0 {
            self.process_linked_list(NodeRef(n.w[5]), 3);
        }
        self.emit_expr(NodeRef(n.w[4]), 1);
        let class = (n.w[1] & 0x300) >> 8;
        if class == 1 || class == 2 {
            self.emit_opcode2(0x3af + class as usize, n.w[8] as u16);
            return DONE;
        }
        let opcode = if class == 0 { 0x3ae } else { 0x3af };
        let size = if n.w[6] != 0 { self.emit_get_type_size3(n.w[6]) } else { 0 };
        if (n.w[1] >> 8) & 0x80 == 0 {
            self.emit_opcode2(opcode, size as u16);
        } else {
            self.emit_opcode2(0x3f2, size as u16);
            let s7 = self.emit_get_type_size3(n.w[7]);
            self.emit_word2(s7 as u16);
        }
        self.emit_word2(n.w[8] as u16);
        Step::Done(self.emit_validate_type_operation((n.w[0] as i32) >> 16, 0x17, context))
    }

    // ── Instruction nodes ────────────────────────────────────────────────────

    /// The region bits of a node's first child.
    fn child_region(&self, n: &RawNode) -> u32 {
        self.arena.get(NodeRef(n.w[4])).w[0] & 0xffff_0000
    }

    /// Store/let instruction. Opcode selected by flags and the target's type
    /// region, then per-type validation.
    fn op_store_let(&mut self, node: NodeRef, n: &RawNode, context: u32) -> Step {
        let w8 = (n.w[8] as i16) != 0;
        let opcode = Self::store_let_opcode(n.w[1] & 0xffff, self.child_region(n), w8);
        self.emit_instruction2(node, opcode as usize, false, true);
        Step::Done(self.emit_validate_type_operation(n.type_tag(), 0x17, context))
    }

    /// Opcode for the store/let instruction.
    fn store_let_opcode(f: u32, child_region: u32, w8: bool) -> i32 {
        let sub4 = if f & 0x800 != 0 { 4 } else { 0 };
        if f & 0x8000 != 0 {
            return if f & 0x4000 == 0 { 0x3a1 + w8 as i32 } else { 0x3ab };
        }
        if f & 0x4000 != 0 {
            return 0x3a7;
        }
        if child_region == 0xf_0000 {
            (if w8 { 0x442 } else { 0x441 }) - sub4
        } else {
            0x399 + w8 as i32
        }
    }

    /// Assignment/copy instruction.
    fn op_copy_instruction(&mut self, node: NodeRef, n: &RawNode) -> Step {
        let f = n.w[1] & 0xffff;
        let opcode = Self::copy_instruction_opcode(f, self.child_region(n));
        let has_arg = f & 0x2000 != 0;
        self.emit_instruction2(node, opcode as usize, has_arg, false);
        DONE
    }

    /// Opcode for the assignment/copy instruction.
    fn copy_instruction_opcode(f: u32, child_region: u32) -> i32 {
        if f & 0x8000 != 0 {
            return if f & 0x2000 == 0 {
                0x3a3
            } else {
                ((f & 0x4000) >> 0xb) as i32 | 0x3a4
            };
        }
        let object_child = child_region == 0xf_0000;
        if f & 0x2000 == 0 {
            return 0x443 - if object_child { 0 } else { 0xa8 };
        }
        if object_child {
            0x444
        } else {
            0x39c + if f & 0x4000 != 0 { 0xc } else { 0 }
        }
    }

    /// Instruction taking an argument, selected by flags and the child region.
    fn op_instruction(&mut self, node: NodeRef, n: &RawNode) -> Step {
        let opcode = Self::instruction_opcode(n.w[1] & 0xffff, self.child_region(n));
        self.emit_instruction2(node, opcode as usize, true, false);
        DONE
    }

    /// Opcode for the argument-taking instruction.
    fn instruction_opcode(f: u32, child_region: u32) -> i32 {
        if f & 0x4000 != 0 {
            return ((f & 0x8000) >> 0xd) as i32 | 0x3a9;
        }
        if child_region == 0xf_0000 {
            0x445
        } else {
            0x39d + if f & 0x8000 != 0 { 8 } else { 0 }
        }
    }

    /// Flag-selected instruction taking an argument.
    fn op_flag_instruction(&mut self, node: NodeRef, n: &RawNode) -> Step {
        let f = n.w[1] & 0xffff;
        let opcode = if f & 0x1000 != 0 {
            0x40c
        } else if f & 0x4000 == 0 {
            0x398 + if f & 0x8000 != 0 { 8 } else { 0 }
        } else {
            0x3a6 + if f & 0x8000 != 0 { 4 } else { 0 }
        };
        self.emit_instruction2(node, opcode as usize, true, false);
        DONE
    }
}
