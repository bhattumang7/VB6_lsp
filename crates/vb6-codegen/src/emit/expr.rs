use super::*;
use super::intrinsics::{explicit_conversion_bytes, unary_intrinsic_bytes};

impl<'a> Emitter<'a> {
    // ── Node emission ────────────────────────────────────────────────────────

    /// Emit the runtime P-code for one node. `context` is the usage context
    /// (0 = normal value; 1/2/3/5/6 select specialized emission). The return
    /// value propagates a sub-result for the few cases that compute one (0
    /// otherwise).
    pub fn emit_expr(&mut self, node: NodeRef, context: u32) -> u32 {
        let n = *self.arena.get(node);
        let op = n.opcode();

        // Synthetic typed-load IR nodes (opcodes `>= 0x74`, which the dispatch
        // guard below rejects). The byte form is produced here directly.
        if op == 0x74 || op == 0x76 {
            self.emit_var_load(&n, context);
            return 0;
        }
        // Synthetic ByRef parameter load (opcode 0x75): same layout as a local
        // load node, routed through the ByRef opcode (base + 0x14).
        if op == 0x75 {
            self.emit_byref_param_node(&n);
            return 0;
        }
        // Synthetic module-global load (opcode 0x77): module descriptor in the
        // low u16 of word[4], field offset in the high u16, type context word[5].
        if op == 0x77 {
            self.emit_global_node_load(&n);
            return 0;
        }
        // Synthetic operand coercion (opcode 0x78): widen a binary-operation
        // operand to the operation type. Node type tag = target type; word[4] =
        // the operand. Emits the operand then the widening conversion opcode.
        if op == 0x78 {
            self.emit_coerce_node(&n);
            return 0;
        }
        // Synthetic string literal (opcode 0x79): push a pooled string constant by
        // index — `0x1b <pool index>` (value-emitter index 0x3bf). word[4] = index.
        if op == 0x79 {
            self.emit_opcode2(0x3bf, n.w[4] as u16);
            return 0;
        }
        // Synthetic Static-local load (opcode 0x7b): read from the procedure's
        // static block — `0x5f <module_desc> 0x0004 <load-op> <static offset>`.
        if op == 0x7b {
            self.emit_static_load(&n);
            return 0;
        }
        // Synthetic explicit type-conversion (opcode 0x7c): an explicit conversion
        // intrinsic (CInt/CLng/CSng/CDbl/CCur/CStr). Emit the child operand, then
        // the explicit-conversion opcode for (dest = node type tag, src = child
        // type tag). This family differs from the implicit assignment coercion.
        if op == 0x7c {
            self.emit_expr(NodeRef(n.w[4]), 2);
            let dest = n.type_tag();
            let src = self.arena.get(NodeRef(n.w[4])).type_tag();
            for &b in explicit_conversion_bytes(dest, src) {
                self.stream.emit_byte(b);
            }
            return 0;
        }
        // Synthetic function-call-in-expression (opcode 0x7e): a pre-emitted call
        // byte blob (result left on the stack) at blob offset word[4], length
        // word[5]. Emit the blob verbatim.
        if op == 0x7e {
            let bytes = self.arena.blob(n.w[4], n.w[5] as usize).to_vec();
            self.stream.emit_bytes(&bytes);
            return 0;
        }
        // Synthetic dedicated-opcode unary intrinsic (opcode 0x7d): Len/Abs/Sgn/
        // Int/Fix. Emit the argument, then the intrinsic opcode selected by the
        // argument's type (kind in word[5]).
        if op == 0x7d {
            self.emit_expr(NodeRef(n.w[4]), 2);
            let arg = self.arena.get(NodeRef(n.w[4])).type_tag();
            for &b in unary_intrinsic_bytes(n.w[5], arg) {
                self.stream.emit_byte(b);
            }
            return 0;
        }

        // Guard: opcodes outside `1..=0x73` emit nothing (opcode 0 wraps to
        // 0xffffffff and is also rejected).
        if (op - 1) as u32 > 0x72 {
            return 0;
        }

        let type_tag = n.type_tag();
        // Flag byte: byte 1 of word[1].
        let flag_byte = (n.w[1] >> 8) & 0xff;
        // 3-bit operation sub-class: word[1] bits 8..=10.
        let op_class = (n.w[1] >> 8) & 7;
        // Node type tag in the high 16 bits, masked for `== 0xNN0000` tests.
        let node_hi = n.w[0] & 0xffff_0000;

        // The opcode emitted by the common tail for cases that fall through.
        // Arms that complete emission `return` directly; arms that fall through
        // evaluate to the trailing opcode.
        let tail_opcode: i32 = match op {
            // case 1: integer literal
            1 => match type_tag {
                3 | 5 | 6 => {
                    let v = n.w[4] as i32;
                    if -0x81 < v && v < 0x80 {
                        self.emit_value2(0x41a);
                        self.stream.emit_byte(v as u8);
                    } else {
                        self.emit_opcode2(0x3b5, n.w[4] as u16);
                    }
                    return 0;
                }
                8 | 0x10 => {
                    self.emit_value2(0x3b8);
                    self.emit_dword(n.w[4]);
                    return 0;
                }
                0x16 => 0x163,
                _ => return 0,
            },
            // case 2: 8-byte literal — Currency (0x3bb → 0xf6) or, when tagged Date
            // (0xc), an OLE date serial (0x3bd → 0xfa).
            2 => {
                let push = if type_tag == 0xc { 0x3bd } else { 0x3bb };
                self.emit_value2(push);
                self.stream.emit_bytes(&n.literal8());
                return 0;
            }
            // case 3: floating-point literal
            3 => return self.emit_float_literal(&n, context),
            // case 4: string literal
            4 => {
                self.emit_string_literal(&n);
                return 0;
            }
            // case 5: typed node. Below the 0x12 opcode boundary the child opcode
            // is propagated as the sub-result; otherwise emit the object guard
            // (flag byte 0x40) and the per-type validation. (The `EbGetTypeSize3`
            // call the runtime makes here is a pure read with a discarded result.)
            5 => {
                let child = *self.arena.get(NodeRef(n.w[4]));
                let child_op = child.w[0] & 0xffff;
                if child_op < 0x12 {
                    return child_op;
                }
                if (n.w[1] >> 8) & 0x40 != 0 {
                    self.emit_value2(0x202);
                }
                return self.emit_validate_type_operation(0xf, 0, context);
            }
            // case 0xb: unary minus / negate (type-class-dependent opcode).
            0xb => {
                self.emit_expr(n.lhs(), 2);
                let lhs = *self.arena.get(n.lhs());
                let lhs_tag = (lhs.w[0] as i32) >> 16;
                let mut iv = RT_TYPE_OFFSET[lhs_tag as usize];
                if iv == 10 {
                    iv = 0xf6;
                } else {
                    if iv == 9 {
                        iv = 1;
                    }
                    iv += 0xf2;
                }
                iv
            }
            // cases 0xc / 0xd: expression-code sub-emission.
            0xc => {
                self.emit_expression_code2(false, node, context);
                return 0;
            }
            0xd => {
                self.emit_expression_code2(true, node, context);
                return 0;
            }
            // case 0xe: load / assign / object-reference path.
            // For a non-object target the body always runs; an object target with
            // certain flags needs the object-reference emit path (deferred). After
            // the inner dispatch the trailing per-type validation is reached.
            0xe => {
                let lhs_kind = (n.w[0] as i32) >> 16;
                // Object target: needs the object-reference emit path.
                if node_hi == 0xf0000 {
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
                match op_class {
                    // Plain assignment: emit the store opcode, then fall through
                    // to the shared trailing validation below.
                    0 => self.emit_assign_op(&n),
                    // UDT copy: sized copy opcode, then object-type validation.
                    1 => {
                        let size = self.emit_get_type_size3(n.w[6]);
                        self.emit_opcode2(0x2fe, size as u16);
                        return self.emit_object_type(&n, context);
                    }
                    // Set (assign reference): fixed opcode, then validation.
                    2 => {
                        self.emit_value2(0x2fd);
                        return self.emit_object_type(&n, context);
                    }
                    // Set with addref: sized opcode, then validation.
                    3 => {
                        let size = self.emit_get_type_size3(n.w[6]);
                        self.emit_opcode2(0x2f9, size as u16);
                        return self.emit_object_type(&n, context);
                    }
                    // Set with release: sized opcode, then validation.
                    4 => {
                        let size = self.emit_get_type_size3(n.w[6]);
                        self.emit_opcode2(0x2fa, size as u16);
                        return self.emit_object_type(&n, context);
                    }
                    // UDT move: sized opcode, then validation.
                    5 => {
                        let size = self.emit_get_type_size3(n.w[6]);
                        self.emit_opcode2(0x2fc, size as u16);
                        return self.emit_object_type(&n, context);
                    }
                    // Array Set/copy: builds and emits a synthetic type node —
                    // needs the type-node construction path.
                    6 => unimplemented!(
                        "assignment op-class 6 (array Set/copy): needs type-node \
                         construction; Phase 4"
                    ),
                    // Me assignment: fixed opcode, then validation.
                    7 => {
                        self.emit_value2(0x41b);
                        return self.emit_object_type(&n, context);
                    }
                    _ => unreachable!(),
                }
                // Trailing tail (op-class 0 only): a 0x202 guard when object-typed
                // with the flag-byte bit set, then per-type validation.
                if flag_byte & 0x80 != 0 && node_hi == 0xf0000 {
                    self.emit_value2(0x202);
                }
                return self.emit_validate_type_operation(lhs_kind, 0, context);
            }
            // case 0xf: name / coerce path — sub-dispatch on the 3-bit op class.
            0xf => match op_class {
                0 => {
                    // Flag 0x8000 clear: emit the child directly with context 1.
                    if n.w[1] & 0x8000 == 0 {
                        self.emit_expr(n.lhs(), 1);
                        return 0;
                    }
                    // Flag set: inspect the child's opcode / flag byte.
                    let child = n.lhs();
                    let cn = *self.arena.get(child);
                    let c_op = cn.w[0] & 0xffff;
                    let c_flag = (cn.w[1] >> 8) & 0xff;
                    if ((c_op != 0x60 || c_flag & 0x20 == 0)
                        && (c_op != 0x69 || c_flag & 0x80 == 0))
                        && c_op != 0x5e
                    {
                        self.emit_expr(child, 3);
                        return 0;
                    }
                    self.emit_expr(child, 1);
                    0x18c
                }
                1 => unimplemented!(
                    "typed name reference: needs symbol / type resolution; Phase 3/4"
                ),
                2 => {
                    self.emit_expr(n.lhs(), 3);
                    return 0;
                }
                // sized name reference: emit the child (context 3), then the
                // sized coercion opcode 0x2c6 with the node's resolved type size.
                3 => {
                    self.emit_expr(n.lhs(), 3);
                    let size = self.emit_get_type_size3(n.w[6]);
                    self.emit_opcode2(0x2c6, size as u16);
                    return 0;
                }
                // in-place name reference: sets the child's flag bit 0x1000 before
                // emitting it (context 3). The emit-time arena is immutable, so the
                // in-place flag mutation is not yet supported (infra, not a
                // symbol-heap dependency).
                4 => unimplemented!(
                    "in-place name reference: needs in-place child-node flag \
                     mutation, unsupported by the immutable emit arena"
                ),
                _ => return 0,
            },
            // case 0x10: emit child, then opcode 0x135.
            0x10 => {
                self.emit_expr(n.lhs(), 1);
                0x135
            }
            // case 0x11: emit the wrapped typed node with context 5.
            0x11 => {
                self.emit_typed_node(NodeRef(n.w[4]), 5);
                return 0;
            }
            // case 0x12: member dereference.
            0x12 => {
                let child = n.lhs();
                let cn = *self.arena.get(child);
                if node_hi == 0x160000 && (cn.w[0] & 0xffff) == 0x67 {
                    self.emit_expr(NodeRef(cn.w[4]), 5);
                    let opcode = if context == 6 { 0x2f7 } else { 0x2f6 };
                    let v = self.type_pool.extract_type_value2(cn.w[5]);
                    self.emit_opcode2(opcode, v);
                    return 0;
                }
                if node_hi == 0xf0000 && (cn.w[0] & 0xffff_0000) == 0x170000 {
                    self.emit_expr(child, 2);
                    return 0;
                }
                if node_hi != 0x170000 || (cn.w[0] & 0xffff_0000) != 0x170000 {
                    if (cn.w[0] & 0xffff) != 0x3f {
                        return 0;
                    }
                    self.emit_expr(child, context);
                    return 0;
                }
                self.emit_expr(child, context);
                0x15c
            }
            // case 0x13: emit child, then opcode 0x397.
            0x13 => {
                self.emit_expr(n.rhs(), 1);
                0x397
            }
            // case 0x14: object cast (self-contained) or type-library item.
            0x14 => {
                if node_hi == 0xf0000 {
                    self.emit_expr(n.lhs(), 3);
                    0x3f9
                } else {
                    if node_hi != 0x170000 {
                        return 0;
                    }
                    unimplemented!("type-library item load; Phase 6");
                }
            }
            // case 0x15: select opcode by the child's type tag.
            0x15 => {
                let lhs = *self.arena.get(n.lhs());
                if (lhs.w[0] & 0xffff_0000) == 0x000f_0000 {
                    0x42e
                } else {
                    0x42d
                }
            }
            // case 0x18: concatenation-style binary op (only the 0xd-type form;
            // every other form is a plain binary operation).
            0x18 => {
                if node_hi != 0x000d_0000 {
                    return self.emit_binary_operation(&n, context);
                }
                self.emit_expr(n.lhs(), 2);
                self.emit_expr(n.rhs(), 2);
                let rhs = *self.arena.get(n.rhs());
                if (rhs.w[0] & 0xffff_0000) == 0x0006_0000 {
                    0xd1
                } else {
                    0xb3
                }
            }
            // case 0x1a: power operator.
            0x1a => {
                self.emit_expr(n.lhs(), 2);
                self.emit_expr(n.rhs(), 2);
                if node_hi == 0x000f_0000 {
                    let size = self.emit_get_type_size3(n.w[6]);
                    self.emit_opcode2(0xce, size as u16);
                } else {
                    self.emit_value2(0xcf);
                }
                return self.emit_validate_type_operation(type_tag, 0, context);
            }
            // no-op group → emit nothing.
            0x1b | 0x1c | 0x2e | 0x35 | 0x3c | 0x3d | 0x40 | 0x5b | 0x62 | 0x64 | 0x6f
            | 0x70 | 0x71 => return 0,
            // case 0x24: `Is` operator.
            0x24 => {
                self.emit_expr(n.lhs(), 1);
                self.emit_expr(n.rhs(), 1);
                if type_tag == 0xf {
                    let size = self.emit_get_type_size3(n.w[6]);
                    self.emit_opcode2(0xef, size as u16);
                    self.emit_validate_type_operation(0xf, 0x17, context);
                    return 0;
                }
                if type_tag == 0x10 {
                    self.emit_value2(0xf0);
                }
                self.emit_validate_type_operation(type_tag, 0x17, context);
                return 0;
            }
            // case 0x2c: assignment statement.
            0x2c => {
                // EbEmitAssignmentStmt common scalar path: emit the RHS value,
                // resolve the LHS reference, then emit its store (nOp 4).
                let lhs = NodeRef(n.w[4]);
                let rhs = NodeRef(n.w[5]);
                let flags = n.w[1] & 0xffff;
                let ln = *self.arena.get(lhs);
                let rn = *self.arena.get(rhs);
                let lhs_op = ln.w[0] & 0xffff;
                let lhs_region = ln.w[0] & 0xffff_0000;
                let rhs_region = rn.w[0] & 0xffff_0000;

                // Dispatch-binding LHS (node+5 bit 1 with a 0x60 LHS).
                if flags & 0x200 != 0 && lhs_op == 0x60 {
                    unimplemented!(
                        "EbEmitAssignmentStmt dispatch-binding LHS \
                         (EbIsDispatchBinding/EbCheckDispatchProperty); Phase 6"
                    );
                }

                let uvar5 = self.emit_expr(rhs, 2);

                // 0x400 flag with a 0x69 LHS (compound-op store).
                if flags & 0x400 != 0 && lhs_op == 0x69 {
                    self.emit_compound_op_store(lhs);
                    return 0;
                }
                // Array / special LHS class.
                if flags & 0x6000 != 0 {
                    unimplemented!(
                        "EbEmitAssignmentStmt array/special LHS (flags & 0x6000); \
                         Phase 6"
                    );
                }
                // Stack-arg pass-through.
                if flags & 0x8000 != 0 && lhs_region == 0xf_0000 && rhs_region == 0xf_0000 {
                    unimplemented!("EbEmitAssignmentStmt stack-arg pass-through; Phase 6");
                }
                // Object-assignment region.
                if lhs_region == 0x12_0000 {
                    unimplemented!(
                        "EbEmitAssignmentStmt object-assignment (LHS region \
                         0x120000); Phase 6"
                    );
                }
                // ByRef-init store.
                if lhs_op == 0x69 && ln.w[1] & 0x8000 != 0 {
                    unimplemented!("EbEmitAssignmentStmt ByRef-init store; Phase 6");
                }

                let sym = self.sym.as_ref().expect(
                    "0x2c assignment needs the LHS symbol context \
                     (Emitter::with_symbol_context)",
                );
                if ln.w[4] != 0 {
                    unimplemented!("EbEmitAssignmentStmt LHS member sub-expression; Phase 6");
                }
                let desc = crate::resolver::resolve_reference2(
                    self.arena,
                    lhs,
                    &sym.heap,
                    sym.member_off,
                    sym.ctx_flag_c,
                    sym.binding,
                );

                let mut f_flags = flags & 0xff00;
                if uvar5 & 1 != 0 {
                    f_flags |= 0x80;
                }
                let mut also_emit_lhs = false;
                if (n.w[0] & 0xffff_0000) != 0x2_0000 {
                    if context == 6 {
                        also_emit_lhs = true;
                    } else {
                        f_flags |= 0x40;
                    }
                }
                let n_type = (lhs_region >> 16) as i32;
                self.emit_reference(&desc, 4, f_flags, n_type);
                if also_emit_lhs {
                    self.emit_expr(lhs, context);
                }
                return 0;
            }
            // case 0x2d: typed assignment / error recovery. When the flag byte is
            // clear and the assigned child is not object-typed, a child of type
            // `0x10` emits its source (context 2) then the sized coercion opcode
            // `0x2c7`; every other child type is a source-level type mismatch
            // that the compiler reports as a diagnostic (no p-code). The
            // remaining shapes re-emit through the assignment statement and need
            // the resolver.
            0x2d => {
                let child = *self.arena.get(NodeRef(n.w[4]));
                let byte5 = (n.w[1] >> 8) & 0xff;
                if byte5 == 0 && (child.w[0] & 0xffff_0000) != 0xf0000 {
                    let child_tag = (child.w[0] as i32) >> 16;
                    if child_tag == 0x10 {
                        self.emit_expr(NodeRef(n.w[5]), 2);
                        let size = self.emit_get_type_size3(child.w[5]);
                        self.emit_opcode2(0x2c7, size as u16);
                        return 0;
                    }
                    unimplemented!(
                        "typed-assignment type mismatch (child type {child_tag}): a \
                         source-level type error, not valid bound input"
                    );
                }
                unimplemented!(
                    "typed-assignment fallthrough (object child / flag set): re-emits \
                     through the assignment statement (resolver); Phase 5"
                );
            }
            // cases 0x2f / 0x30 / 0x31: sequence — emit both children.
            0x2f | 0x30 | 0x31 => {
                let mut result = 0;
                if n.w[4] != 0 {
                    result = self.emit_expr(n.lhs(), context);
                }
                if n.w[5] == 0 {
                    return result;
                }
                return self.emit_expr(n.rhs(), context);
            }
            // case 0x32: type coercion (opcode by flag byte bit 0x80).
            0x32 => {
                let target = if (n.w[1] >> 8) & 0x80 == 0 { 0x40d } else { 0x40e };
                self.emit_type_coercion4(target, node);
                return 0;
            }
            // case 0x33: traverse the child list, emitting each element.
            0x33 => {
                self.traverse_node_tree(node, 1);
                return 0;
            }
            // case 0x34: type coercion (0x40f).
            0x34 => {
                self.emit_type_coercion4(0x40f, node);
                return 0;
            }
            // case 0x36: `Like` operator.
            0x36 => {
                self.emit_expr(n.lhs(), 1);
                self.emit_expr(n.rhs(), 1);
                0xd2
            }
            // case 0x37: process the forward statement list.
            0x37 => {
                self.process_linked_list(node, context);
                return 0;
            }
            // case 0x38: nested member-access size. Emit the inner list element,
            // then the member size opcode 0x20d and per-type validation, both
            // taken from the doubly-nested type node.
            0x38 => {
                let inner = *self.arena.get(NodeRef(n.w[5]));
                self.emit_expr(NodeRef(inner.w[5]), 1);
                let mid = *self.arena.get(NodeRef(inner.w[4]));
                let piv = *self.arena.get(NodeRef(mid.w[4]));
                let sz = self.emit_get_type_size3(piv.w[5]);
                self.emit_opcode2(0x20d, sz as u16);
                return self.emit_validate_type_operation((piv.w[0] as i32) >> 16, 0x17, 0);
            }
            // case 0x39: emit the sized end-with opcode 0x20e.
            0x39 => {
                let inner = *self.arena.get(NodeRef(n.w[5]));
                let size = self.emit_get_type_size3(inner.w[5]);
                self.emit_opcode2(0x20e, size as u16);
                return 0;
            }
            // case 0x3a: traverse child list, then opcode 0x1ca.
            0x3a => {
                self.traverse_node_tree(NodeRef(n.w[5]), 1);
                0x1ca
            }
            // case 0x3b: traverse child list, then opcode 0x1cb.
            0x3b => {
                self.traverse_node_tree(NodeRef(n.w[5]), 1);
                0x1cb
            }
            // case 0x3e: process the inner list, then a fixed "no member" opcode
            // or a pooled member-value opcode.
            0x3e => {
                let inner = *self.arena.get(NodeRef(n.w[5]));
                self.process_linked_list(NodeRef(inner.w[5]), 1);
                if inner.w[4] == 0 {
                    self.emit_value2(0x1c3);
                    return 0;
                }
                let mut p = *self.arena.get(NodeRef(inner.w[4]));
                if p.w[0] & 0xffff == 0x11 {
                    p = *self.arena.get(NodeRef(p.w[4]));
                }
                let opcode = if n.w[1] & 0x8000 != 0 { 0x360 + 0xd1 } else { 0x360 };
                let v = self.type_pool.extract_type_value2(p.w[4]);
                self.emit_opcode2(opcode, v);
                return 0;
            }
            // case 0x3f: process the list, then the member size opcode 0x20f.
            0x3f => {
                self.process_linked_list(NodeRef(n.w[5]), 1);
                let child = *self.arena.get(NodeRef(n.w[4]));
                let size = self.emit_get_type_size3(child.w[5]);
                self.emit_opcode2(0x20f, size as u16);
                return 0;
            }
            // case 0x41: argument-list emission. Emit the list opcode (selected by
            // op-class and flag 0x8000), with the argument count when > 1, then each
            // argument's resolved type size as a trailing word.
            0x41 => {
                let flags = n.w[1] & 0xffff;
                let adj = if flags & 0x8000 != 0 { 0x119 } else { 0 };
                let opcode = match (flags & 0x300) >> 8 {
                    0 => 0x3b2 - adj,
                    1 => 0x3b3 - adj,
                    2 => 0x3b4 - adj,
                    _ => unimplemented!(
                        "argument-list op-class 3: opcode undefined in the decompile"
                    ),
                };
                let mut count: u16 = 1;
                let mut cur = NodeRef(n.w[5]);
                if flags & 0x8000 == 0 {
                    let n5 = *self.arena.get(cur);
                    count = self.arena.get(NodeRef(n5.w[4])).w[4] as u16;
                    cur = NodeRef(n5.w[5]);
                }
                if count < 2 {
                    self.emit_value2(opcode as usize);
                } else {
                    self.emit_opcode2(opcode as usize, count * 2);
                }
                while cur.0 != 0 {
                    let c = *self.arena.get(cur);
                    let (mut elem, next) = if c.w[0] & 0xffff == 0x37 {
                        (NodeRef(c.w[4]), NodeRef(c.w[5]))
                    } else {
                        (cur, NodeRef(0))
                    };
                    while self.arena.get(elem).w[0] & 0xffff == 0x11 {
                        elem = NodeRef(self.arena.get(elem).w[4]);
                    }
                    let sz = self.emit_get_type_size3(self.arena.get(elem).w[5]);
                    self.emit_word2(sz as u16);
                    cur = next;
                }
                return 0;
            }
            // cases 0x42 / 0x43: dispatch-type resolution. The type value comes
            // from the first child (`word[5].word[4]`, unwrapping an `0x11`
            // wrapper); the emitted reference is the second child
            // (`word[5].word[5]`, taken as `.word[4]`). An object member (`0x60` child after
            // unwrapping `0x11`/`0x12`, with node type tag 2) takes a
            // dispatch-binding path that reads a compiled binding record and
            // stays gated; every other shape uses the common typed path.
            0x42 | 0x43 => {
                let w5 = *self.arena.get(NodeRef(n.w[5]));
                let mut a = *self.arena.get(NodeRef(w5.w[4]));
                if a.w[0] & 0xffff == 0x11 {
                    a = *self.arena.get(NodeRef(a.w[4]));
                }
                let type_value = self.type_pool.extract_type_value2(a.w[4]);
                // Walk 0x11/0x12 wrappers on the second child for the dispatch test.
                let mut p6 = NodeRef(w5.w[5]);
                loop {
                    let k = self.arena.get(p6).w[0] & 0xffff;
                    if k == 0x11 || k == 0x12 {
                        p6 = NodeRef(self.arena.get(p6).w[4]);
                    } else {
                        break;
                    }
                }
                if (self.arena.get(p6).w[0] & 0xffff) == 0x60 && ((n.w[0] as i32) >> 16) == 2 {
                    unimplemented!(
                        "dispatch-type resolution dispatch-binding path: reads a \
                         compiled binding record from the module symbol heap; Phase 5"
                    );
                }
                let second = *self.arena.get(NodeRef(w5.w[5]));
                self.emit_typed_node(NodeRef(second.w[4]), 5);
                self.emit_opcode2(0x42a, type_value);
                return self.emit_validate_type_operation((n.w[0] as i32) >> 16, 0x17, 1);
            }
            // cases 0x44..=0x47: type conversion, then operand dispatch (depth 2)
            // for results that are not the 0x20000 form.
            0x44 | 0x45 | 0x46 | 0x47 => {
                let target = match op {
                    0x44 => 0x3f5,
                    0x45 => {
                        if n.w[1] & 0x8000 != 0 {
                            0x437
                        } else {
                            0x2ca
                        }
                    }
                    0x46 => 0x15d,
                    _ => 0x15e,
                };
                self.emit_type_conversion2(target, node, true);
                if node_hi != 0x20000 {
                    self.emit_dispatch_opcode(NodeRef(n.w[5]), 2, type_tag);
                }
                return 0;
            }
            // cases 0x48..=0x4b: traverse list, emit a fixed opcode; results that
            // are not the 0x20000 form additionally dispatch the operand (depth 1).
            0x48 | 0x49 | 0x4a | 0x4b => {
                self.traverse_node_tree(NodeRef(n.w[5]), 1);
                let value = match op {
                    0x48 => 0x158,
                    0x49 => 0x15a,
                    0x4a => 0x159,
                    _ => 0x15b,
                };
                self.emit_value2(value);
                if node_hi != 0x20000 {
                    self.emit_dispatch_opcode(NodeRef(n.w[5]), 1, type_tag);
                }
                return 0;
            }
            // case 0x4c: type conversion (0x35d).
            0x4c => {
                self.emit_type_conversion2(0x35d, node, true);
                return 0;
            }
            // case 0x4d: emit child, then opcode 0x23d.
            0x4d => {
                self.emit_expr(NodeRef(n.w[5]), 1);
                0x23d
            }
            // case 0x4e: opcode 0x23e.
            0x4e => 0x23e,
            // case 0x4f: traverse list, then opcode 0xfb.
            0x4f => {
                self.traverse_node_tree(NodeRef(n.w[5]), 1);
                0xfb
            }
            // case 0x50: traverse list, then opcode 0xfa.
            0x50 => {
                self.traverse_node_tree(NodeRef(n.w[5]), 1);
                0xfa
            }
            // cases 0x51 / 0x52: operator classification. The op-class (word[1]
            // bits 8..10) selects the opcode `K` and the emission mode: a plain
            // value emit (`typed` false) or a typed emit with a pooled type-value
            // operand drawn from the nested operand. For `0x51` the opcode is
            // `K - 2`, for `0x52` it is `K`.
            0x51 | 0x52 => {
                let op_class = (n.w[1] >> 8) & 7;
                let (k, typed): (i32, bool) = match op_class {
                    0 => (0x177, false),
                    1 => (0x178, false),
                    2 => (0x418, true),
                    3 => (0x419, true),
                    4 => (0x414, false),
                    5 => (0x415, false),
                    // op-class 6/7 take their opcode from a caller-passed value
                    // that is the node pointer itself — never a real operator.
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
                    let inner = *self.arena.get(NodeRef(n.w[5]));
                    self.traverse_node_tree(NodeRef(inner.w[5]), 1);
                    let p = *self.arena.get(NodeRef(inner.w[4]));
                    let typeval = self.type_pool.extract_type_value2(p.w[4]);
                    self.emit_opcode2(opcode as usize, typeval);
                }
                return 0;
            }
            // case 0x53: traverse list, then opcode 0x1c0 / 0x1bf by flag bit 0x40.
            0x53 => {
                self.traverse_node_tree(NodeRef(n.w[5]), 1);
                if flag_byte & 0x40 == 0 {
                    0x1c0
                } else {
                    0x1bf
                }
            }
            // case 0x54: type coercion (0x410).
            0x54 => {
                self.emit_type_coercion4(0x410, node);
                return 0;
            }
            // case 0x55: type conversion (0x35e).
            0x55 => {
                self.emit_type_conversion2(0x35e, node, true);
                return 0;
            }
            // case 0x56: type conversion (flag 0x4000 clear) or traverse +
            // flag-selected opcode (set).
            0x56 => {
                if n.w[1] & 0x4000 == 0 {
                    let target = ((n.w[1] & 0x8000) | 0x6f_0000) >> 0xe;
                    self.emit_type_conversion2(target as i32, node, true);
                    return 0;
                }
                let value = (if n.w[1] & 0x8000 != 0 { 2 } else { 0 }) + 0x1bb;
                self.traverse_node_tree(NodeRef(n.w[5]), 1);
                value
            }
            // case 0x57: traverse list, then a fixed opcode + validation (flag
            // 0x4000 clear) or a sized opcode + object-type validation (set).
            0x57 => {
                self.traverse_node_tree(NodeRef(n.w[5]), 1);
                let flags16 = n.w[1] & 0xffff;
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
                return 0;
            }
            // case 0x58: traverse list, then a sized opcode + object-type
            // validation (flag bit 0x40 set) or the fixed opcode 0x3ff (clear).
            0x58 => {
                self.traverse_node_tree(NodeRef(n.w[5]), 1);
                if flag_byte & 0x40 != 0 {
                    let size = self.emit_get_type_size3(n.w[6]);
                    self.emit_opcode2(0x400, size as u16);
                    self.emit_validate_type_operation(type_tag, 0x17, context);
                    return 0;
                }
                0x3ff
            }
            // case 0x59: traverse list, then opcode 0x1c1.
            0x59 => {
                self.traverse_node_tree(NodeRef(n.w[5]), 1);
                0x1c1
            }
            // case 0x5a: complex binary operation.
            0x5a => {
                self.emit_complex_binary_op(node);
                return 0;
            }
            // case 0x5c: opcode 0x162.
            0x5c => 0x162,
            // case 0x5d: type-library-driven cast.
            0x5d => unimplemented!("type-library-driven cast; Phase 6"),
            // case 0x5e: emit child, then opcode 0x40a.
            0x5e => {
                self.emit_expr(NodeRef(n.w[5]), 1);
                0x40a
            }
            // case 0x5f: emit child, then opcode 0x40b.
            0x5f => {
                self.emit_expr(NodeRef(n.w[5]), context);
                0x40b
            }
            // case 0x60: member-reference coercion.
            0x60 | 0x69 => {
                // EbEmitStatement's 0x60/0x69 case (stmt_case_0fab1ddf): resolve
                // the reference, then emit it via the value-emitter with nOp =
                // context, fFlags = 0, nType = node[0] >> 16. For 0x69 the resolver
                // setup (EbSetupBinaryOperation) first traverses + emits the two
                // operands; for 0x60 resolve_reference2 reads the symbol heap.
                // resolve_reference2 itself gates the member-sub-expression and
                // method-binding sub-paths.
                let n_type = (n.w[0] >> 16) as i32;
                let desc = if op == 0x69 {
                    if n.w[1] & 0x8000 != 0 {
                        unimplemented!(
                            "EbSetupBinaryOperation ByRef stack-init (0x69, node+5 \
                             bit 0x80); Phase 6"
                        );
                    }
                    self.emit_setup_binary_operation(node)
                } else {
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
                };
                self.emit_reference(&desc, context as i32, 0, n_type);
                return 0;
            }
            // case 0x61: call site. Assemble a `CallDescriptor` from the bound
            // call node and dispatch to `emit_call`. The convention kind and
            // by-reference mode are resolved by the binder (the values the
            // runtime derives from the callee's compiled type record) and carried
            // on the node: `word[2]` = kind, `word[3]` = byref. The remaining
            // fields mirror the runtime call node — `word[0]` type/region,
            // `word[1]` flags, `word[5]` arg list, `word[6]` callee, `word[7]`
            // low half the member-dispatch id, `word[8]` the callee's type
            // descriptor (resolved to a size only when the dispatch record asks
            // for one), `word[9]` the finalize dispatch-word type value (interned
            // on the emit-mode-1 path).
            0x61 => {
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
                return self.emit_call(&desc, context);
            }
            // case 0x63: emit the child reference, then a pooled member opcode
            // (0x38d for a non-object child, 0x16f for an object one).
            0x63 => {
                self.emit_expr(NodeRef(n.w[4]), 1);
                let child = *self.arena.get(NodeRef(n.w[4]));
                let opcode = if child.w[0] & 0xffff_0000 != 0xf_0000 { 0x38d } else { 0x16f };
                let v = self.type_pool.extract_type_value2(n.w[5]);
                self.emit_opcode2(opcode, v);
                return 0;
            }
            // case 0x65: forward to child.
            0x65 => return self.emit_expr(n.lhs(), context),
            // case 0x66: pooled member-value opcode 0x2f4.
            0x66 => {
                let v = self.type_pool.extract_type_value2(n.w[5]);
                self.emit_opcode2(0x2f4, v);
                return 0;
            }
            // case 0x67: emit child (context 5), then pooled member opcode 0x2f5.
            0x67 => {
                self.emit_expr(NodeRef(n.w[4]), 5);
                let v = self.type_pool.extract_type_value2(n.w[5]);
                self.emit_opcode2(0x2f5, v);
                return 0;
            }
            // case 0x68: emit child, then (context 6) opcode 0x29f, else a
            // type-class-selected opcode. For a `0x160000`-region node with a
            // `0x160000`-region child the opcode is `0x2f2` followed by the
            // pooled type value (`word[5]`). An object-typed child selects its
            // opcode from the type descriptor's attribute / optional flags, which
            // need the type-descriptor attribute model and stay gated.
            0x68 => {
                let child = n.lhs();
                self.emit_expr(child, 1);
                if context != 6 {
                    let cn = *self.arena.get(child);
                    let child_hi = cn.w[0] & 0xffff_0000;
                    let mut needs_word = false;
                    let mut value = context;
                    if node_hi == 0x160000 {
                        if child_hi == 0xf0000 {
                            unimplemented!(
                                "member reference (object child): opcode from the \
                                 type-descriptor attribute / optional flags \
                                 (EbGetAttributeFlags); needs the type-descriptor \
                                 model; Phase 5"
                            );
                        } else if child_hi == 0x160000 {
                            needs_word = true;
                            value = 0x2f2;
                        }
                    }
                    self.emit_value2(value as usize);
                    if needs_word {
                        let w = self.type_pool.extract_type_value2(n.w[5]);
                        self.emit_word2(w);
                    }
                    return 0;
                }
                0x29f
            }
            // case 0x6a: member-call instruction. Process the argument list, emit
            // the callee reference (context 1), then an op-class-selected
            // instruction opcode (word[1] bits 8..9): classes 1/2 emit a sized
            // opcode (0x3b0/0x3b1) with the word[8] operand and finish; classes
            // 0/3 emit 0x3ae/0x3af with the word[6] type size (or 0x3f2 + the
            // word[7] size when the flag byte bit 0x80 is set), the word[8]
            // trailing word, and the per-type validation.
            0x6a => {
                if n.w[5] != 0 {
                    self.process_linked_list(NodeRef(n.w[5]), 3);
                }
                self.emit_expr(NodeRef(n.w[4]), 1);
                let opcode = match (n.w[1] & 0x300) >> 8 {
                    0 => 0x3ae,
                    1 => {
                        self.emit_opcode2(0x3b0, n.w[8] as u16);
                        return 0;
                    }
                    2 => {
                        self.emit_opcode2(0x3b1, n.w[8] as u16);
                        return 0;
                    }
                    3 => 0x3af,
                    _ => return 0,
                };
                let size = if n.w[6] != 0 { self.emit_get_type_size3(n.w[6]) } else { 0 };
                if (n.w[1] >> 8) & 0x80 == 0 {
                    self.emit_opcode2(opcode, size as u16);
                } else {
                    self.emit_opcode2(0x3f2, size as u16);
                    let s7 = self.emit_get_type_size3(n.w[7]);
                    self.emit_word2(s7 as u16);
                }
                self.emit_word2(n.w[8] as u16);
                return self.emit_validate_type_operation((n.w[0] as i32) >> 16, 0x17, context);
            }
            // case 0x6b: store/let instruction. Opcode selected by flags and the
            // target's type region, then per-type validation.
            0x6b => {
                let f = n.w[1] & 0xffff;
                let sub4 = if f & 0x800 != 0 { 4 } else { 0 };
                let w8 = (n.w[8] as i16) != 0;
                let opcode: i32 = if f & 0x8000 == 0 {
                    if f & 0x4000 == 0 {
                        let child_region = self.arena.get(NodeRef(n.w[4])).w[0] & 0xffff_0000;
                        if child_region == 0xf_0000 {
                            (if w8 { 0x442 } else { 0x441 }) - sub4
                        } else {
                            0x399 + w8 as i32
                        }
                    } else {
                        0x3a7
                    }
                } else if f & 0x4000 == 0 {
                    0x3a1 + w8 as i32
                } else {
                    0x3ab
                };
                self.emit_instruction2(node, opcode as usize, false, true);
                return self.emit_validate_type_operation(type_tag, 0x17, context);
            }
            // case 0x6c: assignment/copy instruction.
            0x6c => {
                let f = n.w[1] & 0xffff;
                let opcode: i32 = if f & 0x8000 != 0 {
                    if f & 0x2000 == 0 {
                        0x3a3
                    } else {
                        ((f & 0x4000) >> 0xb) as i32 | 0x3a4
                    }
                } else if f & 0x2000 == 0 {
                    let child_region = self.arena.get(NodeRef(n.w[4])).w[0] & 0xffff_0000;
                    0x443 - if child_region != 0xf_0000 { 0xa8 } else { 0 }
                } else {
                    let child_region = self.arena.get(NodeRef(n.w[4])).w[0] & 0xffff_0000;
                    if child_region != 0xf_0000 {
                        0x39c + if f & 0x4000 != 0 { 0xc } else { 0 }
                    } else {
                        0x444
                    }
                };
                let has_arg = f & 0x2000 != 0;
                self.emit_instruction2(node, opcode as usize, has_arg, false);
                return 0;
            }
            // case 0x6d: instruction.
            0x6d => {
                let f = n.w[1] & 0xffff;
                let opcode: i32 = if f & 0x4000 != 0 {
                    ((f & 0x8000) >> 0xd) as i32 | 0x3a9
                } else {
                    let child_region = self.arena.get(NodeRef(n.w[4])).w[0] & 0xffff_0000;
                    if child_region != 0xf_0000 {
                        0x39d + if f & 0x8000 != 0 { 8 } else { 0 }
                    } else {
                        0x445
                    }
                };
                self.emit_instruction2(node, opcode as usize, true, false);
                return 0;
            }
            // case 0x6e: instruction.
            0x6e => {
                let f = n.w[1] & 0xffff;
                let opcode: i32 = if f & 0x1000 != 0 {
                    0x40c
                } else if f & 0x4000 == 0 {
                    0x398 + if f & 0x8000 != 0 { 8 } else { 0 }
                } else {
                    0x3a6 + if f & 0x8000 != 0 { 4 } else { 0 }
                };
                self.emit_instruction2(node, opcode as usize, true, false);
                return 0;
            }
            // case 0x72: member type-node coercion, then binary operation.
            0x72 => unimplemented!(
                "member type-node coercion: needs type-node construction; Phase 4"
            ),
            // case 0x73: emit opcode 0x266 with the child's word[4] low half.
            0x73 => {
                let cn = *self.arena.get(n.lhs());
                self.emit_opcode2(0x266, (cn.w[4] & 0xffff) as u16);
                return 0;
            }
            // default: every other binary op (arithmetic, logical, comparison).
            _ => return self.emit_binary_operation(&n, context),
        };

        // Common tail: emit the trailing opcode, then return 0.
        self.emit_value2(tail_opcode as usize);
        0
    }
}
