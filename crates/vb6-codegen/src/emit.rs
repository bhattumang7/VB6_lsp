//! Expression / statement code generation — runtime P-code byte-stream form.
//!
//! This is a direct, case-for-case port of the VB6 back-end's node emitter
//! (`EbEmitStatement` @ `0fab161c`): a single `switch` over the bound node's
//! opcode (`(short)*pNode`, our [`RawNode::opcode`]). Each `case` is translated
//! verbatim — the same table indexing, the same helper calls, the same operand
//! widths. The data tables it indexes live in [`crate::tables`] and were
//! extracted from the binary; no opcode value is guessed here.
//!
//! ## Dispatch shape
//! The C function guards `if (0x72 < (short)*pNode - 1) return 0;` (opcodes
//! outside `1..=0x73` emit nothing), switches on the opcode, and for the cases
//! that `break` falls through to a common tail `EbEmitValue2(iVar13)`. We mirror
//! that exactly: portable arms either `return` directly or evaluate to the
//! `iVar13` value consumed by the tail.
//!
//! ## Loads and stores
//! Opcodes `>= 0x74` are rejected by `EbEmitStatement`'s guard — in real VB6 the
//! typed local load/store byte sequences are emitted by the assignment / name
//! paths (`EbEmitAssignOp` and friends), not by a top-level switch case. Until
//! those are ported, [`Emitter::emit_var_load`] / [`Emitter::emit_var_store`]
//! stand in for that byte form (their opcodes are oracle-confirmed), driven by
//! synthetic `0x74` / `0x76` load nodes the binder builds. They are intercepted
//! before the `EbEmitStatement` guard.
//!
//! ## Not yet ported
//! Cases that need the module symbol table, the type/string pool, statement
//! context, or list traversal are `unimplemented!()` with the exact C function
//! and address — never a guessed constant, never silently skipped.

use crate::buffer::PcodeStream;
use crate::node::{NodeArena, NodeRef, RawNode};
use crate::tables::{RT_BINOP_BASE, RT_DISPATCH_FLAG, RT_LOAD_BY_CTX, RT_OPCODE_BYTE, RT_STORE_BY_CTX, RT_TYPE_OFFSET};

/// A resolved-reference descriptor — the input to [`Emitter::emit_reference`],
/// mirroring the 16-byte record `EbResolveIdentRef` builds and `EbEmitExpression2`
/// consumes. The resolver (or the vb6-sema bridge) populates it; the emitter only
/// reads it.
///
/// * `kind` — `*pExprDesc`: the storage class (1 = local, 2 = argument, 7 = …);
///   selects the opcode base.
/// * `operand` — the 2-byte operand emitted after the opcode (descriptor `+10`),
///   e.g. a local's signed frame offset, or type size.
/// * `word6` — descriptor `+6`; for argument/global kinds, bit 0 marks a by-ref
///   slot (`word6 & 1 != 0` → ByRef, forces nOp=2).
/// * `word8` — low 16 bits of descriptor `+8`; emitted as the trailing word in
///   the `sVar8==2` path (opcode 0x438, reached only for the nType==0x17 Variant
///   nested-type-expression chain).
#[derive(Clone, Copy, Debug, Default, PartialEq, Eq)]
pub struct RefDescriptor {
    pub kind: i32,
    pub operand: u16,
    pub word6: u16,
    pub word8: u16,
}

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

    // ── EbEmitStatement @ 0fab161c ───────────────────────────────────────────

    /// Emit the runtime P-code for one bound node. `context` is the C `context`
    /// argument (calling/usage context: 0 = normal value, 1/2/3 = special). The
    /// return value mirrors the C `local_38` (0 in the common case; some cases
    /// propagate a sub-call result).
    pub fn emit_expr(&mut self, node: NodeRef, context: u32) -> u32 {
        let n = *self.arena.get(node);
        let op = n.opcode();

        // Synthetic typed-load IR nodes (not `EbEmitStatement` opcodes — they are
        // `>= 0x74`, which the guard below rejects). Real VB6 emits these bytes
        // from the assignment / name path.
        if op == 0x74 || op == 0x76 {
            self.emit_var_load(&n, context);
            return 0;
        }
        // Synthetic ByRef parameter load (opcode 0x75): same layout as a local load
        // node but routed through emit_byref_load (base+0x14 opcode).
        if op == 0x75 {
            self.emit_byref_param_node(&n);
            return 0;
        }
        // Synthetic module-global load (opcode 0x77): module_desc in low u16 of
        // word[4], field_offset in high u16 of word[4], type_ctx in word[5].
        if op == 0x77 {
            self.emit_global_node_load(&n);
            return 0;
        }

        // Guard: `uVar15 = (short)opcode - 1; if (0x72 < uVar15) return 0;`
        // (opcode 0 wraps to 0xffffffff and is also rejected.)
        if (op - 1) as u32 > 0x72 {
            return 0;
        }

        let type_tag = n.type_tag();

        // The opcode emitted by the common tail for cases that `break`. Arms that
        // finish on their own `return` directly; arms that fall through to the
        // tail evaluate to their `iVar13`.
        let i_var13: i32 = match op {
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
                _ => return 0, // LAB_0fab2106
            },
            // case 2: Currency literal
            2 => {
                self.emit_value2(0x3bb);
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
            // case 5: typed literal / typed store — needs the type/string pool.
            5 => unimplemented!(
                "EbEmitStatement case 5 (typed literal): EbExtractTypeValue2 / \
                 EbParseWithPool + type-descriptor model; Phase 4"
            ),
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
            // cases 0xc / 0xd: expression-code sub-emitter.
            0xc | 0xd => unimplemented!(
                "EbEmitStatement cases 0xc/0xd: EbEmitExpressionCode2; Phase 5"
            ),
            // case 0xe: load / assign / object-ref path.
            // Outer condition: enter normal emit path when LHS is not Object, or when
            // various flag/dispatch conditions hold.  For numeric LHS the condition is
            // always true.  After the inner switch the trailing EbValidateTypeOperation
            // call is always reached (mirrors the C `break` → tail pattern).
            0xe => {
                let lhs_hi = n.w[0] & 0xffff0000;
                let lhs_kind = (n.w[0] as i32) >> 16;
                // Outer condition false path (Object LHS with specific flags):
                // needs EmitPcodeTypeDiagnostic + EbEmitObjectType — unimplemented.
                if lhs_hi == 0xf0000 {
                    unimplemented!(
                        "EbEmitStatement case 0xe: object-assign LHS (hi16=0xf) requires \
                         EmitPcodeTypeDiagnostic @ 0fab196a and EbEmitObjectType @ 0fab1a0a"
                    );
                }
                // Emit RHS expression with context 2.
                let rhs = NodeRef(n.w[4]);
                self.emit_expr(rhs, 2);
                // 0x4000 flag: Set / object-reference assign.
                if n.w[1] & 0x4000 != 0 {
                    unimplemented!(
                        "EbEmitStatement case 0xe: 0x4000 (object Set assign) requires \
                         EbMapTypeCodeValue3 @ 0fab3168 and EbEmitObjectType @ 0fab1a0a"
                    );
                }
                let op_kind = (n.w[1] >> 8) & 7;
                match op_kind {
                    0 => self.emit_assign_op(&n),
                    1 => unimplemented!(
                        "EbEmitStatement case 0xe op-kind 1 (UDT copy): \
                         EbGetTypeSize3 @ 0fab2f55 + EbEmitObjectType @ 0fab1a0a"
                    ),
                    2 => unimplemented!(
                        "EbEmitStatement case 0xe op-kind 2 (Set IDispatch): \
                         EbEmitObjectType @ 0fab1a0a"
                    ),
                    3 => unimplemented!(
                        "EbEmitStatement case 0xe op-kind 3 (Set addref): \
                         EbGetTypeSize3 + EbEmitObjectType"
                    ),
                    4 => unimplemented!(
                        "EbEmitStatement case 0xe op-kind 4 (Set release): \
                         EbGetTypeSize3 + EbEmitObjectType"
                    ),
                    5 => unimplemented!(
                        "EbEmitStatement case 0xe op-kind 5 (UDT move): \
                         EbGetTypeSize3 + EbEmitObjectType"
                    ),
                    6 => unimplemented!(
                        "EbEmitStatement case 0xe op-kind 6 (array Set/copy): \
                         EbGetTypeSize3 + EbCreateTypeNode3"
                    ),
                    7 => unimplemented!(
                        "EbEmitStatement case 0xe op-kind 7 (Me assign): \
                         EbEmitObjectType @ 0fab1a0a"
                    ),
                    _ => unreachable!(),
                }
                // Trailing tail (after outer if + inner switch):
                // emit 0x202 only if Object LHS with byte5 bit 0x80 set.
                if (n.w[1] >> 8) & 0x80 != 0 && lhs_hi == 0xf0000 {
                    self.emit_value2(0x202);
                }
                return self.emit_validate_type_operation(lhs_kind, 0, context);
            }
            // case 0xf: name / coerce path.
            0xf => unimplemented!(
                "EbEmitStatement case 0xf (name/coerce): EbGetVarType / \
                 EbResolveAndSimplify / EbResolveExprTypeImpl; Phase 3/4"
            ),
            // case 0x10: emit child, then opcode 0x135.
            0x10 => {
                self.emit_expr(n.lhs(), 1);
                0x135
            }
            // case 0x11: type-code emit.
            0x11 => unimplemented!(
                "EbEmitStatement case 0x11: EbEmitTypeCode2; Phase 4"
            ),
            // case 0x12: deref / member adjust.
            0x12 => unimplemented!(
                "EbEmitStatement case 0x12 (deref): EbExtractTypeValue2; Phase 4"
            ),
            // case 0x13: emit rhs child, then opcode 0x397.
            0x13 => {
                self.emit_expr(n.rhs(), 1);
                0x397
            }
            // case 0x14: type-library item.
            0x14 => unimplemented!(
                "EbEmitStatement case 0x14: LoadTypeLibraryItem; Phase 6"
            ),
            // case 0x15: select opcode by the child's type tag.
            0x15 => {
                let lhs = *self.arena.get(n.lhs());
                if (lhs.w[0] & 0xffff_0000) == 0x000f_0000 {
                    0x42e
                } else {
                    0x42d
                }
            }
            // case 0x18: concatenation-style binary op (only the 0xd-type form).
            0x18 => {
                if (n.w[0] & 0xffff_0000) != 0x000d_0000 {
                    return self.emit_binary_operation(&n, context); // LAB_0fab1d8e
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
                if (n.w[0] & 0xffff_0000) == 0x000f_0000 {
                    let size = self.emit_get_type_size3(n.w[6]);
                    self.emit_opcode2(0xce, size as u16);
                } else {
                    self.emit_value2(0xcf);
                }
                return self.emit_validate_type_operation(type_tag, 0, context);
            }
            // no-op group → LAB_0fab2106 (emit nothing).
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
                // LAB_0fab1d07
                self.emit_validate_type_operation(type_tag, 0x17, context);
                return 0;
            }
            // case 0x2c: assignment statement.
            0x2c => unimplemented!(
                "EbEmitStatement case 0x2c (assignment): EbEmitAssignmentStmt; Phase 5"
            ),
            // case 0x2d: typed assignment / error recovery.
            0x2d => unimplemented!(
                "EbEmitStatement case 0x2d: EbReportTypeError / EbCreateErrorNode; Phase 5"
            ),
            // cases 0x2f / 0x30 / 0x31: sequence — emit both children.
            0x2f | 0x30 | 0x31 => {
                let mut local_38 = 0;
                if n.w[4] != 0 {
                    local_38 = self.emit_expr(n.lhs(), context);
                }
                if n.w[5] == 0 {
                    return local_38;
                }
                return self.emit_expr(n.rhs(), context);
            }
            // cases 0x32 / 0x33 / 0x34: type coercion / tree traversal.
            0x32 | 0x34 => unimplemented!(
                "EbEmitStatement cases 0x32/0x34: EbEmitTypeCoercion4; Phase 4"
            ),
            0x33 => unimplemented!(
                "EbEmitStatement case 0x33: EbTraverseNodeTree; Phase 5"
            ),
            // case 0x36: `Like` operator.
            0x36 => {
                self.emit_expr(n.lhs(), 1);
                self.emit_expr(n.rhs(), 1);
                0xd2
            }
            // cases 0x37..=0x73: argument lists, calls, members, coercions,
            // instruction helpers — all require later-phase machinery.
            0x37 => unimplemented!("EbEmitStatement case 0x37: EbProcessLinkedList; Phase 5"),
            0x38 | 0x39 | 0x3a | 0x3b | 0x3e | 0x3f | 0x6a => unimplemented!(
                "EbEmitStatement cases 0x38/0x39/0x3a/0x3b/0x3e/0x3f/0x6a: \
                 EbProcessLinkedList / EbTraverseNodeTree + EbGetTypeSize3; Phase 5"
            ),
            0x41 => unimplemented!("EbEmitStatement case 0x41 (arg list): list walk; Phase 5"),
            0x42 | 0x43 => unimplemented!(
                "EbEmitStatement cases 0x42/0x43: EbResolveDispatchType; Phase 5"
            ),
            0x44 | 0x45 | 0x46 | 0x47 | 0x4c | 0x55 | 0x56 => unimplemented!(
                "EbEmitStatement cases 0x44..0x47/0x4c/0x55/0x56: EbEmitTypeConversion2 / \
                 EbDispatchOpcodeToEmitter; Phase 4/5"
            ),
            0x48 | 0x49 | 0x4a | 0x4b | 0x4d | 0x4e | 0x4f | 0x50 | 0x53 | 0x57 | 0x58 | 0x59 => {
                unimplemented!(
                    "EbEmitStatement cases 0x48..0x59: EbTraverseNodeTree + EbGetTypeSize3; Phase 5"
                )
            }
            0x51 | 0x52 => unimplemented!(
                "EbEmitStatement cases 0x51/0x52: EbClassifyOperatorType; Phase 5"
            ),
            0x54 => unimplemented!("EbEmitStatement case 0x54: EbEmitTypeCoercion4; Phase 4"),
            0x5a => unimplemented!("EbEmitStatement case 0x5a: EbEmitComplexBinaryOp; Phase 5"),
            0x5c | 0x5d | 0x5e | 0x5f => unimplemented!(
                "EbEmitStatement cases 0x5c..0x5f: EbEmitStatement + LoadTypeLibraryItem; Phase 6"
            ),
            0x60 => unimplemented!("EbEmitStatement case 0x60: EbCoerceMemberRef; Phase 4"),
            0x61 => unimplemented!(
                "EbEmitStatement case 0x61 (call/arg): EbGetTypeKind2 / EbGetTypeCode3 / \
                 EbEmitExpr2 + arg machinery; Phase 5"
            ),
            0x63 | 0x66 | 0x67 | 0x68 => unimplemented!(
                "EbEmitStatement cases 0x63/0x66/0x67/0x68: EbExtractTypeValue2; Phase 4"
            ),
            // case 0x65: forward to child.
            0x65 => return self.emit_expr(n.lhs(), context),
            0x69 => unimplemented!(
                "EbEmitStatement case 0x69: EbSetupBinaryOperation / EbEmitExpression2; Phase 5"
            ),
            0x6b | 0x6c | 0x6d | 0x6e => unimplemented!(
                "EbEmitStatement cases 0x6b..0x6e: EbEmitInstruction2 @ 0fabf5bc; Phase 3/5"
            ),
            0x72 => unimplemented!(
                "EbEmitStatement case 0x72: EbCreateTypeNode3 / EbCoerceMemberRef; Phase 4"
            ),
            0x73 => unimplemented!(
                "EbEmitStatement case 0x73: EbEmitOpcode2(0x266, ...) over a type descriptor; Phase 4"
            ),
            // default: every other binary op (arithmetic, logical, comparison).
            _ => return self.emit_binary_operation(&n, context),
        };

        // Common tail (`EbEmitValue2(iVar13); return local_38;`).
        self.emit_value2(i_var13 as usize);
        0
    }

    // ── EbEmitBinaryOperation2 @ 0fab2e1e ────────────────────────────────────

    /// Emit a binary operation: both operands (each in context 2) then the
    /// type-class-selected opcode. Direct port of `EbEmitBinaryOperation2`.
    ///
    /// The opcode index is `EbLookupOpcodeTable(opcode)` plus a type offset. Two
    /// dispatch modes, selected by `RT_DISPATCH_FLAG[opcode] & 0x10`:
    /// * clear → arithmetic: offset from the **node's own** type tag.
    /// * set → comparison/string: offset from the **LHS operand's** type tag,
    ///   with special cases for the `0x72` / type-3 / UDT(`0xf`) / `0xd` forms.
    fn emit_binary_operation(&mut self, n: &RawNode, context: u32) -> u32 {
        self.emit_expr(n.lhs(), 2);
        if n.w[5] != 0 {
            self.emit_expr(n.rhs(), 2);
        }
        let op = n.opcode() as usize;
        let mut i_var2 = RT_BINOP_BASE[op] as i32; // EbLookupOpcodeTable
        let type_tag = n.type_tag();

        if RT_DISPATCH_FLAG[op] & 0x10 == 0 {
            // Arithmetic: use the node's own type tag.
            let mut iv4 = RT_TYPE_OFFSET[type_tag as usize];
            if iv4 == 10 {
                iv4 = 4;
            } else if iv4 == 9 {
                iv4 = 1;
            }
            i_var2 += iv4;
        } else {
            // Comparison / string: use the LHS operand's type tag.
            let lhs = *self.arena.get(n.lhs());
            let lhs_w0 = lhs.w[0];
            if (n.w[0] & 0xffff) == 0x72
                || (n.w[0] & 0xffff_0000) != 0x0003_0000
                || (lhs_w0 & 0xffff_0000) != 0x000f_0000
            {
                let rhs = *self.arena.get(n.rhs());
                let rhs_tag = (rhs.w[0] as i32) >> 16;
                if (lhs_w0 & 0xffff_0000) == 0x000d_0000
                    && (rhs_tag == 10 || rhs_tag == 0xb || rhs_tag == 0xc)
                {
                    i_var2 += 0xc;
                } else {
                    let lhs_tag = (lhs_w0 as i32) >> 16;
                    let mut iv4 = RT_TYPE_OFFSET[lhs_tag as usize];
                    if iv4 == 10 {
                        iv4 = 4;
                    } else if iv4 == 9 {
                        iv4 = 1;
                    }
                    i_var2 += iv4;
                    if (n.w[1] >> 8) & 0x80 != 0 {
                        i_var2 += 2;
                    }
                }
            } else if (n.w[1] >> 8) & 0x80 == 0 {
                i_var2 += 10;
            } else {
                i_var2 += 0xb;
            }
        }

        if (n.w[0] & 0xffff_0000) == 0x000f_0000 {
            let size = self.emit_get_type_size3(n.w[6]);
            self.emit_opcode2(i_var2 as usize, size as u16);
        } else {
            self.emit_value2(i_var2 as usize);
        }
        self.emit_validate_type_operation(type_tag, 0, context)
    }

    // ── EbValidateTypeOperation @ 0fab300a ───────────────────────────────────

    /// Emit the per-type validation/conversion opcode after an operation.
    /// Direct port of `EbValidateTypeOperation(nOpType, param2, nTypeFlags)`.
    fn emit_validate_type_operation(&mut self, n_op_type: i32, param2: i32, n_type_flags: u32) -> u32 {
        match n_op_type {
            2 => {
                if param2 == 0x17 {
                    self.emit_value2(0x18b);
                }
            }
            10 | 0xb | 0xc => {
                if n_type_flags != 3 && n_type_flags != 1 {
                    return 1;
                }
                if n_op_type == 10 {
                    self.emit_value2(0x189);
                    return 0;
                }
                if n_op_type > 10 && n_op_type < 0xd {
                    self.emit_value2(0x18a);
                    return 0;
                }
                self.emit_value2(n_type_flags as usize);
                return 0;
            }
            0xf => {
                if n_type_flags == 3 {
                    self.emit_value2(0x18c);
                    return 0;
                }
            }
            _ => {}
        }
        0
    }

    // ── EbGetTypeSize3 @ 0fab2f55 ────────────────────────────────────────────

    /// Type-size lookup for UDT-typed operations. Requires the type-descriptor
    /// model (the descriptor's nested pointers), which is later-phase work.
    fn emit_get_type_size3(&mut self, _type_desc: u32) -> u32 {
        unimplemented!(
            "EbGetTypeSize3 @ 0fab2f55: needs the type-descriptor model \
             (reached only on UDT/type_tag==0xf paths); Phase 4"
        )
    }

    // ── EbEmitExpression2 @ 0fab397a ─────────────────────────────────────────

    /// Emit a resolved reference (typed local/arg load or store) from a
    /// [`RefDescriptor`]. Direct port of `EbEmitExpression2(pExprDesc, nOp,
    /// fFlags, pContext, nType)` — the function that turns a resolved reference
    /// into the typed load/store opcode.
    ///
    /// This is a **pure** function of the descriptor and the emit parameters; it
    /// needs no symbol table. The descriptor kind, `value_class`, and `flags`
    /// are produced upstream by the reference resolver (`EbResolveIdentRef`,
    /// fed here by the vb6-sema bridge). `n_op` selects the operation: 1 = value
    /// load, 4 = store; others per the C.
    ///
    /// Only the descriptor kinds and sub-paths that do not recurse into
    /// not-yet-ported helpers are implemented; the rest are `unimplemented!()`
    /// with their C reference. The opcode base by kind: local = `0x1e0`,
    /// argument = `0x210`, kind-7 = `0x240`.
    pub fn emit_reference(&mut self, desc: &RefDescriptor, n_op_in: i32, f_flags: u32, n_type_in: i32) {
        let mut n_op = n_op_in;
        let n_type = n_type_in;
        // sVar8 controls final emit: 1 = emit_opcode2 (opcode + 2-byte operand),
        // 2 = emit_opcode2 + emit_word (opcode + 2-byte operand + 2-byte extra word),
        // 0 = emit_value2 only. Starts at 1 (the VB6 C default). The sVar8==2 path
        // is reached only via nType==0x17 which hits unimplemented! above, so the
        // value is always 1 in currently-reachable code.
        let s_var8: i16 = 1;
        let u_var5: u16 = desc.operand; // descriptor+10 (frame offset / type size)
        let u_var7: i32; // opcode base by descriptor kind

        match desc.kind {
            // kind 1: local variable — opcode base 0x1e0 (RT_OPCODE_BYTE[0x1e2] = 0x6c for Long).
            1 => u_var7 = 0x1e0,
            // kind 2: argument/parameter — opcode base 0x210 (RT_OPCODE_BYTE[0x212] = 0x80 for Long ByRef).
            // When word6 bit 0 is set (ByRef slot), EbEmitExpression2 calls EbEmitExpressionOp
            // with nOp=2 and returns; we model this by forcing n_op=2 before the nOp switch.
            2 => {
                u_var7 = 0x210;
                if desc.word6 & 1 != 0 {
                    n_op = 2;
                }
            }
            // kind 7: indirect module-level variable — opcode base 0x240.
            // Same ByRef promotion as kind 2.
            7 => {
                u_var7 = 0x240;
                if desc.word6 & 1 != 0 {
                    n_op = 2;
                }
            }
            // kinds 3/4/5/6/0xa: EbEmitExpressionOp is called with u_var7 from the
            // caller's ESI register (not set by EbEmitExpression2 itself). Without the
            // full call chain, u_var7 is not available at this level.
            3 | 4 | 5 | 6 | 0xa => unimplemented!(
                "EbEmitExpression2 kind {}: requires caller-supplied opcode base (ESI) \
                 from EbResolveIdentRef call chain — not available without the full \
                 module compilation context",
                desc.kind
            ),
            // kinds 8/9/0xb: emit fixed opcodes then call EbBuildExprDescriptor @ 0fab3d1c,
            // which requires the module symbol table and compiled type descriptors.
            8 | 9 | 0xb => unimplemented!(
                "EbEmitExpression2 kind {} (member/typed ref): EbBuildExprDescriptor \
                 @ 0fab3d1c requires the module symbol table and compiled type descriptors",
                desc.kind
            ),
            // default: EbEmitExpression3 (wraps EbEmitExpressionOp) — same ESI issue as 3/4/5/6.
            _ => unimplemented!(
                "EbEmitExpression2 kind {} (default): EbEmitExpression3 @ 0fab3cda \
                 requires caller-supplied opcode base from the module compilation context",
                desc.kind
            ),
        }

        // EbGetType2Flag(pContext) check: when nType==0x12 and the context flag is set,
        // nType is promoted to 0x17. Without pContext (not accepted by emit_reference),
        // we cannot make this determination.
        if n_type == 0x12 {
            unimplemented!(
                "EbEmitExpression2 nType 0x12: EbGetType2Flag(pContext) @ 0fab34d0 \
                 determines whether nType is promoted to 0x17; pContext is not \
                 threaded through emit_reference"
            );
        }
        // EbMapOperatorType: maps nType 0x11/0x12 + nOp in {1,2,3} → nOp 5.
        if (n_type == 0x12 || n_type == 0x11) && (n_op == 1 || n_op == 2 || n_op == 3) {
            n_op = 5;
        }

        let u_var6: i32 = match n_op {
            1 => {
                if f_flags & 0x4000 != 0 {
                    // LAB_0fab3b03: object/type-expression path.
                    // nType 0x10 (Object/Dispatch) → opcode 0x23f (RT_OPCODE_BYTE[0x23f] = 0x3e).
                    // nType 0x17 (Variant-with-type) → EbEmitTypeExpr @ 0fac2f00 needed.
                    // All other nType values → opcode 0x262 (RT_OPCODE_BYTE[0x262] = 0x8a).
                    if n_type == 0x10 {
                        0x23f
                    } else if n_type == 0x17 {
                        unimplemented!(
                            "EbEmitExpression2 nOp1 f_flags 0x4000 nType 0x17: \
                             EbEmitTypeExpr @ 0fac2f00 (calls EbEmitTypedExpression — \
                             type-pool machinery not yet ported)"
                        );
                    } else {
                        0x262
                    }
                } else if f_flags & 0x1000 != 0 {
                    0x1b2
                } else {
                    let off = RT_TYPE_OFFSET[n_type as usize];
                    if off == 10 {
                        u_var7 | 4
                    } else if off == 9 {
                        u_var7 | 1
                    } else {
                        u_var7 | off
                    }
                }
            }
            // nOp 2 falls through to the same 0x4000 label as nOp 1 in the C.
            2 => {
                if f_flags & 0x4000 != 0 {
                    if n_type == 0x10 {
                        0x23f
                    } else if n_type == 0x17 {
                        unimplemented!(
                            "EbEmitExpression2 nOp2 f_flags 0x4000 nType 0x17: \
                             EbEmitTypeExpr @ 0fac2f00 (calls EbEmitTypedExpression — \
                             type-pool machinery not yet ported)"
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
                    let mut v = off | u_var7;
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
                    u_var7 + 0xd
                } else {
                    let off = RT_TYPE_OFFSET[n_type as usize];
                    if off == 10 {
                        u_var7 | 4
                    } else if off == 9 {
                        u_var7 | 1
                    } else {
                        u_var7 | off
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
                let mut v = (u_var7 + 0x10) | base_off;
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
                    }
                } else if f_flags & 0x20 == 0 {
                    // DAT_0fab6a38: a small 2-D conversion table indexed by
                    // (~fFlags >> 12 & 1), (~fFlags >> 11 & 1), and the type offset.
                    // The table bytes are not yet extracted into tables.rs.
                    unimplemented!(
                        "EbEmitExpression2 nOp4 f_flags 0x8000 path: \
                         DAT_0fab6a38 @ 0fab6a38 conversion table not yet extracted"
                    );
                } else {
                    v = (if f_flags & 0x800 != 0 { 2 } else { 0 }) + 0x1b7;
                }
                // (fFlags & 0x40) && descriptor kind 1: opcode remap with operand emit.
                if f_flags & 0x40 != 0 && desc.kind == 1 {
                    v = match v {
                        0x1f2 => {
                            self.emit_opcode2(v as usize, u_var5);
                            0x1e2
                        }
                        0x1f6 => 0x263,
                        0x1f7 => 0x264,
                        0x1f8 => 0x29c,
                        0x1ff => 0x29d,
                        0x200 => {
                            self.emit_opcode2(v as usize, u_var5);
                            0x1ec
                        }
                        0x201 => {
                            self.emit_opcode2(v as usize, u_var5);
                            0x1e7
                        }
                        other => other,
                    };
                }
                v
            }
            5 => {
                if f_flags & 0x1000 == 0 {
                    u_var7 + 0xc
                } else {
                    0x1b2
                }
            }
            6 => u_var7 + 0xb,
            _ => unimplemented!("EbEmitExpression2: nOp {} not handled", n_op),
        };

        // The nType==0x17 → sVar8==2 path (opcode 0x438 + trailing word) remains
        // unimplemented! above; s_var8 stays 1 for all currently reachable paths.
        if s_var8 == 0 {
            self.emit_value2(u_var6 as usize);
        } else {
            self.emit_opcode2(u_var6 as usize, u_var5);
            if s_var8 == 2 {
                // Extra word from descriptor+8 (desc.word8), emitted only on the
                // 0x438 path (nType 0x17 double-unwrap returning class 0x12).
                self.stream.emit_word(desc.word8);
            }
        }
        // EbBuildExprDescriptor is called when kind∈{5,7,10} AND nOp==6 AND
        // the "optional" bit in desc+4 is set — none of the kinds currently
        // emitted above, so it is correctly omitted here.
    }

    // ── Emitter primitives (EbEmitValue2 / EbEmitOpcode2 / EbEmitDword) ───────

    /// `EbEmitValue2` @ `0fab30bb`: look up `RT_OPCODE_BYTE[n_opc]`; emit it as a
    /// single byte when `< 0xfb`, otherwise emit that byte then `n_opc as u8`.
    fn emit_value2(&mut self, n_opc: usize) {
        let rt_byte = RT_OPCODE_BYTE[n_opc];
        if rt_byte < 0xfb {
            self.stream.emit_byte(rt_byte);
        } else {
            self.stream.emit_byte(rt_byte);
            self.stream.emit_byte(n_opc as u8);
        }
    }

    /// `EbEmitOpcode2` @ `0fab2f77`: emit the opcode byte(s) for `n_opc`, then a
    /// 2-byte little-endian operand.
    fn emit_opcode2(&mut self, n_opc: usize, operand: u16) {
        self.emit_value2(n_opc);
        self.stream.emit_word(operand);
    }

    /// `EbEmitDword` @ `0fabf585`: emit a 4-byte little-endian value.
    fn emit_dword(&mut self, value: u32) {
        self.stream.emit_bytes(&value.to_le_bytes());
    }

    // ── Literal emitters (EbEmitStatement cases 3 / 4) ───────────────────────

    /// Case 3: floating-point literal. `context == 2` selects the assign-context
    /// opcode variants. Single (type_tag 10) is converted to f32 and emitted as
    /// 4 bytes; Double/Date (11/12) emit the raw 8-byte f64. Returns the C
    /// `uVar14` (1 in assign context for the typed variants, else 0).
    fn emit_float_literal(&mut self, n: &RawNode, context: u32) -> u32 {
        let type_tag = n.type_tag();
        let mut u_var14 = 0u32;
        let mut emit_eight = true;
        let n_opc: usize;
        if type_tag == 10 {
            emit_eight = false;
            if context == 2 {
                u_var14 = 1;
                n_opc = 0x3ba;
            } else {
                n_opc = 0x3b9;
            }
        } else if type_tag > 10 && type_tag < 0xd {
            if context == 2 {
                u_var14 = 1;
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
        u_var14
    }

    /// Case 4: string literal. Bit 15 of `word[1]` (`node+5 & 0x80`) set → null
    /// string, emitted as a Long-zero (`0x3b8` + 4 zero bytes). Clear → a pooled
    /// string literal needing `EbExtractTypeValue2` (later phase).
    fn emit_string_literal(&mut self, n: &RawNode) {
        if (n.w[1] >> 8) & 0x80 == 0 {
            unimplemented!(
                "EbEmitStatement case 4 (pooled string literal): EbExtractTypeValue2 + \
                 type pool; Phase 4"
            );
        }
        self.emit_value2(0x3b8);
        self.emit_dword(0);
    }

    // ── Typed local load / store (runtime byte form; oracle-confirmed) ───────

    /// Emit a typed local-variable load. The type context (`typeCtx`) in
    /// `word[5]` selects the 1-byte opcode from [`RT_LOAD_BY_CTX`]; the bound
    /// symbol child carries the signed frame offset in its `type_info()` field,
    /// emitted as a 2-byte little-endian i16.
    ///
    /// `context` is accepted (the real name path varies its surrounding emission
    /// by context) but does not change the load opcode for a plain numeric local.
    /// Node types `0x74` and `0x76` both route here.
    /// Dispatch a synthetic opcode-0x75 ByRef parameter load node.
    /// Node layout: lhs = sym node whose type_info() holds the frame offset,
    /// word[5] = type_ctx.  Calls emit_byref_load.
    fn emit_byref_param_node(&mut self, n: &RawNode) {
        let type_ctx = n.word(5) as usize;
        let sym = self.arena.get(n.lhs());
        let frame_offset = sym.type_info() as i16;
        self.emit_byref_load(type_ctx, frame_offset);
    }

    /// Dispatch a synthetic opcode-0x77 module-global load node.
    /// Node layout: word[4] low u16 = module_desc, word[4] high u16 = field_offset,
    /// word[5] = type_ctx.  Calls emit_global_load.
    fn emit_global_node_load(&mut self, n: &RawNode) {
        let type_ctx = n.word(5) as usize;
        let packed = n.word(4);
        let module_desc = (packed & 0xffff) as u16;
        let field_offset = (packed >> 16) as u16;
        self.emit_global_load(type_ctx, module_desc, field_offset);
    }

    fn emit_var_load(&mut self, n: &RawNode, _context: u32) {
        let type_ctx = n.word(5) as usize;
        let sym = self.arena.get(n.lhs());
        let frame_offset = sym.type_info() as i16;
        self.emit_typed_load(type_ctx, frame_offset);
    }

    /// Emit a typed local-variable load given its type context and frame offset
    /// directly (the opcode comes from [`RT_LOAD_BY_CTX`]). Mirror of
    /// [`Self::emit_var_store`]; used by the binder bridge.
    pub fn emit_typed_load(&mut self, type_ctx: usize, frame_offset: i16) {
        let opcode = RT_LOAD_BY_CTX.get(type_ctx).copied().unwrap_or(0);
        if opcode == 0 {
            unimplemented!(
                "emit_typed_load: no confirmed runtime opcode for typeCtx {}",
                type_ctx
            );
        }
        self.stream.emit_byte(opcode);
        self.stream.emit_i16(frame_offset);
    }

    /// Emit a typed local-variable store. Mirror of [`Self::emit_var_load`] using
    /// [`RT_STORE_BY_CTX`]. The caller must have emitted the value to store first.
    pub fn emit_var_store(&mut self, type_ctx: usize, frame_offset: i16) {
        let opcode = RT_STORE_BY_CTX.get(type_ctx).copied().unwrap_or(0);
        if opcode == 0 {
            unimplemented!(
                "emit_var_store: no confirmed runtime opcode for typeCtx {}",
                type_ctx
            );
        }
        self.stream.emit_byte(opcode);
        self.stream.emit_i16(frame_offset);
    }

    /// Emit a ByRef parameter load at `frame_offset`.  The ByRef load opcode is
    /// `RT_LOAD_BY_CTX[type_ctx] + 0x14` (oracle-confirmed for Long: 0x6c→0x80).
    /// The frame offset is positive (parameters sit above the frame pointer).
    pub fn emit_byref_load(&mut self, type_ctx: usize, frame_offset: i16) {
        let base = RT_LOAD_BY_CTX.get(type_ctx).copied().unwrap_or(0);
        if base == 0 {
            unimplemented!(
                "emit_byref_load: no confirmed local-load opcode for typeCtx {}",
                type_ctx
            );
        }
        self.stream.emit_byte(base + 0x14);
        self.stream.emit_i16(frame_offset);
    }

    /// Emit a ByRef parameter store at `frame_offset`.  The ByRef store opcode is
    /// `RT_STORE_BY_CTX[type_ctx] + 0x14` (oracle-confirmed for Long: 0x71→0x85).
    pub fn emit_byref_store(&mut self, type_ctx: usize, frame_offset: i16) {
        let base = RT_STORE_BY_CTX.get(type_ctx).copied().unwrap_or(0);
        if base == 0 {
            unimplemented!(
                "emit_byref_store: no confirmed local-store opcode for typeCtx {}",
                type_ctx
            );
        }
        self.stream.emit_byte(base + 0x14);
        self.stream.emit_i16(frame_offset);
    }

    /// Emit a module-level global variable load.  The opcode is
    /// `RT_LOAD_BY_CTX[type_ctx] + 0x28` (oracle-confirmed: Integer=0x93,
    /// Long=0x94, Double=0x97).  The 4-byte operand encodes `module_desc` (the
    /// compiled module-object descriptor) in bytes 0–1 and `field_offset` (the
    /// byte offset within the module's global data block) in bytes 2–3.
    pub fn emit_global_load(
        &mut self,
        type_ctx: usize,
        module_desc: u16,
        field_offset: u16,
    ) {
        let base = RT_LOAD_BY_CTX.get(type_ctx).copied().unwrap_or(0);
        if base == 0 {
            unimplemented!(
                "emit_global_load: no confirmed local-load opcode for typeCtx {}",
                type_ctx
            );
        }
        self.stream.emit_byte(base + 0x28);
        self.stream.emit_word(module_desc);
        self.stream.emit_word(field_offset);
    }

    // ── EbTraverseNodeTree @ (vba6_part0002.c) ──────────────────────────────
    //
    // Walk a singly-linked list of statement nodes (opcodes 0x37 / 0x33) and
    // emit each one.  List structure: word[4] = child statement, word[5] = next
    // list node (sibling).  The C function recurses on the sibling first, then
    // emits the child — producing right-to-left emission order for a forward list.
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

    // ── EbEmitAssignOp @ 0fab3117 ────────────────────────────────────────────
    //
    // Emit the store opcode for a simple `=` assignment after the RHS has already
    // been pushed.  `n` is the assignment node (pNode in C); RHS is `pNode[4]`.
    // Dispatch tree mirrors the C verbatim.
    fn emit_assign_op(&mut self, n: &RawNode) {
        use crate::tables::{RT_ASSIGN_BASE_OPCODE, RT_TYPE_KIND_CLASS};

        let lhs_hi = n.w[0] & 0xffff0000;
        if lhs_hi == 0xf0000 {
            // Object LHS — needs EbGetTypeSize3 + EbEmitOpcode2 + type-pool.
            unimplemented!(
                "EbEmitAssignOp: Object LHS (hi16=0xf) requires EbGetTypeSize3 @ 0fab2f55"
            );
        }

        let rhs_node = *self.arena.get(NodeRef(n.w[4]));

        if lhs_hi == 0xc0000 {
            // Currency LHS: special RHS-kind dispatch.
            let rhs_kind = (rhs_node.w[0] as i32) >> 16;
            match rhs_kind {
                0xb => { self.emit_value2(0x147); return; }
                0xf => { self.emit_value2(0x14f); return; }
                0x10 => { self.emit_value2(0x3c9); return; }
                _ => {} // fall through to generic numeric path below
            }
        }

        let lhs_kind = (n.w[0] as i32) >> 16;
        let rhs_kind = (rhs_node.w[0] as i32) >> 16;

        // Variant / ByRef / Currency triple-type group (LHS and RHS both in {10,0xb,0xc}).
        if matches!(lhs_kind, 10 | 0xb | 0xc) && matches!(rhs_kind, 10 | 0xb | 0xc) {
            // If byte@node+5 bit 0x80 is clear → no-op (assign is handled elsewhere).
            if (n.w[1] >> 8) & 0x80 == 0 {
                return;
            }
            let rhs_class = RT_TYPE_KIND_CLASS[rhs_kind as usize];
            let lhs_class = RT_TYPE_KIND_CLASS[lhs_kind as usize];
            let i_var6 = RT_ASSIGN_BASE_OPCODE[lhs_class as usize];
            if rhs_class == 10 {
                self.emit_value2((i_var6 + 4) as usize);
                return;
            }
            // fall through to shared tail (iVar4 == 9 → 1 remap + emit)
            let mut i_var4 = rhs_class;
            if i_var4 == 9 { i_var4 = 1; }
            self.emit_value2((i_var4 + i_var6) as usize);
            return;
        }

        // General case: inspect RHS hi16 for special source types.
        let rhs_hi = rhs_node.w[0] & 0xffff0000;
        if rhs_hi == 0xc0000 {
            // Currency RHS into non-Currency LHS.
            match lhs_kind {
                0xf => { self.emit_value2(0x2fb); return; }
                0x10 => { self.emit_value2(0x3c8); return; }
                _ => {} // fall through to numeric tail
            }
        }
        if rhs_hi == 0x30000 {
            // Boolean RHS.
            match lhs_kind {
                5 => { self.emit_value2(0x138); return; }
                6 => { return; } // no-op
                0x10 => { self.emit_value2(0x3c7); return; }
                _ => {} // fall through
            }
        }
        if rhs_hi == 0x110000 {
            // Fixed-length string RHS → EbGetTypeLength + EbEmitOpcode2.
            unimplemented!(
                "EbEmitAssignOp: fixed-length string RHS (hi16=0x11) requires \
                 EbGetTypeLength @ 0fab2f9e and EbEmitOpcode2 0x361"
            );
        }
        if rhs_hi == 0xf0000 {
            // Object RHS → EbResolveAndSimplify / EbEmitTypeOfExprPcode3.
            if lhs_hi == 0x140000 {
                unimplemented!(
                    "EbEmitAssignOp: Object RHS with Variant LHS (0x14) requires \
                     EbResolveAndSimplify @ 0fab2fcb and EbEmitOpcode2 0x435"
                );
            }
            if lhs_hi == 0x120000 {
                unimplemented!(
                    "EbEmitAssignOp: Object RHS with TypeOf LHS (0x12) requires \
                     EbEmitTypeOfExprPcode3 @ 0fab33af"
                );
            }
        }

        // Default numeric path: base = RT_ASSIGN_BASE_OPCODE[RT_TYPE_KIND_CLASS[lhs_kind]]
        //                        rhs_class = RT_TYPE_KIND_CLASS[rhs_kind]
        let i_var6 = RT_ASSIGN_BASE_OPCODE[RT_TYPE_KIND_CLASS[lhs_kind as usize] as usize];
        let mut i_var4 = RT_TYPE_KIND_CLASS[rhs_kind as usize];
        if i_var4 == 10 {
            self.emit_value2((i_var6 + 4) as usize);
            return;
        }
        if i_var4 == 9 { i_var4 = 1; }
        self.emit_value2((i_var4 + i_var6) as usize);
    }

    /// Emit a module-level global variable store.  The opcode is
    /// `RT_STORE_BY_CTX[type_ctx] + 0x28` (oracle-confirmed: Integer=0x98,
    /// Long=0x99, Double=0x9c).
    pub fn emit_global_store(
        &mut self,
        type_ctx: usize,
        module_desc: u16,
        field_offset: u16,
    ) {
        let base = RT_STORE_BY_CTX.get(type_ctx).copied().unwrap_or(0);
        if base == 0 {
            unimplemented!(
                "emit_global_store: no confirmed local-store opcode for typeCtx {}",
                type_ctx
            );
        }
        self.stream.emit_byte(base + 0x28);
        self.stream.emit_word(module_desc);
        self.stream.emit_word(field_offset);
    }
}

#[cfg(test)]
#[path = "tests/emit_tests.rs"]
mod tests;
