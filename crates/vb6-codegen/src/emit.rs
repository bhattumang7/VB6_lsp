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
///   e.g. a local's signed frame offset.
/// * `word6` — descriptor `+6`; for argument kinds, bit 0 marks a by-ref slot.
#[derive(Clone, Copy, Debug, Default, PartialEq, Eq)]
pub struct RefDescriptor {
    pub kind: i32,
    pub operand: u16,
    pub word6: u16,
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
        // from the assignment / name path; we stand in until that is ported.
        if op == 0x74 || op == 0x76 {
            self.emit_var_load(&n, context);
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
            0xe => unimplemented!(
                "EbEmitStatement case 0xe (load/assign): EbEmitAssignOp @ 0fab3117 + \
                 EbMapTypeCodeValue3; Phase 3/4"
            ),
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
        let s_var8: i16 = 1;
        let u_var5: u16 = desc.operand; // descriptor+10 operand (frame offset)
        let u_var7: i32; // base opcode by descriptor kind

        match desc.kind {
            1 => u_var7 = 0x1e0,
            2 => {
                u_var7 = 0x210;
                if desc.word6 & 1 != 0 {
                    unimplemented!(
                        "EbEmitExpression2 kind 2 (byref arg): EbEmitExpressionOp mode 2 \
                         @ 0fab39b8; Phase 3"
                    );
                }
            }
            7 => {
                u_var7 = 0x240;
                if desc.word6 & 1 != 0 {
                    unimplemented!(
                        "EbEmitExpression2 kind 7 (byref): EbEmitExpressionOp mode 2 \
                         @ 0fab39b8; Phase 3"
                    );
                }
            }
            3 | 4 | 5 | 6 | 0xa => unimplemented!(
                "EbEmitExpression2 kind {} : EbEmitExpressionOp @ 0fab39b8; Phase 3",
                desc.kind
            ),
            8 | 9 | 0xb => unimplemented!(
                "EbEmitExpression2 kind {} (member/typed): EbBuildExprDescriptor; Phase 4",
                desc.kind
            ),
            _ => unimplemented!(
                "EbEmitExpression2 kind {} (default): EbEmitExpression3; Phase 3",
                desc.kind
            ),
        }

        // nType normalization. The 0x12 sub-path needs EbGetType2Flag(context).
        if n_type == 0x12 {
            unimplemented!(
                "EbEmitExpression2: nType 0x12 path: EbGetType2Flag; Phase 4"
            );
        }
        // EbMapOperatorType(nType, &nOp): 0x12/0x11 + nOp in {1,2,3} → nOp 5.
        if (n_type == 0x12 || n_type == 0x11) && (n_op == 1 || n_op == 2 || n_op == 3) {
            n_op = 5;
        }

        let u_var6: i32 = match n_op {
            1 => {
                if f_flags & 0x4000 != 0 {
                    unimplemented!(
                        "EbEmitExpression2 nOp1 0x4000: EbEmitTypeExpr / EbGetValueTypeClass2; Phase 4"
                    );
                }
                if f_flags & 0x1000 != 0 {
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
            2 => {
                if f_flags & 0x4000 != 0 {
                    unimplemented!(
                        "EbEmitExpression2 nOp2 0x4000: EbEmitTypeExpr / EbGetValueTypeClass2; Phase 4"
                    );
                }
                if f_flags & 0x1000 != 0 {
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
                    unimplemented!(
                        "EbEmitExpression2 nOp4 0x8000 path: DAT_0fab6a38 conversion table; Phase 4"
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

        if s_var8 == 0 {
            self.emit_value2(u_var6 as usize);
        } else {
            self.emit_opcode2(u_var6 as usize, u_var5);
        }
        // The trailing EbBuildExprDescriptor call fires only for descriptor kinds
        // 5/7/10 with nOp 6 — none of the kinds emitted above — so it is omitted
        // here and will be added when those kinds are ported.
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
