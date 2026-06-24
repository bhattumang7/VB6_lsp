//! Expression / statement code generation: the runtime P-code byte stream.
//!
//! [`Emitter::emit_expr`] walks a bound expression/statement tree ([`NodeArena`])
//! and writes the dense little-endian P-code byte stream the runtime interpreter
//! executes. Dispatch is a single `match` on the node opcode
//! (`(short)word[0]`, [`RawNode::opcode`]); opcodes outside `1..=0x73` emit
//! nothing. Cases that finish emission `return`; cases that share the common tail
//! evaluate to the trailing opcode the tail emits.
//!
//! Opcode bytes are never hard-coded: every opcode is resolved through the
//! runtime opcode-byte table ([`RT_OPCODE_BYTE`]) by [`Emitter::emit_value2`] /
//! [`Emitter::emit_opcode2`], which apply the 1-or-2-byte escape encoding.
//!
//! ## Synthetic load/store nodes
//! Typed local / argument / global load and store byte sequences are driven by
//! synthetic node opcodes `0x74`..=`0x77` that the lowering pass builds; they are
//! handled ahead of the `1..=0x73` dispatch guard.
//!
//! ## Deferred cases
//! Cases that depend on machinery not yet built — the module symbol table, the
//! type/string pool, the type-conversion / instruction emitters — are
//! `unimplemented!()` with a description of what they need. They never emit a
//! guessed byte and never silently fall through.

use crate::buffer::PcodeStream;
use crate::node::{NodeArena, NodeRef, RawNode};
use crate::tables::{RT_BINOP_BASE, RT_CALL_TYPECODE, RT_DISPATCH_FLAG, RT_LOAD_BY_CTX, RT_OPCODE_BYTE, RT_RESULT_TYPE, RT_STORE_BY_CTX, RT_TYPE_OFFSET};
use crate::type_pool::TypePool;

/// Compute the call type-code for a call site: index [`RT_CALL_TYPECODE`] by the
/// reference-vs-value path.
///
/// * value path (`is_ref` false): index `(callee_type != 1) + (mask) * 2` → 0..3.
/// * reference path (`is_ref` true): index `(callee_type != 1) + 4` → 4..5.
///
/// `callee_type` and `mask` come from the callee's resolved descriptor (the
/// symbol-table model supplies them); this kernel is a pure function of its
/// inputs and the extracted table.
pub fn call_type_code(callee_type: i32, is_ref: bool, mask: bool) -> u16 {
    let idx = if is_ref {
        (callee_type != 1) as usize + 4
    } else {
        (callee_type != 1) as usize + (mask as usize) * 2
    };
    RT_CALL_TYPECODE[idx]
}

/// Map a call type-code to its runtime call opcode.
pub fn map_call_type_code(code: u16) -> u16 {
    match code {
        0x300 => 0x16a,
        0x310 => 0x169,
        0x320 => 0x34f,
        0x340 => 0x16b,
        0x350 => 0x16c,
        _ => 0x446,
    }
}

/// A resolved-reference descriptor — the input to [`Emitter::emit_reference`].
/// A reference resolver (the vb6-sema bridge) populates it; the emitter only
/// reads it to choose the typed load/store opcode.
///
/// * `kind` — storage class (1 = local, 2 = argument, 7 = indirect module-level);
///   selects the opcode base.
/// * `operand` — the 2-byte operand emitted after the opcode (a local's signed
///   frame offset, or a type size).
/// * `word6` — for argument / global kinds, bit 0 marks a ByRef slot (forces the
///   ByRef operation path). For the operator-reference kinds (8/9/0xb) it is the
///   opcode operand word (descriptor `+6`).
/// * `word8` — the descriptor's `+8` low word: the trailing word for kind 8, the
///   opcode operand for kind 0xb, and the extra word on the nested-type-expression
///   path (reached via the Variant-with-type chain).
/// * `flags1` — the descriptor's `+4` flag byte; bit `0x04` gates the finalize
///   tail for the operator-reference kinds.
#[derive(Clone, Copy, Debug, Default, PartialEq, Eq)]
pub struct RefDescriptor {
    pub kind: i32,
    pub operand: u16,
    pub word6: u16,
    pub word8: u16,
    pub flags1: u8,
}

/// A resolved call site — the input to [`Emitter::emit_call`]. The binder
/// (symbol-table model) populates it from the resolved callee; the emitter only
/// reads it. Fields mirror the bound call node the runtime emitter consumes:
///
/// * `kind` — the callee's call-convention kind (4..8); selects the dispatch
///   record and the call type-code.
/// * `byref` — 0 = by-value / value-returning, 1 = by-reference (method/Sub).
/// * `flags` — the call node's flag word (`word[1]`): bits `0x8000` / `0x800` /
///   `0x2000` / `0x1000` / `0x200` select the emission path; byte 1 carries the
///   `0x20` / `0x40` sub-flags.
/// * `node_word0` — the call node's `word[0]` (type tag in the high half, type
///   region in the high 16 bits).
/// * `callee` — the callee reference sub-node (`word[6]`), emitted with context 6.
/// * `arg_list` — the argument-list node (`word[5]`), or null.
/// * `member_id` — the callee's member-dispatch id (`word[7]`), emitted as a
///   trailing word on the dispatch path.
/// * `size` — the callee's resolved type size (from `word[8]`), emitted only
///   when the dispatch record requests a size operand.
#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub struct CallDescriptor {
    pub kind: i32,
    pub byref: i32,
    pub flags: u32,
    pub node_word0: u32,
    pub callee: NodeRef,
    pub arg_list: NodeRef,
    pub member_id: u16,
    pub size: u16,
}

/// Drives [`Emitter::emit_expr`] over a [`NodeArena`], writing the runtime
/// P-code byte stream.
/// The module symbol context a member-reference (`0x60`) emission needs: the
/// compiled records heap the resolver reads, the member's byte offset within it
/// (`EbGetExprContext`), the compiler-context flag byte (`in_ECX + 0xc`), and the
/// binder-resolved convention `(kind, byref)` for the category-4 path.
#[derive(Clone, Debug, Default)]
pub struct SymbolContext {
    pub heap: Vec<u8>,
    pub member_off: usize,
    pub ctx_flag_c: u8,
    pub binding: Option<(i32, i32)>,
}

pub struct Emitter<'a> {
    arena: &'a NodeArena,
    stream: PcodeStream,
    type_pool: TypePool,
    sym: Option<SymbolContext>,
}

impl<'a> Emitter<'a> {
    pub fn new(arena: &'a NodeArena) -> Self {
        Self {
            arena,
            stream: PcodeStream::new(),
            type_pool: TypePool::new(),
            sym: None,
        }
    }

    /// Attach the module symbol context used by member-reference (`0x60`)
    /// emission.
    pub fn with_symbol_context(mut self, sym: SymbolContext) -> Self {
        self.sym = Some(sym);
        self
    }

    /// The type-intern pool accumulated during emission (for inspection/tests).
    pub fn type_pool(&self) -> &TypePool {
        &self.type_pool
    }

    /// The runtime P-code bytes emitted so far.
    pub fn bytes(&self) -> &[u8] {
        self.stream.bytes()
    }

    /// Consume the emitter, yielding the full runtime P-code byte stream.
    pub fn into_bytes(self) -> Vec<u8> {
        self.stream.into_bytes()
    }

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
                    unimplemented!(
                        "EbEmitAssignmentStmt compound-op store (flag 0x400 + 0x69 \
                         LHS); Phase 6"
                    );
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
            // for one).
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

    // ── Binary operations ────────────────────────────────────────────────────

    /// Emit a binary operation: both operands (each in context 2) then the
    /// type-class-selected opcode.
    ///
    /// The opcode index is the operation's base ([`RT_BINOP_BASE`]) plus a type
    /// offset. Two dispatch modes, selected by `RT_DISPATCH_FLAG[op] & 0x10`:
    /// * clear → arithmetic: offset from the **node's own** type tag.
    /// * set → comparison / string: offset from the **left operand's** type tag,
    ///   with special cases for the `0x72` / type-3 / object(`0xf`) / `0xd` forms.
    fn emit_binary_operation(&mut self, n: &RawNode, context: u32) -> u32 {
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
    fn emit_validate_type_operation(&mut self, op_type: i32, variant: i32, type_flags: u32) -> u32 {
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
    fn emit_get_type_size3(&self, type_desc: u32) -> u32 {
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
    fn emit_object_type(&mut self, n: &RawNode, context: u32) -> u32 {
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
    fn emit_calculate_struct_size(&self, list: NodeRef, flag: bool) -> i32 {
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
    fn emit_type_coercion4(&mut self, target: i32, src: NodeRef) {
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
    fn emit_type_conversion2(&mut self, target: i32, src: NodeRef, explicit: bool) {
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
    fn emit_expression_code2(&mut self, store: bool, node: NodeRef, type_info: u32) {
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
    fn emit_typed_node(&mut self, node: NodeRef, mode: u32) {
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
    fn complex_binop_type_word(&mut self, p: NodeRef, fallback: u32) -> u16 {
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
    fn emit_complex_binary_op(&mut self, node: NodeRef) {
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
    fn emit_dispatch_opcode(&mut self, start: NodeRef, depth: i32, type_tag: i32) {
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
    fn emit_find_actual_node(&mut self, mut node: NodeRef, flags: u32, mut depth: i32) -> NodeRef {
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
    fn emit_traverse_expr_tree3(&mut self, mut node: NodeRef, mut depth: i32) {
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

    fn emit_instruction2(&mut self, node: NodeRef, opcode: usize, has_arg: bool, is_call: bool) {
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
    fn emit_setup_binary_operation(&mut self, node: NodeRef) -> RefDescriptor {
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
            // kinds 3/4/5/6/0xa: the opcode base is supplied by the resolver's
            // call chain, not available without the full module context.
            3 | 4 | 5 | 6 | 0xa => unimplemented!(
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
                // (f_flags & 0x40) with a local descriptor: opcode remap, some
                // variants emitting an extra operand first.
                if f_flags & 0x40 != 0 && desc.kind == 1 {
                    v = match v {
                        0x1f2 => {
                            self.emit_opcode2(v as usize, operand_word);
                            0x1e2
                        }
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
    fn map_operator_type(n_type: i32, n_op: &mut i32) {
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
    fn expr2_finalize_tail(&mut self, n_op: i32, _f_flags: u32, _n_type: i32, flags1: u8) {
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

    // ── Output primitives ────────────────────────────────────────────────────

    /// Resolve `n_opc` through [`RT_OPCODE_BYTE`] and emit it: a single byte when
    /// `< 0xfb`, otherwise the escape byte followed by `n_opc as u8`.
    fn emit_value2(&mut self, n_opc: usize) {
        let rt_byte = RT_OPCODE_BYTE[n_opc];
        if rt_byte < 0xfb {
            self.stream.emit_byte(rt_byte);
        } else {
            self.stream.emit_byte(rt_byte);
            self.stream.emit_byte(n_opc as u8);
        }
    }

    /// Emit the opcode byte(s) for `n_opc`, then a 2-byte little-endian operand.
    fn emit_opcode2(&mut self, n_opc: usize, operand: u16) {
        self.emit_value2(n_opc);
        self.stream.emit_word(operand);
    }

    /// Emit a 4-byte little-endian value.
    fn emit_dword(&mut self, value: u32) {
        self.stream.emit_bytes(&value.to_le_bytes());
    }

    // ── Statement-list walks ─────────────────────────────────────────────────

    /// Walk a forward-linked statement list (opcodes `0x37` / `0x33`), emitting
    /// each list element's child (`word[4]`) in order with `mode`, then emit the
    /// trailing non-list node, if any.
    fn process_linked_list(&mut self, list: NodeRef, mode: u32) {
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
    fn emit_float_literal(&mut self, n: &RawNode, context: u32) -> u32 {
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
    fn emit_string_literal(&mut self, n: &RawNode) {
        if (n.w[1] >> 8) & 0x80 == 0 {
            unimplemented!(
                "pooled string literal: needs the type/string pool; Phase 4"
            );
        }
        self.emit_value2(0x3b8);
        self.emit_dword(0);
    }

    // ── Typed local / argument / global load and store ───────────────────────

    /// Dispatch a synthetic ByRef parameter load node (opcode 0x75): the bound
    /// symbol child carries the frame offset in `type_info()`; `word[5]` is the
    /// type context.
    fn emit_byref_param_node(&mut self, n: &RawNode) {
        let type_ctx = n.word(5) as usize;
        let sym = self.arena.get(n.lhs());
        let frame_offset = sym.type_info() as i16;
        self.emit_byref_load(type_ctx, frame_offset);
    }

    /// Dispatch a synthetic module-global load node (opcode 0x77): `word[4]` low
    /// u16 = module descriptor, high u16 = field offset; `word[5]` = type context.
    fn emit_global_node_load(&mut self, n: &RawNode) {
        let type_ctx = n.word(5) as usize;
        let packed = n.word(4);
        let module_desc = (packed & 0xffff) as u16;
        let field_offset = (packed >> 16) as u16;
        self.emit_global_load(type_ctx, module_desc, field_offset);
    }

    /// Emit a typed local-variable load. The synthetic node carries the frame
    /// offset in the bound symbol child's `type_info()` and the type context in
    /// `word[5]`. Node types `0x74` and `0x76` both route here.
    fn emit_var_load(&mut self, n: &RawNode, _context: u32) {
        let type_ctx = n.word(5) as usize;
        let sym = self.arena.get(n.lhs());
        let frame_offset = sym.type_info() as i16;
        self.emit_typed_load(type_ctx, frame_offset);
    }

    /// Emit a typed local-variable load from its type context and frame offset.
    /// The opcode comes from [`RT_LOAD_BY_CTX`]; the frame offset follows as a
    /// 2-byte signed little-endian value. Mirror of [`Self::emit_var_store`].
    pub fn emit_typed_load(&mut self, type_ctx: usize, frame_offset: i16) {
        // Byte (ctx 7) has a 2-byte escape-paged load opcode (`fc e0`) that the
        // single-byte RT_LOAD_BY_CTX shortcut cannot hold; emit via the value-
        // emitter load index 0x1e0 (RT_OPCODE_BYTE[0x1e0] = 0xfc → escape).
        if type_ctx == 7 {
            self.emit_value2(0x1e0);
            self.stream.emit_i16(frame_offset);
            return;
        }
        // String (ctx 8): the BSTR-pointer load (0x6c) via the value-emitter load
        // index 0x1e7 (string value-class).
        if type_ctx == 8 {
            self.emit_value2(0x1e7);
            self.stream.emit_i16(frame_offset);
            return;
        }
        let opcode = RT_LOAD_BY_CTX.get(type_ctx).copied().unwrap_or(0);
        if opcode == 0 {
            unimplemented!("no load opcode for type context {}", type_ctx);
        }
        self.stream.emit_byte(opcode);
        self.stream.emit_i16(frame_offset);
    }

    /// Emit a typed local-variable store. Mirror of [`Self::emit_var_load`] using
    /// [`RT_STORE_BY_CTX`]. The caller must have emitted the value to store first.
    pub fn emit_var_store(&mut self, type_ctx: usize, frame_offset: i16) {
        // Byte (ctx 7): 2-byte escape store opcode (`fc f0`) via value-emitter
        // store index 0x1f0 (RT_OPCODE_BYTE[0x1f0] = 0xfc → escape).
        if type_ctx == 7 {
            self.emit_value2(0x1f0);
            self.stream.emit_i16(frame_offset);
            return;
        }
        // String (ctx 8): the refcounted BSTR assign store (0x43) via index 0x201.
        if type_ctx == 8 {
            self.emit_value2(0x201);
            self.stream.emit_i16(frame_offset);
            return;
        }
        // String move-store (ctx 9): store a freshly-produced string temp (e.g. a
        // concat result) without addref — opcode 0x31 via index 0x1f7.
        if type_ctx == 9 {
            self.emit_value2(0x1f7);
            self.stream.emit_i16(frame_offset);
            return;
        }
        let opcode = RT_STORE_BY_CTX.get(type_ctx).copied().unwrap_or(0);
        if opcode == 0 {
            unimplemented!("no store opcode for type context {}", type_ctx);
        }
        self.stream.emit_byte(opcode);
        self.stream.emit_i16(frame_offset);
    }

    /// Emit a ByRef parameter load at `frame_offset`. The ByRef load opcode is
    /// `RT_LOAD_BY_CTX[type_ctx] + 0x14`; the offset is positive (parameters sit
    /// above the frame pointer).
    pub fn emit_byref_load(&mut self, type_ctx: usize, frame_offset: i16) {
        let base = RT_LOAD_BY_CTX.get(type_ctx).copied().unwrap_or(0);
        if base == 0 {
            unimplemented!("no load opcode for type context {}", type_ctx);
        }
        self.stream.emit_byte(base + 0x14);
        self.stream.emit_i16(frame_offset);
    }

    /// Emit a ByRef parameter store at `frame_offset`. The ByRef store opcode is
    /// `RT_STORE_BY_CTX[type_ctx] + 0x14`.
    pub fn emit_byref_store(&mut self, type_ctx: usize, frame_offset: i16) {
        let base = RT_STORE_BY_CTX.get(type_ctx).copied().unwrap_or(0);
        if base == 0 {
            unimplemented!("no store opcode for type context {}", type_ctx);
        }
        self.stream.emit_byte(base + 0x14);
        self.stream.emit_i16(frame_offset);
    }

    /// Emit a module-level global variable load. The opcode is
    /// `RT_LOAD_BY_CTX[type_ctx] + 0x28`; the 4-byte operand encodes the module
    /// descriptor in bytes 0–1 and the field offset (byte offset within the
    /// module's global data block) in bytes 2–3.
    pub fn emit_global_load(
        &mut self,
        type_ctx: usize,
        module_desc: u16,
        field_offset: u16,
    ) {
        let base = RT_LOAD_BY_CTX.get(type_ctx).copied().unwrap_or(0);
        if base == 0 {
            unimplemented!("no load opcode for type context {}", type_ctx);
        }
        self.stream.emit_byte(base + 0x28);
        self.stream.emit_word(module_desc);
        self.stream.emit_word(field_offset);
    }

    // ── Store-opcode selection for `=` assignment ────────────────────────────

    /// Store-opcode base for a destination type tag: the entry of
    /// [`RT_ASSIGN_STORE_OPCODE`] at the destination's type-offset class.
    /// Type tags whose class falls outside the store-opcode table are not valid
    /// assignment destinations on this path.
    fn assign_store_base(dest_tag: i32) -> i32 {
        use crate::tables::RT_ASSIGN_STORE_OPCODE;
        let class = RT_TYPE_OFFSET[dest_tag as usize] as usize;
        if class >= RT_ASSIGN_STORE_OPCODE.len() {
            unimplemented!(
                "assignment store for type-offset class {class}: outside the \
                 store-opcode table; Phase 4"
            );
        }
        RT_ASSIGN_STORE_OPCODE[class]
    }

    /// Source-class adjustment added to the store base: the source's type-offset
    /// class, with `10 -> 4` and `9 -> 1` applied.
    fn assign_source_adjust(src_tag: i32) -> i32 {
        match RT_TYPE_OFFSET[src_tag as usize] {
            10 => 4,
            9 => 1,
            c => c,
        }
    }

    /// Emit the store opcode for a simple `=` assignment after the value has
    /// already been pushed. `n` is the assignment node; the source is `word[4]`.
    ///
    /// The general store opcode is `assign_store_base(dest) + assign_source_adjust(src)`,
    /// with direct opcodes for specific Variant / Currency / Boolean / object
    /// type pairs.
    /// Dispatch a synthetic operand-coercion node (opcode 0x78): emit the child
    /// operand, then the conversion opcode that widens it to the node's (target)
    /// type. The conversion opcode index is
    /// `assign_store_base(target) + assign_source_adjust(src)` — the same
    /// store/coerce opcode family the `=` store uses (RT_ASSIGN_STORE_OPCODE
    /// indexed by RT_TYPE_OFFSET[target], plus the source-class adjust). E.g.
    /// Integer→Long → 0x11c+1 = 0x11d (byte 0xe7); Long→Double → 0x12c+2 = 0x12e
    /// (byte 0xec).
    fn emit_coerce_node(&mut self, n: &RawNode) {
        self.emit_expr(NodeRef(n.w[4]), 2);
        let target_tag = n.type_tag();
        let src_tag = self.arena.get(NodeRef(n.w[4])).type_tag();
        // A Date destination uses dedicated conversion opcodes rather than the
        // base+adjust store family (a Date carries an OLE serial with its own
        // range/validity conversion). A Single source has already been widened to
        // the common float representation by its load, so it converts as Double.
        if target_tag == 0xc {
            match src_tag {
                0xa | 0xb => {
                    self.emit_value2(0x147);
                    return;
                }
                0xf => {
                    self.emit_value2(0x14f);
                    return;
                }
                0x10 => {
                    self.emit_value2(0x3c9);
                    return;
                }
                _ => {}
            }
        }
        let opcode = Self::assign_store_base(target_tag) + Self::assign_source_adjust(src_tag);
        self.emit_value2(opcode as usize);
    }

    fn emit_assign_op(&mut self, n: &RawNode) {
        let source = *self.arena.get(NodeRef(n.w[4]));
        let dest_hi = n.w[0] & 0xffff_0000;
        let dest_tag = (n.w[0] as i32) >> 16;
        let src_tag = (source.w[0] as i32) >> 16;

        // Object destination: a sized store. Sources whose type tag is in
        // [3, 0x17] go through a per-source-type sub-dispatch that emits size
        // operand words (needs the type-descriptor model); any other source
        // uses the store table with a trailing size operand.
        if dest_hi == 0xf0000 {
            let size = self.emit_get_type_size3(n.w[6]);
            if ((src_tag - 3) as u32) < 0x15 {
                unimplemented!(
                    "sized object/UDT store (per-source-type sub-dispatch emitting \
                     size operand words); needs the type-descriptor model; Phase 4"
                );
            }
            let opcode = Self::assign_source_adjust(src_tag) + Self::assign_store_base(dest_tag);
            self.emit_opcode2(opcode as usize, size as u16);
            return;
        }

        // Currency destination: direct opcodes for specific source kinds.
        if dest_hi == 0xc0000 {
            match src_tag {
                0xb => { self.emit_value2(0x147); return; }
                0xf => { self.emit_value2(0x14f); return; }
                0x10 => { self.emit_value2(0x3c9); return; }
                _ => {}
            }
        }

        // Both sides Variant / ByRef-Variant / Currency: guarded table store.
        if matches!(dest_tag, 10 | 0xb | 0xc) && matches!(src_tag, 10 | 0xb | 0xc) {
            // Flag-byte bit 0x80 clear → handled elsewhere (no-op here).
            if (n.w[1] >> 8) & 0x80 == 0 {
                return;
            }
            let base = Self::assign_store_base(dest_tag);
            if RT_TYPE_OFFSET[src_tag as usize] == 10 {
                self.emit_value2((base + 4) as usize);
                return;
            }
            self.emit_value2((Self::assign_source_adjust(src_tag) + base) as usize);
            return;
        }

        // Otherwise inspect the source's type region for special stores.
        let src_hi = source.w[0] & 0xffff_0000;
        if src_hi == 0xc0000 {
            // Currency source into a non-Currency destination.
            match dest_tag {
                0xf => { self.emit_value2(0x2fb); return; }
                0x10 => { self.emit_value2(0x3c8); return; }
                _ => {}
            }
        }
        if src_hi == 0x30000 {
            match dest_tag {
                5 => { self.emit_value2(0x138); return; }
                6 => { return; }
                0x10 => { self.emit_value2(0x3c7); return; }
                _ => {}
            }
        }
        if src_hi == 0x110000 {
            // Fixed-length string source: needs the type-length lookup.
            unimplemented!(
                "fixed-length string source store: needs the type-length lookup; Phase 4"
            );
        }
        if src_hi == 0xf0000 {
            // Object source into a Variant / TypeOf destination.
            if dest_hi == 0x140000 {
                unimplemented!(
                    "object source into a Variant target: needs the object-reference \
                     resolution path; Phase 4"
                );
            }
            if dest_hi == 0x120000 {
                unimplemented!(
                    "object source into a TypeOf target: needs the object-reference \
                     resolution path; Phase 4"
                );
            }
        }

        // Generic store: base from destination, adjustment from source.
        let base = Self::assign_store_base(dest_tag);
        if RT_TYPE_OFFSET[src_tag as usize] == 10 {
            self.emit_value2((base + 4) as usize);
            return;
        }
        self.emit_value2((Self::assign_source_adjust(src_tag) + base) as usize);
    }

    /// Emit a module-level global variable store. The opcode is
    /// `RT_STORE_BY_CTX[type_ctx] + 0x28`.
    pub fn emit_global_store(
        &mut self,
        type_ctx: usize,
        module_desc: u16,
        field_offset: u16,
    ) {
        let base = RT_STORE_BY_CTX.get(type_ctx).copied().unwrap_or(0);
        if base == 0 {
            unimplemented!("no store opcode for type context {}", type_ctx);
        }
        self.stream.emit_byte(base + 0x28);
        self.stream.emit_word(module_desc);
        self.stream.emit_word(field_offset);
    }

    // ── Call sites ───────────────────────────────────────────────────────────

    /// Emit a bare 2-byte little-endian value (no opcode byte).
    fn emit_word2(&mut self, value: u16) {
        self.stream.emit_word(value);
    }

    /// Coercion type-code map: `10 -> 4`, `9 -> 1`, everything else unchanged.
    fn map_type_code3(t: i32) -> i32 {
        match t {
            10 => 4,
            9 => 1,
            x => x,
        }
    }

    /// The behaviour-flag byte (`+0x1d`) of the dispatch record selected for a
    /// call of the given convention kind / by-reference mode.
    fn call_record_flag_1d(kind: i32, byref: i32) -> u8 {
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

        let mut call_op = call_type_code(kind, byref == 1, flags & 0x2000 != 0) as i32;
        let mut emit_mode = 0i32;

        if flags & 0x8000 == 0 {
            if flags & 0x800 != 0 {
                unimplemented!(
                    "call type-expression argument (flag 0x800): needs the type-pool \
                     expression path; Phase 5"
                );
            }
            let region = desc.node_word0 & 0xffff_0000;
            if region == 0x20000 {
                if flags & 0x200 != 0 && byref == 1 {
                    unimplemented!(
                        "call coercion sequence (0x20000 + flag 0x200, ByRef): needs \
                         the coercion path; Phase 5"
                    );
                }
                call_op += 0xc;
            } else if region == 0x120000 {
                call_op |= Self::map_type_code3(2);
            } else {
                let type_tag = (desc.node_word0 as i32) >> 16;
                if matches!(type_tag, 10 | 0xb | 0xc) && rec_1d & 8 != 0 {
                    let mut v = RT_TYPE_OFFSET[type_tag as usize];
                    if v == 10 {
                        v = 4;
                    } else if v == 9 {
                        v = 1;
                    }
                    let m = Self::map_type_code3(v);
                    call_op |= m;
                    if m == 3 || m == 4 {
                        call_op += 6;
                    }
                } else {
                    let m = Self::map_type_code3(RT_TYPE_OFFSET[type_tag as usize]);
                    call_op |= m;
                }
            }
        } else if flags & 0x1000 == 0 || call_op != 800 {
            call_op = map_call_type_code(call_op as u16) as i32;
            emit_mode = 1;
        } else {
            call_op = 0x41e;
            emit_mode = 1;
        }

        // Sized prefix (dispatch record requests a size operand).
        if rec_1d & 0x10 != 0 {
            self.emit_opcode2(0x265, desc.size);
        }
        // Argument list.
        if desc.arg_list.0 != 0 {
            if byte5 & 0x40 == 0 {
                self.process_linked_list(desc.arg_list, 3);
            } else {
                self.traverse_node_tree(desc.arg_list, 3);
            }
        }
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
        match emit_mode {
            // Common path: the finalize step emits the member-id word, then the
            // per-type tail.
            0 => self.finalize_call(desc, rec_1d, desc.member_id, context),
            // Dispatch path: the member-id word is emitted here; the finalize
            // step's trailing word is a type-pool lookup of node[9], which the
            // pool subsystem (not yet built) must supply.
            1 => {
                self.emit_word2(desc.member_id);
                unimplemented!(
                    "call dispatch finalize word (emit mode 1): the trailing word \
                     is a type-pool lookup of node[9]; needs the type pool; Phase 5"
                );
            }
            _ => unimplemented!(
                "call result-descriptor path (emit mode 2): needs the struct-size / \
                 member-type model; Phase 5"
            ),
        }
        0
    }

    /// The call-site finalize step: emit the caller-supplied `trailing` word,
    /// then — driven by the call node's type region — either re-enter the
    /// emitter on a synthetic dispatch-type node (gated; needs the type-pool
    /// allocator) or emit the per-type validation opcode. The `0xf0000` /
    /// dispatch-record `0x08` cases validate; everything else terminates the
    /// call site after the trailing word.
    fn finalize_call(&mut self, desc: &CallDescriptor, rec_1d: u8, trailing: u16, context: u32) {
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

#[cfg(test)]
#[path = "tests/emit_tests.rs"]
mod tests;
