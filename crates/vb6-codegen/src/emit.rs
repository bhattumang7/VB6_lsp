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
use crate::tables::{RT_BINOP_BASE, RT_CALL_TYPECODE, RT_DISPATCH_FLAG, RT_LOAD_BY_CTX, RT_OPCODE_BYTE, RT_STORE_BY_CTX, RT_TYPE_OFFSET};

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
///   ByRef operation path).
/// * `word8` — extra trailing word, emitted only on the nested-type-expression
///   path (reached via the Variant-with-type chain).
#[derive(Clone, Copy, Debug, Default, PartialEq, Eq)]
pub struct RefDescriptor {
    pub kind: i32,
    pub operand: u16,
    pub word6: u16,
    pub word8: u16,
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
            // case 5: typed literal / typed store. Every branch first resolves a
            // type size, which needs the type-descriptor model and type/string
            // pool.
            5 => unimplemented!(
                "typed literal / store: needs the type-descriptor model and the \
                 type/string pool; Phase 4"
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
            // cases 0xc / 0xd: expression-code sub-emission.
            0xc => unimplemented!("expression-code emission (form 0); Phase 5"),
            0xd => unimplemented!("expression-code emission (form 1); Phase 5"),
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
                3 => unimplemented!(
                    "sized name reference: needs the type-descriptor model; Phase 4"
                ),
                4 => unimplemented!(
                    "in-place name reference: needs node-flag mutation; Phase 3"
                ),
                _ => return 0,
            },
            // case 0x10: emit child, then opcode 0x135.
            0x10 => {
                self.emit_expr(n.lhs(), 1);
                0x135
            }
            // case 0x11: type-code emission.
            0x11 => unimplemented!("type-code emission; Phase 4"),
            // case 0x12: member dereference.
            0x12 => {
                let child = n.lhs();
                let cn = *self.arena.get(child);
                if node_hi == 0x160000 && (cn.w[0] & 0xffff) == 0x67 {
                    unimplemented!(
                        "typed member dereference: needs the type/string pool; Phase 4"
                    );
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
            0x2c => unimplemented!("assignment statement; Phase 5"),
            // case 0x2d: typed assignment / error recovery (type error report +
            // node mutation + assignment statement).
            0x2d => unimplemented!("typed assignment / error recovery; Phase 5"),
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
            0x32 => unimplemented!("type coercion (0x40d/0x40e); Phase 4"),
            // case 0x33: traverse the child list, emitting each element.
            0x33 => {
                self.traverse_node_tree(node, 1);
                return 0;
            }
            // case 0x34: type coercion (0x40f).
            0x34 => unimplemented!("type coercion (0x40f); Phase 4"),
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
            // case 0x38: nested member access size: needs the type-descriptor
            // model.
            0x38 => unimplemented!(
                "nested member-access size: needs the type-descriptor model; Phase 4"
            ),
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
            // case 0x3e: argument list + member value: needs the type/string pool.
            0x3e => unimplemented!(
                "argument list + member value: needs the type/string pool; Phase 4"
            ),
            // case 0x3f: argument list + member size: needs the type-descriptor
            // model.
            0x3f => unimplemented!(
                "argument list + member size: needs the type-descriptor model; Phase 4"
            ),
            // case 0x41: argument-list emission with per-argument type sizes.
            0x41 => unimplemented!(
                "argument-list emission: needs per-argument type sizes; Phase 4"
            ),
            // cases 0x42 / 0x43: dispatch-type resolution.
            0x42 => unimplemented!("dispatch-type resolution (form 0); Phase 5"),
            0x43 => unimplemented!("dispatch-type resolution (form 1); Phase 5"),
            // cases 0x44..=0x47: type conversion + operand dispatch.
            0x44 | 0x45 | 0x46 | 0x47 => unimplemented!(
                "type conversion + operand dispatch; Phase 5"
            ),
            // cases 0x48..=0x4b: traverse list, emit a fixed opcode; results that
            // are not the 0x20000 form additionally dispatch the operand.
            0x48 | 0x49 | 0x4a | 0x4b => {
                self.traverse_node_tree(NodeRef(n.w[5]), 1);
                let value = match op {
                    0x48 => 0x158,
                    0x49 => 0x15a,
                    0x4a => 0x159,
                    _ => 0x15b,
                };
                self.emit_value2(value);
                if node_hi == 0x20000 {
                    return 0;
                }
                unimplemented!("operand dispatch (non-0x20000 result); Phase 5");
            }
            // case 0x4c: type conversion (0x35d).
            0x4c => unimplemented!("type conversion (0x35d); Phase 5"),
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
            // cases 0x51 / 0x52: operator classification.
            0x51 | 0x52 => unimplemented!("operator classification; Phase 5"),
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
            0x54 => unimplemented!("type coercion (0x410); Phase 4"),
            // case 0x55: type conversion (0x35e).
            0x55 => unimplemented!("type conversion (0x35e); Phase 5"),
            // case 0x56: type conversion (flag 0x4000 clear) or traverse +
            // flag-selected opcode (set).
            0x56 => {
                if n.w[1] & 0x4000 == 0 {
                    unimplemented!("type conversion; Phase 5");
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
            0x5a => unimplemented!("complex binary operation; Phase 5"),
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
            0x60 => unimplemented!("member-reference coercion; Phase 4"),
            // case 0x61: call / argument machinery — needs the symbol table and
            // dispatch tables.
            0x61 => unimplemented!(
                "call / argument emission: needs the symbol table and dispatch \
                 tables; Phase 5"
            ),
            // case 0x63: member-reference value: needs the type/string pool.
            0x63 => unimplemented!(
                "member-reference value: needs the type/string pool; Phase 4"
            ),
            // case 0x65: forward to child.
            0x65 => return self.emit_expr(n.lhs(), context),
            // case 0x66: member-reference value: needs the type/string pool.
            0x66 => unimplemented!(
                "member-reference value (0x2f4): needs the type/string pool; Phase 4"
            ),
            // case 0x67: member-reference value: needs the type/string pool.
            0x67 => unimplemented!(
                "member-reference value (0x2f5): needs the type/string pool; Phase 4"
            ),
            // case 0x68: emit child, then (context 6) opcode 0x29f, else a
            // type-class-selected opcode.
            0x68 => {
                let child = n.lhs();
                self.emit_expr(child, 1);
                if context != 6 {
                    let cn = *self.arena.get(child);
                    let mut needs_word = false;
                    let mut value = context;
                    if node_hi == 0x160000 {
                        if (cn.w[0] & 0xffff_0000) == 0xf0000 {
                            unimplemented!(
                                "member reference (object child): needs the symbol \
                                 table; Phase 5"
                            );
                        } else if (cn.w[0] & 0xffff_0000) == 0x160000 {
                            needs_word = true;
                            value = 0x2f2;
                        }
                    }
                    self.emit_value2(value as usize);
                    if !needs_word {
                        return 0;
                    }
                    unimplemented!(
                        "member reference (trailing word): needs the type/string \
                         pool; Phase 4"
                    );
                }
                0x29f
            }
            // case 0x69: binary-operation setup + reference emission.
            0x69 => unimplemented!(
                "binary-operation setup + reference emission; Phase 5"
            ),
            // case 0x6a: member-call instruction: needs the type-descriptor model.
            0x6a => unimplemented!(
                "member-call instruction: needs the type-descriptor model; Phase 5"
            ),
            // cases 0x6b..=0x6e: instruction emitter.
            0x6b | 0x6c | 0x6d | 0x6e => unimplemented!("instruction emitter; Phase 5"),
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

    // ── Resolved-reference emission ──────────────────────────────────────────

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
            // kinds 8/9/0xb: member / typed reference — needs the module symbol
            // table and compiled type descriptors.
            8 | 9 | 0xb => unimplemented!(
                "reference kind {} (member/typed): needs the module symbol table \
                 and compiled type descriptors",
                desc.kind
            ),
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
                    // A small conversion table indexed by two inverted flag bits
                    // and the type offset; that table is not yet available.
                    unimplemented!(
                        "store conversion (0x8000 path): conversion table not yet \
                         available; Phase 4"
                    );
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
