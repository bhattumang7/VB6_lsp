//! Reference resolution: turning a name/member expression node into the
//! value-emitter descriptor ([`RefDescriptor`]) that drives the typed
//! load/store opcode.
//!
//! This is the consumer half of the resolver — the logic that, given a resolved
//! binding (its by-reference mode and storage class) and the expression node,
//! builds the 4-word descriptor. The binding facts themselves come from the
//! module symbol records (see [`crate::sym_record`]); the values in those records
//! are produced by VB6's own member-numbering machinery, ported separately.

use crate::emit::RefDescriptor;
use crate::node::{NodeArena, NodeRef};
use crate::tables::{
    RT_CALL_CONV_RECORDS, RT_CALL_KIND_CLASS, RT_CALL_SPECIAL_RECORD, RT_RESOLVER_CLASS_FLAG,
    RT_RESOLVER_TYPE_MAP, RT_TYPE_OFFSET,
};

/// Resolve the byte size of a UDT / object type from its type descriptor — the
/// shared reader used across the emit and resolve paths (port of
/// `EbGetTypeSize3`).
///
/// `type_desc` is an arena reference. A fixed-size descriptor (`word[0] == 4`)
/// carries its byte size in the low half of `word[4]`; any other kind — or a
/// null reference — has no fixed size and yields the `0xffff_ffff` sentinel.
pub(crate) fn get_type_size3(arena: &NodeArena, type_desc: u32) -> u32 {
    if type_desc == 0 {
        return 0xffff_ffff;
    }
    let desc = arena.get(NodeRef(type_desc));
    if desc.w[0] == 4 {
        desc.w[4] & 0xffff
    } else {
        0xffff_ffff
    }
}

/// Build the value-emitter descriptor for a resolved reference (port of
/// `EbInitExprDescriptor`).
///
/// * `by_ref` — the binding is passed/stored by reference.
/// * `optional` — the reference is an optional-argument slot (suppresses the
///   `0x04` usage flag and the `0x160000`-region marker).
///
/// The descriptor kind encodes the storage class together with `by_ref`:
/// small/8-byte types use `2 - by_ref` (→ kind 2 by-value, kind 1 by-reference),
/// other types use `7 - by_ref` (→ kind 7 by-value, kind 6 by-reference). The
/// type size (via `EbGetTypeSize3`) only feeds that kind selection (the
/// small-vs-other / size-8 override) — it is not what reaches the emitted
/// bytecode. The value-emitter's kind-1/2 opcode takes the descriptor's
/// `+10` operand as its literal p-code operand word, and the runtime decode of
/// that opcode (confirmed independently against the real engine, and against
/// `e2e_two_sequential_long_assigns`'s oracle-verified bytes) is a **frame-
/// relative offset**, not a size. So `+10` here is the resolved reference's
/// frame offset — supplied by the front end via `expr`'s own `word[7]` (our own
/// convention: `lower.rs` computes the combined offset for a resolved
/// local/member reference and stores it there; there is no COM/type-library
/// storage to consult for a plain scalar). The `0x2000`-flag path (a by-value
/// slot with an out-of-line size) marks `word6` bit 0 and stores the secondary
/// size in `word8`. The type-library variant (`0x4000` flag with a `0x170000`-
/// region node) needs the type-library attribute path and is gated.
pub fn init_expr_descriptor(
    arena: &NodeArena,
    expr: NodeRef,
    by_ref: bool,
    optional: bool,
) -> RefDescriptor {
    let e = *arena.get(expr);
    let size = get_type_size3(arena, e.w[5]) as i16;
    let flags = e.w[1];
    let kind = if flags & 0x100 == 0 || size == 8 {
        2 - by_ref as i32
    } else {
        7 - by_ref as i32
    };
    let mut desc = RefDescriptor {
        kind,
        operand: e.w[7] as u16,
        ..RefDescriptor::default()
    };
    if flags & 0x2000 == 0 || flags & 1 != 0 {
        if flags & 0x4000 != 0 && (e.w[0] & 0xffff_0000) == 0x0017_0000 {
            unimplemented!(
                "type-library reference descriptor (EbSetExprTypeFlags): needs the \
                 type-library attribute path; Phase 6"
            );
        }
    } else {
        // By-value slot with an out-of-line size: mark word6 bit 0 and store the
        // secondary size from word[6].
        desc.word6 |= 1;
        desc.word8 = get_type_size3(arena, e.w[6]) as u16;
    }
    if !optional {
        desc.flags1 |= 4;
        if (e.w[0] & 0xffff_0000) == 0x0016_0000 {
            desc.flags1 |= 1;
        }
    }
    desc
}

/// The compiler context object the resolver consults while classifying a
/// binding. Only the fields the resolver reads are modelled.
///
/// * `kind` — the context discriminator (`*ctx`): 4/5/6 mark a member-access
///   context that carries a member byte offset.
/// * `member_offset` — `ctx[3]`: the member's byte offset into the symbol base.
#[derive(Clone, Copy, Debug, Default, PartialEq, Eq)]
pub struct CompileContext {
    pub kind: i32,
    pub member_offset: i32,
}

/// Port of `EbGetExprContext`: the member byte offset for a member-access
/// context (`kind` 4/5/6), or `-1` for any other context.
pub fn get_expr_context(ctx: &CompileContext) -> i32 {
    if ctx.kind == 4 || ctx.kind == 5 || ctx.kind == 6 {
        ctx.member_offset
    } else {
        -1
    }
}

/// Port of `EbGetCurrentExpression`: given the operand's p-code bytes, return the
/// byte offset of the "current expression" start the classifier inspects.
///
/// A `0x1d`-class opcode (low 6 bits `0x1d`) that carries an inline type word
/// whose following opcode (4 bytes on) is not a `0x25`-class marker is skipped
/// past — the expression start is 4 bytes in. Every other operand starts where it
/// is (offset 0).
pub fn current_expression_offset(pcode: &[u8]) -> usize {
    if pcode[0] & 0x3f == 0x1d && pcode[4] & 0x3f != 0x25 {
        4
    } else {
        0
    }
}

/// Port of `EbIsPcodeTerminator`: a p-code opcode whose low 6 bits are `0x1b` or
/// `0x1c` terminates an operand run.
pub fn is_pcode_terminator(opcode: u8) -> bool {
    let lo = opcode & 0x3f;
    lo == 0x1b || lo == 0x1c
}

/// Port of `EbExtractTypeInfo`: read the inline type word a p-code operand
/// carries. A `0x0d`-class opcode has the sentinel `0xfffe`; a `0x1d`-class
/// opcode carries the type word two bytes in (low bit cleared). Any other
/// opcode has no inline type word here (the reference emitter dereferences a
/// null pointer, i.e. this is never reached for those).
pub fn extract_type_info(pcode: &[u8]) -> u16 {
    let lo = pcode[0] & 0x3f;
    if lo == 0x0d {
        return 0xfffe;
    }
    if lo == 0x1d {
        return u16::from_le_bytes([pcode[2], pcode[3]]) & 0xfffe;
    }
    unreachable!("EbExtractTypeInfo on an opcode without an inline type word")
}

/// The class-flag a p-code opcode carries in `EbGetExpressionType2`'s first gate
/// (`DAT_0fab5b10[opcode & 0x3f]`). The classifier treats the operand as
/// carrying an inline type only when this flag is `0` and the opcode's low 6 bits
/// are neither `0x1e` nor `0x1f`; otherwise the value class is forced to `0`.
pub fn resolver_class_flag(opcode: u8) -> u8 {
    RT_RESOLVER_CLASS_FLAG[(opcode & 0x3f) as usize]
}

/// `EbGetExpressionType2`'s first gate: whether the just-emitted operand carries
/// an inline type the classifier should inspect (class-flag `0` and low 6 bits
/// not `0x1e`/`0x1f`). When this is false the value class collapses to `0`.
pub fn resolver_inspects_operand(opcode: u8) -> bool {
    let lo = opcode & 0x3f;
    resolver_class_flag(opcode) == 0 && lo != 0x1e && lo != 0x1f
}

/// The final type-category lookup of `EbGetExpressionType2`:
/// `(char)DAT_0fc10aa8[(value_class + (member_byte1 & 7) * 3) * 2 + (op_byte1 & 7)]`.
///
/// * `value_class` — the operand's value class (`0`/`1`/`2`), derived by
///   [`expression_type2`] from the operand's p-code.
/// * `member_byte1` — byte `+1` of the resolved member record.
/// * `op_byte1` — the second byte of the operand's p-code opcode.
///
/// Returned signed (the source casts the table byte through `char`).
pub fn resolver_type_category(value_class: i32, member_byte1: u8, op_byte1: u8) -> i32 {
    let idx = (value_class + (member_byte1 & 7) as i32 * 3) * 2 + (op_byte1 & 7) as i32;
    RT_RESOLVER_TYPE_MAP[idx as usize] as i8 as i32
}

/// Port of `EbGetMaskedValue`: read the member record's masked `+0xc` value, used
/// to locate where the record's operand p-code lives in the records heap. With
/// record byte `+0` bit `0x40` set the raw `+0xc` dword is returned; otherwise its
/// low bit is cleared (an all-but-low-bit value of `0xffff_fffe` is an internal
/// error in the source).
pub fn get_masked_value(record: &[u8]) -> u32 {
    let at_c = u32::from_le_bytes([record[0xc], record[0xd], record[0xe], record[0xf]]);
    if record[0] & 0x40 != 0 {
        at_c
    } else {
        let m = at_c & 0xffff_fffe;
        if m == 0xffff_fffe {
            unimplemented!("EbGetMaskedValue internal-error path (masked value 0xfffffffe)");
        }
        m
    }
}

/// Port of `EbResolveAttributePointer` (with `nIndex == 0`): resolve the byte
/// offset, within the records heap, of the operand p-code for the member record
/// at `record_off`. `None` is the null result (masked value `0xffff_ffff`).
///
/// With record byte `+0` bit `0x40` set the operand is inline at `record + 0xc`;
/// otherwise it lives at the masked `+0xc` offset from the heap base.
pub fn resolve_attribute_pointer(heap: &[u8], record_off: usize) -> Option<usize> {
    let masked = get_masked_value(&heap[record_off..]);
    if masked == 0xffff_ffff {
        return None;
    }
    if heap[record_off] & 0x40 != 0 {
        Some(record_off + 0xc)
    } else {
        Some(masked as usize)
    }
}

/// Port of `EbCheckMethodFlags`: a value class of `1` when the operand's byte `+1`
/// has bit `0x40` set and bit `0x20` clear, else `0`.
pub fn method_flags_class(pcode: &[u8]) -> i32 {
    let f = pcode[1];
    (f & 0x40 != 0 && f & 0x20 == 0) as i32
}

/// The result of [`expression_type2`].
#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub struct ExpressionType {
    /// `*pResult` — the type category written back for the caller (`local_30` in
    /// `EbResolveIdentRef`).
    pub category: i32,
    /// The function's own return value (`local_8`): `0` on every path reached
    /// here; a negative error / slot-resolution code only on the gated paths.
    pub code: i32,
}

/// Port of `EbGetExpressionType2`: classify the binding for the member record at
/// `member_off` in the records `heap`.
///
/// The classifier resolves the record's operand p-code
/// ([`resolve_attribute_pointer`] + [`current_expression_offset`]) and derives a
/// value class (`0`/`1`/`2`) from that operand's opcode, then indexes
/// [`resolver_type_category`] with the record's byte `+1` and the operand's second
/// byte. Two deeper sub-paths are faithfully gated: the `0x1a` operand-skip
/// reread, and the `0x1d` slot-type path (which reads the `symbol_base[0x6c]` slot
/// descriptor tables via `EbMapSlotType`).
pub fn expression_type2(heap: &[u8], member_off: usize) -> ExpressionType {
    let member_byte1 = heap[member_off + 1];
    let ptr0 = resolve_attribute_pointer(heap, member_off)
        .expect("EbGetExpressionType2: null operand pointer for member record");
    let ptr = ptr0 + current_expression_offset(&heap[ptr0..]);
    let op = heap[ptr];
    let op_byte1 = heap[ptr + 1];
    let value_class = if resolver_inspects_operand(op) {
        let lo = op & 0x3f;
        if lo == 0x1a {
            unimplemented!(
                "EbGetExpressionType2 0x1a operand-skip path: needs the \
                 terminator-skip reread; Phase 6"
            );
        }
        if lo == 0x1d {
            if op_byte1 & 0x40 == 0 {
                unimplemented!(
                    "EbGetExpressionType2 0x1d slot-type path: needs the \
                     symbol_base[0x6c] slot descriptor tables (EbMapSlotType); Phase 6"
                );
            }
            method_flags_class(&heap[ptr..])
        } else {
            2
        }
    } else {
        0
    };
    ExpressionType {
        category: resolver_type_category(value_class, member_byte1, op_byte1),
        code: 0,
    }
}

/// Port of the value-mapping tail of `EbMapSlotType`: collapse a resolved slot
/// type to a value class. `0 → 0`, `1 → 1`, `3/4/5/8 → 2`; `2`, `6`, `7`, and
/// `>=9` are type mismatches (no value class). (The `EbGetSlotType` lookup that
/// produces the slot type reads the member descriptor and is supplied by the
/// caller.)
pub fn map_slot_type_value(slot: i32) -> Option<i32> {
    if slot == 0 {
        return Some(0);
    }
    if slot == 1 {
        return Some(1);
    }
    if slot < 3 {
        return None; // slot == 2
    }
    if slot > 5 && slot != 8 {
        return None; // 6, 7, >=9
    }
    Some(2) // 3, 4, 5, 8
}

/// Read a little-endian dword from the records heap at `off`.
fn heap_dword(heap: &[u8], off: usize) -> u32 {
    u32::from_le_bytes([heap[off], heap[off + 1], heap[off + 2], heap[off + 3]])
}

/// Port of `EbResolveIdentRef`: resolve a `0x60` name-reference node into its
/// value-emitter descriptor.
///
/// The resolver classifies the binding ([`expression_type2`]), resolves the
/// member record's operand pointer ([`resolve_attribute_pointer`]), and dispatches
/// on the resulting type category to build the descriptor (mostly via
/// [`init_expr_descriptor`]) plus two trailing flag adjustments.
///
/// * `node` — the `0x60` reference node; its `word[5]` is the member-access
///   context node [`init_expr_descriptor`] reads for kind selection, and its
///   `word[7]` is the front-end-resolved frame offset [`init_expr_descriptor`]
///   copies into the descriptor operand.
/// * `heap` — the module records heap (`*symbol_base`).
/// * `member_off` — the member record's byte offset into `heap` (the
///   [`get_expr_context`] result).
/// * `ctx_flag_c` — byte `+0xc` of the compiler context (`in_ECX`); only bit `2`
///   is read, gating the attribute flag on the value cases.
///
/// Gated (each `unimplemented!` rather than a guessed byte): the method/object
/// binding path (record byte `+0` bit `0x80` with `+1 & 7 == 4`), category 4
/// (`EbResolveExprNode`), and categories `0xd`/`0xe`/`0xf` (the binding-emit tail
/// `EbFillBindingDesc` / `EbEmitBinaryOpCode`, which read the COM slot tables and
/// emit into the stream).
pub fn resolve_ident_ref(
    arena: &NodeArena,
    node: NodeRef,
    heap: &[u8],
    member_off: usize,
    ctx_flag_c: u8,
    binding: Option<(i32, i32)>,
) -> RefDescriptor {
    let n = *arena.get(node);
    let type_offset = RT_TYPE_OFFSET[n.type_tag() as usize];
    let rec0 = heap[member_off];
    let rec1 = heap[member_off + 1];

    // Method / object member binding: needs the COM method-binding subsystem.
    if rec0 & 0x80 != 0 && rec1 & 7 == 4 {
        unimplemented!(
            "EbResolveIdentRef method/object binding (record +0 bit 0x80, +1&7==4): \
             needs the method-binding / object-reference emit subsystem; Phase 6"
        );
    }

    let category = expression_type2(heap, member_off).category;
    // The operand pointer the value cases inspect (un-skipped, unlike the
    // classifier's own copy). Null only on a malformed record.
    let operand_off = resolve_attribute_pointer(heap, member_off);

    let mut desc = match category {
        1 | 2 | 3 => {
            let op = heap_dword(heap, operand_off.expect("null operand pointer"));
            let optional = if category == 3 || (op & 0x3f == 0x1b && op & 0x800 != 0) {
                false
            } else if n.w[0] & 0xffff_0000 == 0x0017_0000 {
                unimplemented!(
                    "EbResolveIdentRef type-library reference (region 0x170000, \
                     LoadTypeLibraryItem); Phase 6"
                );
            } else {
                true
            };
            let mut d = init_expr_descriptor(arena, node, true, optional);
            // Attribute flag: set bit 0x02 when the record is flagged and the
            // compiler context does not suppress it.
            if rec1 & 0x20 != 0 && ctx_flag_c & 2 == 0 {
                d.flags1 |= 2;
            }
            d
        }
        7 | 8 => {
            let op = heap_dword(heap, operand_off.expect("null operand pointer"));
            init_expr_descriptor(arena, node, false, op & 0x3f != 0x1b)
        }
        9 => init_expr_descriptor(arena, node, true, false),
        0xc => init_expr_descriptor(arena, node, false, false),
        4 => match binding {
            // The binder (EbResolveExprNode → EbGetTypeKind2/EbGetByRefFlag,
            // ported in `crate::binder`) supplies the resolved convention
            // kind/byref; the record selection + descriptor follow.
            Some((kind, byref)) => call_conv_descriptor(arena, node, type_offset, kind, byref),
            None => unimplemented!(
                "EbResolveIdentRef category 4: needs the binder-resolved (kind, \
                 byref) — call resolve_ident_ref with the EbResolveExprNode result"
            ),
        },
        0xd | 0xe | 0xf => unimplemented!(
            "EbResolveIdentRef categories 0xd/0xe/0xf (binding-emit tail: \
             EbFillBindingDesc / EbEmitBinaryOpCode); Phase 6"
        ),
        // Default: a zeroed descriptor (categories 5/6/0xa/0xb and out of range).
        _ => RefDescriptor::default(),
    };

    // Shared tail (`EbSetTypeFlag2`): a type-offset class of 0xe sets bit 0x08.
    if type_offset == 0xe {
        desc.flags1 |= 8;
    }
    desc
}

/// Port of `EbResolveIdentRef`'s category-4 descriptor selection (the
/// call-convention path).
///
/// Given the resolved callee's convention `kind` (`EbGetTypeKind2`) and
/// by-reference mode `byref` (`EbGetByRefFlag`), select the dispatch record — the
/// `ByRef`+kind-4 special record, else `RT_CALL_CONV_RECORDS[RT_CALL_KIND_CLASS
/// [kind] & 3]` — and read its word at the node's type-offset class
/// (`type_offset = RT_TYPE_OFFSET[type_tag]`). The descriptor's by-reference flag
/// is the complement of that word's bit `0x4000`; the reference is built optional
/// (`EbInitExprDescriptor(node, by_ref, 1)`).
///
/// The shared `EbSetTypeFlag2` tail (a `type_offset == 0xe` setting bit `0x08`)
/// is applied by [`resolve_ident_ref`], not here.
pub fn call_conv_descriptor(
    arena: &NodeArena,
    node: NodeRef,
    type_offset: i32,
    kind: i32,
    byref: i32,
) -> RefDescriptor {
    let record: &[u8] = if byref == 1 && kind == 4 {
        &RT_CALL_SPECIAL_RECORD
    } else {
        let class = (RT_CALL_KIND_CLASS[kind as usize] & 3) as usize;
        &RT_CALL_CONV_RECORDS[class * 0x1e..class * 0x1e + 0x1e]
    };
    let wi = (type_offset * 2) as usize;
    let word = u16::from_le_bytes([record[wi], record[wi + 1]]);
    let by_ref = word & 0x4000 == 0;
    init_expr_descriptor(arena, node, by_ref, true)
}

/// Port of the `EbResolveReference2` dispatcher: route a reference node to its
/// resolver by opcode. `0x60` name references resolve via [`resolve_ident_ref`];
/// `0x69` binary-operation setups (`EbSetupBinaryOperation`) are gated.
///
/// A `0x60` node carrying a member sub-expression (`word[4] != 0`) needs
/// `EbSimplifyMemberExpr` to produce the type descriptor first — also gated.
pub fn resolve_reference2(
    arena: &NodeArena,
    node: NodeRef,
    heap: &[u8],
    member_off: usize,
    ctx_flag_c: u8,
    binding: Option<(i32, i32)>,
) -> RefDescriptor {
    let n = *arena.get(node);
    match n.w[0] & 0xffff {
        0x60 => {
            if n.w[4] != 0 {
                unimplemented!(
                    "EbResolveReference2: 0x60 node with a member sub-expression \
                     (word[4] != 0) needs EbSimplifyMemberExpr; Phase 6"
                );
            }
            resolve_ident_ref(arena, node, heap, member_off, ctx_flag_c, binding)
        }
        0x69 => unimplemented!(
            "EbResolveReference2: 0x69 binary-operation setup \
             (EbSetupBinaryOperation); Phase 6"
        ),
        _ => RefDescriptor::default(),
    }
}

#[cfg(test)]
#[path = "tests/resolver_tests.rs"]
mod tests;
