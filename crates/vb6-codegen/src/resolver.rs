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
use crate::tables::{RT_RESOLVER_CLASS_FLAG, RT_RESOLVER_TYPE_MAP};

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
/// `+10` operand carries the type size. The `0x2000`-flag path (a by-value slot
/// with an out-of-line size) marks `word6` bit 0 and stores the secondary size
/// in `word8`. The type-library variant (`0x4000` flag with a `0x170000`-region
/// node) needs the type-library attribute path and is gated.
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
        operand: size as u16,
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
/// * `value_class` — the operand's value class (`0`/`1`/`2`), as derived from the
///   emitted operand (the derivation reads the live emit buffer and the slot
///   table through registers the decompile dropped, so it is supplied here rather
///   than recomputed).
/// * `member_byte1` — byte `+1` of the resolved member record.
/// * `op_byte1` — the second byte of the operand's p-code opcode.
///
/// Returned signed (the source casts the table byte through `char`).
pub fn resolver_type_category(value_class: i32, member_byte1: u8, op_byte1: u8) -> i32 {
    let idx = (value_class + (member_byte1 & 7) as i32 * 3) * 2 + (op_byte1 & 7) as i32;
    RT_RESOLVER_TYPE_MAP[idx as usize] as i8 as i32
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

#[cfg(test)]
#[path = "tests/resolver_tests.rs"]
mod tests;
