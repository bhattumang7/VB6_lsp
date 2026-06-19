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

#[cfg(test)]
#[path = "tests/resolver_tests.rs"]
mod tests;
