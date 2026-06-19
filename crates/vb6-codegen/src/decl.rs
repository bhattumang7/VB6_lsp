//! Declaration compiler — building the compiled member record ("bag") for a
//! module declaration, the record the resolver later reads.
//!
//! This ports `EbBuildDeclaration`'s scalar-member path: allocate the bag, write
//! its fields from the declaration descriptor, and attach the inline type node
//! produced by `EbBuildTypeNode2`'s base-type path. The parameter/slot loop, the
//! `0x19` (has-parameters) form, the record-type (`type_flags == 8`) check, the
//! symbol-list threading, and the COM `ITypeInfo` tail are gated.

use crate::heap::HeapContext;
use crate::typenode::build_inline_type_node;

/// The declaration descriptor `EbBuildDeclaration` reads (its `piContext`).
/// Only the fields the scalar path consumes are modelled; field names follow
/// the descriptor's dword/short slots.
#[derive(Clone, Copy, Debug, Default)]
pub struct DeclContext {
    /// `ctx[0]` — the declaration's signature word (written to record `+8`).
    pub sig: i32,
    /// `ctx[3]` — the kind discriminator: `<2` ⇒ method bag, `2..=4` ⇒ interface
    /// bag.
    pub kind_disc: i32,
    /// `ctx[4]` — the type mode/flags (size class for `EbSetTypeMode2`).
    pub type_flags: i32,
    /// `ctx[5]` — low 3 bits seed record `+0x3a`.
    pub field5: i32,
    /// `ctx[6]` — slot (parameter/field) count.
    pub slot_count: i16,
    /// `ctx[7]` — the member-dispatch id (written to a method bag's `+0x30`).
    pub member_id: i16,
    /// `ctx[9]` — `0x19` marks the has-parameters form; `0x18` suppresses the
    /// record `+0` `0x20` bit.
    pub flag9: i16,
    /// low 16 bits of `ctx[0xc]` — declaration flag bits.
    pub flags_c: u16,
    /// `*(short*)(ctx + 0x1a)` — optional-parameter / default marker.
    pub field_1a: i16,
    /// The return/element type word (`puTypeCode[1]`) fed to `EbBuildTypeNode2`.
    pub return_type_word: u16,
}

/// `EbBuildDeclaration`'s "bad declaration" status (`0x80028ca1`).
const EB_BAD_DECL: i32 = 0x80028ca1u32 as i32;

/// Port of `EbSetTypeMode2`: pack the size-class nibble for `mode`
/// (`1→1, 2→2, 4→4, 8→8`, else `0`) into the high nibble of record byte `+1`.
fn set_type_mode2(heap: &mut HeapContext, rec: u32, mode: i32) {
    let o = (rec + 1) as usize;
    let nibble = (((((mode == 4) as u8 | ((mode == 8) as u8) << 1) << 1 | (mode == 2) as u8) << 1)
        | (mode == 1) as u8)
        << 4;
    heap.mem[o] = nibble | (heap.mem[o] & 0xf);
}

fn put_u16(heap: &mut HeapContext, off: u32, v: u16) {
    let o = off as usize;
    heap.mem[o] = (v & 0xff) as u8;
    heap.mem[o + 1] = (v >> 8) as u8;
}

fn get_u16(heap: &HeapContext, off: u32) -> u16 {
    let o = off as usize;
    u16::from_le_bytes([heap.mem[o], heap.mem[o + 1]])
}

/// Port of `EbBuildDeclaration` for a scalar member (no parameters/slots).
///
/// Allocates the member record, writes its fields from `ctx`, and stores the
/// inline type node for `ctx.return_type_word`. Returns the record's offset in
/// the heap.
///
/// Gated (each `unimplemented!` rather than a guessed byte): a kind
/// discriminator `>= 5` (no bag is allocated in the source), the `0x19`
/// has-parameters form, a non-zero `slot_count` (the parameter/field loop), and
/// the `type_flags == 8` record-type check. The symbol-list insertion (record
/// `+4` / context list head) and the COM `ITypeInfo` tail are the caller's
/// concern.
pub fn build_declaration_scalar(heap: &mut HeapContext, ctx: &DeclContext) -> Result<u32, i32> {
    if ctx.kind_disc < 0 {
        return Err(EB_BAD_DECL);
    }

    let rec = if ctx.kind_disc < 2 {
        let rec = heap.allocate_method_bag()?;
        // record +1 low 3 bits = 1; +0 bit 0x40 for kind 1; +0x30 = member id.
        let o1 = (rec + 1) as usize;
        heap.mem[o1] = (heap.mem[o1] & 0xf8) ^ 1;
        if ctx.kind_disc == 1 {
            heap.mem[rec as usize] |= 0x40;
        }
        put_u16(heap, rec + 0x30, ctx.member_id as u16);
        rec
    } else if ctx.kind_disc < 5 {
        let rec = heap.allocate_interface_bag()?;
        let o1 = (rec + 1) as usize;
        let base = heap.mem[o1] & 0xf8;
        heap.mem[o1] = match ctx.kind_disc {
            2 => base,
            3 => base ^ 2,
            _ => base ^ 3,
        };
        rec
    } else {
        unimplemented!(
            "EbBuildDeclaration kind discriminator >= 5: the source allocates no \
             bag on this path; Phase 6"
        );
    };

    // ── Common field writes ──────────────────────────────────────────────────
    let sig = ctx.sig;
    // +0x18 &= 0xbfff
    put_u16(heap, rec + 0x18, get_u16(heap, rec + 0x18) & 0xbfff);
    // +8 = sig
    let o8 = (rec + 8) as usize;
    heap.mem[o8..o8 + 4].copy_from_slice(&sig.to_le_bytes());
    // +0x18 = (+0x18 & 0xdfff) | ((sig < 1) << 13)
    let bit13 = if sig < 1 { 1u16 << 13 } else { 0 };
    put_u16(heap, rec + 0x18, (get_u16(heap, rec + 0x18) & 0xdfff) | bit13);
    // EbSetTypeMode2(type_flags) → +1 high nibble
    set_type_mode2(heap, rec, ctx.type_flags);
    // +2 |= 0xfffe
    put_u16(heap, rec + 2, get_u16(heap, rec + 2) | 0xfffe);

    let flags_c = ctx.flags_c;
    if flags_c & 1 != 0 {
        let o1 = (rec + 1) as usize;
        heap.mem[o1] |= 8;
    }
    if flags_c & 0x80 != 0 {
        let o = (rec + 0x3c) as usize;
        heap.mem[o] |= 0x40;
    }
    // +0x3b = (+0x3b & 0xbf) | ((flags_c >> 15) << 6)
    {
        let o = (rec + 0x3b) as usize;
        heap.mem[o] = (heap.mem[o] & 0xbf) | (((flags_c >> 15) as u8) << 6);
    }
    // +0x3a low 3 bits = field5
    {
        let o = (rec + 0x3a) as usize;
        heap.mem[o] = ((ctx.field5 as u8 ^ heap.mem[o]) & 7) ^ heap.mem[o];
    }

    // ── Slot-count / optional-marker gates ───────────────────────────────────
    let sc = ctx.slot_count;
    if !(0..0x3d).contains(&sc) {
        return Err(EB_BAD_DECL);
    }
    let f1a = ctx.field_1a;
    if !(-1..0x3d).contains(&f1a) {
        return Err(EB_BAD_DECL);
    }
    // +0x3c low 6 bits = f1a (or 0x3f when f1a == -1)
    {
        let o = (rec + 0x3c) as usize;
        let b = heap.mem[o];
        heap.mem[o] = if f1a == -1 {
            (b & 0xc0) ^ 0x3f
        } else {
            ((f1a as u8 ^ b) & 0x3f) ^ b
        };
    }
    // +0x3d &= 0xf9
    {
        let o = (rec + 0x3d) as usize;
        heap.mem[o] &= 0xf9;
    }

    if ctx.flag9 == 0x19 {
        unimplemented!("EbBuildDeclaration has-parameters form (ctx[9] == 0x19); Phase 6");
    }

    // ── Return/element type node (EbBuildTypeNode2 base-type path) ────────────
    let node = build_inline_type_node(ctx.return_type_word);
    let o2c = (rec + 0x2c) as usize;
    heap.mem[o2c..o2c + 4].copy_from_slice(&node.to_le_bytes());
    // +0x3a bit 3 = inline flag (the type node reports inline ⇒ 1)
    {
        let o = (rec + 0x3a) as usize;
        heap.mem[o] = (1 << 3) | (heap.mem[o] & 0xf7);
    }
    // +0 |= 0x20 (clearing 0x10) unless ctx[9] == 0x18
    if ctx.flag9 != 0x18 {
        let o = rec as usize;
        heap.mem[o] = (heap.mem[o] & 0xef) | 0x20;
    }

    // ── Slot loop prologue (LAB_0fac43d7) ────────────────────────────────────
    // +0x3b low 6 bits = slot count
    {
        let o = (rec + 0x3b) as usize;
        heap.mem[o] = ((sc as u8 ^ heap.mem[o]) & 0x3f) ^ heap.mem[o];
    }
    if sc > 0 {
        unimplemented!("EbBuildDeclaration parameter/field slot loop (slot_count > 0); Phase 6");
    }
    // Post-loop: +0x3a bit 4 = local_3c (0 on the scalar path).
    {
        let o = (rec + 0x3a) as usize;
        heap.mem[o] &= 0xef;
    }
    if ctx.type_flags == 8 {
        unimplemented!("EbBuildDeclaration record-type check (type_flags == 8); Phase 6");
    }

    Ok(rec)
}

#[cfg(test)]
#[path = "tests/decl_tests.rs"]
mod tests;
