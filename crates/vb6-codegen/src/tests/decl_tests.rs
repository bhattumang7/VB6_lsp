//! Tests for the declaration compiler ([`crate::decl`]).
//!
//! Expected record bytes are derived directly from `EbBuildDeclaration`'s
//! scalar-member write sequence, not from any emitted-p-code sample.

use crate::decl::{build_declaration_scalar, build_property_slot_scalar, DeclContext};
use crate::heap::{HeapContext, NIL};
use crate::resolver::expression_type2;

/// The type category the ported resolver classifier produces for a scalar
/// property-bag slot (record byte `+1` low3 = 3, inline Long operand at `+0xc`):
/// category 4 (the call-convention / `EbResolveExprNode` path). Locked as a
/// regression of the ported declaration → resolver interop.
const EXPECTED_LONG_SLOT_CATEGORY: i32 = 4;

/// A heap seeded with one big free block at offset 0 (in-place mode, 2-byte
/// align) so allocations run without the gated grow path.
fn seeded_heap(len: usize) -> HeapContext {
    let mut h = HeapContext {
        mem: vec![0u8; len],
        free_head: 0,
        flags: 1,
        buffer_flag: 0,
    };
    h.mem[4..6].copy_from_slice(&((len - 8) as u16).to_le_bytes());
    h
}

#[test]
fn scalar_method_member_record_bytes() {
    let mut h = seeded_heap(0x200);
    let ctx = DeclContext {
        sig: 0,
        kind_disc: 0, // method bag
        type_flags: 4, // size class 4
        field5: 0,
        slot_count: 0,
        member_id: 0x1234,
        flag9: 0,
        flags_c: 0,
        field_1a: -1,
        return_type_word: 8, // e.g. Long
    };

    let rec = build_declaration_scalar(&mut h, &ctx, &[]).unwrap();
    let r = rec as usize;
    let b = &h.mem[r..r + 0x40];

    // Expected, computed from the write sequence:
    let mut want = [0u8; 0x40];
    want[0x00] = 0x24; // method tag 4, then +0 |= 0x20
    want[0x01] = 0x41; // low3 = 1; high nibble = size-class 4
    want[0x02] = 0xfe; // +2 |= 0xfffe
    want[0x03] = 0xff;
    want[0x18] = 0x00; // +0x18 bit 0x2000 set (sig < 1)
    want[0x19] = 0x20;
    want[0x2c] = 0x08; // inline type node opcode = Long (8)
    want[0x30] = 0x34; // member id 0x1234
    want[0x31] = 0x12;
    want[0x3a] = 0x08; // bit3 = inline flag
    want[0x3c] = 0x3f; // field_1a == -1 → low6 = 0x3f

    assert_eq!(b, &want, "scalar method record mismatch");
}

#[test]
fn scalar_interface_member_kind3_low3_is_2() {
    let mut h = seeded_heap(0x200);
    let ctx = DeclContext {
        sig: 1, // sig >= 1 → +0x18 bit 0x2000 clear
        kind_disc: 3, // interface bag, low3 = 2
        type_flags: 2, // size class 2
        field5: 0,
        slot_count: 0,
        member_id: 0, // interface bag has no +0x30 write
        flag9: 0,
        flags_c: 0,
        field_1a: 0, // low6 = 0
        return_type_word: 6, // e.g. Integer
    };

    let rec = build_declaration_scalar(&mut h, &ctx, &[]).unwrap();
    let r = rec as usize;

    assert_eq!(h.mem[r], 0x20); // +0 |= 0x20 (no method tag on interface bag)
    assert_eq!(h.mem[r + 1], 0x22); // low3 = 2, size-class nibble 2
    assert_eq!(h.mem[r + 8], 1); // sig
    assert_eq!(h.mem[r + 0x18], 0x00); // sig >= 1 → bit 0x2000 clear
    assert_eq!(h.mem[r + 0x19], 0x00);
    assert_eq!(h.mem[r + 0x2c], 6); // Integer type node
    assert_eq!(h.mem[r + 0x3a], 0x08); // inline flag
    assert_eq!(h.mem[r + 0x3c], 0x00); // field_1a == 0
}

#[test]
fn negative_kind_disc_is_bad_decl() {
    let mut h = seeded_heap(0x80);
    let ctx = DeclContext {
        kind_disc: -1,
        ..DeclContext::default()
    };
    assert_eq!(build_declaration_scalar(&mut h, &ctx, &[]), Err(0x80028ca1u32 as i32));
}

#[test]
fn out_of_range_slot_count_is_bad_decl() {
    let mut h = seeded_heap(0x200);
    let ctx = DeclContext {
        kind_disc: 0,
        slot_count: 0x40, // > 0x3c
        field_1a: -1,
        ..DeclContext::default()
    };
    assert_eq!(build_declaration_scalar(&mut h, &ctx, &[]), Err(0x80028ca1u32 as i32));
}

#[test]
fn has_parameters_form_is_gated() {
    let mut h = seeded_heap(0x200);
    let ctx = DeclContext {
        kind_disc: 0,
        flag9: 0x19,
        field_1a: -1,
        ..DeclContext::default()
    };
    let r = std::panic::catch_unwind(std::panic::AssertUnwindSafe(|| {
        let _ = build_declaration_scalar(&mut h, &ctx, &[]);
    }));
    assert!(r.is_err());
}

/// Full ported-code round trip: the declaration compiler builds a scalar slot
/// record, and the ported resolver classifier reads it back cleanly (no gated
/// path) — the two halves of the compiler agree on the record layout.
#[test]
fn scalar_slot_record_flows_through_resolver() {
    let mut h = seeded_heap(0x400);
    // Build a parent method record first so the slot links somewhere real.
    let parent = build_declaration_scalar(
        &mut h,
        &DeclContext {
            kind_disc: 0,
            type_flags: 4,
            field_1a: -1,
            return_type_word: 8,
            ..DeclContext::default()
        },
        &[],
    )
    .unwrap();

    let mut tail = NIL;
    let slot = build_property_slot_scalar(&mut h, 8 /* Long */, 0, parent, &mut tail).unwrap();

    // Slot record shape the resolver depends on.
    assert_eq!(h.mem[slot as usize], 0x60); // inline (0x40) + 0x20
    assert_eq!(h.mem[slot as usize + 1], 0x03); // low3 = 3
    assert_eq!(h.mem[slot as usize + 0xc], 8); // inline Long type node
    // Linked under the parent's child list (head at parent+0x28).
    assert_eq!(h.read_dword(parent + 0x28), slot);
    assert_eq!(tail, slot);

    // The resolver classifier reads the operand at +0xc and produces a category
    // without hitting any gated path.
    let et = expression_type2(&h.mem, slot as usize);
    assert_eq!(et.code, 0);
    // The declaration-built operand classifies through the ported tables (the
    // call-convention category); full resolve_ident_ref category 4 needs the
    // EbResolveExprNode path, which is separately gated.
    assert_eq!(et.category, EXPECTED_LONG_SLOT_CATEGORY);
}

/// The record `build_declaration_scalar` produces must resolve through the
/// ported resolver: a scalar member's inline type node sits at `+0x2c`, but the
/// resolver reads the operand pointer at `+0xc`; confirm the record is at least
/// internally consistent for the fields the resolver inspects (`+0`, `+1`).
#[test]
fn produced_record_kind_and_byref_readable() {
    use crate::sym_record::MemberRecord;
    let mut h = seeded_heap(0x200);
    let ctx = DeclContext {
        sig: 0,
        kind_disc: 0,
        type_flags: 4,
        field5: 0,
        slot_count: 0,
        member_id: 0x55,
        flag9: 0,
        flags_c: 0,
        field_1a: -1,
        return_type_word: 8,
    };
    let rec = build_declaration_scalar(&mut h, &ctx, &[]).unwrap();
    let r = rec as usize;

    // Copy the produced bytes into a MemberRecord and read the call-path fields.
    let mut mr = MemberRecord::new();
    for (i, &byte) in h.mem[r..r + 0x40].iter().enumerate() {
        mr.set_byte(i, byte);
    }
    assert_eq!(mr.member_id(), 0x55);
    // +0x3a low3 == 0 ⇒ convention kind 0 (no 0x10 bit set here).
    assert_eq!(mr.kind(), 0);
}

// ── Record (`Type...End Type`) declaration: the field slot loop ─────────────

/// `Type Point : X As Long : Y As Long : End Type` — the milestone-1 fixture's
/// declaration. `type_flags == 8` routes straight to the slot loop (no single
/// scalar type node at `+0x2c`); each field becomes a property-bag slot
/// record via `build_property_slot_scalar`, linked under the declaration's
/// child list.
#[test]
fn record_declaration_builds_one_field_slot_per_entry() {
    let mut h = seeded_heap(0x400);
    let ctx = DeclContext {
        sig: 0,
        kind_disc: 0,
        type_flags: 8, // record type
        field5: 0,
        slot_count: 2, // two fields
        member_id: 0,
        flag9: 0,
        flags_c: 0,
        field_1a: -1,
        return_type_word: 0, // unused on the record path
    };
    let rec = build_declaration_scalar(&mut h, &ctx, &[8, 8] /* Long, Long */).unwrap();
    let r = rec as usize;

    // No single scalar type node written at +0x2c (the record path skips it).
    assert_eq!(h.mem[r + 0x2c], 0);
    // +0x3b low6 = slot count (2).
    assert_eq!(h.mem[r + 0x3b] & 0x3f, 2);

    // Two field records linked under the declaration's child list (head at
    // rec+0x28), in declaration order.
    let field_x = h.read_dword(rec + 0x28);
    assert_ne!(field_x, NIL);
    assert_eq!(h.mem[field_x as usize + 1], 0x03); // property-slot low3 = 3
    assert_eq!(h.mem[field_x as usize + 0xc], 8); // inline Long type node

    let field_y = h.read_dword(field_x + 0x14); // next-link
    assert_ne!(field_y, NIL);
    assert_eq!(h.mem[field_y as usize + 1], 0x03);
    assert_eq!(h.mem[field_y as usize + 0xc], 8);

    // Each field resolves cleanly through the ported classifier.
    assert_eq!(expression_type2(&h.mem, field_x as usize).category, EXPECTED_LONG_SLOT_CATEGORY);
    assert_eq!(expression_type2(&h.mem, field_y as usize).category, EXPECTED_LONG_SLOT_CATEGORY);
}

#[test]
#[should_panic(expected = "field_type_words must have one entry")]
fn record_declaration_rejects_mismatched_field_count() {
    let mut h = seeded_heap(0x400);
    let ctx = DeclContext {
        kind_disc: 0,
        type_flags: 8,
        slot_count: 2,
        field_1a: -1,
        ..DeclContext::default()
    };
    // Only one type word for a 2-field declaration.
    let _ = build_declaration_scalar(&mut h, &ctx, &[8]);
}

#[test]
fn record_declaration_with_no_fields_leaves_prior_gate_intact() {
    // type_flags == 8 with slot_count == 0: no fields to build, but this must
    // no longer hit the old "record-type check" gate — it just skips both the
    // scalar type node and the (empty) loop.
    let mut h = seeded_heap(0x200);
    let ctx = DeclContext {
        kind_disc: 0,
        type_flags: 8,
        slot_count: 0,
        field_1a: -1,
        ..DeclContext::default()
    };
    let rec = build_declaration_scalar(&mut h, &ctx, &[]).unwrap();
    assert_eq!(h.mem[rec as usize + 0x2c], 0);
}
