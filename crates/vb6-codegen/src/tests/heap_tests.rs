//! Tests for the module symbol-heap allocator ([`crate::heap`]).
//!
//! The free-block layout under test: each 8-byte header is
//! `[next: u32][size: u16][pad: u16]`, where `size` is the usable bytes (total
//! block size minus the 8-byte header) and `next` is `NIL` at the list tail.

use crate::heap::{HeapContext, EB_ALLOC_FAILED, NIL};

/// Build a heap with `len` bytes of zeroed backing and the in-place coalescing
/// mode enabled (flag bit 0), 2-byte alignment (flag bit 1 clear).
fn heap(len: usize) -> HeapContext {
    HeapContext {
        mem: vec![0u8; len],
        free_head: NIL,
        flags: 1,
        buffer_flag: 0,
    }
}

/// Lay a free-block header (next + size) at `off`.
fn put_block(h: &mut HeapContext, off: u32, next: u32, size: u16) {
    let o = off as usize;
    h.mem[o..o + 4].copy_from_slice(&next.to_le_bytes());
    h.mem[o + 4..o + 6].copy_from_slice(&size.to_le_bytes());
}

#[test]
fn align_size8_two_byte_mode() {
    let h = heap(0);
    // flag bit 1 clear → align to 2, minimum 8.
    assert_eq!(h.align_size8(0), 8);
    assert_eq!(h.align_size8(7), 8);
    assert_eq!(h.align_size8(8), 8);
    assert_eq!(h.align_size8(9), 10);
    assert_eq!(h.align_size8(10), 10);
    assert_eq!(h.align_size8(11), 12);
}

#[test]
fn align_size8_eight_byte_mode() {
    let mut h = heap(0);
    h.flags = 1 | 0b10; // align-to-8 mode
    assert_eq!(h.align_size8(0), 8);
    assert_eq!(h.align_size8(8), 8);
    assert_eq!(h.align_size8(9), 16);
    assert_eq!(h.align_size8(16), 16);
    assert_eq!(h.align_size8(17), 24);
}

#[test]
fn coalesce_into_empty_list_appends_head() {
    let mut h = heap(0x100);
    h.coalesce_memory(0x10, 0x20);
    assert_eq!(h.free_head, 0x10);
    assert_eq!(h.block_next(0x10), NIL);
    // size = total - 8 = 0x20 - 8 = 0x18
    assert_eq!(h.block_size(0x10), 0x18);
}

#[test]
fn coalesce_inserts_before_non_adjacent_block() {
    let mut h = heap(0x200);
    put_block(&mut h, 0x100, NIL, 0x10);
    h.free_head = 0x100;

    h.coalesce_memory(0x10, 0x20);

    // New block links ahead of the existing one; neither is adjacent so no merge.
    assert_eq!(h.free_head, 0x10);
    assert_eq!(h.block_next(0x10), 0x100);
    assert_eq!(h.block_size(0x10), 0x18);
    assert_eq!(h.block_next(0x100), NIL);
    assert_eq!(h.block_size(0x100), 0x10);
}

#[test]
fn coalesce_merges_onto_preceding_block_at_tail() {
    let mut h = heap(0x200);
    // Block 0x10 spans 0x10..0x30 (size 0x18 + 8-byte header = 0x20 total).
    put_block(&mut h, 0x10, NIL, 0x18);
    h.free_head = 0x10;

    // Free 0x30..0x50 — contiguous with the tail of block 0x10.
    h.coalesce_memory(0x30, 0x20);

    assert_eq!(h.free_head, 0x10);
    assert_eq!(h.block_next(0x10), NIL);
    // Grown by the full freed region size: 0x18 + 0x20 = 0x38.
    assert_eq!(h.block_size(0x10), 0x38);
}

#[test]
fn coalesce_merges_following_block_onto_new_region() {
    let mut h = heap(0x200);
    // Block 0x50 spans 0x50..0x70 (size 0x18).
    put_block(&mut h, 0x50, NIL, 0x18);
    h.free_head = 0x50;

    // Free 0x30..0x50 — contiguous with the head of block 0x50.
    h.coalesce_memory(0x30, 0x20);

    // New block absorbs the successor: size = succ.size + 8 + new.size.
    assert_eq!(h.free_head, 0x30);
    assert_eq!(h.block_next(0x30), NIL);
    assert_eq!(h.block_size(0x30), 0x18 + 8 + 0x18);
}

#[test]
fn find_free_block_splits_large_block() {
    let mut h = heap(0x200);
    // One free block at 0x10, usable size 0x40 (total 0x48).
    put_block(&mut h, 0x10, NIL, 0x40);
    h.free_head = 0x10;

    let off = h.find_free_block(0x10); // aligns to 0x10

    assert_eq!(off, 0x10);
    // Remainder block carved at 0x10 + 0x10 = 0x20, size = 0x40 - 0x10 = 0x30.
    assert_eq!(h.free_head, 0x20);
    assert_eq!(h.block_next(0x20), NIL);
    assert_eq!(h.block_size(0x20), 0x30);
}

#[test]
fn find_free_block_consumes_block_whole_when_remainder_too_small() {
    let mut h = heap(0x200);
    // Usable 0x08 → total 0x10; request 0x10 leaves 0 spare → no split.
    put_block(&mut h, 0x10, NIL, 0x08);
    h.free_head = 0x10;

    let off = h.find_free_block(0x10);

    assert_eq!(off, 0x10);
    assert_eq!(h.free_head, NIL); // list emptied
}

#[test]
fn find_free_block_returns_nil_when_nothing_fits() {
    let mut h = heap(0x200);
    put_block(&mut h, 0x10, NIL, 0x04); // total 0x0c < requested 0x10
    h.free_head = 0x10;

    let off = h.find_free_block(0x10);

    assert_eq!(off, NIL);
    // List untouched.
    assert_eq!(h.free_head, 0x10);
    assert_eq!(h.block_next(0x10), NIL);
    assert_eq!(h.block_size(0x10), 0x04);
}

#[test]
fn find_free_block_skips_too_small_then_splits_second() {
    let mut h = heap(0x200);
    put_block(&mut h, 0x10, 0x40, 0x04); // too small, links to 0x40
    put_block(&mut h, 0x40, NIL, 0x40); // fits
    h.free_head = 0x10;

    let off = h.find_free_block(0x10);

    assert_eq!(off, 0x40);
    // First block now points at the carved remainder at 0x50.
    assert_eq!(h.block_next(0x10), 0x50);
    assert_eq!(h.block_size(0x50), 0x30);
    assert_eq!(h.block_next(0x50), NIL);
}

/// Seed a heap with a single free block at offset 0 spanning the whole buffer,
/// so the free-list-hit allocation path can run without the gated grow path.
fn seeded(len: usize) -> HeapContext {
    let mut h = heap(len);
    put_block(&mut h, 0, NIL, (len - 8) as u16);
    h.free_head = 0;
    h
}

#[test]
fn allocate_heap_space_serves_from_free_list() {
    let mut h = seeded(0x200);
    let off = h.allocate_heap_space(0x20).unwrap();
    assert_eq!(off, 0);
    // Remainder free block carved just past the 0x20 allocation.
    assert_eq!(h.free_head, 0x20);
}

#[test]
fn allocate_heap_space_fails_when_oversized_and_growth_disabled() {
    let mut h = seeded(0x200);
    h.buffer_flag = 1; // growth disabled
    // 0x10000+ aligns past the in-line limit → straight to failure.
    assert_eq!(h.allocate_heap_space(0x10000), Err(EB_ALLOC_FAILED));
}

#[test]
fn allocate_heap_space_fails_when_no_fit_and_growth_disabled() {
    let mut h = heap(0x40);
    put_block(&mut h, 0, NIL, 0x04); // tiny block, nothing fits
    h.free_head = 0;
    h.buffer_flag = 1; // growth disabled
    assert_eq!(h.allocate_heap_space(0x20), Err(EB_ALLOC_FAILED));
}

#[test]
fn allocate_heap_space_grow_path_is_gated() {
    let mut h = heap(0x40); // empty free list, growth enabled (buffer_flag 0)
    let r = std::panic::catch_unwind(std::panic::AssertUnwindSafe(|| {
        h.allocate_heap_space(0x20).ok();
    }));
    assert!(r.is_err());
}

#[test]
fn method_bag_sets_method_tag() {
    let mut h = seeded(0x200);
    let off = h.allocate_method_bag().unwrap();
    assert_eq!(off, 0);
    assert_eq!(h.mem[off as usize], 4);
    // Remaining record bytes zeroed.
    assert!(h.mem[off as usize + 1..off as usize + 0x40].iter().all(|&b| b == 0));
}

#[test]
fn interface_bag_is_all_zero() {
    let mut h = seeded(0x200);
    let off = h.allocate_interface_bag().unwrap();
    assert!(h.mem[off as usize..off as usize + 0x40].iter().all(|&b| b == 0));
}

#[test]
fn property_bag_is_all_zero() {
    let mut h = seeded(0x200);
    let off = h.allocate_property_bag().unwrap();
    assert!(h.mem[off as usize..off as usize + 0x1c].iter().all(|&b| b == 0));
}

#[test]
fn parameter_bag_sets_tag_and_marker_bytes() {
    let mut h = seeded(0x200);
    let off = h.allocate_parameter_bag().unwrap();
    let o = off as usize;
    assert_eq!(h.mem[o], 2);
    assert_eq!(h.mem[o + 0x12], 0xff);
    assert_eq!(h.mem[o + 0x13], 0xff);
    // Every other byte of the 0x28 record is zero.
    for (i, &b) in h.mem[o..o + 0x28].iter().enumerate() {
        if i == 0 || i == 0x12 || i == 0x13 {
            continue;
        }
        assert_eq!(b, 0, "byte +{i:#x}");
    }
}

#[test]
fn type_descriptor_zeroes_first_0x20_only() {
    let mut h = seeded(0x200);
    // Dirty the trailing 4 bytes the allocator must leave untouched.
    h.mem[0x20..0x24].fill(0xaa);
    let off = h.allocate_type_descriptor().unwrap();
    assert_eq!(off, 0);
    assert!(h.mem[0..0x20].iter().all(|&b| b == 0));
    assert_eq!(&h.mem[0x20..0x24], &[0xaa; 4]);
}

#[test]
fn link_list_node3_first_node_sets_head_and_tail() {
    let mut h = heap(0x100);
    let mut tail = NIL;
    h.link_list_node3(0x40, 0x10, &mut tail);
    assert_eq!(tail, 0x40);
    assert_eq!(h.read_dword(0x10), 0x40); // head pointer written
}

#[test]
fn link_list_node3_appends_after_existing_tail() {
    let mut h = heap(0x100);
    let mut tail = NIL;
    h.link_list_node3(0x40, 0x10, &mut tail); // first
    h.link_list_node3(0x80, 0x10, &mut tail); // second
    assert_eq!(tail, 0x80);
    // Head still points at the first node; first node's +0x14 links to second.
    assert_eq!(h.read_dword(0x10), 0x40);
    assert_eq!(h.read_dword(0x40 + 0x14), 0x80);
}

#[test]
fn coalesce_deferred_free_mode_is_gated() {
    let mut h = heap(0x100);
    h.flags = 0; // bit 0 clear → deferred-free path
    let r = std::panic::catch_unwind(std::panic::AssertUnwindSafe(|| {
        h.coalesce_memory(0x10, 0x20);
    }));
    assert!(r.is_err());
}
