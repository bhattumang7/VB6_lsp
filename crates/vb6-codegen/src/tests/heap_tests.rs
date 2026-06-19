//! Tests for the module symbol-heap allocator ([`crate::heap`]).
//!
//! The free-block layout under test: each 8-byte header is
//! `[next: u32][size: u16][pad: u16]`, where `size` is the usable bytes (total
//! block size minus the 8-byte header) and `next` is `NIL` at the list tail.

use crate::heap::{HeapContext, NIL};

/// Build a heap with `len` bytes of zeroed backing and the in-place coalescing
/// mode enabled (flag bit 0), 2-byte alignment (flag bit 1 clear).
fn heap(len: usize) -> HeapContext {
    HeapContext {
        mem: vec![0u8; len],
        free_head: NIL,
        flags: 1,
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
fn coalesce_deferred_free_mode_is_gated() {
    let mut h = heap(0x100);
    h.flags = 0; // bit 0 clear → deferred-free path
    let r = std::panic::catch_unwind(std::panic::AssertUnwindSafe(|| {
        h.coalesce_memory(0x10, 0x20);
    }));
    assert!(r.is_err());
}
