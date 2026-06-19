//! Tests for the type-node flag mutators ([`crate::typenode`]).

use crate::typenode::{
    build_inline_type_node, process_type3_simple, set_type_flag4, toggle_bitfield, toggle_type_flag,
};

#[test]
fn toggle_bitfield_to_0x1d_stamps_follow_on_opcode() {
    let mut node = [0u8; 8];
    toggle_bitfield(&mut node, 0x1d);
    assert_eq!(node[0], 0x1d);
    // word[0] copied to +4/+5, then +4 = (+4 & 0xe5) | 0x25.
    assert_eq!(node[4], 0x25);
    assert_eq!(node[5], 0x00);
}

#[test]
fn toggle_bitfield_non_0x1d_leaves_follow_on_untouched() {
    let mut node = [0u8; 8];
    node[4] = 0xaa;
    toggle_bitfield(&mut node, 0x0c);
    assert_eq!(node[0], 0x0c);
    assert_eq!(node[4], 0xaa); // untouched
}

#[test]
fn set_type_flag4_sets_bit3_on_byte1_only_for_plain_node() {
    let mut node = [0u8; 8]; // +0 == 0, not the 0x1d form
    set_type_flag4(&mut node, true);
    assert_eq!(node[1], 0x08);
    assert_eq!(node[5], 0x00); // secondary slot untouched
    set_type_flag4(&mut node, false);
    assert_eq!(node[1], 0x00);
}

#[test]
fn set_type_flag4_updates_secondary_slot_for_0x1d_node() {
    let mut node = [0u8; 8];
    node[0] = 0x1d; // 0x1d form
    node[4] = 0x00; // low6 != 0x25 → secondary slot live
    set_type_flag4(&mut node, true);
    assert_eq!(node[1], 0x08);
    assert_eq!(node[5], 0x08);
}

#[test]
fn toggle_type_flag_sets_low3_kind_preserving_high_bits() {
    let mut node = [0u8; 8];
    node[1] = 0xf0; // high bits must survive
    toggle_type_flag(&mut node, 3);
    assert_eq!(node[1], 0xf3); // low3 = 3, high bits kept
    // The op sets (not toggles): re-applying the same value is idempotent.
    toggle_type_flag(&mut node, 3);
    assert_eq!(node[1], 0xf3);
    // A different value replaces the low 3 bits.
    toggle_type_flag(&mut node, 0);
    assert_eq!(node[1], 0xf0);
}

#[test]
fn build_inline_type_node_encodes_base_type_code() {
    // A base type code becomes the node opcode (low 6 bits of byte 0).
    assert_eq!(build_inline_type_node(0x08), 0x08); // e.g. Long
    assert_eq!(build_inline_type_node(0x06), 0x06); // e.g. Integer
    assert_eq!(build_inline_type_node(0x0b), 0x0b); // e.g. Double
}

#[test]
fn build_inline_type_node_remaps_variant_like_codes_to_object_form() {
    assert_eq!(build_inline_type_node(0x0a), 3);
    assert_eq!(build_inline_type_node(0x16), 3);
    assert_eq!(build_inline_type_node(0x19), 3);
}

#[test]
fn process_type3_simple_op_bits() {
    let mut n = [0u8; 8];
    process_type3_simple(&mut n, 6).unwrap();
    assert_eq!(u16::from_le_bytes([n[2], n[3]]), 0x0880);

    let mut n = [0u8; 8];
    process_type3_simple(&mut n, 7).unwrap();
    assert_eq!(u16::from_le_bytes([n[2], n[3]]), 0x0180);

    let mut n = [0u8; 8];
    process_type3_simple(&mut n, 8).unwrap();
    assert_eq!(u16::from_le_bytes([n[2], n[3]]), 0x0440);

    let mut n = [0u8; 8];
    process_type3_simple(&mut n, 5).unwrap(); // fallthrough: only 0x80
    assert_eq!(u16::from_le_bytes([n[2], n[3]]), 0x0080);
}

#[test]
fn process_type3_simple_clears_then_sets() {
    let mut n = [0u8; 8];
    n[2] = 0xff;
    n[3] = 0xff;
    process_type3_simple(&mut n, 6).unwrap();
    // 0xffff & 0xf2ff = 0xf2ff, then | 0x0880 = 0xfaff.
    assert_eq!(u16::from_le_bytes([n[2], n[3]]), 0xfaff);
}
