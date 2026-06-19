use super::*;
use crate::emit::RefDescriptor;
use crate::node::NodeArena;

/// A fixed-size type descriptor (`word[0] == 4`) carrying `size` in word[4].
fn size_desc(a: &mut NodeArena, size: u16) -> NodeRef {
    a.alloc(NodeArena::node(4, 0, size as u32, 0, 0, 0))
}

/// Build a reference expression node: region tag in word[0] high half, flag word
/// in word[1], the size descriptor in word[5], and the secondary descriptor in
/// word[6].
fn ref_node(
    a: &mut NodeArena,
    region_tag: u16,
    flags1word: u32,
    w5_desc: NodeRef,
    w6_desc: NodeRef,
) -> NodeRef {
    let mut n = NodeArena::node(0x60, region_tag, 0, w5_desc.0, w6_desc.0, 0);
    n.w[1] = flags1word;
    a.alloc(n)
}

#[test]
fn small_type_by_value_is_kind_2() {
    let mut a = NodeArena::new();
    let d = size_desc(&mut a, 4);
    let n = ref_node(&mut a, 0, 0, d, NodeRef(0));
    let desc = init_expr_descriptor(&a, n, false, true);
    assert_eq!(
        desc,
        RefDescriptor { kind: 2, operand: 4, word6: 0, word8: 0, flags1: 0 }
    );
}

#[test]
fn small_type_by_reference_is_kind_1() {
    let mut a = NodeArena::new();
    let d = size_desc(&mut a, 4);
    let n = ref_node(&mut a, 0, 0, d, NodeRef(0));
    let desc = init_expr_descriptor(&a, n, true, true);
    assert_eq!(desc.kind, 1);
    assert_eq!(desc.operand, 4);
}

#[test]
fn other_type_by_value_is_kind_7() {
    // flags bit 0x100 set and size != 8 → "other" class.
    let mut a = NodeArena::new();
    let d = size_desc(&mut a, 2);
    let n = ref_node(&mut a, 0, 0x100, d, NodeRef(0));
    let desc = init_expr_descriptor(&a, n, false, true);
    assert_eq!(desc.kind, 7);
    assert_eq!(desc.operand, 2);
}

#[test]
fn other_type_by_reference_is_kind_6() {
    let mut a = NodeArena::new();
    let d = size_desc(&mut a, 2);
    let n = ref_node(&mut a, 0, 0x100, d, NodeRef(0));
    let desc = init_expr_descriptor(&a, n, true, true);
    assert_eq!(desc.kind, 6);
}

#[test]
fn size_eight_forces_small_class_even_with_flag() {
    // flags bit 0x100 set but size == 8 → still the small class (kind 2/1).
    let mut a = NodeArena::new();
    let d = size_desc(&mut a, 8);
    let n = ref_node(&mut a, 0, 0x100, d, NodeRef(0));
    let desc = init_expr_descriptor(&a, n, false, true);
    assert_eq!(desc.kind, 2);
    assert_eq!(desc.operand, 8);
}

#[test]
fn non_optional_sets_usage_flag_and_region_marker() {
    // Not optional → flags1 bit 0x04; a 0x160000-region node also sets bit 0x01.
    let mut a = NodeArena::new();
    let d = size_desc(&mut a, 4);
    let n = ref_node(&mut a, 0x16, 0, d, NodeRef(0)); // region 0x160000
    let desc = init_expr_descriptor(&a, n, false, false);
    assert_eq!(desc.flags1, 0x05);
}

#[test]
fn non_optional_non_region_sets_only_usage_flag() {
    let mut a = NodeArena::new();
    let d = size_desc(&mut a, 4);
    let n = ref_node(&mut a, 0, 0, d, NodeRef(0));
    let desc = init_expr_descriptor(&a, n, false, false);
    assert_eq!(desc.flags1, 0x04);
}

#[test]
fn out_of_line_size_path_sets_word6_and_word8() {
    // flags bit 0x2000 set with bit 0 clear → out-of-line size path: word6 bit 0
    // and word8 = secondary descriptor size (word[6]).
    let mut a = NodeArena::new();
    let d5 = size_desc(&mut a, 4);
    let d6 = size_desc(&mut a, 0x10);
    let n = ref_node(&mut a, 0, 0x2000, d5, d6);
    let desc = init_expr_descriptor(&a, n, false, true);
    assert_eq!(desc.word6 & 1, 1);
    assert_eq!(desc.word8, 0x10);
}

#[test]
fn get_expr_context_returns_offset_for_member_kinds() {
    for kind in [4, 5, 6] {
        let ctx = CompileContext { kind, member_offset: 0x18 };
        assert_eq!(get_expr_context(&ctx), 0x18);
    }
    let ctx = CompileContext { kind: 1, member_offset: 0x18 };
    assert_eq!(get_expr_context(&ctx), -1);
}

#[test]
fn pcode_terminator_matches_0x1b_0x1c() {
    assert!(is_pcode_terminator(0x1b));
    assert!(is_pcode_terminator(0x1c));
    assert!(is_pcode_terminator(0x40 | 0x1b)); // high bits ignored
    assert!(!is_pcode_terminator(0x1a));
    assert!(!is_pcode_terminator(0x1d));
}

#[test]
fn extract_type_info_class_0xd_and_0x1d() {
    assert_eq!(extract_type_info(&[0x0d, 0, 0, 0]), 0xfffe);
    // 0x1d class: type word two bytes in, low bit cleared.
    assert_eq!(extract_type_info(&[0x1d, 0x00, 0x57, 0x12]), 0x1256 & 0xfffe);
}

#[test]
fn map_slot_type_value_mapping() {
    assert_eq!(map_slot_type_value(0), Some(0));
    assert_eq!(map_slot_type_value(1), Some(1));
    assert_eq!(map_slot_type_value(2), None);
    assert_eq!(map_slot_type_value(3), Some(2));
    assert_eq!(map_slot_type_value(5), Some(2));
    assert_eq!(map_slot_type_value(8), Some(2));
    assert_eq!(map_slot_type_value(6), None);
    assert_eq!(map_slot_type_value(7), None);
    assert_eq!(map_slot_type_value(9), None);
}

#[test]
fn current_expression_offset_skips_typed_0x1d() {
    // 0x1d-class opcode with following opcode != 0x25 → skip 4.
    assert_eq!(current_expression_offset(&[0x1d, 0, 0, 0, 0x10, 0, 0, 0]), 4);
    // 0x1d-class but the following opcode is a 0x25-class marker → no skip.
    assert_eq!(current_expression_offset(&[0x1d, 0, 0, 0, 0x25, 0, 0, 0]), 0);
    // high bits of either opcode are masked off before the comparison.
    assert_eq!(current_expression_offset(&[0xc0 | 0x1d, 0, 0, 0, 0x40 | 0x25, 0, 0, 0]), 0);
    // not a 0x1d-class opcode → no skip.
    assert_eq!(current_expression_offset(&[0x0d, 0, 0, 0, 0x10, 0, 0, 0]), 0);
}

#[test]
fn resolver_class_flag_table_exact() {
    // Spot the logical class-flag entries (low 0x28) and the call-conv aliasing
    // beyond it that the `& 0x3f` accessor still reaches.
    assert_eq!(resolver_class_flag(0x00), 0);
    assert_eq!(resolver_class_flag(0x02), 1);
    assert_eq!(resolver_class_flag(0x11), 1);
    assert_eq!(resolver_class_flag(0x20), 1);
    assert_eq!(resolver_class_flag(0x21), 1);
    assert_eq!(resolver_class_flag(0x27), 0);
    // High bits of the opcode are masked away.
    assert_eq!(resolver_class_flag(0xc0 | 0x02), 1);
    // Index 0x28..0x3f aliases the call-convention records (first byte 0x13).
    assert_eq!(resolver_class_flag(0x28), 0x13);
    assert_eq!(resolver_class_flag(0x3f), 0x2d);
}

#[test]
fn resolver_inspects_operand_gate() {
    // class-flag 0 and low6 not 0x1e/0x1f → inspect.
    assert!(resolver_inspects_operand(0x00));
    assert!(resolver_inspects_operand(0x1d));
    // class-flag 1 → no inspection.
    assert!(!resolver_inspects_operand(0x02));
    // low6 0x1e / 0x1f are explicitly excluded even though their flag is 0.
    assert!(!resolver_inspects_operand(0x1e));
    assert!(!resolver_inspects_operand(0x1f));
}

#[test]
fn resolver_type_category_index_math() {
    // value_class 0, member byte1 0, op byte1 0 → index 0 → 0x0d.
    assert_eq!(resolver_type_category(0, 0, 0), 0x0d);
    // index 2 → 0x0e (value_class 1).
    assert_eq!(resolver_type_category(1, 0, 0), 0x0e);
    // index (0 + 1*3)*2 + 0 = 6 → 0x00.
    assert_eq!(resolver_type_category(0, 1, 0), 0x00);
    // index (2 + 1*3)*2 + 1 = 11 → 0x00.
    assert_eq!(resolver_type_category(2, 1, 1), 0x00);
    // Only the low 3 bits of each byte index in.
    assert_eq!(resolver_type_category(0, 0xf8, 0xf8), 0x0d);
    // Max reachable index 0x35 (value_class 2, both nibbles 7) → 0x00.
    assert_eq!(resolver_type_category(2, 7, 7), 0x00);
    // index 0x12 → 0x04 (value_class 0, member byte1 3, op byte1 0).
    assert_eq!(resolver_type_category(0, 3, 0), 0x04);
}

#[test]
fn type_library_descriptor_is_gated() {
    // flags 0x4000 with a 0x170000-region node hits the gated type-library path.
    let mut a = NodeArena::new();
    let d = size_desc(&mut a, 4);
    let n = ref_node(&mut a, 0x17, 0x4000, d, NodeRef(0)); // region 0x170000
    let r = std::panic::catch_unwind(std::panic::AssertUnwindSafe(|| {
        init_expr_descriptor(&a, n, false, true)
    }));
    assert!(r.is_err());
}
