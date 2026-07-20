use super::*;
use crate::emit::RefDescriptor;
use crate::node::NodeArena;

/// A fixed-size type descriptor (`word[0] == 4`) carrying `size` in word[4].
fn size_desc(a: &mut NodeArena, size: u16) -> NodeRef {
    a.alloc(NodeArena::node(4, 0, size as u32, 0, 0, 0))
}

/// Build a reference expression node: region tag in word[0] high half, flag word
/// in word[1], the size descriptor in word[5], and the secondary descriptor in
/// word[6]. `word[7]` defaults to 0 (no front-end-resolved frame offset); use
/// [`ref_node_with_offset`] when the test cares about the emitted operand.
fn ref_node(
    a: &mut NodeArena,
    region_tag: u16,
    flags1word: u32,
    w5_desc: NodeRef,
    w6_desc: NodeRef,
) -> NodeRef {
    ref_node_with_offset(a, region_tag, flags1word, w5_desc, w6_desc, 0)
}

/// [`ref_node`] plus an explicit `word[7]` — the front-end-resolved frame
/// offset `init_expr_descriptor` copies verbatim into the descriptor operand
/// (a value deliberately distinct from the word[5] type size, so a test
/// asserting on `operand` proves the two are not conflated).
fn ref_node_with_offset(
    a: &mut NodeArena,
    region_tag: u16,
    flags1word: u32,
    w5_desc: NodeRef,
    w6_desc: NodeRef,
    offset: u16,
) -> NodeRef {
    let mut n = NodeArena::node(0x60, region_tag, 0, w5_desc.0, w6_desc.0, offset as u32);
    n.w[1] = flags1word;
    a.alloc(n)
}

#[test]
fn small_type_by_value_is_kind_2() {
    let mut a = NodeArena::new();
    let d = size_desc(&mut a, 4);
    // Offset (0x20) deliberately differs from the word[5] size (4): operand
    // must track the offset, not the size.
    let n = ref_node_with_offset(&mut a, 0, 0, d, NodeRef(0), 0x20);
    let desc = init_expr_descriptor(&a, n, false, true);
    assert_eq!(
        desc,
        RefDescriptor { kind: 2, operand: 0x20, word6: 0, word8: 0, flags1: 0 }
    );
}

#[test]
fn small_type_by_reference_is_kind_1() {
    let mut a = NodeArena::new();
    let d = size_desc(&mut a, 4);
    let n = ref_node_with_offset(&mut a, 0, 0, d, NodeRef(0), 0x20);
    let desc = init_expr_descriptor(&a, n, true, true);
    assert_eq!(desc.kind, 1);
    assert_eq!(desc.operand, 0x20);
}

#[test]
fn other_type_by_value_is_kind_7() {
    // flags bit 0x100 set and size != 8 → "other" class.
    let mut a = NodeArena::new();
    let d = size_desc(&mut a, 2);
    let n = ref_node_with_offset(&mut a, 0, 0x100, d, NodeRef(0), 0x20);
    let desc = init_expr_descriptor(&a, n, false, true);
    assert_eq!(desc.kind, 7);
    assert_eq!(desc.operand, 0x20);
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
    // Offset (0x20) deliberately differs from the word[5] size (8).
    let mut a = NodeArena::new();
    let d = size_desc(&mut a, 8);
    let n = ref_node_with_offset(&mut a, 0, 0x100, d, NodeRef(0), 0x20);
    let desc = init_expr_descriptor(&a, n, false, true);
    assert_eq!(desc.kind, 2);
    assert_eq!(desc.operand, 0x20);
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

// ── Reader chain (EbGetExpressionType2 and its helpers) ──────────────────────

/// Build a records heap and place a member record at `record_off` with the given
/// `+0` flag byte, `+1` byte, and `+0xc` masked dword. Operand p-code bytes are
/// written separately by each test.
fn heap_with_record(record_off: usize, flag0: u8, byte1: u8, masked: u32) -> Vec<u8> {
    let mut h = vec![0u8; 0x60];
    h[record_off] = flag0;
    h[record_off + 1] = byte1;
    h[record_off + 0xc..record_off + 0x10].copy_from_slice(&masked.to_le_bytes());
    h
}

#[test]
fn get_masked_value_inline_and_indirect() {
    // 0x40 set → raw +0xc dword.
    let mut r = vec![0u8; 0x10];
    r[0] = 0x40;
    r[0xc..0x10].copy_from_slice(&0x1234_5679u32.to_le_bytes());
    assert_eq!(get_masked_value(&r), 0x1234_5679);
    // 0x40 clear → low bit cleared.
    r[0] = 0x00;
    assert_eq!(get_masked_value(&r), 0x1234_5678);
}

#[test]
fn resolve_attribute_pointer_inline_vs_indirect() {
    // Inline (0x40 set): operand at record + 0xc.
    let h = heap_with_record(0x10, 0x40, 0, 0x0000_0002);
    assert_eq!(resolve_attribute_pointer(&h, 0x10), Some(0x1c));
    // Indirect (0x40 clear): operand at the masked heap offset.
    let h = heap_with_record(0x10, 0x00, 0, 0x0000_0028);
    assert_eq!(resolve_attribute_pointer(&h, 0x10), Some(0x28));
    // Null: masked value 0xffffffff (0x40 set).
    let h = heap_with_record(0x10, 0x40, 0, 0xffff_ffff);
    assert_eq!(resolve_attribute_pointer(&h, 0x10), None);
}

#[test]
fn method_flags_class_bit_pattern() {
    assert_eq!(method_flags_class(&[0x00, 0x40]), 1); // 0x40 set, 0x20 clear
    assert_eq!(method_flags_class(&[0x00, 0x60]), 0); // 0x20 also set
    assert_eq!(method_flags_class(&[0x00, 0x00]), 0); // 0x40 clear
}

#[test]
fn expression_type2_uninspected_operand_class_zero() {
    // Inline operand opcode 0x02 has class-flag 1 → not inspected → value_class 0.
    // member byte1 = 0, op byte1 = 0 → type-map index 0 → 0x0d.
    let mut h = heap_with_record(0x10, 0x40, 0, 0x0000_0002);
    h[0x1c] = 0x02; // operand opcode
    h[0x1d] = 0x00; // operand byte 1
    assert_eq!(expression_type2(&h, 0x10), ExpressionType { category: 0x0d, code: 0 });
}

#[test]
fn expression_type2_inspected_operand_class_two() {
    // Inline operand opcode 0x00 has class-flag 0, not 0x1a/0x1d/0x1e/0x1f →
    // value_class 2. member byte1 = 1, op byte1 = 0 → index (2 + 3)*2 = 10 → 0x12.
    let mut h = heap_with_record(0x10, 0x40, 1, 0x0000_0000);
    h[0x1c] = 0x00;
    h[0x1d] = 0x00;
    assert_eq!(expression_type2(&h, 0x10).category, 0x12);
}

#[test]
fn expression_type2_method_flag_class_one() {
    // 0x1d operand with byte1 bit 0x40 set (0x20 clear) → method_flags_class = 1.
    // The +4 byte is a 0x25-class marker so the current-expression skip is
    // suppressed. member byte1 = 0, op byte1 (0x40) & 7 = 0 → index 2 → 0x0e.
    let mut h = heap_with_record(0x10, 0x40, 0, 0x0000_401d);
    h[0x1c] = 0x1d;
    h[0x1d] = 0x40;
    h[0x20] = 0x25; // suppress the +4 skip
    assert_eq!(expression_type2(&h, 0x10).category, 0x0e);
}

#[test]
fn expression_type2_applies_current_expression_skip() {
    // 0x1d operand whose +4 opcode is not 0x25 → skip +4 to a class-flag-1 opcode
    // (value_class 0). member byte1 = 0, op byte1 at 0x21 = 0 → index 0 → 0x0d.
    let mut h = heap_with_record(0x10, 0x40, 0, 0x0000_001d);
    h[0x1c] = 0x1d;
    h[0x20] = 0x05; // +4 opcode (not 0x25) → skip lands here
    h[0x21] = 0x00;
    assert_eq!(expression_type2(&h, 0x10).category, 0x0d);
}

#[test]
fn expression_type2_indirect_operand_offset() {
    // Indirect record (0x40 clear): operand lives at the masked offset 0x28.
    // opcode 0x05 (class-flag 1) → value_class 0; member byte1 = 3, op byte1 = 2
    // → index (0 + 9)*2 + 2 = 20 → 0x05.
    let mut h = heap_with_record(0x10, 0x00, 3, 0x0000_0028);
    h[0x28] = 0x05;
    h[0x29] = 0x02;
    assert_eq!(expression_type2(&h, 0x10).category, 0x05);
}

#[test]
fn expression_type2_slot_path_is_gated() {
    // 0x1d operand with byte1 bit 0x40 clear → the slot-type path (gated).
    let mut h = heap_with_record(0x10, 0x40, 0, 0x0000_001d);
    h[0x1c] = 0x1d;
    h[0x1d] = 0x00; // 0x40 clear
    h[0x20] = 0x25; // suppress skip so we reach the slot branch
    let r = std::panic::catch_unwind(|| expression_type2(&h, 0x10));
    assert!(r.is_err());
}

// ── EbResolveIdentRef dispatcher ─────────────────────────────────────────────

/// A records heap with an inline member record at 0x10 (operand at 0x1c). `rec0`
/// / `rec1` are record bytes `+0` / `+1`; `op` / `op_byte1` are the operand's two
/// bytes. The record's `+0xc` masked dword coincides with the operand bytes
/// (inline), which is non-`-1` for the values used here.
fn ident_heap(rec0: u8, rec1: u8, op: u8, op_byte1: u8) -> Vec<u8> {
    let mut h = vec![0u8; 0x40];
    h[0x10] = rec0;
    h[0x11] = rec1;
    h[0x1c] = op;
    h[0x1d] = op_byte1;
    h
}

/// A clean 0x60 reference node: type tag in the high half, the given flags, a
/// fixed-size (4-byte) type descriptor in word[5], and a front-end-resolved
/// frame offset (0x20 — deliberately distinct from the word[5] size) in
/// word[7], so a test asserting on `operand` proves it tracks the offset.
fn ident_node(a: &mut NodeArena, type_tag: u16, flags: u32) -> NodeRef {
    let d = size_desc(a, 4);
    ref_node_with_offset(a, type_tag, flags, d, NodeRef(0), 0x20)
}

#[test]
fn resolve_ident_ref_category_7_value_load() {
    // rec1=3, op byte1=1, value_class 0 → category 7. Operand 0x02 (!= 0x1b) →
    // optional. type tag 5 → type-offset 0 (no tail flag).
    let mut a = NodeArena::new();
    let n = ident_node(&mut a, 5, 0);
    let h = ident_heap(0x40, 3, 0x02, 1);
    let desc = resolve_ident_ref(&a, n, &h, 0x10, 0, None, RawNode::default());
    assert_eq!(
        desc,
        RefDescriptor { kind: 2, operand: 0x20, word6: 0, word8: 0, flags1: 0 }
    );
}

#[test]
fn resolve_ident_ref_category_0xc_non_optional() {
    // rec1=3, op byte1=5 → category 0xc → init(by_ref=false, optional=false) →
    // flags1 bit 0x04.
    let mut a = NodeArena::new();
    let n = ident_node(&mut a, 5, 0);
    let h = ident_heap(0x40, 3, 0x02, 5);
    let desc = resolve_ident_ref(&a, n, &h, 0x10, 0, None, RawNode::default());
    assert_eq!(desc, RefDescriptor { kind: 2, operand: 0x20, word6: 0, word8: 0, flags1: 4 });
}

#[test]
fn resolve_ident_ref_category_9_by_reference() {
    // rec1=3, op byte1=4 → category 9 → init(by_ref=true, optional=false) → kind 1.
    let mut a = NodeArena::new();
    let n = ident_node(&mut a, 5, 0);
    let h = ident_heap(0x40, 3, 0x02, 4);
    let desc = resolve_ident_ref(&a, n, &h, 0x10, 0, None, RawNode::default());
    assert_eq!(desc, RefDescriptor { kind: 1, operand: 0x20, word6: 0, word8: 0, flags1: 4 });
}

#[test]
fn resolve_ident_ref_category_1_sets_attribute_flag() {
    // rec1=0x24 (low 3 = 4, bit 0x20 set), op byte1=0 → category 1. The 0x20 record
    // flag with ctx flag bit 2 clear sets descriptor flags1 bit 0x02.
    let mut a = NodeArena::new();
    let n = ident_node(&mut a, 5, 0);
    let h = ident_heap(0x40, 0x24, 0x02, 0);
    let desc = resolve_ident_ref(&a, n, &h, 0x10, 0, None, RawNode::default());
    assert_eq!(desc, RefDescriptor { kind: 1, operand: 0x20, word6: 0, word8: 0, flags1: 2 });
    // ctx flag bit 2 set suppresses the attribute flag.
    let desc = resolve_ident_ref(&a, n, &h, 0x10, 2, None, RawNode::default());
    assert_eq!(desc.flags1, 0);
}

#[test]
fn resolve_ident_ref_type_offset_0xe_sets_tail_flag() {
    // type tag 17 → type-offset 0xe → the shared tail sets flags1 bit 0x08.
    // Category 0xc gives the 0x04 flag; combined → 0x0c.
    let mut a = NodeArena::new();
    let n = ident_node(&mut a, 17, 0);
    let h = ident_heap(0x40, 3, 0x02, 5);
    let desc = resolve_ident_ref(&a, n, &h, 0x10, 0, None, RawNode::default());
    assert_eq!(desc.flags1, 0x0c);
}

#[test]
fn resolve_ident_ref_method_binding_zero_slot_gated() {
    // record +0 bit 0x80 with +1 & 7 == 4, AND the record's own has-slot check
    // comes up zero (slot_id = -1 -> (slot_id+1)&0xfff == 0) -> the genuinely
    // COM-dependent zero-slot sub-path (gated).
    let mut a = NodeArena::new();
    let n = ident_node(&mut a, 5, 0);
    let mut h = ident_heap(0xc0, 4, 0x02, 0);
    h[0x10 + 10] = 0xff;
    h[0x10 + 11] = 0xff; // slot_id = -1 -> has_slot == false
    let r = std::panic::catch_unwind(std::panic::AssertUnwindSafe(|| {
        resolve_ident_ref(&a, n, &h, 0x10, 0, None, RawNode::default())
    }));
    assert!(r.is_err());
}

#[test]
fn resolve_ident_ref_method_binding_nonzero_slot_remaps_to_0xd_e_f() {
    // Same method-binding flags, but the has-slot check is NONZERO (slot_id =
    // 0, the ident_heap default) -> the gate is narrowed: this does NOT hit
    // the zero-slot COM path at all. The category remaps (1/2/3 -> 0xd/0xe/0xf)
    // and falls through to the (leaf-ported, not-yet-wired) binding-emit
    // tail, which still gates -- but with the NEW, more specific message
    // (proving the narrowing actually took effect, not just re-gating the
    // same way).
    let mut a = NodeArena::new();
    let n = ident_node(&mut a, 5, 0);
    let h = ident_heap(0xc0, 4, 0x02, 0); // slot_id defaults to 0 -> has_slot == true
    let r = std::panic::catch_unwind(std::panic::AssertUnwindSafe(|| {
        resolve_ident_ref(&a, n, &h, 0x10, 0, None, RawNode::default())
    }));
    let err = r.unwrap_err();
    let msg = err
        .downcast_ref::<String>()
        .map(String::as_str)
        .or_else(|| err.downcast_ref::<&str>().copied())
        .unwrap_or("");
    assert!(
        msg.contains("0xd/0xe/0xf"),
        "expected the binding-emit-tail gate, got: {msg}"
    );
}

#[test]
fn resolve_ident_ref_category_4_gated() {
    // rec1=3, op byte1=0 → category 4 (gated: EbResolveExprNode).
    let mut a = NodeArena::new();
    let n = ident_node(&mut a, 5, 0);
    let h = ident_heap(0x40, 3, 0x02, 0);
    let r = std::panic::catch_unwind(std::panic::AssertUnwindSafe(|| {
        resolve_ident_ref(&a, n, &h, 0x10, 0, None, RawNode::default())
    }));
    assert!(r.is_err());
}

#[test]
fn resolve_reference2_dispatches_0x60() {
    // A 0x60 node with no member sub-expression routes to resolve_ident_ref.
    let mut a = NodeArena::new();
    let n = ident_node(&mut a, 5, 0); // ref_node builds opcode 0x60
    let h = ident_heap(0x40, 3, 0x02, 1);
    let via_ref2 = resolve_reference2(&a, n, &h, 0x10, 0, None);
    let direct = resolve_ident_ref(&a, n, &h, 0x10, 0, None, RawNode::default());
    assert_eq!(via_ref2, direct);
}

#[test]
fn resolve_reference2_0x69_and_member_subexpr_gated() {
    let mut a = NodeArena::new();
    let h = ident_heap(0x40, 3, 0x02, 1);
    // 0x69 binary-operation setup is gated.
    let mut b = NodeArena::node(0x69, 5, 0, 0, 0, 0);
    b.w[5] = size_desc(&mut a, 4).0;
    let n69 = a.alloc(b);
    let r = std::panic::catch_unwind(std::panic::AssertUnwindSafe(|| {
        resolve_reference2(&a, n69, &h, 0x10, 0, None)
    }));
    assert!(r.is_err());
    // A 0x60 node with a member sub-expression (word[4] != 0) is gated.
    let d = size_desc(&mut a, 4);
    let mut m = NodeArena::node(0x60, 5, 0x99, d.0, 0, 0); // word[4] != 0
    m.w[5] = d.0;
    let nmem = a.alloc(m);
    let r = std::panic::catch_unwind(std::panic::AssertUnwindSafe(|| {
        resolve_reference2(&a, nmem, &h, 0x10, 0, None)
    }));
    assert!(r.is_err());
}

#[test]
fn call_conv_descriptor_selects_record_and_builds() {
    // Exercises every record-selection branch (kind 4..8, and ByRef+kind-4 →
    // special record). All dispatch-record words have bit 0x4000 clear, so the
    // by-reference flag is always set; for a small (size-4) type that yields
    // descriptor kind 1.
    let mut a = NodeArena::new();
    let d = size_desc(&mut a, 4);
    let n = ref_node_with_offset(&mut a, 0, 0, d, NodeRef(0), 0x20);
    for &(kind, byref) in &[(4, 0), (5, 0), (6, 0), (7, 0), (8, 0), (4, 1)] {
        let desc = call_conv_descriptor(&a, n, 6, kind, byref);
        assert_eq!(desc.kind, 1, "kind={kind} byref={byref}");
        assert_eq!(desc.operand, 0x20);
        // optional = true ⇒ no usage flag set.
        assert_eq!(desc.flags1, 0);
    }
}

#[test]
fn call_conv_descriptor_uses_special_record_for_byref_kind4() {
    // The ByRef+kind-4 special record is a distinct table; confirm it is indexed
    // without panicking across the full type-offset range it supports (0..14).
    let mut a = NodeArena::new();
    let d = size_desc(&mut a, 4);
    let n = ref_node(&mut a, 0, 0, d, NodeRef(0));
    for off in 0..15 {
        let desc = call_conv_descriptor(&a, n, off, 4, 1);
        assert_eq!(desc.kind, 1);
    }
}

#[test]
fn category4_resolves_with_binder_binding() {
    // A member record that classifies to category 4 (the call-convention path):
    // record byte+1 low3 = 3, an inline Long operand at +0xc (class-flag 1 →
    // value-class 0, op-byte1 0). With the binder-supplied (kind, byref) the
    // resolver now produces a descriptor instead of gating.
    let mut a = NodeArena::new();
    let d = size_desc(&mut a, 4);
    let n = ref_node_with_offset(&mut a, 8, 0, d, NodeRef(0), 0x20); // type tag 8

    let mut h = vec![0u8; 0x40];
    h[0x10] = 0x40; // +0 bit 6 (inline operand at +0xc)
    h[0x11] = 0x03; // +1 low3 = 3
    h[0x10 + 0xc] = 8; // inline Long operand opcode

    assert_eq!(expression_type2(&h, 0x10).category, 4);

    // Without a binding the category-4 path is still gated.
    let gated = std::panic::catch_unwind(std::panic::AssertUnwindSafe(|| {
        resolve_ident_ref(&a, n, &h, 0x10, 0, None, RawNode::default())
    }));
    assert!(gated.is_err());

    // With the binder-resolved (kind 4, byref 0) it resolves to a descriptor.
    let desc = resolve_ident_ref(&a, n, &h, 0x10, 0, Some((4, 0)), RawNode::default());
    assert_eq!(desc.kind, 1);
    assert_eq!(desc.operand, 0x20);
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

// ── fill_binding_desc (2026-07-14 port) ──────────────────────────────────────

#[test]
fn fill_binding_desc_common_branch_yields_kind4() {
    // binding_flag_5 bit1 clear, and rec0&7 != 2 (the "not a specific slot
    // kind" leg of the OR) -> the common branch: kind 4, word3 = the member
    // record's +8 word.
    let mut heap = vec![0u8; 0x30];
    heap[0x10] = 0x01; // rec0 & 7 == 1, != 2
    heap[0x10 + 8] = 0x34;
    heap[0x10 + 9] = 0x12; // +8 word = 0x1234
    let result = fill_binding_desc(0, &heap, 0x10, 0);
    assert_eq!(result, FillBindingDesc::Kind4 { word3: 0x1234 });
}

#[test]
fn fill_binding_desc_common_branch_via_rec_0x10_bit0() {
    // rec0&7 == 2 (fails the first OR leg) but rec_0x10&1 != 0 (second OR leg
    // true) -> still the common branch.
    let mut heap = vec![0u8; 0x30];
    heap[0x10] = 0x02; // rec0 & 7 == 2
    heap[0x10 + 0x10] = 0x01; // rec_0x10 & 1 != 0
    heap[0x10 + 8] = 0x78;
    heap[0x10 + 9] = 0x56;
    let result = fill_binding_desc(0, &heap, 0x10, 0);
    assert_eq!(result, FillBindingDesc::Kind4 { word3: 0x5678 });
}

#[test]
fn fill_binding_desc_common_branch_via_rec_0x14_not_neg2() {
    // rec0&7==2 and rec_0x10&1==0 (both OR legs so far false) but rec_0x14 !=
    // -2 (third OR leg true) -> still the common branch.
    let mut heap = vec![0u8; 0x30];
    heap[0x10] = 0x02;
    heap[0x10 + 0x14..0x10 + 0x18].copy_from_slice(&0i32.to_le_bytes()); // != -2
    heap[0x10 + 8] = 0x11;
    heap[0x10 + 9] = 0x00;
    let result = fill_binding_desc(0, &heap, 0x10, 0);
    assert_eq!(result, FillBindingDesc::Kind4 { word3: 0x0011 });
}

#[test]
fn fill_binding_desc_binding_flag_forces_slot_table_branch() {
    // binding_flag_5 bit1 set short-circuits the common branch even though
    // the record shape would otherwise qualify -- with ctx_flag_c bit0 set,
    // this lands on the COM-bypass edge case (kind 5, sentinel 0xffff).
    let mut heap = vec![0u8; 0x30];
    heap[0x10] = 0x01; // would qualify for the common branch on its own
    let result = fill_binding_desc(2, &heap, 0x10, 1);
    assert_eq!(result, FillBindingDesc::Kind5Bypass);
}

#[test]
fn fill_binding_desc_slot_kind_with_flags_set_is_bypass() {
    // rec0&7==2 AND rec_0x10&1==0 AND rec_0x14==-2 (all three OR legs false)
    // -> the "is a specific slot kind" condition holds, common branch NOT
    // taken; ctx_flag_c bit0 set bypasses the COM slot-table lookup.
    let mut heap = vec![0u8; 0x30];
    heap[0x10] = 0x02;
    heap[0x10 + 0x14..0x10 + 0x18].copy_from_slice(&(-2i32).to_le_bytes());
    let result = fill_binding_desc(0, &heap, 0x10, 1);
    assert_eq!(result, FillBindingDesc::Kind5Bypass);
}

#[test]
#[should_panic(expected = "EbFillBindingDesc slot-table path")]
fn fill_binding_desc_slot_kind_without_bypass_is_gated() {
    // Same record shape as the bypass test, but ctx_flag_c bit0 clear ->
    // needs the live COM slot table (EbBuildSlotTable) -- must loud-gate.
    let mut heap = vec![0u8; 0x30];
    heap[0x10] = 0x02;
    heap[0x10 + 0x14..0x10 + 0x18].copy_from_slice(&(-2i32).to_le_bytes());
    let _ = fill_binding_desc(0, &heap, 0x10, 0);
}
