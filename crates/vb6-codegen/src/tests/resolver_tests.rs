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
