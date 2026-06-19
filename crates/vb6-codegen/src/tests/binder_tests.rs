//! Tests for the binder front-half ([`crate::binder`]).
//!
//! Bind structures are hand-laid in a byte buffer; values are derived from the
//! decompiled `EbResolveExprNode`/`EbSetupResult`/`EbMapTypeToCallConv` logic.

use crate::binder::{
    get_bind_result, handle_byref, handle_kind, map_type_to_call_conv, resolve_expr_node,
    setup_result, BindHandle,
};
use crate::sym_record::MemberRecord;

fn put_u32(m: &mut [u8], off: usize, v: u32) {
    m[off..off + 4].copy_from_slice(&v.to_le_bytes());
}

#[test]
fn map_type_to_call_conv_maps_low3_bits() {
    // symbol @0x40, type-desc ptr at +0x20 → 0x80, flags word at type_desc+0x12.
    let mut m = vec![0u8; 0x100];
    put_u32(&mut m, 0x40 + 0x20, 0x80);
    for (bits, want) in [(0u32, 4), (1, 1), (2, 2), (3, 6), (4, 4), (7, 4)] {
        put_u32(&mut m, 0x80 + 0x12, bits);
        assert_eq!(map_type_to_call_conv(&m, 0x40), want, "bits={bits}");
    }
}

#[test]
fn setup_result_builds_disc1_handle() {
    let mut m = vec![0u8; 0x100];
    // symbol @0x40 with node-type nibble 8 (not 9).
    m[0x40 + 8] = 0x08;
    let h = setup_result(&m, 0x40, 0x1234, 0x5678);
    assert_eq!(
        h,
        BindHandle { disc: 1, w3: 0x40, w4: 0x1234, w5: 0x5678, w1: 0, w2: 0 }
    );
}

#[test]
fn setup_result_type9_suppresses_aux_slots() {
    let mut m = vec![0u8; 0x100];
    m[0x40 + 8] = 0x09; // node-type nibble 9
    let h = setup_result(&m, 0x40, 0x1234, 0x5678);
    assert_eq!(h.disc, 1);
    assert_eq!(h.w3, 0x40);
    assert_eq!(h.w4, 0);
    assert_eq!(h.w5, 0);
}

#[test]
fn setup_result_null_node_clears() {
    let m = vec![0u8; 0x100];
    assert_eq!(setup_result(&m, 0, 1, 2), BindHandle::default());
}

#[test]
fn get_bind_result_requires_nonzero_disc() {
    assert_eq!(get_bind_result(BindHandle::default()), None);
    let h = BindHandle { disc: 1, w3: 0x40, ..BindHandle::default() };
    assert_eq!(get_bind_result(h), Some(h));
}

#[test]
fn handle_kind_disc1_uses_map() {
    let mut m = vec![0u8; 0x100];
    put_u32(&mut m, 0x40 + 0x20, 0x80);
    put_u32(&mut m, 0x80 + 0x12, 3); // → 6
    let h = BindHandle { disc: 1, w3: 0x40, ..BindHandle::default() };
    assert_eq!(handle_kind(&h, &m, None), 6);
    assert_eq!(handle_byref(&h, None), 0);
}

#[test]
fn handle_kind_disc5_reads_member_record() {
    let mut rec = MemberRecord::new();
    rec.set_byte(0x3a, 0x10); // bit 0x10 → kind 8
    rec.set_byte(0x3d, 0); // &6 != 6 → byref reads +1
    rec.set_byte(1, 0x02); // low3 = 2
    let h = BindHandle { disc: 5, ..BindHandle::default() };
    let m = vec![0u8; 4];
    assert_eq!(handle_kind(&h, &m, Some(&rec)), 8);
    assert_eq!(handle_byref(&h, Some(&rec)), 2);
}

#[test]
fn handle_kind_other_disc_is_default() {
    let h = BindHandle { disc: 3, ..BindHandle::default() };
    let m = vec![0u8; 4];
    assert_eq!(handle_kind(&h, &m, None), 4);
    assert_eq!(handle_byref(&h, None), 0);
}

#[test]
fn resolve_expr_node_disc1_name_reference() {
    // context @0x10: [0]=1 (disc), [3]@0x1c = 0x30 (ptr-to-symbol-ptr),
    // [4]@0x20 = p4. *0x30 = 0x40 (symbol). symbol+8 nibble = 8.
    let mut m = vec![0u8; 0x100];
    put_u32(&mut m, 0x10, 1); // disc
    put_u32(&mut m, 0x10 + 12, 0x30); // pNode[3]
    put_u32(&mut m, 0x10 + 16, 0xabcd); // pNode[4]
    put_u32(&mut m, 0x30, 0x40); // *pNode[3] = symbol
    m[0x40 + 8] = 0x08; // symbol node-type nibble
    let h = resolve_expr_node(&m, 0x10).unwrap();
    assert_eq!(h.disc, 1);
    assert_eq!(h.w3, 0x40);
    assert_eq!(h.w4, 0xabcd);
}

#[test]
fn resolve_expr_node_disc1_node_type_10_is_empty() {
    let mut m = vec![0u8; 0x100];
    put_u32(&mut m, 0x10, 1);
    put_u32(&mut m, 0x10 + 12, 0x30);
    put_u32(&mut m, 0x30, 0x40);
    m[0x40 + 8] = 0x0a; // node-type nibble 10 → empty
    assert_eq!(resolve_expr_node(&m, 0x10), None);
}

#[test]
fn resolve_expr_node_disc2_is_empty() {
    let mut m = vec![0u8; 0x40];
    put_u32(&mut m, 0x10, 2);
    assert_eq!(resolve_expr_node(&m, 0x10), None);
}

#[test]
fn resolve_expr_node_disc3_initializes_context() {
    // disc 3: handle is disc-2 carrying *(*(pNode[1]+0x24)+0x5c).
    let mut m = vec![0u8; 0x100];
    put_u32(&mut m, 0x10, 3); // disc
    put_u32(&mut m, 0x10 + 4, 0x30); // pNode[1]
    put_u32(&mut m, 0x30 + 0x24, 0x60); // *(pNode[1]+0x24)
    put_u32(&mut m, 0x60 + 0x5c, 0xdead); // +0x5c init value
    let h = resolve_expr_node(&m, 0x10).unwrap();
    assert_eq!(h.disc, 2);
    assert_eq!(h.w3, 0xdead);
    // disc 2/3 → plain-value convention kind 4.
    assert_eq!(handle_kind(&h, &m, None), 4);
}

#[test]
fn resolve_expr_node_disc5_and_6_use_context() {
    let mut m = vec![0u8; 0x40];
    for disc in [5u32, 6] {
        put_u32(&mut m, 0x10, disc);
        put_u32(&mut m, 0x10 + 4, 0xc0de); // pNode[1] = context
        let h = resolve_expr_node(&m, 0x10).unwrap();
        assert_eq!(h.disc, 3, "disc {disc}");
        assert_eq!(h.w1, 0xc0de);
        assert_eq!(handle_kind(&h, &m, None), 4);
    }
}

#[test]
fn resolve_expr_node_disc4_context_path() {
    // disc 4 with pNode[4] == -1 and pNode[5] == -1 → context handle.
    let mut m = vec![0u8; 0x40];
    put_u32(&mut m, 0x10, 4);
    put_u32(&mut m, 0x10 + 4, 0xbeef); // pNode[1]
    put_u32(&mut m, 0x10 + 16, 0xffff_ffff); // pNode[4] = -1
    put_u32(&mut m, 0x10 + 20, 0xffff_ffff); // pNode[5] = -1
    let h = resolve_expr_node(&m, 0x10).unwrap();
    assert_eq!(h.disc, 3);
    assert_eq!(h.w1, 0xbeef);
}

#[test]
fn resolve_expr_node_disc4_exception_path_is_empty() {
    let mut m = vec![0u8; 0x40];
    put_u32(&mut m, 0x10, 4);
    put_u32(&mut m, 0x10 + 20, 0x1234); // pNode[5] != -1 → error path
    assert_eq!(resolve_expr_node(&m, 0x10), None);
}

#[test]
fn resolve_expr_node_disc4_slot_path_is_gated() {
    let mut m = vec![0u8; 0x40];
    put_u32(&mut m, 0x10, 4);
    put_u32(&mut m, 0x10 + 16, 0x40); // pNode[4] != -1 → COM slot
    put_u32(&mut m, 0x10 + 20, 0xffff_ffff); // pNode[5] == -1
    let r = std::panic::catch_unwind(std::panic::AssertUnwindSafe(|| {
        resolve_expr_node(&m, 0x10)
    }));
    assert!(r.is_err());
}
