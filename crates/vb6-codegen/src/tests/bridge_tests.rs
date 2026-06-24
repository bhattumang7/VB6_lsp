use super::*;
use crate::bind::ProcFrame;
use crate::emit::Emitter;
use crate::node::NodeArena;
use vb6_sema::sema::VbaType;

// ── Type mapping ─────────────────────────────────────────────────────────────

#[test]
fn type_ctx_maps_confirmed_types() {
    assert_eq!(type_ctx(&VbaType::Integer), Some(1));
    assert_eq!(type_ctx(&VbaType::Boolean), Some(1));
    assert_eq!(type_ctx(&VbaType::Byte), Some(1));
    assert_eq!(type_ctx(&VbaType::Long), Some(2));
    assert_eq!(type_ctx(&VbaType::Single), Some(3));
    assert_eq!(type_ctx(&VbaType::Double), Some(4));
    assert_eq!(type_ctx(&VbaType::String), Some(5));
    assert_eq!(type_ctx(&VbaType::Currency), Some(6));
    assert_eq!(type_ctx(&VbaType::Object), Some(0));
}

#[test]
fn type_ctx_none_for_unconfirmed() {
    // Date is now confirmed (Double-backed, class 4).
    assert_eq!(type_ctx(&VbaType::Date), Some(4));
    assert_eq!(type_ctx(&VbaType::Variant), None);
    assert_eq!(type_ctx(&VbaType::Decimal), None);
    assert_eq!(type_ctx(&VbaType::UserDefined(0)), None);
}

#[test]
fn load_store_ctx_maps_numeric_primitives() {
    assert_eq!(load_store_ctx(&VbaType::Integer), Some(1));
    assert_eq!(load_store_ctx(&VbaType::Long), Some(2));
    assert_eq!(load_store_ctx(&VbaType::Single), Some(3));
    assert_eq!(load_store_ctx(&VbaType::Double), Some(4));
    assert_eq!(load_store_ctx(&VbaType::Currency), Some(6));
    // Boolean is operated on / stored as Integer; Date as Double; Byte uses the
    // ctx-7 escape-paged load/store path.
    assert_eq!(load_store_ctx(&VbaType::Boolean), Some(1));
    assert_eq!(load_store_ctx(&VbaType::Date), Some(4));
    assert_eq!(load_store_ctx(&VbaType::Byte), Some(7));
}

#[test]
fn load_store_ctx_none_for_non_simple_types() {
    // String assigns via runtime-helper sequences; the rest are unconfirmed.
    assert_eq!(load_store_ctx(&VbaType::String), None);
    assert_eq!(load_store_ctx(&VbaType::Object), None);
    assert_eq!(load_store_ctx(&VbaType::Variant), None);
}

// ── Bridge emit (type → emit_reference) ──────────────────────────────────────

fn load_bytes(ty: &VbaType, offset: i16) -> Vec<u8> {
    let arena = NodeArena::new();
    let mut e = Emitter::new(&arena);
    emit_local_load(&mut e, ty, offset).unwrap();
    e.into_bytes()
}

fn store_bytes(ty: &VbaType, offset: i16) -> Vec<u8> {
    let arena = NodeArena::new();
    let mut e = Emitter::new(&arena);
    emit_local_store(&mut e, ty, offset).unwrap();
    e.into_bytes()
}

#[test]
fn bridge_load_long() {
    assert_eq!(load_bytes(&VbaType::Long, 0xff78u16 as i16), &[0x6c, 0x78, 0xff]);
}

#[test]
fn bridge_store_long() {
    assert_eq!(store_bytes(&VbaType::Long, 0xff70u16 as i16), &[0x71, 0x70, 0xff]);
}

#[test]
fn bridge_load_integer() {
    assert_eq!(load_bytes(&VbaType::Integer, -4), &[0x6b, 0xfc, 0xff]);
}

#[test]
fn bridge_store_currency() {
    assert_eq!(store_bytes(&VbaType::Currency, -8), &[0x72, 0xf8, 0xff]);
}

#[test]
fn bridge_load_single() {
    // Single local load → 0x6e (RT_LOAD_BY_CTX[3]).
    assert_eq!(load_bytes(&VbaType::Single, 0xff78u16 as i16), &[0x6e, 0x78, 0xff]);
}

#[test]
fn bridge_store_single() {
    // Single store → 0x73 (RT_STORE_BY_CTX[3]).
    assert_eq!(store_bytes(&VbaType::Single, 0xff74u16 as i16), &[0x73, 0x74, 0xff]);
}

#[test]
fn bridge_load_double() {
    // Double local load → 0x6f (RT_LOAD_BY_CTX[4]).
    assert_eq!(load_bytes(&VbaType::Double, 0xff74u16 as i16), &[0x6f, 0x74, 0xff]);
}

#[test]
fn bridge_store_double() {
    // Double store → 0x74 (RT_STORE_BY_CTX[4]).
    assert_eq!(store_bytes(&VbaType::Double, 0xff6cu16 as i16), &[0x74, 0x6c, 0xff]);
}

#[test]
fn bridge_load_unsupported_type_errors() {
    // String assigns via a runtime-helper sequence, not a single load opcode.
    let arena = NodeArena::new();
    let mut e = Emitter::new(&arena);
    assert_eq!(emit_local_load(&mut e, &VbaType::String, -4), Err(UnsupportedType));
}

// ── End-to-end: ProcFrame (VB6-exact offsets) + bridge ──────────────────────
//
// This ties the two halves: a Long declared through ProcFrame gets the
// probe-confirmed frame offset (-136 for the first local), and the bridge maps
// VbaType::Long → value class 8 → opcode 0x6c, producing the oracle load bytes.

#[test]
fn procframe_offset_drives_bridge_load() {
    let mut frame = ProcFrame::new();
    let a = frame.declare_local("a", type_ctx(&VbaType::Long).unwrap()).unwrap();
    assert_eq!(a.frame_offset, -136); // 0xff78
    assert_eq!(load_bytes(&VbaType::Long, a.frame_offset), &[0x6c, 0x78, 0xff]);
}

#[test]
fn procframe_offsets_drive_bridge_store() {
    // Dim a As Long, b As Long, r As Long  →  r at -144 (0xff70).
    let mut frame = ProcFrame::new();
    frame.declare_local("a", 2).unwrap();
    frame.declare_local("b", 2).unwrap();
    let r = frame.declare_local("r", 2).unwrap();
    assert_eq!(r.frame_offset, -144);
    assert_eq!(store_bytes(&VbaType::Long, r.frame_offset), &[0x71, 0x70, 0xff]);
}

// ── Resolution → emit path (binder local_idx → frame slot → opcode) ──────────
//
// Mirrors how a NameResolution::Local{local_idx} from vb6-sema is lowered:
// allocate the proc frame from the declared local types (declaration order),
// then emit by index.

#[test]
fn frame_from_local_types_matches_probe_offsets() {
    // Dim a As Long, b As Long, r As Long
    let types = [VbaType::Long, VbaType::Long, VbaType::Long];
    let slots = frame_from_local_types(&types).unwrap();
    assert_eq!(slots[0].frame_offset, -136); // a
    assert_eq!(slots[1].frame_offset, -140); // b
    assert_eq!(slots[2].frame_offset, -144); // r
}

#[test]
fn frame_from_local_types_unsupported_errors() {
    // A Variant local has no confirmed frame size → refuse the whole layout.
    let types = [VbaType::Long, VbaType::Variant];
    assert_eq!(frame_from_local_types(&types), Err(UnsupportedType));
}

#[test]
fn resolved_local_load_by_index() {
    // Dim a As Long, b As Long, r As Long → load local 1 (b) → 0x6c at -140.
    let types = [VbaType::Long, VbaType::Long, VbaType::Long];
    let slots = frame_from_local_types(&types).unwrap();
    let arena = NodeArena::new();
    let mut e = Emitter::new(&arena);
    emit_resolved_local_load(&mut e, 1, &types, &slots).unwrap();
    assert_eq!(e.into_bytes(), &[0x6c, 0x74, 0xff]); // -140 = 0xff74
}

#[test]
fn resolved_local_store_by_index() {
    // Store into local 2 (r) → 0x71 at -144.
    let types = [VbaType::Long, VbaType::Long, VbaType::Long];
    let slots = frame_from_local_types(&types).unwrap();
    let arena = NodeArena::new();
    let mut e = Emitter::new(&arena);
    emit_resolved_local_store(&mut e, 2, &types, &slots).unwrap();
    assert_eq!(e.into_bytes(), &[0x71, 0x70, 0xff]); // -144 = 0xff70
}

#[test]
fn resolved_local_mixed_type_frame() {
    // Dim n As Integer, x As Double → Integer at -134 (0xff7a), Double aligned
    // to -144 (0xff70). Load both by index; opcodes 0x6b (Integer), 0x6f (Double).
    let types = [VbaType::Integer, VbaType::Double];
    let slots = frame_from_local_types(&types).unwrap();
    assert_eq!(slots[0].frame_offset, -134);
    assert_eq!(slots[1].frame_offset, -144);
    let arena = NodeArena::new();
    let mut e = Emitter::new(&arena);
    emit_resolved_local_load(&mut e, 0, &types, &slots).unwrap();
    emit_resolved_local_load(&mut e, 1, &types, &slots).unwrap();
    assert_eq!(e.into_bytes(), &[0x6b, 0x7a, 0xff, 0x6f, 0x70, 0xff]);
}

// ── ByVal parameter bridge (oracle-verified) ──────────────────────────────────

#[test]
fn byval_param_long_load() {
    // Oracle: `Sub Foo(ByVal p As Long)` → r = p → `6c 0c 00` (p at +12). ✓
    let arena = NodeArena::new();
    let mut e = Emitter::new(&arena);
    emit_byval_param_load(&mut e, &VbaType::Long, 12).unwrap();
    assert_eq!(e.into_bytes(), &[0x6c, 0x0c, 0x00]);
}

#[test]
fn byval_param_long_store() {
    // Oracle: `p = r` (ByVal Long p at +12) → `71 0c 00`. ✓
    let arena = NodeArena::new();
    let mut e = Emitter::new(&arena);
    emit_byval_param_store(&mut e, &VbaType::Long, 12).unwrap();
    assert_eq!(e.into_bytes(), &[0x71, 0x0c, 0x00]);
}

#[test]
fn byval_param_single_load() {
    // Oracle: `Sub Foo(ByVal p As Single)` → `6e 0c 00` (p at +12). ✓
    let arena = NodeArena::new();
    let mut e = Emitter::new(&arena);
    emit_byval_param_load(&mut e, &VbaType::Single, 12).unwrap();
    assert_eq!(e.into_bytes(), &[0x6e, 0x0c, 0x00]);
}

#[test]
fn byval_param_double_load() {
    // Oracle: `Sub Foo(ByVal p As Double)` → `6f 0c 00` (p at +12). ✓
    let arena = NodeArena::new();
    let mut e = Emitter::new(&arena);
    emit_byval_param_load(&mut e, &VbaType::Double, 12).unwrap();
    assert_eq!(e.into_bytes(), &[0x6f, 0x0c, 0x00]);
}

#[test]
fn byval_param_second_long_at_16() {
    // Oracle: `(ByVal p As Long, ByVal q As Long)` → q at +16.
    // Load q: `6c 10 00`. ✓
    let arena = NodeArena::new();
    let mut e = Emitter::new(&arena);
    emit_byval_param_load(&mut e, &VbaType::Long, 16).unwrap();
    assert_eq!(e.into_bytes(), &[0x6c, 0x10, 0x00]);
}

// ── ByRef parameter bridge (oracle-verified) ──────────────────────────────────

#[test]
fn byref_param_long_load() {
    // Oracle: `Sub Foo(ByRef p As Long)` → r = p → `80 0c 00`. ✓
    // 0x80 = RT_LOAD_BY_CTX[Long=2] (0x6c) + 0x14.
    let arena = NodeArena::new();
    let mut e = Emitter::new(&arena);
    emit_byref_param_load(&mut e, &VbaType::Long, 12).unwrap();
    assert_eq!(e.into_bytes(), &[0x80, 0x0c, 0x00]);
}

#[test]
fn byref_param_long_store() {
    // Oracle: `p = r` (ByRef Long p at +12) → `85 0c 00`. ✓
    // 0x85 = RT_STORE_BY_CTX[Long=2] (0x71) + 0x14.
    let arena = NodeArena::new();
    let mut e = Emitter::new(&arena);
    emit_byref_param_store(&mut e, &VbaType::Long, 12).unwrap();
    assert_eq!(e.into_bytes(), &[0x85, 0x0c, 0x00]);
}

// ── Resolved-param path: param_frame_from_types → emit_resolved_param_* ───────

#[test]
fn param_frame_from_types_two_longs() {
    // `Sub Foo(ByVal p As Long, ByVal q As Long)` → p at +12, q at +16.
    let types = [VbaType::Long, VbaType::Long];
    let byref = [false, false];
    let slots = param_frame_from_types(&types, &byref).unwrap();
    assert_eq!(slots[0].frame_offset, 12);
    assert_eq!(slots[1].frame_offset, 16);
}

#[test]
fn emit_resolved_byval_param_load_by_index() {
    // Load the second ByVal Long param (index 1) → `6c 10 00`. ✓
    let types = [VbaType::Long, VbaType::Long];
    let byref = [false, false];
    let slots = param_frame_from_types(&types, &byref).unwrap();
    let arena = NodeArena::new();
    let mut e = Emitter::new(&arena);
    emit_resolved_param_load(&mut e, 1, &types, &slots).unwrap();
    assert_eq!(e.into_bytes(), &[0x6c, 0x10, 0x00]);
}

#[test]
fn emit_resolved_byref_param_load_by_index() {
    // Load a ByRef Long param at index 0 → `80 0c 00`. ✓
    let types = [VbaType::Long];
    let byref = [true];
    let slots = param_frame_from_types(&types, &byref).unwrap();
    let arena = NodeArena::new();
    let mut e = Emitter::new(&arena);
    emit_resolved_param_load(&mut e, 0, &types, &slots).unwrap();
    assert_eq!(e.into_bytes(), &[0x80, 0x0c, 0x00]);
}

// ── Module global bridge (oracle-verified) ────────────────────────────────────

#[test]
fn global_long_load() {
    // Oracle: `Public g As Long; r = g` → `94 08 00 00 00`. ✓
    // Opcode 0x94 = RT_LOAD_BY_CTX[Long=2] (0x6c) + 0x28.
    // module_desc=0x0008, field_offset=0x0000.
    let arena = NodeArena::new();
    let mut e = Emitter::new(&arena);
    emit_global_var_load(&mut e, &VbaType::Long, 0x0008, 0x0000).unwrap();
    assert_eq!(e.into_bytes(), &[0x94, 0x08, 0x00, 0x00, 0x00]);
}

#[test]
fn global_long_store() {
    // Oracle: `g = r` → `99 08 00 00 00`. ✓
    // Opcode 0x99 = RT_STORE_BY_CTX[Long=2] (0x71) + 0x28.
    let arena = NodeArena::new();
    let mut e = Emitter::new(&arena);
    emit_global_var_store(&mut e, &VbaType::Long, 0x0008, 0x0000).unwrap();
    assert_eq!(e.into_bytes(), &[0x99, 0x08, 0x00, 0x00, 0x00]);
}

#[test]
fn global_second_long_field_offset_4() {
    // Oracle: second `Public b As Long` (after a) → `94 08 00 04 00`. ✓
    // field_offset=0x0004 because a occupies the first 4 bytes.
    let arena = NodeArena::new();
    let mut e = Emitter::new(&arena);
    emit_global_var_load(&mut e, &VbaType::Long, 0x0008, 0x0004).unwrap();
    assert_eq!(e.into_bytes(), &[0x94, 0x08, 0x00, 0x04, 0x00]);
}

#[test]
fn global_integer_load() {
    // Oracle: `Public g As Integer; r = g` → `93 08 00 00 00`. ✓
    // Opcode 0x93 = RT_LOAD_BY_CTX[Integer=1] (0x6b) + 0x28.
    let arena = NodeArena::new();
    let mut e = Emitter::new(&arena);
    emit_global_var_load(&mut e, &VbaType::Integer, 0x0008, 0x0000).unwrap();
    assert_eq!(e.into_bytes(), &[0x93, 0x08, 0x00, 0x00, 0x00]);
}

#[test]
fn global_integer_store() {
    // Oracle: `g = r` (Integer) → `98 08 00 00 00`. ✓
    let arena = NodeArena::new();
    let mut e = Emitter::new(&arena);
    emit_global_var_store(&mut e, &VbaType::Integer, 0x0008, 0x0000).unwrap();
    assert_eq!(e.into_bytes(), &[0x98, 0x08, 0x00, 0x00, 0x00]);
}

#[test]
fn global_double_load() {
    // Oracle: `Public g As Double; r = g` → `97 08 00 00 00`. ✓
    // Opcode 0x97 = RT_LOAD_BY_CTX[Double=4] (0x6f) + 0x28.
    let arena = NodeArena::new();
    let mut e = Emitter::new(&arena);
    emit_global_var_load(&mut e, &VbaType::Double, 0x0008, 0x0000).unwrap();
    assert_eq!(e.into_bytes(), &[0x97, 0x08, 0x00, 0x00, 0x00]);
}

#[test]
fn global_double_store() {
    // Oracle: `g = r` (Double) → `9c 08 00 00 00`. ✓
    // Opcode 0x9c = RT_STORE_BY_CTX[Double=4] (0x74) + 0x28.
    let arena = NodeArena::new();
    let mut e = Emitter::new(&arena);
    emit_global_var_store(&mut e, &VbaType::Double, 0x0008, 0x0000).unwrap();
    assert_eq!(e.into_bytes(), &[0x9c, 0x08, 0x00, 0x00, 0x00]);
}
