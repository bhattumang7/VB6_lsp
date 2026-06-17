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
    assert_eq!(type_ctx(&VbaType::Date), None);
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
}

#[test]
fn load_store_ctx_none_for_non_simple_types() {
    // String/Byte assign via runtime-helper sequences; the rest are unconfirmed.
    assert_eq!(load_store_ctx(&VbaType::String), None);
    assert_eq!(load_store_ctx(&VbaType::Byte), None);
    assert_eq!(load_store_ctx(&VbaType::Date), None);
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
