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
fn value_class_maps_simple_path_types() {
    assert_eq!(value_class(&VbaType::Integer), Some(6));
    assert_eq!(value_class(&VbaType::Long), Some(8));
    assert_eq!(value_class(&VbaType::Currency), Some(0xc));
}

#[test]
fn value_class_none_for_complex_path_types() {
    // Single/Double/String resolve through the not-yet-ported value-class
    // expression branch, so no class is fabricated here.
    assert_eq!(value_class(&VbaType::Single), None);
    assert_eq!(value_class(&VbaType::Double), None);
    assert_eq!(value_class(&VbaType::String), None);
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
fn bridge_load_unsupported_type_errors() {
    let arena = NodeArena::new();
    let mut e = Emitter::new(&arena);
    assert_eq!(emit_local_load(&mut e, &VbaType::Single, -4), Err(UnsupportedType));
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
