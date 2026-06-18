use super::*;

#[test]
fn kind_reads_low_three_bits_unless_0x10_set() {
    let mut m = MemberRecord::new();
    m.set_byte(0x3a, 0x05); // 0x10 clear → low 3 bits
    assert_eq!(m.kind(), 5);
    m.set_byte(0x3a, 0x01);
    assert_eq!(m.kind(), 1);
    m.set_byte(0x3a, 0x15); // 0x10 set → 8 (low bits ignored)
    assert_eq!(m.kind(), 8);
    m.set_byte(0x3a, 0x10);
    assert_eq!(m.kind(), 8);
}

#[test]
fn byref_reads_record1_low_bits_unless_mode_is_6() {
    let mut m = MemberRecord::new();
    m.set_byte(1, 0x01); // low 3 bits = 1
    m.set_byte(0x3d, 0x00); // & 6 != 6
    assert_eq!(m.byref(), 1);
    m.set_byte(1, 0x03);
    assert_eq!(m.byref(), 3);
    m.set_byte(0x3d, 0x06); // & 6 == 6 → 0
    assert_eq!(m.byref(), 0);
    m.set_byte(0x3d, 0x07); // & 6 == 6 → 0
    assert_eq!(m.byref(), 0);
    m.set_byte(0x3d, 0x02); // & 6 == 2 != 6 → low bits
    assert_eq!(m.byref(), 3);
}

#[test]
fn member_id_round_trips_at_0x30() {
    let mut m = MemberRecord::new();
    m.set_member_id(0x1234);
    assert_eq!(m.member_id(), 0x1234);
    assert_eq!(m.byte(0x30), 0x34);
    assert_eq!(m.byte(0x31), 0x12);
}

#[test]
fn pack_type_size_class_sets_high_nibble_preserves_low() {
    for (tc, nibble) in [(1, 0x10u8), (2, 0x20), (4, 0x40), (8, 0x80), (3, 0x00), (0, 0x00)] {
        let mut m = MemberRecord::new();
        m.set_byte(1, 0x07); // a populated low nibble (byref class 7) to preserve
        m.pack_type_size_class(tc);
        assert_eq!(m.byte(1), nibble | 0x07, "type code {tc}");
    }
}

#[test]
fn callee_type_info_resolved_member_reads_record() {
    let mut m = MemberRecord::new();
    m.set_byte(0x3a, 0x04); // kind 4
    m.set_byte(1, 0x02); // byref class 2
    let ti = CalleeTypeInfo::ResolvedMember(&m);
    assert_eq!(ti.kind(), 4);
    assert_eq!(ti.byref(), 2);
}

#[test]
fn callee_type_info_default_is_value_kind_four() {
    let ti = CalleeTypeInfo::Default;
    assert_eq!(ti.kind(), 4);
    assert_eq!(ti.byref(), 0);
}

#[test]
#[should_panic(expected = "FUN_0fbe1daa")]
fn callee_type_info_typelib_kind_is_gated() {
    let _ = CalleeTypeInfo::TypeLib.kind();
}
