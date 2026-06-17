use super::*;

#[test]
fn emit_word_is_little_endian_and_advances() {
    let mut s = PcodeStream::new();
    let p = s.emit_word(0x1234);
    assert_eq!(p, BytePos(0));
    assert_eq!(s.emit_word(0x00ff), BytePos(2));
    assert_eq!(s.bytes(), &[0x34, 0x12, 0xff, 0x00]);
}

#[test]
fn emit_byte_advances_by_one() {
    let mut s = PcodeStream::new();
    let p0 = s.emit_byte(0xab);
    assert_eq!(p0, BytePos(0));
    let p1 = s.emit_byte(0xcd);
    assert_eq!(p1, BytePos(1));
    assert_eq!(s.bytes(), &[0xab, 0xcd]);
}

#[test]
fn emit_i16_writes_signed_le() {
    let mut s = PcodeStream::new();
    s.emit_i16(-140); // 0xff74 as i16
    assert_eq!(s.bytes(), &[0x74, 0xff]);
}

#[test]
fn emit_load_store_writes_three_bytes() {
    // Double-load opcode 0x6f at frame offset 0xff74 (-140 as i16).
    let mut s = PcodeStream::new();
    s.emit_load_store(0x6f, -140);
    assert_eq!(s.bytes(), &[0x6f, 0x74, 0xff]);
}

#[test]
fn emit_word4_writes_two_words_in_order() {
    let mut s = PcodeStream::new();
    s.emit_word4(0x1234, 0x5678);
    assert_eq!(s.bytes(), &[0x34, 0x12, 0x78, 0x56]);
}

#[test]
fn emit_pcode3_writes_opcode_then_two_operands() {
    let mut s = PcodeStream::new();
    s.emit_pcode3(0x00e0, 0x0001, 0x0002);
    assert_eq!(s.bytes(), &[0xe0, 0x00, 0x01, 0x00, 0x02, 0x00]);
}

#[test]
fn emit_literal8_writes_opcode_then_eight_bytes() {
    let mut s = PcodeStream::new();
    let payload = 10_000_i64.to_le_bytes();
    s.emit_literal8(0x00a9, payload);
    let mut expect = vec![0xa9, 0x00];
    expect.extend_from_slice(&payload);
    assert_eq!(s.bytes(), expect.as_slice());
}

#[test]
fn emit_word_and_data_even_length_string() {
    let mut s = PcodeStream::new();
    s.emit_word_and_data(0x00b6, 2, b"AB");
    assert_eq!(s.bytes(), &[0xb6, 0x00, 0x02, 0x00, 0x41, 0x42]);
}

#[test]
fn emit_word_and_data_records_logical_len_but_copies_rounded() {
    let mut s = PcodeStream::new();
    s.emit_word_and_data(0x00b6, 3, b"ABC\0");
    assert_eq!(s.bytes(), &[0xb6, 0x00, 0x03, 0x00, 0x41, 0x42, 0x43, 0x00]);
}

#[test]
fn patch_word_backpatches_in_place() {
    let mut s = PcodeStream::new();
    s.emit_word(0x00e0);
    let target = s.emit_word(0xffff);
    s.emit_word(0x0001);
    s.patch_word(target, 0x0042);
    assert_eq!(s.bytes(), &[0xe0, 0x00, 0x42, 0x00, 0x01, 0x00]);
}

#[test]
fn byte_and_word_interleave_without_alignment_constraint() {
    // Runtime stream mixes 1-byte opcodes and 2-byte operands freely.
    let mut s = PcodeStream::new();
    s.emit_byte(0xf4); // push-small-int opcode
    s.emit_byte(0x02); // immediate 2
    s.emit_byte(0xeb); // coerce opcode
    s.emit_load_store(0x74, -140i16); // store-Double at 0xff74
    assert_eq!(s.bytes(), &[0xf4, 0x02, 0xeb, 0x74, 0x74, 0xff]);
}
