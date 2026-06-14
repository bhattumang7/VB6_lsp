use vb6_core::frx::{FrxError, FrxReader};

// --- seek -------------------------------------------------------------------

#[test]
fn seek_to_zero_on_new_reader() {
    let r = FrxReader::new(&[0x01, 0x02]);
    assert_eq!(r.pos(), 0);
    assert_eq!(r.remaining(), 2);
}

#[test]
fn seek_to_valid_offset() {
    let mut r = FrxReader::new(&[0x00; 8]);
    r.seek(4).unwrap();
    assert_eq!(r.pos(), 4);
    assert_eq!(r.remaining(), 4);
}

#[test]
fn seek_to_end_is_ok() {
    let data = [0u8; 4];
    let mut r = FrxReader::new(&data);
    r.seek(4).unwrap();
    assert_eq!(r.remaining(), 0);
}

#[test]
fn seek_past_end_is_err() {
    let data = [0u8; 4];
    let mut r = FrxReader::new(&data);
    assert_eq!(r.seek(5), Err(FrxError::SeekOutOfRange { target: 5, len: 4 }));
    assert_eq!(r.pos(), 0, "cursor must be unchanged after a failed seek");
}

// --- read_u8 ----------------------------------------------------------------

#[test]
fn read_u8_basic() {
    let mut r = FrxReader::new(&[0xAB, 0xCD]);
    assert_eq!(r.read_u8().unwrap(), 0xAB);
    assert_eq!(r.read_u8().unwrap(), 0xCD);
    assert_eq!(r.remaining(), 0);
}

#[test]
fn read_u8_eof() {
    let mut r = FrxReader::new(&[]);
    assert_eq!(
        r.read_u8(),
        Err(FrxError::UnexpectedEof { pos: 0, needed: 1, available: 0 })
    );
}

// --- read_u16_le ------------------------------------------------------------

#[test]
fn read_u16_le_basic() {
    let mut r = FrxReader::new(&[0x34, 0x12]);
    assert_eq!(r.read_u16_le().unwrap(), 0x1234);
}

#[test]
fn read_u16_le_eof_partial() {
    let mut r = FrxReader::new(&[0x01]);
    let e = r.read_u16_le().unwrap_err();
    assert_eq!(e, FrxError::UnexpectedEof { pos: 0, needed: 2, available: 1 });
    assert_eq!(r.pos(), 0, "cursor must not advance on EOF");
}

// --- read_u32_le ------------------------------------------------------------

#[test]
fn read_u32_le_basic() {
    let mut r = FrxReader::new(&[0x78, 0x56, 0x34, 0x12]);
    assert_eq!(r.read_u32_le().unwrap(), 0x12345678);
}

#[test]
fn read_u32_le_eof() {
    let mut r = FrxReader::new(&[0x01, 0x02, 0x03]);
    assert!(r.read_u32_le().is_err());
}

// --- read_bytes -------------------------------------------------------------

#[test]
fn read_bytes_exact() {
    let data = [1u8, 2, 3, 4, 5];
    let mut r = FrxReader::new(&data);
    assert_eq!(r.read_bytes(3).unwrap(), &[1, 2, 3]);
    assert_eq!(r.pos(), 3);
}

#[test]
fn read_zero_bytes_ok() {
    let mut r = FrxReader::new(&[]);
    assert_eq!(r.read_bytes(0).unwrap(), &[] as &[u8]);
}

// --- peek_u16_le ------------------------------------------------------------

#[test]
fn peek_u16_le_does_not_advance() {
    let mut r = FrxReader::new(&[0x02, 0x01]);
    let v = r.peek_u16_le().unwrap();
    assert_eq!(v, 0x0102);
    assert_eq!(r.pos(), 0, "peek must not advance cursor");
    assert_eq!(r.read_u16_le().unwrap(), 0x0102, "read after peek must return same value");
}

// --- read_len_prefixed_bytes ------------------------------------------------

#[test]
fn len_prefixed_zero_bytes() {
    let mut r = FrxReader::new(&[0x00, 0x00, 0x00, 0x00]);
    assert_eq!(r.read_len_prefixed_bytes().unwrap(), &[] as &[u8]);
}

#[test]
fn len_prefixed_three_bytes() {
    let mut r = FrxReader::new(&[0x03, 0x00, 0x00, 0x00, 0xAA, 0xBB, 0xCC]);
    assert_eq!(r.read_len_prefixed_bytes().unwrap(), &[0xAA, 0xBB, 0xCC]);
    assert_eq!(r.remaining(), 0);
}

#[test]
fn len_prefixed_overflow() {
    // declares 10 bytes but only 2 remain after the length field
    let mut r = FrxReader::new(&[0x0A, 0x00, 0x00, 0x00, 0x01, 0x02]);
    assert!(matches!(r.read_len_prefixed_bytes(), Err(FrxError::LengthOverflow { .. })));
}

// --- require_magic ----------------------------------------------------------

#[test]
fn require_magic_ok() {
    let mut r = FrxReader::new(&[0x6C, 0x74, 0xFF]);
    r.require_magic(0x746C).unwrap();
    assert_eq!(r.pos(), 2);
}

#[test]
fn require_magic_mismatch_rewinds() {
    let mut r = FrxReader::new(&[0x01, 0x02]);
    let e = r.require_magic(0xDEAD).unwrap_err();
    assert_eq!(e, FrxError::BadMagic { pos: 0, expected: 0xDEAD, got: 0x0201 });
    assert_eq!(r.pos(), 0, "cursor must rewind after bad magic");
}

// --- sequential reads -------------------------------------------------------

#[test]
fn sequential_reads_advance_correctly() {
    // 1-byte, 2-byte, 4-byte read in sequence
    let data = [0x42u8, 0x34, 0x12, 0x78, 0x56, 0x34, 0x12];
    let mut r = FrxReader::new(&data);
    assert_eq!(r.read_u8().unwrap(), 0x42);
    assert_eq!(r.read_u16_le().unwrap(), 0x1234);
    assert_eq!(r.read_u32_le().unwrap(), 0x12345678);
    assert_eq!(r.remaining(), 0);
}

// --- seek + read round-trip -------------------------------------------------

#[test]
fn seek_then_read_at_arbitrary_offset() {
    let data: Vec<u8> = (0u8..=9).collect();
    let mut r = FrxReader::new(&data);
    r.seek(5).unwrap();
    assert_eq!(r.read_u8().unwrap(), 5);
    assert_eq!(r.read_u8().unwrap(), 6);
    r.seek(0).unwrap();
    assert_eq!(r.read_u8().unwrap(), 0);
}
