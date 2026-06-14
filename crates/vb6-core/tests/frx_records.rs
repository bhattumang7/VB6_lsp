use vb6_core::frx::{FrxError, FrxReader, FrxRecord, RecordKind};
use vb6_core::frx::records::{
    Clsid, ControlCreateRecord, ControlPersistRecord, ItemDataRecord, LenPrefixedBytes,
    ListRecord, OcxBagRecord, OleStorageRecord, PictureRecord, PropertyBagRecord,
    PropertyPagesRecord, StdFontRecord, StringShortBytes,
};
use vb6_core::frx::records::ole_storage::PersistMechanism;

// ============================================================================
// helpers
// ============================================================================

fn reader(data: &[u8]) -> FrxReader<'_> {
    FrxReader::new(data)
}

fn len4(data: &[u8]) -> Vec<u8> {
    let mut v = (data.len() as u32).to_le_bytes().to_vec();
    v.extend_from_slice(data);
    v
}

// ============================================================================
// CLSID
// ============================================================================

#[test]
fn clsid_round_trip() {
    let clsid = Clsid {
        data1: 0x12345678, data2: 0xABCD, data3: 0xEF01,
        data4: [1, 2, 3, 4, 5, 6, 7, 8],
    };
    let mut buf = Vec::new();
    clsid.write(&mut buf);
    assert_eq!(buf.len(), 16);
    let mut r = reader(&buf);
    let parsed = Clsid::read(&mut r).unwrap();
    assert_eq!(parsed, clsid);
}

#[test]
fn clsid_braced_string() {
    let clsid = Clsid::IID_IPICTURE;
    let s = clsid.to_braced_string();
    assert!(s.starts_with('{'));
    assert!(s.ends_with('}'));
    assert_eq!(s.len(), 38);
}

#[test]
fn clsid_eof() {
    let mut r = reader(&[0u8; 15]);
    assert!(matches!(Clsid::read(&mut r), Err(FrxError::UnexpectedEof { .. })));
}

// ============================================================================
// StringShort (1-byte prefix, no $ in .frm)
// ============================================================================

#[test]
fn string_short_round_trip() {
    let payload = b"Hello";
    let mut buf = vec![payload.len() as u8];
    buf.extend_from_slice(payload);
    let mut r = reader(&buf);
    let s = StringShortBytes::read(&mut r).unwrap();
    assert_eq!(s.data, payload.as_ref());
    let mut out = Vec::new();
    StringShortBytes::write(s.data, &mut out);
    assert_eq!(out, buf);
}

#[test]
fn string_short_empty() {
    let buf = [0u8]; // len = 0
    let mut r = reader(&buf);
    let s = StringShortBytes::read(&mut r).unwrap();
    assert_eq!(s.data.len(), 0);
}

#[test]
fn string_short_via_frxrecord() {
    let payload = b"Caption";
    let mut buf = vec![payload.len() as u8];
    buf.extend_from_slice(payload);
    let mut r = reader(&buf);
    let rec = FrxRecord::read(RecordKind::StringShort, &mut r).unwrap();
    assert!(matches!(rec, FrxRecord::StringShort(_)));
}

// ============================================================================
// BinaryString / LenPrefixedBytes (4-byte prefix, $ form)
// ============================================================================

#[test]
fn binary_string_round_trip() {
    let payload = b"Hello, VB6 World!";
    let buf = len4(payload);
    let mut r = reader(&buf);
    let rec = FrxRecord::read(RecordKind::BinaryString, &mut r).unwrap();
    match rec {
        FrxRecord::BinaryString(s) => assert_eq!(s.data, payload.as_ref()),
        _ => panic!("wrong variant"),
    }
    let mut out = Vec::new();
    FrxRecord::BinaryString(LenPrefixedBytes { data: payload.as_ref() }).write(&mut out);
    assert_eq!(out, buf);
}

#[test]
fn binary_string_empty() {
    let buf = [0u8; 4];
    let mut r = reader(&buf);
    let rec = FrxRecord::read(RecordKind::BinaryString, &mut r).unwrap();
    match rec {
        FrxRecord::BinaryString(s) => assert_eq!(s.data.len(), 0),
        _ => panic!(),
    }
}

// ============================================================================
// StdFont
//
// Wire layout: [u8 version=1][u16 charset][u8 flags][u16 weight][u32 size][u8 name_len][name]
// Total header = 11 bytes before name.
// ============================================================================

fn font_bytes(charset: u16, flags: u8, weight: u16, size: u32, name: &[u8]) -> Vec<u8> {
    let mut v = vec![StdFontRecord::VERSION];
    v.extend_from_slice(&charset.to_le_bytes()); // u16 charset — 2 bytes
    v.push(flags);
    v.extend_from_slice(&weight.to_le_bytes());
    v.extend_from_slice(&size.to_le_bytes());
    v.push(name.len() as u8);
    v.extend_from_slice(name);
    v
}

#[test]
fn font_normal() {
    let bytes = font_bytes(0, 0, 400, 82500, b"Arial");
    let mut r = reader(&bytes);
    let f = StdFontRecord::read(&mut r).unwrap();
    assert_eq!(f.charset, 0);
    assert!(!f.is_bold());
    assert!(!f.is_italic());
    assert_eq!(f.weight, 400);
    assert_eq!(f.size_times_10k, 82500);
    assert_eq!(f.name, b"Arial");
}

#[test]
fn font_bold_italic() {
    let bytes = font_bytes(0, 0b0001, 700, 120000, b"Times New Roman");
    let mut r = reader(&bytes);
    let f = StdFontRecord::read(&mut r).unwrap();
    assert!(f.is_bold());
    assert!(f.is_italic());
    assert!(!f.is_underline());
}

#[test]
fn font_round_trip() {
    let bytes = font_bytes(34, 0b0110, 700, 100000, b"Courier New");
    let mut r = reader(&bytes);
    let f = StdFontRecord::read(&mut r).unwrap();
    assert_eq!(f.charset, 34);
    let mut out = Vec::new();
    f.write(&mut out);
    assert_eq!(out, bytes);
}

#[test]
fn font_bad_version() {
    // version = 2 (not 1) — must fail with BadMagic
    let mut bytes = font_bytes(0, 0, 400, 82500, b"Taho");
    bytes[0] = 0x02;
    let mut r = reader(&bytes);
    assert!(matches!(StdFontRecord::read(&mut r), Err(FrxError::BadMagic { .. })));
}

#[test]
fn font_via_frxrecord() {
    let bytes = font_bytes(0, 0, 400, 82500, b"MS Sans Serif");
    let mut r = reader(&bytes);
    let rec = FrxRecord::read(RecordKind::Font, &mut r).unwrap();
    assert!(matches!(rec, FrxRecord::Font(_)));
}

#[test]
fn font_strikethrough() {
    let bytes = font_bytes(0, 0b0100, 400, 100000, b"Arial");
    let mut r = reader(&bytes);
    let f = StdFontRecord::read(&mut r).unwrap();
    assert!(f.is_strikethrough());
    assert!(!f.is_italic());
    assert!(!f.is_underline());
}

#[test]
fn font_all_flags() {
    let bytes = font_bytes(0, 0b0111, 400, 100000, b"Arial");
    let mut r = reader(&bytes);
    let f = StdFontRecord::read(&mut r).unwrap();
    assert!(f.is_italic());
    assert!(f.is_underline());
    assert!(f.is_strikethrough());
}

#[test]
fn font_empty_name() {
    let bytes = font_bytes(0, 0, 400, 82500, b"");
    let mut r = reader(&bytes);
    let f = StdFontRecord::read(&mut r).unwrap();
    assert_eq!(f.name.len(), 0);
    let mut out = Vec::new();
    f.write(&mut out);
    assert_eq!(out, bytes);
}

// Verify against the bytes of a StdFont stream:
// 'MS Sans Serif' 8.25pt normal — charset=0 (u16), flags=0, weight=400, size=82500
#[test]
fn font_test_bytes() {
    let test_bytes: Vec<u8> = vec![
        0x01,                   // version = 1
        0x00, 0x00,             // charset = 0 (u16 LE)
        0x00,                   // flags = 0
        0x90, 0x01,             // weight = 400 (0x0190)
        0x44, 0x42, 0x01, 0x00, // size = 82500 (0x14244)
        0x0D,                   // name_len = 13
        b'M', b'S', b' ', b'S', b'a', b'n', b's', b' ', b'S', b'e', b'r', b'i', b'f',
    ];
    let mut r = reader(&test_bytes);
    let f = StdFontRecord::read(&mut r).unwrap();
    assert_eq!(f.charset, 0u16);
    assert_eq!(f.weight, 400);
    assert_eq!(f.size_times_10k, 82500);
    assert_eq!(f.name, b"MS Sans Serif");
    assert!(!f.is_bold());
    // Round-trip must produce the exact input bytes.
    let mut out = Vec::new();
    f.write(&mut out);
    assert_eq!(out, test_bytes);
}

// ============================================================================
// PictureRecord
//
// Wire layout: [u32 outer]["lt\0\0"][u32 dataLen][image bytes]
// CLSID variant: [u32 outer][16-byte CLSID]["lt\0\0"][u32 dataLen][image bytes]
// Empty slot:    [u32=8]["lt\0\0"][u32=0]
// ============================================================================

fn picture_bytes(data: &[u8]) -> Vec<u8> {
    // Standard (no CLSID): outer = 8 + data_len
    let mut v = ((data.len() as u32) + 8).to_le_bytes().to_vec();
    v.extend_from_slice(b"lt\0\0");
    v.extend_from_slice(&(data.len() as u32).to_le_bytes());
    v.extend_from_slice(data);
    v
}

fn picture_bytes_with_clsid(clsid: &[u8; 16], data: &[u8]) -> Vec<u8> {
    // CLSID variant: outer = 24 + data_len
    let mut v = ((data.len() as u32) + 24).to_le_bytes().to_vec();
    v.extend_from_slice(clsid);
    v.extend_from_slice(b"lt\0\0");
    v.extend_from_slice(&(data.len() as u32).to_le_bytes());
    v.extend_from_slice(data);
    v
}

fn picture_empty_bytes() -> Vec<u8> {
    // [u32=8]["lt\0\0"][u32=0]
    let mut v = 8u32.to_le_bytes().to_vec();
    v.extend_from_slice(b"lt\0\0");
    v.extend_from_slice(&0u32.to_le_bytes());
    v
}

#[test]
fn picture_none() {
    // Empty slot: data_len=0, no CLSID
    let bytes = picture_empty_bytes();
    let mut r = reader(&bytes);
    let p = PictureRecord::read(&mut r).unwrap();
    assert!(p.data.is_empty());
    assert!(p.clsid.is_none());
    let mut out = Vec::new();
    p.write(&mut out);
    assert_eq!(out, bytes);
}

#[test]
fn picture_bitmap_round_trip() {
    let fake_bmp = [0x42u8, 0x4D, 0xAA, 0xBB, 0xCC]; // "BM" prefix
    let bytes = picture_bytes(&fake_bmp);
    let mut r = reader(&bytes);
    let p = PictureRecord::read(&mut r).unwrap();
    assert!(p.clsid.is_none());
    assert_eq!(p.data, &fake_bmp);
    let mut out = Vec::new();
    p.write(&mut out);
    assert_eq!(out, bytes);
}

#[test]
fn picture_icon() {
    let fake_ico = [0x00u8, 0x00, 0x01, 0x00]; // ICO magic
    let bytes = picture_bytes(&fake_ico);
    let mut r = reader(&bytes);
    let p = PictureRecord::read(&mut r).unwrap();
    assert_eq!(p.data, &fake_ico);
}

#[test]
fn picture_metafile() {
    let fake_wmf = [0xD7u8, 0xCD, 0xC6, 0x9A, 0x00, 0x00];
    let bytes = picture_bytes(&fake_wmf);
    let mut r = reader(&bytes);
    let p = PictureRecord::read(&mut r).unwrap();
    assert_eq!(p.data, &fake_wmf);
    let mut out = Vec::new();
    p.write(&mut out);
    assert_eq!(out, bytes);
}

#[test]
fn picture_enh_metafile() {
    let fake_emf = [0x01u8, 0x00, 0x00, 0x00, 0x58, 0x00, 0x00, 0x00];
    let bytes = picture_bytes(&fake_emf);
    let mut r = reader(&bytes);
    let p = PictureRecord::read(&mut r).unwrap();
    assert_eq!(p.data, &fake_emf);
    let mut out = Vec::new();
    p.write(&mut out);
    assert_eq!(out, bytes);
}

#[test]
fn picture_with_clsid() {
    // ImageList / collection-bag form: CLSID precedes "lt\0\0"
    let clsid: [u8; 16] = [
        0x04, 0x52, 0xE3, 0x0B, 0x91, 0x8F, 0xCE, 0x11,
        0x9D, 0xE3, 0x00, 0xAA, 0x00, 0x4B, 0xB8, 0x51,
    ];
    let fake_ico = [0x00u8, 0x00, 0x01, 0x00];
    let bytes = picture_bytes_with_clsid(&clsid, &fake_ico);
    let mut r = reader(&bytes);
    let p = PictureRecord::read(&mut r).unwrap();
    assert_eq!(p.clsid, Some(clsid));
    assert_eq!(p.data, &fake_ico);
    let mut out = Vec::new();
    p.write(&mut out);
    assert_eq!(out, bytes);
}

#[test]
fn picture_async() {
    let bytes = picture_bytes(&[1, 2, 3]);
    let mut r = reader(&bytes);
    let rec = FrxRecord::read(RecordKind::AsyncPicture, &mut r).unwrap();
    assert!(matches!(rec, FrxRecord::AsyncPicture(_)));
}

#[test]
fn picture_bad_magic() {
    // Build a buffer where: after outer(4), the first 4 bytes are not "lt\0\0"
    // (triggering the CLSID-read path), the next 16 bytes are consumed as CLSID,
    // and the 4 bytes that should be "lt\0\0" are also wrong → BadMagic.
    let mut v = Vec::new();
    v.extend_from_slice(&28u32.to_le_bytes()); // outer (value not checked by reader)
    v.extend_from_slice(&[0x01u8; 16]);        // 16 "CLSID" bytes (0x01010101 ≠ LT_MAGIC)
    v.extend_from_slice(b"XX\0\0");            // bad magic where "lt\0\0" should be
    v.extend_from_slice(&0u32.to_le_bytes());  // 4 dummy bytes for data_len
    let mut r = reader(&v);
    assert!(matches!(PictureRecord::read(&mut r), Err(FrxError::BadMagic { .. })));
}

// ============================================================================
// ListRecord
//
// Wire layout: [u16 count] (stop if 0) [u16 sig] [{u16 len + bytes} × count]
// ============================================================================

fn list_bytes(sig: u16, items: &[&[u8]]) -> Vec<u8> {
    let mut v = (items.len() as u16).to_le_bytes().to_vec();
    if !items.is_empty() {
        v.extend_from_slice(&sig.to_le_bytes());
        for item in items {
            v.extend_from_slice(&(item.len() as u16).to_le_bytes());
            v.extend_from_slice(item);
        }
    }
    v
}

#[test]
fn list_empty() {
    let bytes = list_bytes(0, &[]);
    assert_eq!(bytes.len(), 2); // just the u16 count = 0
    let mut r = reader(&bytes);
    let l = ListRecord::read(&mut r).unwrap();
    assert_eq!(l.items.len(), 0);
    let mut out = Vec::new();
    l.write(&mut out);
    assert_eq!(out, bytes);
}

#[test]
fn list_three_items_no_item_data() {
    let bytes = list_bytes(0x000B, &[b"Alpha", b"Beta", b"Gamma"]);
    let mut r = reader(&bytes);
    let l = ListRecord::read(&mut r).unwrap();
    assert_eq!(l.items.len(), 3);
    assert_eq!(l.items[0], b"Alpha");
    assert_eq!(l.items[1], b"Beta");
    assert_eq!(l.items[2], b"Gamma");
    assert_eq!(l.sig, 0x000B);
}

#[test]
fn list_round_trip() {
    let bytes = list_bytes(0x000B, &[b"Alpha", b"Beta", b"Gamma"]);
    let mut r = reader(&bytes);
    let l = ListRecord::read(&mut r).unwrap();
    let mut out = Vec::new();
    l.write(&mut out);
    assert_eq!(out, bytes);
}

// Sample bytes: count=3, sig=0x000B, items: "Weblogs.com", "blo.gs", "bloglines.c"
#[test]
fn list_corpus_bytes() {
    let bytes: Vec<u8> = vec![
        0x03, 0x00, // count = 3
        0x0B, 0x00, // sig = 11
        0x0B, 0x00, b'W', b'e', b'b', b'l', b'o', b'g', b's', b'.', b'c', b'o', b'm',
        0x06, 0x00, b'b', b'l', b'o', b'.', b'g', b's',
        0x0B, 0x00, b'b', b'l', b'o', b'g', b'l', b'i', b'n', b'e', b's', b'.', b'c',
    ];
    let mut r = reader(&bytes);
    let l = ListRecord::read(&mut r).unwrap();
    assert_eq!(l.items.len(), 3);
    assert_eq!(l.items[0], b"Weblogs.com");
    assert_eq!(l.items[1], b"blo.gs");
    assert_eq!(l.items[2], b"bloglines.c");
    assert_eq!(l.sig, 0x000B);
    // Round-trip
    let mut out = Vec::new();
    l.write(&mut out);
    assert_eq!(out, bytes);
}

// ============================================================================
// ItemDataRecord (separate FRX record, same u16 framing as ListRecord)
// ============================================================================

fn itemdata_bytes(sig: u16, items: &[&[u8]]) -> Vec<u8> {
    let mut v = (items.len() as u16).to_le_bytes().to_vec();
    if !items.is_empty() {
        v.extend_from_slice(&sig.to_le_bytes());
        for item in items {
            v.extend_from_slice(&(item.len() as u16).to_le_bytes());
            v.extend_from_slice(item);
        }
    }
    v
}

#[test]
fn itemdata_empty() {
    let bytes = itemdata_bytes(0, &[]);
    let mut r = reader(&bytes);
    let d = ItemDataRecord::read(&mut r).unwrap();
    assert_eq!(d.items.len(), 0);
    let mut out = Vec::new();
    d.write(&mut out);
    assert_eq!(out, bytes);
}

#[test]
fn itemdata_round_trip() {
    // Three i32 values: 100, 200, 300 — each stored as 4-byte LE
    let v100 = 100i32.to_le_bytes();
    let v200 = 200i32.to_le_bytes();
    let v300 = 300i32.to_le_bytes();
    let bytes = itemdata_bytes(0x000B, &[&v100, &v200, &v300]);
    let mut r = reader(&bytes);
    let d = ItemDataRecord::read(&mut r).unwrap();
    assert_eq!(d.items.len(), 3);
    assert_eq!(ItemDataRecord::item_value(&d.items[0]), 100);
    assert_eq!(ItemDataRecord::item_value(&d.items[1]), 200);
    assert_eq!(ItemDataRecord::item_value(&d.items[2]), 300);
    let mut out = Vec::new();
    d.write(&mut out);
    assert_eq!(out, bytes);
}

#[test]
fn itemdata_via_frxrecord() {
    let v = 42i32.to_le_bytes();
    let bytes = itemdata_bytes(0x000B, &[&v]);
    let mut r = reader(&bytes);
    let rec = FrxRecord::read(RecordKind::ItemData, &mut r).unwrap();
    assert!(matches!(rec, FrxRecord::ItemData(_)));
}

// ============================================================================
// PropertyPagesRecord
// ============================================================================

fn prop_pages_bytes(pages: &[&[u8]]) -> Vec<u8> {
    let mut v = (pages.len() as u32).to_le_bytes().to_vec();
    for page in pages {
        // len includes trailing NUL
        v.extend_from_slice(&((page.len() + 1) as u16).to_le_bytes());
        v.extend_from_slice(page);
        v.push(0);
    }
    v
}

#[test]
fn property_pages_round_trip() {
    let bytes = prop_pages_bytes(&[b"PPGeneral", b"PPCol"]);
    let mut r = reader(&bytes);
    let p = PropertyPagesRecord::read(&mut r).unwrap();
    assert_eq!(p.pages.len(), 2);
    assert_eq!(p.pages[0], b"PPGeneral");
    assert_eq!(p.pages[1], b"PPCol");
    let mut out = Vec::new();
    p.write(&mut out);
    assert_eq!(out, bytes);
}

#[test]
fn property_pages_empty() {
    let bytes = 0u32.to_le_bytes();
    let mut r = reader(&bytes);
    let p = PropertyPagesRecord::read(&mut r).unwrap();
    assert_eq!(p.pages.len(), 0);
}

#[test]
fn property_pages_via_frxrecord() {
    let bytes = prop_pages_bytes(&[b"PPFont"]);
    let mut r = reader(&bytes);
    let rec = FrxRecord::read(RecordKind::PropertyPages, &mut r).unwrap();
    assert!(matches!(rec, FrxRecord::PropertyPages(_)));
}

// ============================================================================
// OcxBagRecord
// ============================================================================

fn ocx_bag_bytes(body: &[u8]) -> Vec<u8> {
    let mut v = (body.len() as u32).to_le_bytes().to_vec();
    v.extend_from_slice(body);
    v
}

#[test]
fn ocx_bag_round_trip() {
    let body = b"vendor-specific-data-here";
    let bytes = ocx_bag_bytes(body);
    let mut r = reader(&bytes);
    let b = OcxBagRecord::read(&mut r).unwrap();
    assert_eq!(b.data, body.as_ref());
    let mut out = Vec::new();
    b.write(&mut out);
    assert_eq!(out, bytes);
}

#[test]
fn ocx_bag_with_clsid() {
    // Body starts with 16 non-printable bytes (looks like a GUID)
    let clsid_bytes: [u8; 16] = [
        0x80, 0x09, 0xF8, 0x7B, 0x92, 0xCC, 0x1B, 0x42,
        0xBE, 0x9E, 0xE9, 0xC5, 0x92, 0xE8, 0x8B, 0x61,
    ];
    let extra = b"extra-vendor-data";
    let mut body = clsid_bytes.to_vec();
    body.extend_from_slice(extra);
    let bytes = ocx_bag_bytes(&body);
    let mut r = reader(&bytes);
    let b = OcxBagRecord::read(&mut r).unwrap();
    assert_eq!(b.clsid, Some(clsid_bytes));
    assert_eq!(b.data, body.as_slice());
    let mut out = Vec::new();
    b.write(&mut out);
    assert_eq!(out, bytes);
}

#[test]
fn ocx_bag_via_frxrecord() {
    let bytes = ocx_bag_bytes(b"opaque");
    let mut r = reader(&bytes);
    let rec = FrxRecord::read(RecordKind::OcxBag, &mut r).unwrap();
    assert!(matches!(rec, FrxRecord::OcxBag(_)));
}

// ============================================================================
// PropertyBag
// ============================================================================

fn bag_bytes(entries: &[(&[u8], &[u8])]) -> Vec<u8> {
    let mut v = (entries.len() as u32).to_le_bytes().to_vec();
    for (name, value) in entries {
        v.extend_from_slice(&(name.len() as u32).to_le_bytes());
        v.extend_from_slice(name);
        v.extend_from_slice(&(value.len() as u32).to_le_bytes());
        v.extend_from_slice(value);
    }
    v
}

#[test]
fn property_bag_round_trip() {
    let entries: &[(&[u8], &[u8])] = &[
        (b"Caption",  b"Hello"),
        (b"Enabled",  b"True"),
        (b"TabIndex", b"0"),
    ];
    let bytes = bag_bytes(entries);
    let mut r = reader(&bytes);
    let bag = PropertyBagRecord::read(&mut r).unwrap();
    assert_eq!(bag.entries.len(), 3);
    assert_eq!(bag.entries[0].name,  b"Caption");
    assert_eq!(bag.entries[0].value, b"Hello");
    let mut out = Vec::new();
    bag.write(&mut out);
    assert_eq!(out, bytes);
}

#[test]
fn property_bag_empty() {
    let bytes = 0u32.to_le_bytes();
    let mut r = reader(&bytes);
    let bag = PropertyBagRecord::read(&mut r).unwrap();
    assert_eq!(bag.entries.len(), 0);
}

// ============================================================================
// OleStorageRecord
// ============================================================================

fn ole_bytes(mechanism: u8, clsid_bytes: &[u8; 16], data: &[u8]) -> Vec<u8> {
    let mut v = vec![mechanism];
    v.extend_from_slice(clsid_bytes);
    v.extend_from_slice(&(data.len() as u32).to_le_bytes());
    v.extend_from_slice(data);
    v
}

fn clsid_bytes(clsid: &Clsid) -> [u8; 16] {
    let mut buf = Vec::new();
    clsid.write(&mut buf);
    buf.try_into().unwrap()
}

#[test]
fn ole_storage_stream_round_trip() {
    let clsid = Clsid::IID_IPERSIST_STREAM;
    let data = [0xDE, 0xAD, 0xBE, 0xEF];
    let bytes = ole_bytes(0x01, &clsid_bytes(&clsid), &data);
    let mut r = reader(&bytes);
    let rec = OleStorageRecord::read(&mut r).unwrap();
    assert_eq!(rec.mechanism, PersistMechanism::Stream);
    assert_eq!(rec.data, &data);
    let mut out = Vec::new();
    rec.write(&mut out);
    assert_eq!(out, bytes);
}

#[test]
fn ole_storage_storage_mechanism_round_trip() {
    let clsid = Clsid::IID_IPERSIST_STORAGE;
    let data = [0x01u8, 0x02, 0x03, 0x04];
    let bytes = ole_bytes(0x02, &clsid_bytes(&clsid), &data);
    let mut r = reader(&bytes);
    let rec = OleStorageRecord::read(&mut r).unwrap();
    assert_eq!(rec.mechanism, PersistMechanism::Storage);
    assert_eq!(rec.clsid, clsid);
    assert_eq!(rec.data, &data);
    let mut out = Vec::new();
    rec.write(&mut out);
    assert_eq!(out, bytes);
}

#[test]
fn ole_storage_bad_mechanism() {
    let bytes = ole_bytes(0x99, &[0u8; 16], &[]);
    let mut r = reader(&bytes);
    assert!(matches!(OleStorageRecord::read(&mut r), Err(FrxError::BadMagic { .. })));
}

// ============================================================================
// ControlCreateRecord
// ============================================================================

fn ctrl_create_bytes(clsid: &Clsid, prog_id: &[u8], license: Option<&[u8]>) -> Vec<u8> {
    let mut v = Vec::new();
    clsid.write(&mut v);
    v.extend_from_slice(&(prog_id.len() as u32).to_le_bytes());
    v.extend_from_slice(prog_id);
    match license {
        None => v.push(0),
        Some(key) => {
            v.push(1);
            v.extend_from_slice(&(key.len() as u32).to_le_bytes());
            v.extend_from_slice(key);
        }
    }
    v
}

#[test]
fn control_create_no_license() {
    let clsid = Clsid::IID_IPICTURE;
    let bytes = ctrl_create_bytes(&clsid, b"MSComCtl.Slider.2", None);
    let mut r = reader(&bytes);
    let rec = ControlCreateRecord::read(&mut r).unwrap();
    assert_eq!(rec.clsid, clsid);
    assert_eq!(rec.prog_id, b"MSComCtl.Slider.2");
    assert!(rec.license.is_none());
}

#[test]
fn control_create_with_license_round_trip() {
    let clsid = Clsid::IID_IFONT;
    let key = b"some-license-key-data";
    let bytes = ctrl_create_bytes(&clsid, b"MSGrid.Grid.1", Some(key));
    let mut r = reader(&bytes);
    let rec = ControlCreateRecord::read(&mut r).unwrap();
    assert_eq!(rec.license.as_ref().unwrap().key, key.as_ref());
    let mut out = Vec::new();
    rec.write(&mut out);
    assert_eq!(out, bytes);
}

// ============================================================================
// ControlPersistRecord
// ============================================================================

#[test]
fn control_persist_stream_round_trip() {
    let clsid = Clsid::IID_IPERSIST_STREAM_INIT;
    let inner_data = [0xCA, 0xFE, 0xBA, 0xBE];
    let mut v = vec![0x01u8];
    clsid.write(&mut v);
    let inner_clsid = Clsid::IID_IPERSIST_STREAM;
    v.push(0x01);
    inner_clsid.write(&mut v);
    v.extend_from_slice(&(inner_data.len() as u32).to_le_bytes());
    v.extend_from_slice(&inner_data);

    let mut r = reader(&v);
    let rec = ControlPersistRecord::read(&mut r).unwrap();
    assert_eq!(rec.clsid(), &clsid);
    let mut out = Vec::new();
    rec.write(&mut out);
    assert_eq!(out, v);
}

#[test]
fn control_persist_storage_round_trip() {
    let outer_clsid = Clsid::IID_IPERSIST_STORAGE;
    let inner_clsid = Clsid::IID_IPERSIST_STORAGE;
    let inner_data = [0xAA, 0xBB, 0xCC, 0xDD];
    let mut v = vec![0x02u8];
    outer_clsid.write(&mut v);
    v.push(0x02);
    inner_clsid.write(&mut v);
    v.extend_from_slice(&(inner_data.len() as u32).to_le_bytes());
    v.extend_from_slice(&inner_data);

    let mut r = reader(&v);
    let rec = ControlPersistRecord::read(&mut r).unwrap();
    assert_eq!(rec.clsid(), &outer_clsid);
    let mut out = Vec::new();
    rec.write(&mut out);
    assert_eq!(out, v);
}

#[test]
fn control_persist_property_bag() {
    let clsid = Clsid::IID_IPERSIST_PROPERTY_BAG;
    let mut v = vec![0x03u8];
    clsid.write(&mut v);
    v.extend_from_slice(&1u32.to_le_bytes());
    v.extend_from_slice(&7u32.to_le_bytes()); v.extend_from_slice(b"Enabled");
    v.extend_from_slice(&4u32.to_le_bytes()); v.extend_from_slice(b"True");

    let mut r = reader(&v);
    let rec = ControlPersistRecord::read(&mut r).unwrap();
    match &rec {
        ControlPersistRecord::PropertyBag(_, bag) => {
            assert_eq!(bag.entries[0].name, b"Enabled");
        }
        _ => panic!("expected PropertyBag"),
    }
}

// ============================================================================
// ControlArray
// ============================================================================

#[test]
fn control_array_visible_and_tab_stop() {
    let bytes = [3u8, 0, 0, 0, 0x03];
    let mut r = reader(&bytes);
    let rec = FrxRecord::read(RecordKind::ControlArray, &mut r).unwrap();
    match rec {
        FrxRecord::ControlArray { index, visible, tab_stop } => {
            assert_eq!(index, 3);
            assert!(visible);
            assert!(tab_stop);
        }
        _ => panic!(),
    }
}

#[test]
fn control_array_write_read_round_trip() {
    let mut out = Vec::new();
    FrxRecord::ControlArray { index: 7, visible: true, tab_stop: false }.write(&mut out);
    let mut r = reader(&out);
    let rec = FrxRecord::read(RecordKind::ControlArray, &mut r).unwrap();
    match rec {
        FrxRecord::ControlArray { index, visible, tab_stop } => {
            assert_eq!(index, 7);
            assert!(visible);
            assert!(!tab_stop);
        }
        _ => panic!(),
    }
}

// ============================================================================
// Menu
// ============================================================================

#[test]
fn menu_separator_bar() {
    let caption = b"-";
    let mut v = (caption.len() as u32).to_le_bytes().to_vec();
    v.extend_from_slice(caption);
    v.push(0);
    let mut r = reader(&v);
    let rec = FrxRecord::read(RecordKind::Menu, &mut r).unwrap();
    match rec {
        FrxRecord::Menu { caption, negotiate_position } => {
            assert_eq!(caption, b"-");
            assert_eq!(negotiate_position, 0);
        }
        _ => panic!(),
    }
}

#[test]
fn menu_round_trip() {
    let mut out = Vec::new();
    FrxRecord::Menu { caption: b"File".to_vec(), negotiate_position: 2 }.write(&mut out);
    let mut r = reader(&out);
    let rec = FrxRecord::read(RecordKind::Menu, &mut r).unwrap();
    match rec {
        FrxRecord::Menu { caption, negotiate_position } => {
            assert_eq!(caption, b"File");
            assert_eq!(negotiate_position, 2);
        }
        _ => panic!(),
    }
}

// ============================================================================
// DataBinding
// ============================================================================

#[test]
fn data_binding_round_trip() {
    let mut out = Vec::new();
    FrxRecord::DataBinding {
        prop_name:  b"Text".to_vec(),
        data_field: b"CustomerName".to_vec(),
    }.write(&mut out);
    let mut r = reader(&out);
    let rec = FrxRecord::read(RecordKind::DataBinding, &mut r).unwrap();
    match rec {
        FrxRecord::DataBinding { prop_name, data_field } => {
            assert_eq!(prop_name,  b"Text");
            assert_eq!(data_field, b"CustomerName");
        }
        _ => panic!(),
    }
}
