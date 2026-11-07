use super::*;

fn record(vt: u16, name: &str, value: &[u8]) -> Vec<u8> {
    let name16: Vec<u8> = name.encode_utf16().flat_map(|u| u.to_le_bytes()).collect();
    let size = (RECORD_HEADER_LEN + name16.len() + value.len()) as u32;
    let mut r = Vec::new();
    r.extend(vt.to_le_bytes());
    r.extend((name.encode_utf16().count() as u16).to_le_bytes());
    r.extend(0u32.to_le_bytes());
    r.extend(size.to_le_bytes());
    r.extend(name16);
    r.extend(value);
    r
}

fn bag(records: &[Vec<u8>]) -> Vec<u8> {
    let mut b = vec![0x00, 0x74, 0xe8, 0xb1, 0xf9, 0x01, 0xd0, 0x11, 0xb3, 0xde, 0x00, 0xa0, 0xc9, 0x0f, 0x27, 0x31];
    b.extend([0x11, 0, 1, 0]);
    b.extend(vec![0u8; RECORDS_START - 20]);
    for r in records { b.extend(r); }
    b
}

#[test]
fn reads_each_value_type() {
    let mut s = 3u32.to_le_bytes().to_vec();
    s.extend("abc".encode_utf16().flat_map(|u| u.to_le_bytes()));
    let b = bag(&[
        record(VT_I2, "a", &7i16.to_le_bytes()),
        record(VT_I4, "b", &(-5i32).to_le_bytes()),
        record(VT_R8, "c", &1.5f64.to_le_bytes()),
        record(VT_BOOL, "d", &(-1i16).to_le_bytes()),
        record(VT_BSTR, "e", &s),
    ]);
    let v = decode(&b).unwrap();
    let r = v["records"].as_array().unwrap();
    assert_eq!(r[0]["value"], 7);
    assert_eq!(r[1]["value"], -5);
    assert_eq!(r[2]["value"], 1.5);
    assert_eq!(r[3]["value"], true);
    assert_eq!(r[4]["value"], "abc");
    assert_eq!(v["unread_bytes"], 0);
}

#[test]
fn rejects_other_streams_and_stops_at_a_bad_record_length() {
    assert!(decode(&[0u8; 64]).is_none());
    let mut b = bag(&[record(VT_I2, "a", &1i16.to_le_bytes())]);
    b.extend(record(VT_I2, "z", &[0, 0]).iter().take(8));
    let v = decode(&b).unwrap();
    assert_eq!(v["records"].as_array().unwrap().len(), 1);
    assert_eq!(v["unread_bytes"], 8);
    assert_eq!(v["tail_hex"].as_str().unwrap().len(), 16);
}

#[test]
fn a_negative_length_opens_a_group_whose_records_follow() {
    let mut group = record(VT_BOOL, "vl", &[]);
    group[8..12].copy_from_slice(&(-40i32).to_le_bytes());
    let header_len = RECORD_HEADER_LEN + 4;
    group.truncate(header_len);
    let mut records = vec![group];
    let mut stream_header = vec![1, 0];
    stream_header.extend(40u32.to_le_bytes());
    stream_header.extend(0u32.to_le_bytes());
    stream_header.extend(STREAM_MARKER.to_le_bytes());
    stream_header.extend(40u32.to_le_bytes());
    records.push(stream_header);
    records.push(record(VT_I2, "n", &2i16.to_le_bytes()));
    let v = decode(&bag(&records)).unwrap();
    let r = v["records"].as_array().unwrap();
    assert_eq!(r[0]["type"], "group");
    assert_eq!(r[0]["value"], 40);
    assert_eq!(r[1]["name"], "n");
    assert_eq!(v["unread_bytes"], 0);
}
