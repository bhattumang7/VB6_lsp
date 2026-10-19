//! Reader for the VB6 persisted PropertyBag stream that VB-authored third-party controls
//! store in a form's `.frx`/`.ctx` for each `OcxBag` property.
//!
//! Layout (little endian), found by comparing real bags:
//!   0x00 class id (16) | 0x10 u16 version | 0x12 u16 | 0x14 u32 body length | 0x18 u32 0 | 0x1C u32 | 0x20 u32 body length
//!   0x24 records, each: u16 vartype | u16 name length in chars | u32 name hash | u32 record length (whole record)
//!        | UTF-16LE name | value (by vartype)
//! Values: I2, I4, R8, BOOL (2 bytes), BSTR (u32 char count, UTF-16LE), object (class id + persisted data).
//! A record whose length reads as a negative number opens a nested group (a column's value list, the closing
//! section marker): its header is reported as a `group` record and the nested records follow in the same list.
//! Bytes left after the last readable record are returned as `tail_hex`.

use serde_json::{json, Value};

pub const PROPERTY_BAG_CLSID: &str = "{B1E87400-01F9-11D0-B3DE-00A0C90F2731}";
const RECORDS_START: usize = 0x24;
const RECORD_HEADER_LEN: usize = 12;

const VT_I2: u16 = 2;
const VT_I4: u16 = 3;
const VT_R4: u16 = 4;
const VT_R8: u16 = 5;
const VT_BSTR: u16 = 8;
const VT_OBJECT: u16 = 9;
const VT_BOOL: u16 = 11;

fn u16_at(b: &[u8], p: usize) -> u16 { u16::from_le_bytes([b[p], b[p + 1]]) }
fn u32_at(b: &[u8], p: usize) -> u32 { u32::from_le_bytes([b[p], b[p + 1], b[p + 2], b[p + 3]]) }

fn hex(b: &[u8]) -> String { b.iter().map(|x| format!("{:02x}", x)).collect() }

fn guid_text(b: &[u8]) -> String {
    format!(
        "{{{:08X}-{:04X}-{:04X}-{:02X}{:02X}-{}}}",
        u32_at(b, 0), u16_at(b, 4), u16_at(b, 6), b[8], b[9], b[10..16].iter().map(|x| format!("{:02X}", x)).collect::<String>()
    )
}

fn utf16(b: &[u8]) -> String {
    let units: Vec<u16> = b.chunks_exact(2).map(|c| u16::from_le_bytes([c[0], c[1]])).collect();
    String::from_utf16_lossy(&units)
}

/// A group's records start with a stream header like the bag's own, without the second version word:
/// u16 version | u32 length | u32 0 | u32 marker | u32 length.
const GROUP_STREAM_HEADER_LEN: usize = 18;
const STREAM_MARKER: u32 = 0xB293;

fn is_group_stream_header(b: &[u8], p: usize) -> bool {
    p + GROUP_STREAM_HEADER_LEN <= b.len() && u32_at(b, p + 10) == STREAM_MARKER && u32_at(b, p + 2) == u32_at(b, p + 14)
}

fn value_json(vt: u16, v: &[u8]) -> (&'static str, Value) {
    match vt {
        VT_I2 if v.len() >= 2 => ("i2", json!(i16::from_le_bytes([v[0], v[1]]))),
        VT_I4 if v.len() >= 4 => ("i4", json!(i32::from_le_bytes([v[0], v[1], v[2], v[3]]))),
        VT_R4 if v.len() >= 4 => ("r4", json!(f32::from_le_bytes([v[0], v[1], v[2], v[3]]))),
        VT_R8 if v.len() >= 8 => ("r8", json!(f64::from_le_bytes(v[..8].try_into().unwrap()))),
        VT_BOOL if v.len() >= 2 => ("bool", json!(i16::from_le_bytes([v[0], v[1]]) != 0)),
        VT_BSTR if v.len() >= 4 => {
            let n = u32_at(v, 0);
            if n == u32::MAX { ("bstr", Value::Null) }
            else if 4 + 2 * (n as usize) <= v.len() { ("bstr", json!(utf16(&v[4..4 + 2 * n as usize]))) }
            else { ("bstr?", json!(hex(v))) }
        }
        VT_OBJECT if v.len() >= 16 => ("object", json!({ "clsid": guid_text(v), "data_hex": hex(&v[16..]) })),
        _ => ("raw", json!({ "vartype": vt, "hex": hex(v) })),
    }
}

/// Decode one bag. `None` when the bytes are not a property-bag stream.
pub fn decode(data: &[u8]) -> Option<Value> {
    if data.len() < RECORDS_START || guid_text(data) != PROPERTY_BAG_CLSID { return None; }
    let mut p = RECORDS_START;
    let mut records = Vec::new();
    while p + RECORD_HEADER_LEN <= data.len() {
        let vt = u16_at(data, p);
        let name_chars = u16_at(data, p + 2) as usize;
        let raw_size = u32_at(data, p + 8);
        if raw_size > i32::MAX as u32 {
            let header_len = RECORD_HEADER_LEN + 2 * name_chars;
            if p + header_len > data.len() { break; }
            let name = utf16(&data[p + RECORD_HEADER_LEN..p + header_len]);
            records.push(json!({ "name": name, "type": "group", "value": (raw_size as i32).unsigned_abs() }));
            p += header_len;
            if is_group_stream_header(data, p) { p += GROUP_STREAM_HEADER_LEN; }
            continue;
        }
        let size = raw_size as usize;
        if size < RECORD_HEADER_LEN + 2 * name_chars || p + size > data.len() { break; }
        let name = utf16(&data[p + RECORD_HEADER_LEN..p + RECORD_HEADER_LEN + 2 * name_chars]);
        let (ty, value) = value_json(vt, &data[p + RECORD_HEADER_LEN + 2 * name_chars..p + size]);
        records.push(json!({ "name": name, "type": ty, "value": value }));
        p += size;
    }
    Some(json!({ "version": u16_at(data, 16), "records": records, "unread_bytes": data.len() - p, "tail_hex": hex(&data[p..]) }))
}

#[cfg(test)]
mod tests {
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
}
