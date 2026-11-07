//! Finding the type library inside a file, on synthetic executable images.

use vb6_typelib::pe::{type_library_bytes, LocateError};

const SECTION_RVA: u32 = 0x1000;
const SECTION_RAW: u32 = 0x400;

fn put_u16(out: &mut Vec<u8>, v: u16) {
    out.extend_from_slice(&v.to_le_bytes());
}

fn put_u32(out: &mut Vec<u8>, v: u32) {
    out.extend_from_slice(&v.to_le_bytes());
}

/// One resource directory table: (named entries, id entries), each (key, value).
fn directory(named: &[(u32, u32)], ids: &[(u32, u32)]) -> Vec<u8> {
    let mut out = Vec::new();
    put_u32(&mut out, 0);
    put_u32(&mut out, 0);
    put_u16(&mut out, 0);
    put_u16(&mut out, 0);
    put_u16(&mut out, named.len() as u16);
    put_u16(&mut out, ids.len() as u16);
    for (key, value) in named.iter().chain(ids) {
        put_u32(&mut out, *key);
        put_u32(&mut out, *value);
    }
    out
}

/// A resource section holding `TYPELIB` resources with the given ids and payloads (and, if
/// `other_type` is set, a numeric-type resource ahead of them).
fn resource_section(resources: &[(u32, &[u8])], other_type: bool) -> Vec<u8> {
    // Layout: root, [other type dir], TYPELIB name dir, per-resource language dirs and data
    // entries, the type name string, then the payloads.
    let high = 0x8000_0000u32;
    let root_size = 16 + 8 * if other_type { 2 } else { 1 };
    let other_dir_at = root_size;
    let other_dir_size = 16 + 8;
    let other_lang_at = other_dir_at + other_dir_size;
    let other_lang_size = 16 + 8;
    let other_data_at = other_lang_at + other_lang_size;
    let other_total = if other_type { other_dir_size + other_lang_size + 16 } else { 0 };
    let names_at = root_size + other_total;
    let names_size = 16 + 8 * resources.len();
    let mut cursor = names_at + names_size;
    let mut language_dirs = Vec::new();
    for _ in resources {
        language_dirs.push(cursor);
        cursor += 16 + 8;
    }
    let mut data_entries = Vec::new();
    for _ in resources {
        data_entries.push(cursor);
        cursor += 16;
    }
    let string_at = cursor;
    cursor += 2 + "TYPELIB".len() * 2;
    cursor = (cursor + 3) & !3;
    let mut payload_at = Vec::new();
    for (_, payload) in resources {
        payload_at.push(cursor);
        cursor += (payload.len() + 3) & !3;
    }

    let mut out = Vec::new();
    // Root: the named TYPELIB entry first, then (when asked) a numeric type (RT_RCDATA = 10).
    let named = [(high | string_at as u32, high | names_at as u32)];
    let ids: Vec<(u32, u32)> = if other_type { vec![(10, high | other_dir_at as u32)] } else { vec![] };
    out.extend(directory(&named, &ids));
    if other_type {
        out.extend(directory(&[], &[(1, high | other_lang_at as u32)]));
        out.extend(directory(&[], &[(1033, other_data_at as u32)]));
        put_u32(&mut out, SECTION_RVA + cursor as u32);
        put_u32(&mut out, 1);
        put_u32(&mut out, 0);
        put_u32(&mut out, 0);
    }
    let name_entries: Vec<(u32, u32)> =
        resources.iter().zip(&language_dirs).map(|((id, _), at)| (*id, high | *at as u32)).collect();
    out.extend(directory(&[], &name_entries));
    for entry in &data_entries {
        out.extend(directory(&[], &[(1033, *entry as u32)]));
    }
    for ((_, payload), at) in resources.iter().zip(&payload_at) {
        put_u32(&mut out, SECTION_RVA + *at as u32);
        put_u32(&mut out, payload.len() as u32);
        put_u32(&mut out, 0);
        put_u32(&mut out, 0);
    }
    put_u16(&mut out, "TYPELIB".len() as u16);
    for unit in "TYPELIB".encode_utf16() {
        put_u16(&mut out, unit);
    }
    while out.len() < payload_at.first().copied().unwrap_or(out.len()) {
        out.push(0);
    }
    for (_, payload) in resources {
        out.extend_from_slice(payload);
        while out.len() % 4 != 0 {
            out.push(0);
        }
    }
    out
}

/// A PE image whose only section is `section`, with its resource directory at the section start.
fn image(section: &[u8], plus: bool, has_resources: bool) -> Vec<u8> {
    let mut out = vec![0u8; 0x40];
    out[0] = b'M';
    out[1] = b'Z';
    out[0x3C..0x40].copy_from_slice(&0x40u32.to_le_bytes());
    out.extend_from_slice(b"PE\0\0");
    put_u16(&mut out, 0x14c);
    put_u16(&mut out, 1);
    put_u32(&mut out, 0);
    put_u32(&mut out, 0);
    put_u32(&mut out, 0);
    let optional_size: u16 = if plus { 240 } else { 224 };
    put_u16(&mut out, optional_size);
    put_u16(&mut out, 0);
    let optional_start = out.len();
    put_u16(&mut out, if plus { 0x20b } else { 0x10b });
    out.resize(optional_start + if plus { 112 } else { 96 }, 0);
    for directory in 0..16 {
        if directory == 2 && has_resources {
            put_u32(&mut out, SECTION_RVA);
            put_u32(&mut out, section.len() as u32);
        } else {
            put_u32(&mut out, 0);
            put_u32(&mut out, 0);
        }
    }
    out.resize(optional_start + optional_size as usize, 0);
    out.extend_from_slice(b".rsrc\0\0\0");
    put_u32(&mut out, section.len() as u32);
    put_u32(&mut out, SECTION_RVA);
    put_u32(&mut out, section.len() as u32);
    put_u32(&mut out, SECTION_RAW);
    out.resize(SECTION_RAW as usize, 0);
    out.extend_from_slice(section);
    out
}

#[test]
fn a_type_library_file_is_its_own_bytes() {
    let data = b"MSFT\x02\x00\x01\x00rest of the file";

    assert_eq!(type_library_bytes(data, None), Ok(&data[..]));
    assert_eq!(type_library_bytes(data, Some(3)), Ok(&data[..]));
}

#[test]
fn a_file_that_is_neither_a_type_library_nor_an_image_is_refused() {
    assert_eq!(type_library_bytes(b"plain text", None), Err(LocateError::NotAnImage));
    assert_eq!(type_library_bytes(b"", None), Err(LocateError::NotAnImage));
    let mut no_pe = vec![0u8; 0x80];
    no_pe[0] = b'M';
    no_pe[1] = b'Z';
    no_pe[0x3C] = 0x40;
    assert_eq!(type_library_bytes(&no_pe, None), Err(LocateError::NotAnImage));
}

#[test]
fn the_first_type_library_resource_of_an_image_is_the_default() {
    let section = resource_section(&[(1, b"FIRST-LIBRARY"), (2, b"SECOND-LIBRARY")], false);
    let data = image(&section, false, true);

    assert_eq!(type_library_bytes(&data, None), Ok(&b"FIRST-LIBRARY"[..]));
    assert_eq!(type_library_bytes(&data, Some(0)), Ok(&b"FIRST-LIBRARY"[..]));
    assert_eq!(type_library_bytes(&data, Some(1)), Ok(&b"FIRST-LIBRARY"[..]));
}

#[test]
fn a_type_library_resource_is_chosen_by_its_index() {
    let section = resource_section(&[(1, b"FIRST-LIBRARY"), (2, b"SECOND-LIBRARY"), (3, b"THIRD")], false);
    let data = image(&section, false, true);

    assert_eq!(type_library_bytes(&data, Some(2)), Ok(&b"SECOND-LIBRARY"[..]));
    assert_eq!(type_library_bytes(&data, Some(3)), Ok(&b"THIRD"[..]));
    assert_eq!(type_library_bytes(&data, Some(4)), Err(LocateError::NoTypeLibrary));
}

#[test]
fn a_64_bit_image_is_read_the_same_way() {
    let section = resource_section(&[(1, b"ONLY-LIBRARY")], false);
    let data = image(&section, true, true);

    assert_eq!(type_library_bytes(&data, None), Ok(&b"ONLY-LIBRARY"[..]));
}

#[test]
fn other_resource_types_are_passed_over() {
    let section = resource_section(&[(1, b"THE-LIBRARY")], true);
    let data = image(&section, false, true);

    assert_eq!(type_library_bytes(&data, None), Ok(&b"THE-LIBRARY"[..]));
}

#[test]
fn an_image_without_resources_has_no_type_library() {
    let data = image(&[0u8; 16], false, false);

    assert_eq!(type_library_bytes(&data, None), Err(LocateError::NoTypeLibrary));
}

#[test]
fn an_image_whose_tables_point_outside_the_file_is_reported_damaged() {
    let section = resource_section(&[(1, b"LIBRARY-BYTES")], false);
    let data = image(&section, false, true);

    let cut = &data[..SECTION_RAW as usize + 20];

    assert!(matches!(type_library_bytes(cut, None), Err(LocateError::Damaged(_))));
    assert!(matches!(type_library_bytes(&data[..0x30], None), Err(LocateError::Damaged(_))));
}

#[test]
fn a_damaged_image_message_names_what_is_damaged() {
    assert_eq!(LocateError::NotAnImage.to_string(), "not a type library or an executable image");
    assert_eq!(LocateError::Damaged("section table").to_string(), "damaged executable image (section table)");
    assert_eq!(LocateError::NoTypeLibrary.to_string(), "no type library resource");
}
