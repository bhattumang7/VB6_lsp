//! Decoding and re-encoding of `.frx` blobs into values: exact results, spans and errors.

use vb6_formats::frx::{
    decode, decode_ocx_bag, decode_span, encode, itemdata_value, FontInfo, FrxError, FrxValue,
    ImageFormat, PropKind,
};

/// A standard-framed picture: `[outer][lt\0\0][len][image]`.
fn picture(image: &[u8]) -> Vec<u8> {
    let mut b = Vec::new();
    b.extend_from_slice(&(image.len() as u32 + 8).to_le_bytes());
    b.extend_from_slice(b"lt\0\0");
    b.extend_from_slice(&(image.len() as u32).to_le_bytes());
    b.extend_from_slice(image);
    b
}

/// A class-id-framed picture: `[outer][clsid][lt\0\0][len][image]`.
fn picture_with_clsid(clsid: [u8; 16], image: &[u8]) -> Vec<u8> {
    let mut b = Vec::new();
    b.extend_from_slice(&(image.len() as u32 + 24).to_le_bytes());
    b.extend_from_slice(&clsid);
    b.extend_from_slice(b"lt\0\0");
    b.extend_from_slice(&(image.len() as u32).to_le_bytes());
    b.extend_from_slice(image);
    b
}

/// `MS Sans Serif` 8.25pt, normal weight, no flags.
fn ms_sans_serif() -> Vec<u8> {
    vec![
        0x01, 0x00, 0x00, 0x00, 0x90, 0x01, 0x44, 0x42, 0x01, 0x00, 0x0D, b'M', b'S', b' ', b'S',
        b'a', b'n', b's', b' ', b'S', b'e', b'r', b'i', b'f',
    ]
}

fn list_blob() -> Vec<u8> {
    let mut b = Vec::new();
    b.extend_from_slice(&2u16.to_le_bytes());
    b.extend_from_slice(&0x000Bu16.to_le_bytes());
    b.extend_from_slice(&3u16.to_le_bytes());
    b.extend_from_slice(b"abc");
    b.extend_from_slice(&2u16.to_le_bytes());
    b.extend_from_slice(b"de");
    b
}

const CLSID: [u8; 16] =
    [0x04, 0x52, 0xE3, 0x0B, 0x91, 0x8F, 0xCE, 0x11, 0x9D, 0xE3, 0x00, 0xAA, 0x00, 0x4B, 0xB8, 0x51];

// ---- pictures ----------------------------------------------------------------------------

#[test]
fn a_standard_picture_decodes_to_its_image_and_format() {
    let blob = picture(&[0x42, 0x4D, 0x00, 0x00]);
    let (value, span) = decode_span(&blob, 0, PropKind::Picture).unwrap();
    assert_eq!(
        value,
        FrxValue::Picture {
            format: ImageFormat::Bmp,
            data: vec![0x42, 0x4D, 0x00, 0x00],
            clsid: None
        }
    );
    assert_eq!(span, blob.len());
}

#[test]
fn a_picture_is_read_at_the_offset_given() {
    let mut blob = vec![0xAA; 5];
    blob.extend(picture(&[0x47, 0x49, 0x46, 0x38]));
    let (value, span) = decode_span(&blob, 5, PropKind::Picture).unwrap();
    assert_eq!(
        value,
        FrxValue::Picture {
            format: ImageFormat::Gif,
            data: vec![0x47, 0x49, 0x46, 0x38],
            clsid: None
        }
    );
    assert_eq!(span, blob.len() - 5);
}

#[test]
fn an_empty_picture_slot_is_an_empty_value() {
    let blob = picture(&[]);
    assert_eq!(blob.len(), 12);
    assert_eq!(decode_span(&blob, 0, PropKind::Picture).unwrap(), (FrxValue::Empty, 12));
}

#[test]
fn a_class_id_framed_picture_keeps_its_class_id() {
    let blob = picture_with_clsid(CLSID, &[0xFF, 0xD8, 0xFF, 0xE0]);
    let (value, span) = decode_span(&blob, 0, PropKind::Picture).unwrap();
    assert_eq!(
        value,
        FrxValue::Picture {
            format: ImageFormat::Jpeg,
            data: vec![0xFF, 0xD8, 0xFF, 0xE0],
            clsid: Some(CLSID)
        }
    );
    assert_eq!(span, 32);
}

#[test]
fn an_empty_class_id_framed_slot_keeps_its_framing() {
    let blob = picture_with_clsid(CLSID, &[]);
    assert_eq!(blob.len(), 28);
    let (value, span) = decode_span(&blob, 0, PropKind::Picture).unwrap();
    assert_eq!(
        value,
        FrxValue::Picture { format: ImageFormat::Unknown, data: Vec::new(), clsid: Some(CLSID) }
    );
    assert_eq!(span, 28);
    assert_eq!(encode(&value, PropKind::Picture), blob);
}

#[test]
fn a_picture_without_the_marker_is_a_bad_header() {
    let mut blob = picture(&[1, 2, 3, 4]);
    blob[4] = b'x';
    assert_eq!(
        decode(&blob, 0, PropKind::Picture).unwrap_err(),
        FrxError::BadPictureHeader { pos: 0 }
    );
}

#[test]
fn a_picture_whose_outer_length_disagrees_is_a_bad_header() {
    let mut blob = picture(&[1, 2, 3, 4]);
    blob[0] = 13; // should be 12
    assert_eq!(
        decode(&blob, 0, PropKind::Picture).unwrap_err(),
        FrxError::BadPictureHeader { pos: 0 }
    );
}

#[test]
fn an_empty_slot_whose_outer_length_disagrees_is_a_bad_header() {
    let mut blob = picture(&[]);
    blob[0] = 9; // should be 8
    assert_eq!(
        decode(&blob, 0, PropKind::Picture).unwrap_err(),
        FrxError::BadPictureHeader { pos: 0 }
    );
}

#[test]
fn a_picture_whose_image_is_cut_short_is_a_length_overflow() {
    let mut blob = picture(&[1, 2, 3, 4]);
    blob.truncate(blob.len() - 2);
    assert_eq!(
        decode(&blob, 0, PropKind::Picture).unwrap_err(),
        FrxError::LengthOverflow { pos: 8, declared: 4, remaining: 2 }
    );
}

#[test]
fn a_picture_blob_too_short_for_a_header_is_a_bad_header() {
    assert_eq!(
        decode(&[0x0C, 0, 0, 0, b'l', b't'], 0, PropKind::Picture).unwrap_err(),
        FrxError::BadPictureHeader { pos: 0 }
    );
}

#[test]
fn every_image_format_is_detected_from_its_leading_bytes() {
    let cases: [(&[u8], ImageFormat); 7] = [
        (&[0x42, 0x4D, 0, 0], ImageFormat::Bmp),
        (&[0, 0, 1, 0], ImageFormat::Ico),
        (&[0, 0, 2, 0], ImageFormat::Cur),
        (&[0x89, 0x50, 0x4E, 0x47], ImageFormat::Png),
        (&[0xD7, 0xCD, 0xC6, 0x9A], ImageFormat::Wmf),
        (&[1, 0, 9, 0], ImageFormat::Wmf),
        (&[9, 9, 9, 9], ImageFormat::Unknown),
    ];
    for (image, expected) in cases {
        let value = decode(&picture(image), 0, PropKind::Picture).unwrap();
        assert_eq!(
            value,
            FrxValue::Picture { format: expected, data: image.to_vec(), clsid: None }
        );
    }

    let mut emf = vec![0u8; 44];
    emf[40..44].copy_from_slice(b" EMF");
    match decode(&picture(&emf), 0, PropKind::Picture).unwrap() {
        FrxValue::Picture { format, .. } => assert_eq!(format, ImageFormat::Emf),
        other => panic!("expected a picture, got {other:?}"),
    }
}

#[test]
fn image_formats_name_their_file_extensions() {
    let named = [
        (ImageFormat::Bmp, "bmp"),
        (ImageFormat::Ico, "ico"),
        (ImageFormat::Cur, "cur"),
        (ImageFormat::Gif, "gif"),
        (ImageFormat::Jpeg, "jpg"),
        (ImageFormat::Png, "png"),
        (ImageFormat::Wmf, "wmf"),
        (ImageFormat::Emf, "emf"),
        (ImageFormat::Unknown, "bin"),
    ];
    for (format, ext) in named {
        assert_eq!(format.ext(), ext);
    }
}

// ---- fonts -------------------------------------------------------------------------------

#[test]
fn a_font_decodes_to_every_field() {
    let blob = ms_sans_serif();
    let (value, span) = decode_span(&blob, 0, PropKind::Font).unwrap();
    assert_eq!(
        value,
        FrxValue::Font(FontInfo {
            name: "MS Sans Serif".to_string(),
            size_pt: 8.25,
            weight: 400,
            charset: 0,
            raw_flags: 0,
            bold: false,
            italic: false,
            underline: false,
            strikethrough: false,
        })
    );
    assert_eq!(span, blob.len());
    assert_eq!(encode(&value, PropKind::Font), blob);
}

#[test]
fn font_styles_come_from_the_weight_and_the_flag_bits() {
    let styled = |weight: u16, flags: u8| {
        let mut b = ms_sans_serif();
        b[3] = flags;
        b[4..6].copy_from_slice(&weight.to_le_bytes());
        match decode(&b, 0, PropKind::Font).unwrap() {
            FrxValue::Font(f) => (f.bold, f.italic, f.underline, f.strikethrough, f.raw_flags),
            other => panic!("expected a font, got {other:?}"),
        }
    };
    assert_eq!(styled(400, 0x00), (false, false, false, false, 0x00));
    assert_eq!(styled(700, 0x00), (true, false, false, false, 0x00));
    assert_eq!(styled(400, 0x02), (false, true, false, false, 0x02));
    assert_eq!(styled(400, 0x04), (false, false, true, false, 0x04));
    assert_eq!(styled(400, 0x08), (false, false, false, true, 0x08));
    assert_eq!(styled(700, 0x0E), (true, true, true, true, 0x0E));
}

#[test]
fn a_font_with_another_version_is_refused() {
    let mut b = ms_sans_serif();
    b[0] = 2;
    assert_eq!(
        decode(&b, 0, PropKind::Font).unwrap_err(),
        FrxError::BadMagic { pos: 0, expected: 1, got: 2 }
    );
}

#[test]
fn a_font_name_is_windows_1252() {
    let mut b = ms_sans_serif();
    b.truncate(11);
    b[10] = 2;
    b.extend_from_slice(&[0x80, 0xE9]); // euro sign, e-acute
    match decode(&b, 0, PropKind::Font).unwrap() {
        FrxValue::Font(f) => assert_eq!(f.name, "\u{20AC}\u{E9}"),
        other => panic!("expected a font, got {other:?}"),
    }
}

// ---- strings -----------------------------------------------------------------------------

#[test]
fn short_and_long_strings_decode_with_their_prefix_width() {
    let mut short = vec![5u8];
    short.extend_from_slice(b"Hello");
    assert_eq!(
        decode_span(&short, 0, PropKind::StringShort).unwrap(),
        (FrxValue::Text("Hello".to_string()), 6)
    );

    let mut long = 5u32.to_le_bytes().to_vec();
    long.extend_from_slice(b"World");
    assert_eq!(
        decode_span(&long, 0, PropKind::StringLong).unwrap(),
        (FrxValue::Text("World".to_string()), 9)
    );
}

#[test]
fn a_string_longer_than_the_file_is_a_length_overflow() {
    assert_eq!(
        decode(&[9, b'a', b'b'], 0, PropKind::StringShort).unwrap_err(),
        FrxError::LengthOverflow { pos: 0, declared: 9, remaining: 2 }
    );
    assert_eq!(
        decode(&[9, 0, 0, 0, b'a'], 0, PropKind::StringLong).unwrap_err(),
        FrxError::LengthOverflow { pos: 0, declared: 9, remaining: 1 }
    );
}

#[test]
fn a_string_is_windows_1252_text() {
    let blob = [3u8, 0x80, 0x99, 0xE9];
    assert_eq!(
        decode(&blob, 0, PropKind::StringShort).unwrap(),
        FrxValue::Text("\u{20AC}\u{2122}\u{E9}".to_string())
    );
}

#[test]
fn a_string_of_utf16_code_units_is_read_as_utf16() {
    let units: Vec<u8> = "Caption".encode_utf16().flat_map(|u| u.to_le_bytes()).collect();
    let mut blob = (units.len() as u32).to_le_bytes().to_vec();
    blob.extend_from_slice(&units);
    assert_eq!(
        decode(&blob, 0, PropKind::StringLong).unwrap(),
        FrxValue::Text("Caption".to_string())
    );
}

#[test]
fn text_is_encoded_back_in_the_width_of_its_property_kind() {
    let text = FrxValue::Text("a\u{E9}".to_string());
    assert_eq!(encode(&text, PropKind::StringShort), vec![2, b'a', 0xE9]);
    assert_eq!(encode(&text, PropKind::StringLong), vec![2, 0, 0, 0, b'a', 0xE9]);
}

#[test]
fn a_character_with_no_windows_1252_byte_is_encoded_as_a_question_mark() {
    let text = FrxValue::Text("a\u{4E2D}".to_string());
    assert_eq!(encode(&text, PropKind::StringShort), vec![2, b'a', b'?']);
}

// ---- lists -------------------------------------------------------------------------------

#[test]
fn a_list_keeps_its_signature_and_items() {
    let blob = list_blob();
    let (value, span) = decode_span(&blob, 0, PropKind::List).unwrap();
    assert_eq!(
        value,
        FrxValue::List { items: vec!["abc".to_string(), "de".to_string()], sig: 0x000B }
    );
    assert_eq!(span, blob.len());
    assert_eq!(encode(&value, PropKind::List), blob);
}

#[test]
fn an_empty_list_is_only_its_count() {
    let blob = [0u8, 0, 0xAA, 0xBB];
    assert_eq!(
        decode_span(&blob, 0, PropKind::List).unwrap(),
        (FrxValue::List { items: Vec::new(), sig: 0 }, 2)
    );
    assert_eq!(
        decode_span(&blob, 0, PropKind::ItemData).unwrap(),
        (FrxValue::ItemData { items: Vec::new(), sig: 0 }, 2)
    );
}

#[test]
fn item_data_keeps_the_stored_bytes_of_each_value() {
    let mut blob = Vec::new();
    blob.extend_from_slice(&3u16.to_le_bytes());
    blob.extend_from_slice(&0x0003u16.to_le_bytes());
    for item in [&[0x30u8][..], &[0x01, 0x02][..], &[0xFF, 0xFF, 0xFF, 0xFF][..]] {
        blob.extend_from_slice(&(item.len() as u16).to_le_bytes());
        blob.extend_from_slice(item);
    }
    let (value, span) = decode_span(&blob, 0, PropKind::ItemData).unwrap();
    assert_eq!(
        value,
        FrxValue::ItemData {
            items: vec![vec![0x30], vec![0x01, 0x02], vec![0xFF, 0xFF, 0xFF, 0xFF]],
            sig: 0x0003
        }
    );
    assert_eq!(span, blob.len());
    assert_eq!(encode(&value, PropKind::ItemData), blob);
    assert_eq!(itemdata_value(&[0x30]), 48);
    assert_eq!(itemdata_value(&[0x01, 0x02]), 0x0201);
    assert_eq!(itemdata_value(&[0xFF, 0xFF, 0xFF, 0xFF]), -1);
}

#[test]
fn a_list_item_longer_than_the_file_is_a_length_overflow() {
    let mut blob = list_blob();
    blob.truncate(blob.len() - 1);
    assert_eq!(
        decode(&blob, 0, PropKind::List).unwrap_err(),
        FrxError::LengthOverflow { pos: 9, declared: 2, remaining: 1 }
    );
}

#[test]
fn property_pages_are_the_names_without_their_terminators() {
    let mut blob = 2u32.to_le_bytes().to_vec();
    blob.extend_from_slice(&10u16.to_le_bytes());
    blob.extend_from_slice(b"PPGeneral\0");
    blob.extend_from_slice(&6u16.to_le_bytes());
    blob.extend_from_slice(b"PPCol\0");
    let (value, span) = decode_span(&blob, 0, PropKind::PropertyPages).unwrap();
    assert_eq!(
        value,
        FrxValue::PropertyPages(vec!["PPGeneral".to_string(), "PPCol".to_string()])
    );
    assert_eq!(span, blob.len());
    assert_eq!(encode(&value, PropKind::PropertyPages), blob);
}

// ---- control bags ------------------------------------------------------------------------

#[test]
fn a_framed_bag_with_a_class_id_exposes_it() {
    let mut body = CLSID.to_vec();
    body.extend_from_slice(&[1, 2, 3, 4]);
    let mut blob = (body.len() as u32).to_le_bytes().to_vec();
    blob.extend_from_slice(&body);
    blob.extend_from_slice(&[0xEE; 3]); // the next blob
    let (value, span) = decode_span(&blob, 0, PropKind::OcxBag).unwrap();
    assert_eq!(value, FrxValue::OcxBag { clsid: Some(CLSID), data: body.clone() });
    assert_eq!(span, 4 + body.len());
    assert_eq!(decode_ocx_bag(&blob, 0), value);
    assert_eq!(encode(&value, PropKind::OcxBag), blob[..4 + body.len()].to_vec());
}

#[test]
fn a_framed_bag_of_printable_bytes_has_no_class_id() {
    let body = b"Hello, this is plain text".to_vec();
    let mut blob = (body.len() as u32).to_le_bytes().to_vec();
    blob.extend_from_slice(&body);
    let (value, span) = decode_span(&blob, 0, PropKind::OcxBag).unwrap();
    assert_eq!(value, FrxValue::OcxBag { clsid: None, data: body.clone() });
    assert_eq!(span, blob.len());
}

#[test]
fn a_bag_whose_length_runs_past_the_file_keeps_what_is_there() {
    let mut blob = 100u32.to_le_bytes().to_vec();
    blob.extend_from_slice(b"short body");
    let (value, span) = decode_span(&blob, 0, PropKind::OcxBag).unwrap();
    assert_eq!(value, FrxValue::OcxBag { clsid: None, data: b"short body".to_vec() });
    assert_eq!(span, blob.len());
}

#[test]
fn a_bag_with_a_zero_length_claims_the_rest_of_the_file() {
    let blob = [0u8, 0, 0, 0, 7, 8, 9];
    let (value, span) = decode_span(&blob, 0, PropKind::OcxBag).unwrap();
    assert_eq!(value, FrxValue::OcxBag { clsid: None, data: blob.to_vec() });
    assert_eq!(span, 7);
}

#[test]
fn a_bag_at_the_end_of_the_file_is_empty() {
    let blob = [1u8, 2, 3];
    let (value, span) = decode_span(&blob, 3, PropKind::OcxBag).unwrap();
    assert_eq!(value, FrxValue::OcxBag { clsid: None, data: Vec::new() });
    assert_eq!(span, 0);
}

#[test]
fn a_bag_offset_past_the_end_of_the_file_is_out_of_range() {
    assert_eq!(
        decode_span(&[1u8, 2, 3], 9, PropKind::OcxBag).unwrap_err(),
        FrxError::SeekOutOfRange { target: 9, len: 3 }
    );
}

// ---- offsets and re-encoding -------------------------------------------------------------

#[test]
fn an_offset_past_the_end_of_the_file_is_out_of_range() {
    assert_eq!(
        decode(&[0u8; 4], 99, PropKind::Picture).unwrap_err(),
        FrxError::SeekOutOfRange { target: 99, len: 4 }
    );
}

#[test]
fn an_offset_at_the_end_of_the_file_has_no_blob_to_read() {
    assert_eq!(
        decode(&[0u8; 4], 4, PropKind::Picture).unwrap_err(),
        FrxError::UnexpectedEof { pos: 4, needed: 4, available: 0 }
    );
}

#[test]
fn a_live_decoded_bag_has_no_byte_form() {
    let bag = FrxValue::DecodedBag { clsid: None, properties: Vec::new() };
    assert_eq!(encode(&bag, PropKind::OcxBag), Vec::<u8>::new());
}

#[test]
fn an_empty_value_encodes_as_an_empty_picture_slot() {
    assert_eq!(encode(&FrxValue::Empty, PropKind::Picture), picture(&[]));
}

#[test]
fn every_kind_round_trips_byte_for_byte() {
    let blobs: [(PropKind, Vec<u8>); 5] = [
        (PropKind::Picture, picture(&[0x42, 0x4D, 0, 0])),
        (PropKind::Picture, picture_with_clsid(CLSID, &[0x89, 0x50, 0x4E, 0x47])),
        (PropKind::Font, ms_sans_serif()),
        (PropKind::List, list_blob()),
        (PropKind::StringLong, {
            let mut b = 4u32.to_le_bytes().to_vec();
            b.extend_from_slice(b"Text");
            b
        }),
    ];
    for (kind, blob) in blobs {
        let (value, span) = decode_span(&blob, 0, kind).unwrap();
        assert_eq!(span, blob.len(), "{kind:?}");
        assert_eq!(encode(&value, kind), blob, "{kind:?}");
    }
}
