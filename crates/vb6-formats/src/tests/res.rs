use super::*;

#[test]
fn test_resource_type_conversion() {
    assert_eq!(ResourceType::from_id(6), ResourceType::String);
    assert_eq!(ResourceType::String.to_id(), Some(6));

    assert_eq!(ResourceType::from_id(3), ResourceType::Icon);
    assert_eq!(ResourceType::Icon.to_id(), Some(3));
}

#[test]
fn test_align_to_dword() {
    assert_eq!(align_to_dword(0), 0);
    assert_eq!(align_to_dword(1), 4);
    assert_eq!(align_to_dword(2), 4);
    assert_eq!(align_to_dword(3), 4);
    assert_eq!(align_to_dword(4), 4);
    assert_eq!(align_to_dword(5), 8);
}

#[test]
fn test_empty_header() {
    let header = ResHeader::empty();
    assert_eq!(header.data_size, 0);

    let buffer = Vec::new();
    let mut writer = PositionWriter::new(buffer);
    header.write_to(&mut writer).unwrap();
    let buffer = writer.into_inner();
    assert!(!buffer.is_empty());
}

#[test]
fn test_resource_entry() {
    let entry = ResourceEntry::new(
        ResourceType::Bitmap,
        ResourceId::Id(200),
        0x0409,
        vec![1, 2, 3, 4],
    );

    assert_eq!(entry.data.len(), 4);
    assert_eq!(entry.language_id, 0x0409);

    let header = entry.to_header();
    assert_eq!(header.data_size, 4);
}

#[test]
fn test_write_and_read_res_file() {
    let entries = vec![
        ResourceEntry::new(
            ResourceType::String,
            ResourceId::Id(1),
            0x0409,
            vec![1, 2, 3, 4],
        ),
        ResourceEntry::new(
            ResourceType::Bitmap,
            ResourceId::Id(100),
            0x0409,
            vec![5, 6, 7, 8],
        ),
    ];

    // Write to temp file
    let temp_path = "test_output.res";
    write_res_file(temp_path, &entries).unwrap();

    // Read back
    let read_entries = read_res_file(temp_path).unwrap();
    assert_eq!(read_entries.len(), 2);

    // Clean up
    let _ = std::fs::remove_file(temp_path);
}
