//! Shared resource JSON / name helpers used by both front ends.

use serde_json::json;
use vb6_formats::res_json::{
    find_string_table_block, parse_resource_type_name, resource_entry_to_json,
    string_entries_to_json, ResourceTypeNameError,
};
use vb6_formats::res::{parse_string_table, ResourceEntry, ResourceId, ResourceType};

#[test]
fn builtin_names_map_to_types() {
    assert_eq!(parse_resource_type_name("Bitmap").unwrap(), ResourceType::Bitmap);
    assert_eq!(parse_resource_type_name("DlgInit").unwrap(), ResourceType::DlgInit);
    assert_eq!(parse_resource_type_name("RcData").unwrap(), ResourceType::RcData);
}

#[test]
fn custom_and_named_types() {
    assert_eq!(parse_resource_type_name("Custom(12)").unwrap(), ResourceType::Custom(12));
    assert_eq!(
        parse_resource_type_name("Named(\"CUSTOM\")").unwrap(),
        ResourceType::Named("CUSTOM".to_string())
    );
}

#[test]
fn invalid_custom_id_and_unknown_name_are_distinguished() {
    assert!(matches!(
        parse_resource_type_name("Custom(abc)"),
        Err(ResourceTypeNameError::InvalidCustomId(_))
    ));
    assert!(matches!(
        parse_resource_type_name("Custom(70000)"),
        Err(ResourceTypeNameError::InvalidCustomId(_))
    ));
    assert!(matches!(
        parse_resource_type_name("Nope"),
        Err(ResourceTypeNameError::Unknown)
    ));
}

#[test]
fn entry_json_for_id_and_name() {
    let by_id = ResourceEntry::new(ResourceType::Bitmap, ResourceId::Id(7), 1033, vec![1, 2, 3]);
    assert_eq!(
        resource_entry_to_json(&by_id),
        json!({
            "resource_type": "Bitmap",
            "name": {"type": "Id", "value": 7},
            "language_id": 1033,
            "data_size": 3,
            "data_base64": "AQID",
        })
    );
    let by_name = ResourceEntry::new(
        ResourceType::Named("X".to_string()),
        ResourceId::Name("ABC".to_string()),
        0,
        vec![],
    );
    assert_eq!(
        resource_entry_to_json(&by_name),
        json!({
            "resource_type": "Named(\"X\")",
            "name": {"type": "Name", "value": "ABC"},
            "language_id": 0,
            "data_size": 0,
            "data_base64": "",
        })
    );
}

#[test]
fn string_table_block_lookup_and_json() {
    // Block 1 holds string ids 16..31; encode ids 16 and 17 ("A", "BC"), rest empty.
    let mut data = Vec::new();
    for s in ["A", "BC"] {
        data.extend_from_slice(&(s.chars().count() as u16).to_le_bytes());
        for c in s.encode_utf16() {
            data.extend_from_slice(&c.to_le_bytes());
        }
    }
    for _ in 2..16 {
        data.extend_from_slice(&0u16.to_le_bytes());
    }
    let resources = vec![
        ResourceEntry::new(ResourceType::Icon, ResourceId::Id(1), 0, vec![]),
        ResourceEntry::new(ResourceType::String, ResourceId::Name("S".to_string()), 0, vec![]),
        ResourceEntry::new(ResourceType::String, ResourceId::Id(2), 0, vec![]),
        ResourceEntry::new(ResourceType::String, ResourceId::Id(1), 0, data),
    ];
    assert!(find_string_table_block(&resources, 3).is_none());
    let block = find_string_table_block(&resources, 1).expect("block 1");
    assert_eq!(block.name, ResourceId::Id(1));
    assert_eq!(block.data.len(), (2 + 2) + (2 + 4) + 28);

    let strings = parse_string_table(&block.data, 1).unwrap();
    let rendered = string_entries_to_json(&strings);
    assert_eq!(rendered.len(), strings.len());
    assert_eq!(rendered[0], json!({"id": strings[0].id, "value": "A"}));
    assert_eq!(rendered[1], json!({"id": strings[1].id, "value": "BC"}));
}
