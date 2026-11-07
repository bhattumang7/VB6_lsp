//! JSON and name-based helpers for resource files, shared by the command-line
//! and MCP front ends.

use base64::Engine as _;

use crate::res::StringTableEntry;
use crate::res::{ResourceEntry, ResourceId, ResourceType};

/// Why a resource type name could not be turned into a [`ResourceType`].
#[derive(Debug)]
pub enum ResourceTypeNameError {
    /// `Custom(<id>)` whose id is not a valid `u16`.
    InvalidCustomId(std::num::ParseIntError),
    /// The name matches no known resource type.
    Unknown,
}

/// Parse a resource type name (e.g. `Bitmap`, `Custom(12)`, `Named("CUSTOM")`).
pub fn parse_resource_type_name(s: &str) -> Result<ResourceType, ResourceTypeNameError> {
    if let Some(known) = parse_builtin_type(s) {
        return Ok(known);
    }
    if s.starts_with("Custom(") {
        let id: u16 = s
            .trim_start_matches("Custom(")
            .trim_end_matches(')')
            .parse()
            .map_err(ResourceTypeNameError::InvalidCustomId)?;
        return Ok(ResourceType::Custom(id));
    }
    if s.starts_with("Named(") {
        let name = s.trim_start_matches("Named(\"").trim_end_matches("\")");
        return Ok(ResourceType::Named(name.to_string()));
    }
    Err(ResourceTypeNameError::Unknown)
}

fn parse_builtin_type(s: &str) -> Option<ResourceType> {
    Some(match s {
        "Cursor" => ResourceType::Cursor,
        "Bitmap" => ResourceType::Bitmap,
        "Icon" => ResourceType::Icon,
        "Menu" => ResourceType::Menu,
        "Dialog" => ResourceType::Dialog,
        "String" => ResourceType::String,
        "FontDir" => ResourceType::FontDir,
        "Font" => ResourceType::Font,
        "Accelerator" => ResourceType::Accelerator,
        "RcData" => ResourceType::RcData,
        "MessageTable" => ResourceType::MessageTable,
        "GroupCursor" => ResourceType::GroupCursor,
        "GroupIcon" => ResourceType::GroupIcon,
        "Version" => ResourceType::Version,
        "DlgInclude" => ResourceType::DlgInclude,
        "PlugPlay" => ResourceType::PlugPlay,
        "Vxd" => ResourceType::Vxd,
        "AniCursor" => ResourceType::AniCursor,
        "AniIcon" => ResourceType::AniIcon,
        "Html" => ResourceType::Html,
        "Manifest" => ResourceType::Manifest,
        "Toolbar" => ResourceType::Toolbar,
        "DlgInit" => ResourceType::DlgInit,
        _ => return None,
    })
}

/// JSON description of one resource entry, with its data encoded as Base64.
pub fn resource_entry_to_json(r: &ResourceEntry) -> serde_json::Value {
    let name = match &r.name {
        ResourceId::Id(id) => serde_json::json!({"type": "Id", "value": id}),
        ResourceId::Name(n) => serde_json::json!({"type": "Name", "value": n}),
    };
    serde_json::json!({
        "resource_type": format!("{:?}", r.resource_type),
        "name": name,
        "language_id": r.language_id,
        "data_size": r.data.len(),
        "data_base64": base64::engine::general_purpose::STANDARD.encode(&r.data),
    })
}

/// Find the string table resource with numeric id `block_id`.
pub fn find_string_table_block(
    resources: &[ResourceEntry],
    block_id: u16,
) -> Option<&ResourceEntry> {
    resources.iter().find(|r| {
        r.resource_type == ResourceType::String
            && matches!(r.name, ResourceId::Id(id) if id == block_id)
    })
}

/// JSON list of `{id, value}` objects for parsed string table entries.
pub fn string_entries_to_json(strings: &[StringTableEntry]) -> Vec<serde_json::Value> {
    strings
        .iter()
        .map(|s| serde_json::json!({"id": s.id, "value": s.value}))
        .collect()
}
