//! FRX Binary Resource File Parser
//!
//! Parses VB6 Form Resource (.frx) files which contain binary data
//! such as images, icons, and list items referenced from .frm files.
//!
//! Based on vb6parse-master FRX parsing.

use std::fs::File;
use std::io::{self, Read, Seek, SeekFrom};
use std::path::Path;

/// Represents a parsed FRX resource file
#[derive(Debug, Clone)]
pub struct FrxFile {
    /// Path to the FRX file
    pub path: std::path::PathBuf,
    /// Resources indexed by offset
    pub resources: Vec<FrxResource>,
}

/// A resource entry in an FRX file
#[derive(Debug, Clone)]
pub struct FrxResource {
    /// Offset in the file where this resource starts
    pub offset: u32,
    /// Type of resource
    pub resource_type: FrxResourceType,
    /// Size of the resource data in bytes
    pub size: u32,
    /// Raw data (optional, loaded on demand)
    pub data: Option<Vec<u8>>,
}

/// Types of resources that can be stored in FRX files
#[derive(Debug, Clone, PartialEq, Eq)]
pub enum FrxResourceType {
    /// Binary blob (image, icon, cursor, etc.)
    BinaryBlob,
    /// String resource (long property values)
    String(String),
    /// List items for ListBox/ComboBox
    ListItems(Vec<String>),
    /// Unknown resource type
    Unknown,
}

impl FrxResource {
    /// Get a description of this resource for hover info
    pub fn description(&self) -> String {
        match &self.resource_type {
            FrxResourceType::BinaryBlob => {
                format!("Binary resource: {} bytes at offset 0x{:04X}", self.size, self.offset)
            }
            FrxResourceType::String(s) => {
                let preview = if s.len() > 50 {
                    format!("{}...", &s[..50])
                } else {
                    s.clone()
                };
                format!("String resource: \"{}\"", preview)
            }
            FrxResourceType::ListItems(items) => {
                format!("List items: {} entries", items.len())
            }
            FrxResourceType::Unknown => {
                format!("Unknown resource: {} bytes at offset 0x{:04X}", self.size, self.offset)
            }
        }
    }
}

impl FrxFile {
    /// Parse an FRX file from disk
    pub fn parse(path: &Path) -> io::Result<Self> {
        let mut file = File::open(path)?;
        let metadata = file.metadata()?;
        let file_size = metadata.len() as u32;

        let mut resources = Vec::new();

        // FRX files don't have a header - they're just concatenated resources
        // Each resource is referenced by offset from the .frm file
        // We scan the file looking for recognizable patterns

        // For now, we just record that the file exists and its size
        // Resources are loaded on-demand when we know the offset from .frm parsing

        resources.push(FrxResource {
            offset: 0,
            resource_type: FrxResourceType::Unknown,
            size: file_size,
            data: None,
        });

        Ok(FrxFile {
            path: path.to_path_buf(),
            resources,
        })
    }

    /// Read a binary blob at a specific offset
    ///
    /// VB6 binary blobs have a 12-byte header:
    /// - Bytes 0-3: Unknown/signature
    /// - Bytes 4-7: Data length (little-endian)
    /// - Bytes 8-11: Unknown/flags
    pub fn read_binary_blob(&self, offset: u32) -> io::Result<FrxResource> {
        let mut file = File::open(&self.path)?;
        file.seek(SeekFrom::Start(offset as u64))?;

        // Read the 12-byte header
        let mut header = [0u8; 12];
        file.read_exact(&mut header)?;

        // Extract data length from bytes 4-7 (little-endian)
        let data_length = u32::from_le_bytes([header[4], header[5], header[6], header[7]]);

        // Read the actual data
        let mut data = vec![0u8; data_length as usize];
        file.read_exact(&mut data)?;

        let total_size = 12 + data_length;

        Ok(FrxResource {
            offset,
            resource_type: FrxResourceType::BinaryBlob,
            size: total_size,
            data: Some(data),
        })
    }

    /// Read a string resource at a specific offset
    ///
    /// String resources are length-prefixed:
    /// - Bytes 0-3: String length (little-endian)
    /// - Bytes 4+: String data (typically Windows-1252 encoded)
    pub fn read_string(&self, offset: u32) -> io::Result<FrxResource> {
        let mut file = File::open(&self.path)?;
        file.seek(SeekFrom::Start(offset as u64))?;

        // Read the 4-byte length prefix
        let mut len_bytes = [0u8; 4];
        file.read_exact(&mut len_bytes)?;
        let string_length = u32::from_le_bytes(len_bytes);

        // Read the string data
        let mut data = vec![0u8; string_length as usize];
        file.read_exact(&mut data)?;

        // Convert from Windows-1252 to UTF-8
        // For simplicity, we'll use lossy conversion
        let string = String::from_utf8_lossy(&data).into_owned();

        let total_size = 4 + string_length;

        Ok(FrxResource {
            offset,
            resource_type: FrxResourceType::String(string),
            size: total_size,
            data: Some(data),
        })
    }

    /// Read list items at a specific offset
    ///
    /// List items are stored as:
    /// - 4 bytes: Number of items
    /// - For each item:
    ///   - 4 bytes: Item length
    ///   - N bytes: Item text
    pub fn read_list_items(&self, offset: u32) -> io::Result<FrxResource> {
        let mut file = File::open(&self.path)?;
        file.seek(SeekFrom::Start(offset as u64))?;

        // Read item count
        let mut count_bytes = [0u8; 4];
        file.read_exact(&mut count_bytes)?;
        let item_count = u32::from_le_bytes(count_bytes);

        let mut items = Vec::with_capacity(item_count as usize);
        let mut total_bytes = 4u32;

        for _ in 0..item_count {
            // Read item length
            let mut len_bytes = [0u8; 4];
            file.read_exact(&mut len_bytes)?;
            let item_length = u32::from_le_bytes(len_bytes);
            total_bytes += 4;

            // Read item text
            let mut data = vec![0u8; item_length as usize];
            file.read_exact(&mut data)?;
            total_bytes += item_length;

            let item_text = String::from_utf8_lossy(&data).into_owned();
            items.push(item_text);
        }

        Ok(FrxResource {
            offset,
            resource_type: FrxResourceType::ListItems(items),
            size: total_bytes,
            data: None,
        })
    }

    /// Get resource at offset (tries to determine type automatically)
    pub fn get_resource(&self, offset: u32) -> io::Result<FrxResource> {
        // Try reading as binary blob first (most common)
        if let Ok(resource) = self.read_binary_blob(offset) {
            return Ok(resource);
        }

        // Fall back to string
        if let Ok(resource) = self.read_string(offset) {
            return Ok(resource);
        }

        // Return unknown resource
        Ok(FrxResource {
            offset,
            resource_type: FrxResourceType::Unknown,
            size: 0,
            data: None,
        })
    }

    /// Check if the FRX file exists
    pub fn exists(&self) -> bool {
        self.path.exists()
    }

    /// Get the file size
    pub fn file_size(&self) -> io::Result<u64> {
        std::fs::metadata(&self.path).map(|m| m.len())
    }
}

/// Parse an FRX offset reference from a .frm property value
///
/// In .frm files, FRX references look like: $"frmMain.frx":0000
/// This extracts the offset (0000 in this example)
pub fn parse_frx_reference(value: &str) -> Option<(String, u32)> {
    // Pattern: $"filename.frx":XXXX where XXXX is hex offset
    if !value.starts_with("$\"") {
        return None;
    }

    let value = value.trim_start_matches("$\"");

    // Find the closing quote and colon
    let parts: Vec<&str> = value.splitn(2, "\":").collect();
    if parts.len() != 2 {
        return None;
    }

    let filename = parts[0].to_string();
    let offset_str = parts[1].trim();

    // Parse hex offset
    let offset = u32::from_str_radix(offset_str, 16).ok()?;

    Some((filename, offset))
}

/// Detect the type of image from its header bytes
pub fn detect_image_type(data: &[u8]) -> Option<&'static str> {
    if data.len() < 8 {
        return None;
    }

    // Check magic bytes
    if data.starts_with(&[0x89, 0x50, 0x4E, 0x47]) {
        Some("PNG")
    } else if data.starts_with(&[0xFF, 0xD8, 0xFF]) {
        Some("JPEG")
    } else if data.starts_with(&[0x47, 0x49, 0x46, 0x38]) {
        Some("GIF")
    } else if data.starts_with(&[0x42, 0x4D]) {
        Some("BMP")
    } else if data.starts_with(&[0x00, 0x00, 0x01, 0x00]) {
        Some("ICO")
    } else if data.starts_with(&[0x00, 0x00, 0x02, 0x00]) {
        Some("CUR")
    } else {
        None
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn test_parse_frx_reference() {
        let result = parse_frx_reference("$\"frmMain.frx\":0000");
        assert!(result.is_some());
        let (filename, offset) = result.unwrap();
        assert_eq!(filename, "frmMain.frx");
        assert_eq!(offset, 0);

        let result2 = parse_frx_reference("$\"frmMain.frx\":00AC");
        assert!(result2.is_some());
        let (_, offset2) = result2.unwrap();
        assert_eq!(offset2, 0xAC);
    }

    #[test]
    fn test_detect_image_type() {
        assert_eq!(detect_image_type(&[0x89, 0x50, 0x4E, 0x47, 0x0D, 0x0A, 0x1A, 0x0A]), Some("PNG"));
        assert_eq!(detect_image_type(&[0xFF, 0xD8, 0xFF, 0xE0]), Some("JPEG"));
        assert_eq!(detect_image_type(&[0x47, 0x49, 0x46, 0x38, 0x39, 0x61]), Some("GIF"));
        assert_eq!(detect_image_type(&[0x42, 0x4D, 0x00, 0x00]), Some("BMP"));
        assert_eq!(detect_image_type(&[0x00, 0x00, 0x01, 0x00]), Some("ICO"));
    }
}
