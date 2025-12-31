//! VB6 FRX (Form Resource) binary file parser
//!
//! FRX files contain binary resource data for VB6 visual designers including:
//! - Images (icons, pictures)
//! - Binary data for controls
//! - List items
//! - Font data
//! - Other UI element resources
//!
//! ## Universal Format
//!
//! The FRX binary format is **identical** for all VB6 visual designer files:
//! - `.frm` files (Forms) → `.frx` companions
//! - `.ctl` files (UserControls) → `.frx` companions
//! - `.pag` files (Property Pages) → `.frx` companions
//! - `.dob` files (UserDocuments) → `.frx` companions
//!
//! The parser works universally for all these file types since they all use the same
//! binary format with variable-length record headers accessed by offset.
//!
//! The format consists of records with variable-length headers and data.
//! There is no overall file header - records are accessed by offset from the parent file.

use std::io;

/// Resolves a resource from a VB6 FRX file at the given offset.
///
/// This function works identically for all VB6 visual designer file types:
/// - Forms (`.frm` → `.frx`)
/// - UserControls (`.ctl` → `.frx`)
/// - Property Pages (`.pag` → `.frx`)
/// - UserDocuments (`.dob` → `.frx`)
///
/// The FRX binary format is universal - VB6 uses the same encoding scheme
/// regardless of the parent file type.
///
/// # Arguments
///
/// * `file_path` - Path to the .frx file (any type)
/// * `offset` - Byte offset into the file where the resource starts
///
/// # Returns
///
/// The resource data as a vector of bytes, or an IO error.
///
/// # Examples
///
/// ```no_run
/// use vb6_lsp::workspace::resource_file_resolver;
///
/// // Works for Form FRX files
/// let form_icon = resource_file_resolver("MyForm.frx", 0x0C).unwrap();
///
/// // Works for UserControl FRX files
/// let control_image = resource_file_resolver("MyControl.frx", 0x18).unwrap();
///
/// // Works for Property Page FRX files
/// let page_resource = resource_file_resolver("MyPage.frx", 0x24).unwrap();
/// ```
///
/// # FRX Record Format
///
/// VB6 FRX files contain multiple record types:
///
/// 1. **12-byte header records** (signature: `lt\0\0` at offset+4):
///    - Bytes 0-3: Total size (including 8-byte prefix)
///    - Bytes 4-7: Magic signature `lt\0\0`
///    - Bytes 8-11: Data size (should equal total_size - 8)
///    - Bytes 12+: Actual data
///
/// 2. **16-bit records** (first byte is 0xFF):
///    - Byte 0: 0xFF marker
///    - Bytes 1-2: Data size (u16 little-endian)
///    - Bytes 3+: Actual data
///
/// 3. **List records** (signature: 0x03 0x00 or 0x07 0x00 at offset+2):
///    - Bytes 0-1: Number of list items (u16)
///    - Bytes 2-3: List type signature
///    - For each item:
///      - 2 bytes: Item size (u16)
///      - N bytes: Item data (no null terminator)
///
/// 4. **4-byte header records** (contains null bytes in first 4 bytes):
///    - Bytes 0-3: Data size (u32 little-endian)
///    - Bytes 4+: Actual data
///
/// 5. **8-bit records** (default/fallback):
///    - Byte 0: Data size (u8)
///    - Bytes 1+: Actual data
///
pub fn resource_file_resolver(file_path: &str, offset: usize) -> Result<Vec<u8>, io::Error> {
    // Load the entire FRX file
    let buffer = std::fs::read(file_path).map_err(|err| {
        io::Error::new(
            io::ErrorKind::NotFound,
            format!("Failed to read resource file {}: {}", file_path, err),
        )
    })?;

    // Validate offset
    if offset >= buffer.len() {
        return Err(io::Error::new(
            io::ErrorKind::InvalidInput,
            format!("Offset {} is out of bounds for resource file {} (size: {})",
                    offset, file_path, buffer.len()),
        ));
    }

    // Check for 12-byte header with "lt\0\0" signature
    if buffer.len() >= offset + 12 {
        let signature = &buffer[offset + 4..offset + 8];
        if signature == b"lt\0\0" {
            return parse_12_byte_header(&buffer, file_path, offset);
        }
    }

    // Check for 16-bit record (starts with 0xFF)
    if buffer[offset] == 0xFF {
        return parse_16bit_record(&buffer, file_path, offset);
    }

    // Check for list record
    if buffer.len() >= offset + 4 {
        let list_signature = &buffer[offset + 2..offset + 4];
        if list_signature == [0x03, 0x00] || list_signature == [0x07, 0x00] {
            return parse_list_record(&buffer, file_path, offset);
        }
    }

    // Check for 4-byte header (contains null bytes)
    if buffer.len() >= offset + 12 && buffer[offset..offset + 4].contains(&0u8) {
        return parse_4_byte_header(&buffer, file_path, offset);
    }

    // Default: 8-bit record
    parse_8bit_record(&buffer, file_path, offset)
}

/// Parse a 12-byte header record with "lt\0\0" signature
fn parse_12_byte_header(buffer: &[u8], file_path: &str, offset: usize) -> Result<Vec<u8>, io::Error> {
    let size_buffer = &buffer[offset..offset + 4];
    let buffer_size_1 = u32::from_le_bytes(
        size_buffer.try_into().map_err(|_| {
            io::Error::new(
                io::ErrorKind::InvalidData,
                format!("Failed to read size buffer for {}", file_path),
            )
        })?
    ) as usize;

    let secondary_size_buffer = &buffer[offset + 8..offset + 12];
    let buffer_size_2 = u32::from_le_bytes(
        secondary_size_buffer.try_into().map_err(|_| {
            io::Error::new(
                io::ErrorKind::InvalidData,
                format!("Failed to read secondary size buffer for {}", file_path),
            )
        })?
    ) as usize;

    // Special case: empty record (e.g., removed icon)
    if buffer_size_1 == 8 && buffer_size_2 == 0 {
        return Ok(vec![]);
    }

    // Validate size consistency
    if buffer_size_2 != buffer_size_1.saturating_sub(8) {
        return Err(io::Error::new(
            io::ErrorKind::InvalidData,
            format!(
                "Record size mismatch in {}: {} != {} - 8. Likely corrupted.",
                file_path, buffer_size_2, buffer_size_1
            ),
        ));
    }

    let record_start = offset + 12;
    let record_end = record_start + buffer_size_2;

    if record_end > buffer.len() {
        return Err(io::Error::new(
            io::ErrorKind::InvalidData,
            format!("Record end {} out of bounds in {}", record_end, file_path),
        ));
    }

    Ok(buffer[record_start..record_end].to_vec())
}

/// Parse a 16-bit record (starts with 0xFF)
fn parse_16bit_record(buffer: &[u8], file_path: &str, offset: usize) -> Result<Vec<u8>, io::Error> {
    let header_size_offset = offset + 1;
    let header_size_end = header_size_offset + 2;

    if header_size_end > buffer.len() {
        return Err(io::Error::new(
            io::ErrorKind::InvalidData,
            format!("Header size element out of bounds in {}", file_path),
        ));
    }

    let size_bytes = &buffer[header_size_offset..header_size_end];
    let mut record_size = u16::from_le_bytes(
        size_bytes.try_into().map_err(|_| {
            io::Error::new(
                io::ErrorKind::InvalidData,
                format!("Failed to read header size in {}", file_path),
            )
        })?
    ) as usize;

    // VB6 off-by-one error: sometimes the size is 1 byte too large
    // This usually happens with string resources
    if header_size_offset + record_size > buffer.len() {
        record_size = record_size.saturating_sub(1);
    }

    let record_offset = header_size_end;
    let record_start = record_offset;
    let record_end = record_offset + record_size;

    if record_start > buffer.len() {
        return Err(io::Error::new(
            io::ErrorKind::InvalidData,
            format!("Record start {} out of bounds in {}", record_start, file_path),
        ));
    }

    if record_end > buffer.len() {
        return Err(io::Error::new(
            io::ErrorKind::InvalidData,
            format!("Record end {} out of bounds in {}", record_end, file_path),
        ));
    }

    Ok(buffer[record_start..record_end].to_vec())
}

/// Parse a list record (list items)
fn parse_list_record(buffer: &[u8], file_path: &str, offset: usize) -> Result<Vec<u8>, io::Error> {
    let item_count_bytes = &buffer[offset..offset + 2];
    let list_item_count = u16::from_le_bytes(
        item_count_bytes.try_into().map_err(|_| {
            io::Error::new(
                io::ErrorKind::InvalidData,
                format!("Failed to read list item count in {}", file_path),
            )
        })?
    ) as usize;

    let header_size = 4;
    let mut record_offset = offset + header_size;

    // Read all list items
    for _ in 0..list_item_count {
        if record_offset + 2 > buffer.len() {
            return Err(io::Error::new(
                io::ErrorKind::InvalidData,
                format!("List item header out of bounds in {}", file_path),
            ));
        }

        let size_bytes = &buffer[record_offset..record_offset + 2];
        let list_item_size = u16::from_le_bytes(
            size_bytes.try_into().map_err(|_| {
                io::Error::new(
                    io::ErrorKind::InvalidData,
                    format!("Failed to read list item size in {}", file_path),
                )
            })?
        ) as usize;

        record_offset += 2 + list_item_size;
    }

    if record_offset > buffer.len() {
        return Err(io::Error::new(
            io::ErrorKind::InvalidData,
            format!("List record end {} out of bounds in {}", record_offset, file_path),
        ));
    }

    Ok(buffer[offset..record_offset].to_vec())
}

/// Parse a 4-byte header record
fn parse_4_byte_header(buffer: &[u8], file_path: &str, offset: usize) -> Result<Vec<u8>, io::Error> {
    let size_bytes = &buffer[offset..offset + 4];
    let record_size = u32::from_le_bytes(
        size_bytes.try_into().map_err(|_| {
            io::Error::new(
                io::ErrorKind::InvalidData,
                format!("Failed to read record size in {}", file_path),
            )
        })?
    ) as usize;

    let header_size = 4;
    let record_start = offset + header_size;
    let record_end = record_start + record_size;

    if record_end > buffer.len() {
        return Err(io::Error::new(
            io::ErrorKind::InvalidData,
            format!("Record end {} out of bounds in {}", record_end, file_path),
        ));
    }

    Ok(buffer[record_start..record_end].to_vec())
}

/// Parse an 8-bit record (default/fallback)
fn parse_8bit_record(buffer: &[u8], file_path: &str, offset: usize) -> Result<Vec<u8>, io::Error> {
    let header_size = 1;
    let mut record_size = buffer[offset] as usize;
    let record_start = offset + header_size;

    // VB6 off-by-one error handling
    let record_end = if record_size >= buffer.len() {
        record_start + record_size - 1
    } else {
        record_start + record_size
    };

    if record_start > buffer.len() {
        return Err(io::Error::new(
            io::ErrorKind::InvalidData,
            format!("Record start {} out of bounds in {}", record_start, file_path),
        ));
    }

    if record_end > buffer.len() {
        return Err(io::Error::new(
            io::ErrorKind::InvalidData,
            format!("Record end {} out of bounds in {}", record_end, file_path),
        ));
    }

    Ok(buffer[record_start..record_end].to_vec())
}

/// Parses a list from FRX binary data.
///
/// Lists are used for ComboBox/ListBox items and similar controls.
///
/// # Format
/// - Bytes 0-1: Number of items (u16 little-endian)
/// - Bytes 2-3: List type signature (0x03 0x00 or 0x07 0x00)
/// - For each item:
///   - 2 bytes: Item size (u16)
///   - N bytes: Item data (UTF-8 or system encoding, no null terminator)
///
/// # Arguments
///
/// * `buffer` - The raw bytes from a list record
///
/// # Returns
///
/// A vector of strings representing the list items.
///
pub fn list_resolver(buffer: &[u8]) -> Vec<String> {
    let mut list_items = Vec::new();

    if buffer.len() < 2 {
        return list_items;
    }

    let item_count_bytes = match buffer[0..2].try_into() {
        Ok(bytes) => bytes,
        Err(_) => return list_items,
    };
    let list_item_count = u16::from_le_bytes(item_count_bytes) as usize;

    let header_size = 4;
    let mut record_offset = header_size;
    let list_item_header_size = 2;

    for _ in 0..list_item_count {
        if record_offset + list_item_header_size > buffer.len() {
            return list_items;
        }

        let size_bytes = match buffer[record_offset..record_offset + list_item_header_size].try_into() {
            Ok(bytes) => bytes,
            Err(_) => return list_items,
        };
        let list_item_size = u16::from_le_bytes(size_bytes) as usize;

        let item_start = record_offset + list_item_header_size;
        let item_end = item_start + list_item_size;

        if item_end > buffer.len() {
            return list_items;
        }

        let item_bytes = &buffer[item_start..item_end];

        // Try to convert to UTF-8, fallback to lossy conversion
        let item_string = String::from_utf8_lossy(item_bytes).to_string();
        list_items.push(item_string);

        record_offset += list_item_header_size + list_item_size;
    }

    list_items
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn test_universal_format_for_all_control_types() {
        // The FRX binary format is identical for .frm, .ctl, .pag, and .dob files
        // This test demonstrates that the parser works regardless of the source file type

        let test_data = vec![
            0xFF,                    // 16-bit record marker
            0x05, 0x00,              // size = 5
            0x41, 0x42, 0x43, 0x44, 0x45,  // "ABCDE"
        ];

        // Parse from a Form's FRX file
        let form_result = parse_16bit_record(&test_data, "MyForm.frx", 0).unwrap();

        // Parse from a UserControl's FRX file
        let usercontrol_result = parse_16bit_record(&test_data, "MyControl.frx", 0).unwrap();

        // Parse from a Property Page's FRX file
        let propertypage_result = parse_16bit_record(&test_data, "MyPage.frx", 0).unwrap();

        // Parse from a UserDocument's FRX file
        let userdocument_result = parse_16bit_record(&test_data, "MyDocument.frx", 0).unwrap();

        // All results should be identical - the format is universal
        assert_eq!(form_result, vec![0x41, 0x42, 0x43, 0x44, 0x45]);
        assert_eq!(usercontrol_result, vec![0x41, 0x42, 0x43, 0x44, 0x45]);
        assert_eq!(propertypage_result, vec![0x41, 0x42, 0x43, 0x44, 0x45]);
        assert_eq!(userdocument_result, vec![0x41, 0x42, 0x43, 0x44, 0x45]);
    }

    #[test]
    fn test_empty_12_byte_header() {
        // Empty record: size1=8, signature="lt\0\0", size2=0
        let data = vec![
            0x08, 0x00, 0x00, 0x00,  // size1 = 8
            b'l', b't', 0x00, 0x00,  // signature
            0x00, 0x00, 0x00, 0x00,  // size2 = 0
        ];

        let result = parse_12_byte_header(&data, "test.frx", 0).unwrap();
        assert_eq!(result.len(), 0);
    }

    #[test]
    fn test_12_byte_header_with_data() {
        // Record with 4 bytes of data
        let data = vec![
            0x0C, 0x00, 0x00, 0x00,  // size1 = 12 (8 + 4)
            b'l', b't', 0x00, 0x00,  // signature
            0x04, 0x00, 0x00, 0x00,  // size2 = 4
            0x01, 0x02, 0x03, 0x04,  // data
        ];

        let result = parse_12_byte_header(&data, "test.frx", 0).unwrap();
        assert_eq!(result, vec![0x01, 0x02, 0x03, 0x04]);
    }

    #[test]
    fn test_16bit_record() {
        // 16-bit record with 5 bytes of data
        let data = vec![
            0xFF,                    // marker
            0x05, 0x00,              // size = 5
            0x41, 0x42, 0x43, 0x44, 0x45,  // "ABCDE"
        ];

        let result = parse_16bit_record(&data, "test.frx", 0).unwrap();
        assert_eq!(result, vec![0x41, 0x42, 0x43, 0x44, 0x45]);
    }

    #[test]
    fn test_list_resolver_empty() {
        let data = vec![
            0x00, 0x00,  // 0 items
            0x03, 0x00,  // signature
        ];

        let result = list_resolver(&data);
        assert_eq!(result.len(), 0);
    }

    #[test]
    fn test_list_resolver_with_items() {
        let data = vec![
            0x02, 0x00,              // 2 items
            0x03, 0x00,              // signature
            0x05, 0x00,              // item1 size = 5
            b'I', b't', b'e', b'm', b'1',  // "Item1"
            0x05, 0x00,              // item2 size = 5
            b'I', b't', b'e', b'm', b'2',  // "Item2"
        ];

        let result = list_resolver(&data);
        assert_eq!(result.len(), 2);
        assert_eq!(result[0], "Item1");
        assert_eq!(result[1], "Item2");
    }

    #[test]
    fn test_8bit_record() {
        // 8-bit record with 3 bytes of data
        let data = vec![
            0x03,                    // size = 3
            0x41, 0x42, 0x43,        // "ABC"
        ];

        let result = parse_8bit_record(&data, "test.frx", 0).unwrap();
        assert_eq!(result, vec![0x41, 0x42, 0x43]);
    }

    #[test]
    fn test_4byte_header() {
        // 4-byte header with 6 bytes of data
        let data = vec![
            0x06, 0x00, 0x00, 0x00,  // size = 6
            0x01, 0x02, 0x03, 0x04, 0x05, 0x06,  // data
        ];

        let result = parse_4_byte_header(&data, "test.frx", 0).unwrap();
        assert_eq!(result, vec![0x01, 0x02, 0x03, 0x04, 0x05, 0x06]);
    }
}
