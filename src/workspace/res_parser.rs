//! Win32 Resource File (.res) Parser and Writer
//!
//! This module provides comprehensive support for reading and writing Win32 compiled
//! resource files (.res) used in VB6 projects. These files contain:
//! - String tables (localized text)
//! - Icons and cursors (with group resources)
//! - Bitmaps and images
//! - Dialog templates
//! - Menus and accelerators
//! - Version information
//! - Custom binary data (RCDATA)
//! - And all other standard Win32 resource types
//!
//! ## File Format
//!
//! .res files consist of a sequence of resource records, each with:
//! 1. A ResHeader (variable length, DWORD-aligned)
//! 2. Resource data (variable length)
//! 3. Padding to next DWORD boundary
//!
//! The first record in a .res file is typically an empty header (DataSize=0).
//!
//! ## Usage
//!
//! ```no_run
//! use vb6_lsp::workspace::res_parser::*;
//!
//! // Read a .res file
//! let resources = read_res_file("myproject.res")?;
//!
//! // Access string table
//! for entry in &resources {
//!     if entry.resource_type == ResourceType::String {
//!         let strings = parse_string_table(&entry.data);
//!         // ... use strings
//!     }
//! }
//!
//! // Modify and write back
//! write_res_file("output.res", &resources)?;
//! ```

use std::io::{self, Read, Write, Cursor, Seek};
use byteorder::{LittleEndian, ReadBytesExt, WriteBytesExt};

/// Standard Win32 resource types
#[derive(Debug, Clone, PartialEq, Eq)]
pub enum ResourceType {
    /// Cursor (RT_CURSOR = 1)
    Cursor,
    /// Bitmap (RT_BITMAP = 2)
    Bitmap,
    /// Icon (RT_ICON = 3)
    Icon,
    /// Menu (RT_MENU = 4)
    Menu,
    /// Dialog (RT_DIALOG = 5)
    Dialog,
    /// String table (RT_STRING = 6)
    String,
    /// Font directory (RT_FONTDIR = 7)
    FontDir,
    /// Font (RT_FONT = 8)
    Font,
    /// Accelerator table (RT_ACCELERATOR = 9)
    Accelerator,
    /// Raw data (RT_RCDATA = 10)
    RcData,
    /// Message table (RT_MESSAGETABLE = 11)
    MessageTable,
    /// Cursor group (RT_GROUP_CURSOR = 12)
    GroupCursor,
    /// Icon group (RT_GROUP_ICON = 14)
    GroupIcon,
    /// Version information (RT_VERSION = 16)
    Version,
    /// Dialog include (RT_DLGINCLUDE = 17)
    DlgInclude,
    /// Plug and Play resource (RT_PLUGPLAY = 19)
    PlugPlay,
    /// VXD (RT_VXD = 20)
    Vxd,
    /// Animated cursor (RT_ANICURSOR = 21)
    AniCursor,
    /// Animated icon (RT_ANIICON = 22)
    AniIcon,
    /// HTML (RT_HTML = 23)
    Html,
    /// Manifest (RT_MANIFEST = 24)
    Manifest,
    /// Toolbar (custom type used by VB6)
    Toolbar,
    /// DlgInit (custom type used by VB6)
    DlgInit,
    /// Custom resource type with numeric ID
    Custom(u16),
    /// Custom resource type with string name
    Named(String),
}

impl ResourceType {
    /// Convert from a numeric resource type ID
    pub fn from_id(id: u16) -> Self {
        match id {
            1 => ResourceType::Cursor,
            2 => ResourceType::Bitmap,
            3 => ResourceType::Icon,
            4 => ResourceType::Menu,
            5 => ResourceType::Dialog,
            6 => ResourceType::String,
            7 => ResourceType::FontDir,
            8 => ResourceType::Font,
            9 => ResourceType::Accelerator,
            10 => ResourceType::RcData,
            11 => ResourceType::MessageTable,
            12 => ResourceType::GroupCursor,
            14 => ResourceType::GroupIcon,
            16 => ResourceType::Version,
            17 => ResourceType::DlgInclude,
            19 => ResourceType::PlugPlay,
            20 => ResourceType::Vxd,
            21 => ResourceType::AniCursor,
            22 => ResourceType::AniIcon,
            23 => ResourceType::Html,
            24 => ResourceType::Manifest,
            _ => ResourceType::Custom(id),
        }
    }

    /// Convert to a numeric resource type ID (if applicable)
    pub fn to_id(&self) -> Option<u16> {
        match self {
            ResourceType::Cursor => Some(1),
            ResourceType::Bitmap => Some(2),
            ResourceType::Icon => Some(3),
            ResourceType::Menu => Some(4),
            ResourceType::Dialog => Some(5),
            ResourceType::String => Some(6),
            ResourceType::FontDir => Some(7),
            ResourceType::Font => Some(8),
            ResourceType::Accelerator => Some(9),
            ResourceType::RcData => Some(10),
            ResourceType::MessageTable => Some(11),
            ResourceType::GroupCursor => Some(12),
            ResourceType::GroupIcon => Some(14),
            ResourceType::Version => Some(16),
            ResourceType::DlgInclude => Some(17),
            ResourceType::PlugPlay => Some(19),
            ResourceType::Vxd => Some(20),
            ResourceType::AniCursor => Some(21),
            ResourceType::AniIcon => Some(22),
            ResourceType::Html => Some(23),
            ResourceType::Manifest => Some(24),
            ResourceType::Custom(id) => Some(*id),
            _ => None,
        }
    }
}

/// Resource name or ID
#[derive(Debug, Clone, PartialEq, Eq)]
pub enum ResourceId {
    /// Numeric ID
    Id(u16),
    /// String name
    Name(String),
}

impl ResourceId {
    /// Check if this is a numeric ID
    pub fn is_id(&self) -> bool {
        matches!(self, ResourceId::Id(_))
    }

    /// Check if this is a string name
    pub fn is_name(&self) -> bool {
        matches!(self, ResourceId::Name(_))
    }

    /// Get the numeric ID if available
    pub fn as_id(&self) -> Option<u16> {
        match self {
            ResourceId::Id(id) => Some(*id),
            _ => None,
        }
    }

    /// Get the string name if available
    pub fn as_name(&self) -> Option<&str> {
        match self {
            ResourceId::Name(name) => Some(name),
            _ => None,
        }
    }
}

/// Memory flags for resource header
#[derive(Debug, Clone, Copy)]
pub struct MemoryFlags(pub u16);

impl MemoryFlags {
    pub const MOVEABLE: u16 = 0x0010;
    pub const PURE: u16 = 0x0020;
    pub const PRELOAD: u16 = 0x0040;
    pub const DISCARDABLE: u16 = 0x1000;

    /// Default flags for most resources
    pub fn default_flags() -> Self {
        MemoryFlags(Self::MOVEABLE | Self::PURE | Self::DISCARDABLE)
    }
}

/// Resource header structure
///
/// This appears at the start of each resource record in a .res file.
#[derive(Debug, Clone)]
pub struct ResHeader {
    /// Size of the resource data
    pub data_size: u32,
    /// Size of this header structure
    pub header_size: u32,
    /// Resource type
    pub resource_type: ResourceType,
    /// Resource name/ID
    pub name: ResourceId,
    /// Data version (usually 0)
    pub data_version: u32,
    /// Memory flags
    pub memory_flags: MemoryFlags,
    /// Language ID (e.g., 0x0409 for US English)
    pub language_id: u16,
    /// Version (usually 0)
    pub version: u32,
    /// Characteristics (usually 0)
    pub characteristics: u32,
}

impl ResHeader {
    /// Create a new resource header with default values
    pub fn new(resource_type: ResourceType, name: ResourceId, language_id: u16) -> Self {
        ResHeader {
            data_size: 0,
            header_size: 32, // Will be recalculated
            resource_type,
            name,
            data_version: 0,
            memory_flags: MemoryFlags::default_flags(),
            language_id,
            version: 0,
            characteristics: 0,
        }
    }

    /// Create an empty header (used as first record in .res files)
    pub fn empty() -> Self {
        ResHeader {
            data_size: 0,
            header_size: 32,
            resource_type: ResourceType::Custom(0),
            name: ResourceId::Id(0),
            data_version: 0,
            memory_flags: MemoryFlags(0),
            language_id: 0,
            version: 0,
            characteristics: 0,
        }
    }

    /// Calculate the actual header size based on type and name
    pub fn calculate_header_size(&self) -> u32 {
        let mut size = 8; // DataSize + HeaderSize

        // Type field
        size += match &self.resource_type {
            ResourceType::Named(name) => {
                let str_len = (name.encode_utf16().count() + 1) * 2; // UTF-16 + null
                align_to_dword(str_len)
            }
            _ => 4, // 0xFFFF marker + u16 ID
        };

        // Name field
        size += match &self.name {
            ResourceId::Name(name) => {
                let str_len = (name.encode_utf16().count() + 1) * 2; // UTF-16 + null
                align_to_dword(str_len)
            }
            ResourceId::Id(_) => 4, // 0xFFFF marker + u16 ID
        };

        // Fixed fields: DataVersion + MemoryFlags + LanguageId + Version + Characteristics
        size += 16;

        align_to_dword(size) as u32
    }

    /// Read a resource header from a byte stream
    pub fn read_from<R: Read + Seek>(reader: &mut R) -> io::Result<Self> {
        // Read DataSize and HeaderSize
        let data_size = reader.read_u32::<LittleEndian>()?;
        let header_size = reader.read_u32::<LittleEndian>()?;

        // Read resource type
        let resource_type = read_id_or_string(reader)?;
        let resource_type = match resource_type {
            ResourceId::Id(id) => ResourceType::from_id(id),
            ResourceId::Name(name) => ResourceType::Named(name),
        };

        // Read resource name
        let name = read_id_or_string(reader)?;

        // Align to DWORD boundary
        align_reader_to_dword(reader)?;

        // Read remaining fixed fields
        let data_version = reader.read_u32::<LittleEndian>()?;
        let memory_flags = MemoryFlags(reader.read_u16::<LittleEndian>()?);
        let language_id = reader.read_u16::<LittleEndian>()?;
        let version = reader.read_u32::<LittleEndian>()?;
        let characteristics = reader.read_u32::<LittleEndian>()?;

        // Align to DWORD boundary
        align_reader_to_dword(reader)?;

        Ok(ResHeader {
            data_size,
            header_size,
            resource_type,
            name,
            data_version,
            memory_flags,
            language_id,
            version,
            characteristics,
        })
    }

    /// Write this resource header to a byte stream
    pub fn write_to<W: Write>(&self, writer: &mut PositionWriter<W>) -> io::Result<()> {
        // Write DataSize and HeaderSize
        writer.write_u32::<LittleEndian>(self.data_size)?;
        writer.write_u32::<LittleEndian>(self.header_size)?;

        // Write resource type
        match &self.resource_type {
            ResourceType::Named(name) => {
                write_string(writer, name)?;
            }
            _ => {
                if let Some(id) = self.resource_type.to_id() {
                    write_id(writer, id)?;
                } else {
                    write_id(writer, 0)?;
                }
            }
        }

        // Write resource name
        match &self.name {
            ResourceId::Name(name) => {
                write_string(writer, name)?;
            }
            ResourceId::Id(id) => {
                write_id(writer, *id)?;
            }
        }

        // Align to DWORD boundary
        align_writer_to_dword(writer)?;

        // Write remaining fixed fields
        writer.write_u32::<LittleEndian>(self.data_version)?;
        writer.write_u16::<LittleEndian>(self.memory_flags.0)?;
        writer.write_u16::<LittleEndian>(self.language_id)?;
        writer.write_u32::<LittleEndian>(self.version)?;
        writer.write_u32::<LittleEndian>(self.characteristics)?;

        // Align to DWORD boundary
        align_writer_to_dword(writer)?;

        Ok(())
    }
}

/// A single resource entry in a .res file
#[derive(Debug, Clone)]
pub struct ResourceEntry {
    /// Resource type
    pub resource_type: ResourceType,
    /// Resource name/ID
    pub name: ResourceId,
    /// Language ID
    pub language_id: u16,
    /// Raw resource data
    pub data: Vec<u8>,
    /// Memory flags
    pub memory_flags: MemoryFlags,
    /// Data version
    pub data_version: u32,
    /// Version
    pub version: u32,
    /// Characteristics
    pub characteristics: u32,
}

impl ResourceEntry {
    /// Create a new resource entry
    pub fn new(
        resource_type: ResourceType,
        name: ResourceId,
        language_id: u16,
        data: Vec<u8>,
    ) -> Self {
        ResourceEntry {
            resource_type,
            name,
            language_id,
            data,
            memory_flags: MemoryFlags::default_flags(),
            data_version: 0,
            version: 0,
            characteristics: 0,
        }
    }

    /// Convert this entry to a ResHeader
    pub fn to_header(&self) -> ResHeader {
        let mut header = ResHeader::new(
            self.resource_type.clone(),
            self.name.clone(),
            self.language_id,
        );
        header.data_size = self.data.len() as u32;
        header.memory_flags = self.memory_flags;
        header.data_version = self.data_version;
        header.version = self.version;
        header.characteristics = self.characteristics;
        header.header_size = header.calculate_header_size();
        header
    }
}

/// Read a complete .res file and return all resource entries
///
/// # Arguments
///
/// * `file_path` - Path to the .res file
///
/// # Returns
///
/// A vector of ResourceEntry objects, or an IO error
///
pub fn read_res_file(file_path: &str) -> io::Result<Vec<ResourceEntry>> {
    let data = std::fs::read(file_path)?;
    let mut cursor = Cursor::new(data);
    let mut entries = Vec::new();

    loop {
        // Try to read a header
        let header = match ResHeader::read_from(&mut cursor) {
            Ok(h) => h,
            Err(e) if e.kind() == io::ErrorKind::UnexpectedEof => break,
            Err(e) => return Err(e),
        };

        // Read the data
        let mut data = vec![0u8; header.data_size as usize];
        if header.data_size > 0 {
            cursor.read_exact(&mut data)?;
        }

        // Align to DWORD boundary
        align_reader_to_dword(&mut cursor)?;

        // Skip empty headers (first record in file)
        if header.data_size == 0 && entries.is_empty() {
            continue;
        }

        // Create entry
        let entry = ResourceEntry {
            resource_type: header.resource_type,
            name: header.name,
            language_id: header.language_id,
            data,
            memory_flags: header.memory_flags,
            data_version: header.data_version,
            version: header.version,
            characteristics: header.characteristics,
        };

        entries.push(entry);
    }

    Ok(entries)
}

/// Write resource entries to a .res file
///
/// # Arguments
///
/// * `file_path` - Path to the output .res file
/// * `entries` - Vector of resource entries to write
///
pub fn write_res_file(file_path: &str, entries: &[ResourceEntry]) -> io::Result<()> {
    let buffer = Vec::new();
    let mut writer = PositionWriter::new(buffer);

    // Write empty header first (standard .res file format)
    let empty_header = ResHeader::empty();
    empty_header.write_to(&mut writer)?;

    // Write all entries
    for entry in entries {
        let header = entry.to_header();
        header.write_to(&mut writer)?;

        // Write data
        writer.write_all(&entry.data)?;

        // Align to DWORD boundary
        align_writer_to_dword(&mut writer)?;
    }

    let buffer = writer.into_inner();
    std::fs::write(file_path, buffer)?;
    Ok(())
}

// =============================================================================
// Helper Functions
// =============================================================================

/// Align a value to the next DWORD (4-byte) boundary
fn align_to_dword(value: usize) -> usize {
    (value + 3) & !3
}

/// Position-tracking reader wrapper for proper DWORD alignment
struct PositionTracker<R> {
    inner: R,
    position: usize,
}

impl<R: Read> PositionTracker<R> {
    fn new(inner: R) -> Self {
        PositionTracker { inner, position: 0 }
    }

    fn position(&self) -> usize {
        self.position
    }

    fn align_to_dword(&mut self) -> io::Result<()> {
        let remainder = self.position % 4;
        if remainder != 0 {
            let padding = 4 - remainder;
            let mut buf = vec![0u8; padding];
            self.inner.read_exact(&mut buf)?;
            self.position += padding;
        }
        Ok(())
    }
}

impl<R: Read> Read for PositionTracker<R> {
    fn read(&mut self, buf: &mut [u8]) -> io::Result<usize> {
        let n = self.inner.read(buf)?;
        self.position += n;
        Ok(n)
    }
}

/// Position-tracking writer wrapper for proper DWORD alignment
pub(crate) struct PositionWriter<W> {
    inner: W,
    position: usize,
}

impl<W: Write> PositionWriter<W> {
    fn new(inner: W) -> Self {
        PositionWriter { inner, position: 0 }
    }

    fn position(&self) -> usize {
        self.position
    }

    fn align_to_dword(&mut self) -> io::Result<()> {
        let remainder = self.position % 4;
        if remainder != 0 {
            let padding = 4 - remainder;
            for _ in 0..padding {
                self.inner.write_u8(0)?;
                self.position += 1;
            }
        }
        Ok(())
    }

    fn into_inner(self) -> W {
        self.inner
    }
}

impl<W: Write> Write for PositionWriter<W> {
    fn write(&mut self, buf: &[u8]) -> io::Result<usize> {
        let n = self.inner.write(buf)?;
        self.position += n;
        Ok(n)
    }

    fn flush(&mut self) -> io::Result<()> {
        self.inner.flush()
    }
}

/// Read an ID or string from the stream
/// Returns ResourceId and the number of bytes consumed
fn read_id_or_string<R: Read>(reader: &mut R) -> io::Result<ResourceId> {
    let first_word = reader.read_u16::<LittleEndian>()?;

    if first_word == 0xFFFF {
        // It's an ID: 0xFFFF followed by u16 ID
        let id = reader.read_u16::<LittleEndian>()?;
        Ok(ResourceId::Id(id))
    } else {
        // It's a UTF-16 string - first_word is the first character
        let mut chars = vec![first_word];

        // Read until null terminator
        loop {
            let ch = reader.read_u16::<LittleEndian>()?;
            if ch == 0 {
                break;
            }
            chars.push(ch);
        }

        // Convert UTF-16 to String
        let string = String::from_utf16(&chars)
            .map_err(|_| io::Error::new(io::ErrorKind::InvalidData, "Invalid UTF-16 string"))?;

        Ok(ResourceId::Name(string))
    }
}

/// Write an ID to the stream (0xFFFF marker + u16 ID)
fn write_id<W: Write>(writer: &mut W, id: u16) -> io::Result<()> {
    writer.write_u16::<LittleEndian>(0xFFFF)?;
    writer.write_u16::<LittleEndian>(id)?;
    Ok(())
}

/// Write a string to the stream (UTF-16 with null terminator)
fn write_string<W: Write>(writer: &mut W, s: &str) -> io::Result<()> {
    for ch in s.encode_utf16() {
        writer.write_u16::<LittleEndian>(ch)?;
    }
    writer.write_u16::<LittleEndian>(0)?; // Null terminator
    Ok(())
}

/// Align reader to DWORD boundary (for Cursor<Vec<u8>>)
fn align_reader_to_dword<R: Read + std::io::Seek>(reader: &mut R) -> io::Result<()> {
    let pos = reader.stream_position()?;
    let remainder = (pos % 4) as usize;
    if remainder != 0 {
        let padding = 4 - remainder;
        reader.seek(std::io::SeekFrom::Current(padding as i64))?;
    }
    Ok(())
}

/// Align writer to DWORD boundary (write padding zeros)
fn align_writer_to_dword<W: Write>(writer: &mut PositionWriter<W>) -> io::Result<()> {
    writer.align_to_dword()
}

// =============================================================================
// Resource Type-Specific Parsers
// =============================================================================

/// A single string table entry
#[derive(Debug, Clone)]
pub struct StringTableEntry {
    /// String ID (0-65535)
    pub id: u16,
    /// String value
    pub value: String,
}

/// Parse a string table resource
///
/// String tables in .res files are organized in blocks of 16 strings.
/// Each block has a resource ID = (string_id / 16) + 1
///
/// Format:
/// - For each of 16 slots:
///   - u16: string length in characters (not bytes)
///   - UTF-16 string data (no null terminator)
///
pub fn parse_string_table(data: &[u8], block_id: u16) -> io::Result<Vec<StringTableEntry>> {
    let mut cursor = Cursor::new(data);
    let mut entries = Vec::new();

    let base_id = (block_id.saturating_sub(1)) * 16;

    for i in 0..16 {
        // Read string length
        let length = match cursor.read_u16::<LittleEndian>() {
            Ok(len) => len as usize,
            Err(_) => break, // End of data
        };

        if length == 0 {
            // Empty slot
            continue;
        }

        // Read UTF-16 characters
        let mut chars = Vec::with_capacity(length);
        for _ in 0..length {
            let ch = cursor.read_u16::<LittleEndian>()?;
            chars.push(ch);
        }

        // Convert to String
        let value = String::from_utf16(&chars)
            .map_err(|_| io::Error::new(io::ErrorKind::InvalidData, "Invalid UTF-16 string"))?;

        entries.push(StringTableEntry {
            id: base_id + i,
            value,
        });
    }

    Ok(entries)
}

/// Create a string table resource data block
///
/// # Arguments
///
/// * `entries` - String entries (must all be in the same block of 16)
///
pub fn create_string_table(entries: &[StringTableEntry]) -> io::Result<Vec<u8>> {
    let mut buffer = Vec::new();

    if entries.is_empty() {
        return Ok(buffer);
    }

    // Determine the block
    let block_id = (entries[0].id / 16) * 16;

    // Create 16 slots
    for i in 0..16 {
        let id = block_id + i;

        // Find entry for this ID
        if let Some(entry) = entries.iter().find(|e| e.id == id) {
            // Write length
            let chars: Vec<u16> = entry.value.encode_utf16().collect();
            buffer.write_u16::<LittleEndian>(chars.len() as u16)?;

            // Write characters
            for ch in chars {
                buffer.write_u16::<LittleEndian>(ch)?;
            }
        } else {
            // Empty slot
            buffer.write_u16::<LittleEndian>(0)?;
        }
    }

    Ok(buffer)
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn test_resource_type_conversion() {
        assert_eq!(ResourceType::from_id(6), ResourceType::String);
        assert_eq!(ResourceType::String.to_id(), Some(6));

        assert_eq!(ResourceType::from_id(3), ResourceType::Icon);
        assert_eq!(ResourceType::Icon.to_id(), Some(3));
    }

    #[test]
    fn test_resource_id() {
        let id = ResourceId::Id(100);
        assert!(id.is_id());
        assert_eq!(id.as_id(), Some(100));

        let name = ResourceId::Name("TEST".to_string());
        assert!(name.is_name());
        assert_eq!(name.as_name(), Some("TEST"));
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
    fn test_string_table_creation() {
        let entries = vec![
            StringTableEntry {
                id: 100,
                value: "Hello".to_string(),
            },
            StringTableEntry {
                id: 101,
                value: "World".to_string(),
            },
        ];

        let data = create_string_table(&entries).unwrap();
        assert!(!data.is_empty());

        // Parse it back
        let block_id = (100 / 16) + 1;
        let parsed = parse_string_table(&data, block_id).unwrap();

        assert_eq!(parsed.len(), 2);
        assert_eq!(parsed[0].id, 100);
        assert_eq!(parsed[0].value, "Hello");
        assert_eq!(parsed[1].id, 101);
        assert_eq!(parsed[1].value, "World");
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
}
