# VB6 Resource File (.res) Support

## Overview

The VB6 LSP supports reading, writing, and modifying Win32 compiled resource files (.res). This enables IDE features and programmatic access to VB6 resources through both the LSP server and MCP server.

## Resource Types Supported

All standard Win32 resource types are supported:

- **Visual Resources**: Bitmap, Icon, Cursor, GroupIcon, GroupCursor, AniIcon, AniCursor
- **UI Resources**: Menu, Dialog, Accelerator, Toolbar, DlgInit
- **String Resources**: String tables with multiple entries per block
- **Data Resources**: RcData (raw data), Font, FontDir, Html, Manifest
- **System Resources**: Version, MessageTable, PlugPlay, Vxd, DlgInclude
- **Custom Resources**: Named resources and numeric custom types

## File Format

VB6 .res files use the standard Win32 compiled resource format:

```
[Empty Header]
[Resource Header][Resource Data][Padding]
[Resource Header][Resource Data][Padding]
...
```

### Key Format Details

- **DWORD Alignment**: All data must be aligned to 4-byte boundaries
- **UTF-16 Encoding**: Resource names and string data use UTF-16 LE
- **Resource IDs**: Can be numeric (0-65535) or string names
- **Language IDs**: Support for multiple languages (e.g., 0x0409 for US English)

## CLI Usage

### Reading Resource Files

```bash
vb6-lsp read-res Game2048.RES
```

Output (JSON):
```json
{
  "resources": [
    {
      "resource_type": "Bitmap",
      "name": {"type": "Id", "value": 101},
      "language_id": 1033,
      "data_size": 5160,
      "data_base64": "KAAAAEAAAABAAAAAAQAIAAAAAAAAEAAAxA4AAMQOAAAAAAAAAAAAAAAAAAAAAgAAAIAAAACAgACA..."
    }
  ]
}
```

### Writing Resource Files

```bash
# Create input.json with resource definitions
vb6-lsp write-res input.json output.res
```

### Parsing String Tables

```bash
vb6-lsp parse-string-table Game2048.RES 1
```

Output:
```json
{
  "strings": [
    {"id": 1, "value": "Welcome"},
    {"id": 2, "value": "Game Over"}
  ]
}
```

## MCP Server Integration

The VB6 MCP server exposes resource file operations to Claude AI:

### Available Tools

#### `vb6_read_res_file`
Read and parse a .res file, returning all resources with metadata and base64-encoded data.

**Input:**
```json
{
  "file_path": "C:\\path\\to\\resources.res"
}
```

**Output:**
```json
{
  "file": "C:\\path\\to\\resources.res",
  "resourceCount": 16,
  "resources": [...]
}
```

#### `vb6_write_res_file`
Write resources to a .res file.

**Input:**
```json
{
  "file_path": "C:\\path\\to\\output.res",
  "resources": [
    {
      "resource_type": "Bitmap",
      "name": {"type": "Id", "value": 200},
      "language_id": 1033,
      "data_base64": "..."
    }
  ]
}
```

#### `vb6_get_string_table`
Parse a specific string table block.

**Input:**
```json
{
  "file_path": "C:\\path\\to\\resources.res",
  "block_id": 1
}
```

**Output:**
```json
{
  "file": "C:\\path\\to\\resources.res",
  "blockId": 1,
  "stringCount": 5,
  "strings": [
    {"id": 1, "value": "String 1"},
    {"id": 2, "value": "String 2"}
  ]
}
```

## Programmatic Usage (Rust)

```rust
use vb6_lsp::workspace::{
    read_res_file, write_res_file, parse_string_table,
    ResourceEntry, ResourceId, ResourceType
};

// Read a resource file
let resources = read_res_file("Game2048.RES")?;

// Find bitmap resources
for res in &resources {
    if res.resource_type == ResourceType::Bitmap {
        println!("Bitmap ID: {:?}, Size: {} bytes", res.name, res.data.len());
    }
}

// Parse string table
if let Some(string_res) = resources.iter()
    .find(|r| r.resource_type == ResourceType::String)
{
    if let ResourceId::Id(block_id) = string_res.name {
        let strings = parse_string_table(&string_res.data, block_id)?;
        for s in strings {
            println!("String {}: {}", s.id, s.value);
        }
    }
}

// Create new resources
let new_resource = ResourceEntry::new(
    ResourceType::Bitmap,
    ResourceId::Id(200),
    1033, // US English
    bitmap_data
);

write_res_file("output.res", &[new_resource])?;
```

## Implementation Details

### Position Tracking

The parser uses `PositionWriter` and `PositionTracker` wrappers to automatically handle DWORD alignment:

```rust
pub(crate) struct PositionWriter<W> {
    inner: W,
    position: usize,
}

impl<W: Write> PositionWriter<W> {
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
}
```

### String Table Format

String tables are stored in blocks of 16 strings. Each block has an ID (1-4095), and strings within the block are numbered (0-15):

```
Block ID 1: Strings 1-16
Block ID 2: Strings 17-32
...
```

Each string entry is:
```
[u16 length][UTF-16 characters]
```

## Testing

Integration tests validate the parser with real VB6 resource files:

```bash
cargo test --test test_real_res
```

Tests include:
- **Round-trip validation**: Read → Write → Read comparison
- **Binary exactness**: Byte-perfect reproduction of original files
- **Multiple file types**: Simple (Game2048.RES) and complex (sheep.res) files
- **Resource type coverage**: Bitmaps, icons, custom resources, and more

## Reference

Based on the RisohEditor resource editor implementation, ensuring compatibility with all edge cases handled by production VB6 resource tools.
