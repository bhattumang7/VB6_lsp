# VB6 MCP Server

An MCP (Model Context Protocol) server that provides VB6 code analysis capabilities to Claude and other AI assistants.

## Features

- **Symbol Discovery** - Get all symbols (variables, functions, classes, etc.) in a file
- **Go to Definition** - Find where a symbol is defined
- **Find References** - Find all usages of a symbol
- **Hover Information** - Get type and signature info at a position
- **Code Completion** - Get completion suggestions at a position
- **Diagnostics** - Get parse errors and warnings
- **Form / Resource Decoding** - Read a `.frm`/`.ctl` design plus its `.frx`/`.ctx` companion resources (pictures, fonts, strings, lists) as structured JSON; optionally decode proprietary control bags live via COM
- **Resource Files** - Read and write Win32 `.res` files and string tables

## Installation

```bash
cd vb6-mcp-server
npm install
npm run build
```

## Usage with Claude Code

Add to your `.mcp.json`:

```json
{
  "mcpServers": {
    "vb6": {
      "command": "node",
      "args": ["C:/projects/VB6_lsp/vb6-mcp-server/dist/index.js"]
    }
  }
}
```

## Available Tools

### `vb6_get_symbols`
Get all symbols defined in a VB6 file.

```json
{
  "file_path": "C:/path/to/file.bas"
}
```

### `vb6_find_definition`
Find where a symbol at a position is defined.

```json
{
  "file_path": "C:/path/to/file.bas",
  "line": 10,
  "column": 5
}
```

### `vb6_find_references`
Find all references to a symbol.

```json
{
  "file_path": "C:/path/to/file.bas",
  "line": 10,
  "column": 5
}
```

### `vb6_get_hover`
Get hover information at a position.

```json
{
  "file_path": "C:/path/to/file.bas",
  "line": 10,
  "column": 5
}
```

### `vb6_get_completions`
Get code completion suggestions.

```json
{
  "file_path": "C:/path/to/file.bas",
  "line": 10,
  "column": 5
}
```

### `vb6_get_diagnostics`
Get parse errors for a file.

```json
{
  "file_path": "C:/path/to/file.bas"
}
```

### `vb6_read_form`
Read a VB6 form/control (`.frm`/`.ctl`/`.pag`/`.dob`) and return its full design as
structured JSON: the control tree, `Object=` type-library declarations, every
`.frx`/`.ctx` companion resource resolved, and a byte-accounting coverage report.

Proprietary vendor control bags (e.g. an MSChart `OleObjectBlob`) are labelled
opaque with their CLSID by default. Set `com_decode: true` to decode them live into
typed properties by hosting the control via COM — **Windows-only, ~2-3s per call**,
and requires the control to be registered with its design license present (otherwise
the bag stays opaque or returns an explicit error). The bridge scripts
(`scripts/com_bag_decode.ps1`, `scripts/ComBag.cs`) must sit next to the `vb6-lsp`
binary, or `VB6_COM_BRIDGE` must point at them.

```json
{
  "file_path": "C:/path/to/Form1.frm",
  "com_decode": false
}
```

### `vb6_read_res_file` / `vb6_write_res_file` / `vb6_get_string_table`
Read a Win32 `.res` file to JSON, write one back from JSON, or extract its string
table. Each takes a `file_path` (see the tool schema for write/table specifics).

> The `vb6-lsp` binary is located relative to this server's module by default
> (`../../target/{release,debug}/vb6-lsp[.exe]`); set the `VB6_LSP_BIN` environment
> variable to point at an explicit path.

## Development

```bash
# Build
npm run build

# Watch mode
npm run dev

# Run directly
npm start
```

## Architecture

The server reuses the tree-sitter VB6 grammar from the parent project and implements:

- `src/analysis/` - Symbol table, scope management, tree-sitter walker
- `src/index.ts` - MCP server entry point with tool handlers

## License

MIT
