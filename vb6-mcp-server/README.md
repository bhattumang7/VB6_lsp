# VB6 MCP Server

An MCP (Model Context Protocol) server that provides VB6 code analysis capabilities to Claude and other AI assistants.

## Features

- **Symbol Discovery** - Get all symbols (variables, functions, classes, etc.) in a file
- **Go to Definition** - Find where a symbol is defined
- **Find References** - Find all usages of a symbol
- **Hover Information** - Get type and signature info at a position
- **Code Completion** - Get completion suggestions at a position
- **Diagnostics** - Get parse errors and warnings

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
