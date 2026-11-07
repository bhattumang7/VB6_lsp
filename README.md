# VB6 Language Server (vb6-lsp)

A Language Server Protocol implementation for Visual Basic 6, written entirely in Rust. Provides IDE features via LSP and VB6 analysis tools via MCP.

## Background

This started as a "scratch your own itch" project — maintaining a large VB6 codebase with no modern tooling is painful, and every existing solution was either incomplete or long abandoned. The parser and analysis engine are built to closely match real VB6 syntax and semantics, not just approximate them.

The most widely-used public VB6 grammar is the ANTLR4 grammar at [antlr/grammars-v4](https://github.com/antlr/grammars-v4/tree/master/vb6). It is a useful reference but has a number of correctness issues — missing `#Const` support, broken statement-separator colon, wrong operator associativity for `^`, accepting VBA7/VB.NET-only syntax like `+=` and inline `For` type declarations, and more. The Rust parser in this project fixes all of them. A full writeup is in [`antlr parser/BUGS.md`](antlr%20parser/BUGS.md).

## Features

- **Syntax Highlighting** — semantic token-based
- **Code Completion** — variables, functions, keywords
- **Hover** — type and signature info
- **Go to Definition** — including form controls
- **Find References** — workspace-wide
- **Document / Workspace Symbols** — outline and project-wide search
- **Diagnostics** — syntax and semantic errors in real time
- **Code Formatting** — block indentation, trailing-whitespace trim, keyword-case normalisation
- **Code Actions** — declare an undeclared variable, create a missing Sub/Function, toggle single-line ↔ block `If`
- **Rename** — safe project-wide symbol renaming
- **Semantic tokens**, **signature help**, **folding ranges**, **call hierarchy**

### Supported file types

| Extension | Description |
|-----------|-------------|
| `.bas` | Standard modules |
| `.cls` | Class modules |
| `.frm` | Forms (companion `.frx` parsed automatically) |
| `.ctl` | UserControls (companion `.frx` parsed automatically) |
| `.pag` | PropertyPages (companion `.frx` parsed automatically) |
| `.dob` | UserDocuments (companion `.frx` parsed automatically) |
| `.vbp` | Project files — list the files analysed together (see [Opening a project](#4--open-a-vb6-project-or-group)) |
| `.vbg` | Project groups — several projects analysed together, each seeing the public classes of the projects it references |
| `.res` | Compiled Win32 resource files (read/write via CLI) |

## Architecture

Everything is pure Rust — no Node.js, no tree-sitter, no TypeScript.

```
src/
├── main.rs          →  vb6-lsp  (LSP server binary, stdio JSON-RPC)
├── mcp_main.rs      →  vb6-mcp  (MCP server binary, stdio JSON-RPC)
├── engine_glue.rs   →  maps vb6-engine results to LSP/MCP wire types
├── lsp/             →  tower-lsp handlers
├── controls/        →  .frx / .frm form-control parsing
├── workspace/       →  .vbp project discovery, .res resource files
└── utils/           →  ANSI/Unicode encoding helpers

crates/
├── vb6-syntax/      →  scanner, parser, AST, diagnostics
├── vb6-sema/        →  name binder, type system
├── vb6-engine/      →  Session API (hover, nav, rename, tokens, …)
├── vb6-core/        →  integration test harness
└── vb6-ast-derive/  →  proc-macro (Children derive)

vscode-vb6/          →  VS Code extension (bundles via esbuild, .vsix)
```

## Building

**Requirements:** Rust 1.70+ with the MSVC toolchain (Windows).

```powershell
# Install Rust (choose MSVC toolchain when prompted)
Invoke-WebRequest -Uri https://win.rustup.rs/x86_64 -OutFile rustup-init.exe
.\rustup-init.exe
```

```powershell
# Clone and build
git clone https://github.com/bhattumang7/VB6_lsp.git
cd VB6_lsp

# Debug build (fast compile, slower binary)
cargo build

# Release build (LTO enabled — use this for day-to-day work)
cargo build --release
```

The release build produces two binaries:

| Binary | Path | Purpose |
|--------|------|---------|
| `vb6-lsp.exe` | `target/release/vb6-lsp.exe` | LSP server for IDEs |
| `vb6-mcp.exe` | `target/release/vb6-mcp.exe` | MCP server for Claude Code |

## Development

```powershell
# Run the LSP server directly (communicates via stdin/stdout)
cargo run

# Run with debug logging
$env:RUST_LOG = "debug"
cargo run

# Format code
cargo fmt

# Lint
cargo clippy

# Run all tests
cargo test

# Run tests for a specific crate
cargo test -p vb6-syntax
cargo test -p vb6-sema
cargo test -p vb6-engine

# Structural suite (AST / parser / scanner / binder)
cargo test-ast          # alias in .cargo/config.toml

# 32-bit suite (vb6-core runtime — requires i686 target)
rustup target add i686-pc-windows-msvc
cargo test-i686

# Run all three suites in sequence
pwsh scripts/test.ps1
```

## VS Code Setup

### 1 — Build the extension

```powershell
cd vscode-vb6
npm install          # one-time: installs esbuild + vsce
npm run build        # bundles extension.js → dist/extension.js
npm run package      # produces vb6-lsp-0.1.0.vsix
```

### 2 — Install the extension

```powershell
code --install-extension vscode-vb6/vb6-lsp-0.1.0.vsix
```

Or: open VS Code → `Extensions` → `…` menu → `Install from VSIX…` → pick the file.

### 3 — Point the extension at the server

Open VS Code settings (`Ctrl+,`) and set:

```json
{
  "vb6.lsp.serverPath": "C:\\projects\\VB6_lsp\\target\\release\\vb6-lsp.exe"
}
```

Or set it in `.vscode/settings.json` inside your VB6 project folder so the path is checked in alongside the code.

### 4 — Open a VB6 project or group

The language server analyses one project (`.vbp`) or group (`.vbg`) at a time: every file the project lists is analysed together, whether or not the editor has it open, so names resolve across files, closing a file goes back to what is on disk, and an edit that affects other files refreshes their diagnostics.

- **Command Palette → `VB6: Open Project or Group…`** lists the projects found in the folder (or browse for one). Opening a project from an empty window opens its folder first.
- **Right-click a `.vbp`/`.vbg` in the Explorer → `Open as VB6 Project`.**
- **Opening a `.vbp`/`.vbg` file** shows a project page (files, references, problems) with **Load Project** and **Edit as Text** buttons.
- The **VB6 sidebar** shows the open project like the VB6 Project Explorer (forms, modules, classes, references; missing files are marked), and the **status bar** shows the open project and its problem count.

The choice is stored in the `vb6.project` setting (relative to the folder). When it is empty and the folder has exactly one group, or exactly one project, that one opens automatically (`vb6.autoLoadProject`).

**References.** `Reference=` and `Object=` lines of a project make that library's types visible to the analysis. On Windows a type library or control library is found by its GUID and version in the registry (the 32-bit registration first, as VB6 uses), else at the file the project names, and read through the system's type-information interfaces. On other platforms there is no registry: the library is read directly from the file the project names (a `.tlb`/`.olb`, or the type library inside a `.dll`/`.ocx`), with `\` separators and the case of the path matched to the files on disk, so the files must be present at the paths the project gives. A library that takes types from another library needs that library's file beside it; when it is not found, a warning on the `.vbp` line says which file, and the types taken from it are not available. A reference to another project (`*\A...`) makes that project's public classes visible, and **Go to Definition** on one of their names opens the declaration in the referenced project's own source.

Problems in the project files themselves — a listed file that is missing, a referenced project that is not found or that references itself, a reference whose types are not available — appear in the Problems panel on the line of the `.vbp`/`.vbg` that declares them.

Other editors can use the same features through the server: pass `initializationOptions: { "project": "<path to .vbp or .vbg>" }`, or run the `workspace/executeCommand` commands `vb6.openProject`, `vb6.closeProject`, `vb6.projectInfo` and `vb6.projectPreview`; the server reports its state with the `vb6/projectStatus` notification.

**Tip:** setting `"vb6.lsp.trace.server": "verbose"` in settings shows the full LSP traffic in the Output panel (select "VB6 Language Server" from the drop-down).

## Claude Code MCP Setup

The `vb6-mcp` binary speaks the [Model Context Protocol](https://modelcontextprotocol.io/) over stdio, giving Claude Code access to VB6 analysis tools.

### Tools exposed to Claude

| Tool | Description |
|------|-------------|
| `vb6_get_symbols` | List all symbols (variables, functions, types) in a file |
| `vb6_find_definition` | Go to where a symbol is defined |
| `vb6_find_references` | Find all usages of a symbol |
| `vb6_get_hover` | Get type / signature info at a position |
| `vb6_get_diagnostics` | Get parse and semantic errors |
| `vb6_read_res_file` | Parse a compiled `.res` file (bitmaps, icons, strings, …) |
| `vb6_write_res_file` | Write resources back to a `.res` file |
| `vb6_get_string_table` | Extract string-table entries from a `.res` file |

### Configure Claude Code

Add the server to your MCP configuration. You can do this at user level (`~/.claude/settings.json`) or per-project (`.mcp.json` in the project root).

**`~/.claude/settings.json`** (user-level, available in every project):
```json
{
  "mcpServers": {
    "vb6": {
      "command": "C:\\projects\\VB6_lsp\\target\\release\\vb6-mcp.exe",
      "args": []
    }
  }
}
```

**`.mcp.json`** (project-level, checked in with the VB6 project):
```json
{
  "mcpServers": {
    "vb6": {
      "command": "C:\\projects\\VB6_lsp\\target\\release\\vb6-mcp.exe",
      "args": []
    }
  }
}
```

After saving, restart Claude Code (or run `/mcp` to reload servers). Claude will then have the VB6 tools available whenever you work in that project.

## CLI commands

The `vb6-lsp` binary doubles as a CLI for resource-file operations:

```powershell
# Dump all Win32 resources from a .res file as JSON
vb6-lsp read-res Game2048.RES

# Write resources from a JSON file back to a .res file
vb6-lsp write-res input.json output.res

# Extract string-table entries (language ID 1 = English)
vb6-lsp parse-string-table Game2048.RES 1

# Control tree of a form with the .frx/.ctx data resolved, as JSON
vb6-lsp read-form Main.frm

# Write a form's pictures, icons and raw control bags to a folder
vb6-lsp extract-form Main.frm out/

# Decode the VB6 property bags that VB-authored controls keep in the .frx:
# every record as name, type and value; nested groups (a column's value list) are flattened
vb6-lsp read-bags Main.frm
```

See [docs/RESOURCE_FILES.md](docs/RESOURCE_FILES.md) for format details.

## Neovim setup

```lua
local lspconfig = require('lspconfig')
local configs   = require('lspconfig.configs')

if not configs.vb6_lsp then
  configs.vb6_lsp = {
    default_config = {
      cmd      = { 'C:\\projects\\VB6_lsp\\target\\release\\vb6-lsp.exe' },
      filetypes = { 'vb', 'bas', 'cls', 'frm', 'ctl', 'pag', 'dob' },
      root_dir  = lspconfig.util.root_pattern('.git', '*.vbp'),
      settings  = {},
    },
  }
end

lspconfig.vb6_lsp.setup {}
```

## Project structure

```
vb6-lsp/
├── src/                        # LSP + MCP glue (Rust)
│   ├── main.rs                 # vb6-lsp entry point + CLI commands
│   ├── mcp_main.rs             # vb6-mcp entry point
│   ├── engine_glue.rs          # vb6-engine → LSP/MCP type mapping
│   ├── lsp/                    # tower-lsp request handlers
│   ├── controls/               # Form control + .frx parsing
│   ├── workspace/              # .vbp discovery, .res read/write
│   └── utils/                  # ANSI/Unicode helpers
│
├── crates/
│   ├── vb6-syntax/             # Scanner, parser, AST, diagnostics
│   ├── vb6-sema/               # Name binder, type resolution
│   ├── vb6-engine/             # Session API (all IDE queries)
│   ├── vb6-core/               # Integration tests
│   └── vb6-ast-derive/         # Children proc-macro
│
├── vscode-vb6/                 # VS Code extension
│   ├── src/extension.js        # Extension entry point
│   └── package.json            # Manifest + build scripts
│
├── tests/                      # Integration tests
│   ├── engine_analysis.rs
│   ├── integration_vbp_parsing.rs
│   ├── lsp_server.rs
│   ├── mcp_server.rs
│   └── fixtures/
│
├── docs/
│   └── RESOURCE_FILES.md
├── scripts/
│   └── test.ps1
├── .cargo/config.toml          # test-ast / test-i686 aliases
└── Cargo.toml                  # Workspace manifest
```

## Limitations

- Visual form layout not integrated into LSP features

## License

MIT
