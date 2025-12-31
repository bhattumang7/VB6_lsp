# VB6 Language Server Protocol (LSP) with Claude AI

A modern Language Server Protocol implementation for Visual Basic 6, written in Rust with Claude Sonnet AI integration for intelligent code assistance.

## Two Ways to Use

| Component | Purpose | Language |
|-----------|---------|----------|
| **vb6-lsp** | Traditional LSP for IDEs (VS Code, Neovim, etc.) | Rust |
| **vb6-mcp-server** | MCP server for Claude Code AI integration | TypeScript |

## Features

### Core LSP Features
- **Syntax Highlighting** - Semantic token-based highlighting
- **Code Completion** - IntelliSense-style completions for variables, functions, keywords
- **Hover Information** - Type information and documentation on hover
- **Go to Definition** - Navigate to symbol definitions (including form controls)
- **Find References** - Find all references to a symbol across the workspace
- **Document Symbols** - Outline view of file structure
- **Diagnostics** - Real-time syntax and semantic error checking
- **Code Formatting** - Automatic code indentation and formatting
- **Rename Refactoring** - Safe symbol renaming across files

### Supported File Types
- **Source Files:** `.bas` (modules), `.cls` (classes), `.frm` (forms), `.ctl` (UserControls), `.pag` (PropertyPages), `.dob` (UserDocuments)
- **Binary Resources:** `.frx` files automatically parsed when opening `.frm`, `.ctl`, `.pag`, or `.dob` files
- **Project Files:** `.vbp` files parsed for workspace-wide symbol resolution
- **Compiled Resources:** `.res` files can be read/written via CLI commands

### Claude AI Integration (LSP)
When `ANTHROPIC_API_KEY` is set, additional AI-powered features:
- **Intelligent Code Completion** - Context-aware suggestions
- **Code Explanations** - Natural language explanations of VB6 code
- **Error Analysis** - AI-powered error diagnosis and fix suggestions
- **Refactoring Suggestions** - Smart code improvement recommendations
- **Documentation Generation** - Automatic comment generation
- **Migration Assistance** - Help converting VB6 to VB.NET/C#

### MCP Server Tools (for Claude Code)
The MCP server exposes these tools to Claude:

**Code Analysis:**
- **vb6_get_symbols** - List all symbols (variables, functions, types) in a file
- **vb6_find_definition** - Go to where a symbol is defined
- **vb6_find_references** - Find all usages of a symbol
- **vb6_get_hover** - Get type/signature info at a position
- **vb6_get_completions** - Get code completion suggestions
- **vb6_get_diagnostics** - Get parse errors and warnings

**Resource Files:**
- **vb6_read_res_file** - Read and parse a compiled .res file (all Win32 resource types: bitmaps, icons, strings, cursors, dialogs, etc.)
- **vb6_write_res_file** - Write resources to a compiled .res file
- **vb6_get_string_table** - Parse string table entries from a .res file

## Architecture

### Option 1: Traditional LSP (for IDEs)
```
┌─────────────────────────────────────────────────────────────────┐
│                     VSCode / Neovim / IDE                        │
│                    (LSP Client via JSON-RPC)                     │
└─────────────────────────────────────────────────────────────────┘
                              │
                              ▼
┌─────────────────────────────────────────────────────────────────┐
│                      vb6-lsp (Rust Server)                       │
│  ┌─────────────┐  ┌─────────────┐  ┌─────────────────────────┐  │
│  │ tower-lsp   │  │ Tree-Sitter │  │   Claude Integration    │  │
│  │ (LSP Core)  │◄─┤   Parser    │  │   (AI Assistance)       │  │
│  │             │  │   (VB6)     │  │                         │  │
│  └─────────────┘  └─────────────┘  └─────────────────────────┘  │
│         │               │                    │                   │
│         ▼               ▼                    ▼                   │
│  ┌─────────────────────────────────────────────────────────────┐│
│  │     Symbol Table + Analyzer (Scopes, Symbols, References)   ││
│  └─────────────────────────────────────────────────────────────┘│
└─────────────────────────────────────────────────────────────────┘
```

### Option 2: MCP Server (for Claude Code)
```
┌─────────────────────────────────────────────────────────────────┐
│                         Claude Code                              │
│                    (MCP Client via stdio)                        │
└─────────────────────────────────────────────────────────────────┘
                              │
                              ▼
┌─────────────────────────────────────────────────────────────────┐
│                 vb6-mcp-server (TypeScript)                      │
│  ┌─────────────┐  ┌─────────────┐  ┌─────────────────────────┐  │
│  │  MCP SDK    │  │ Tree-Sitter │  │     Tool Handlers       │  │
│  │  (Server)   │◄─┤   Parser    │  │  (symbols, refs, etc)   │  │
│  │             │  │   (VB6)     │  │                         │  │
│  └─────────────┘  └─────────────┘  └─────────────────────────┘  │
│         │               │                    │                   │
│         ▼               ▼                    ▼                   │
│  ┌─────────────────────────────────────────────────────────────┐│
│  │     Symbol Table + Analyzer (ported from Rust)              ││
│  └─────────────────────────────────────────────────────────────┘│
└─────────────────────────────────────────────────────────────────┘
```

## Prerequisites

### For Rust LSP Server (vb6-lsp)

1. **Rust** (1.70+) with MSVC toolchain
   ```powershell
   # Install Rust via rustup (choose MSVC toolchain during setup)
   Invoke-WebRequest -Uri https://win.rustup.rs/x86_64 -OutFile rustup-init.exe
   .\rustup-init.exe
   ```

2. **Visual Studio Build Tools 2022**
   - Download from: https://visualstudio.microsoft.com/downloads/#build-tools-for-visual-studio-2022
   - Install with "Desktop development with C++" workload
   - Required components:
     - MSVC v143 build tools
     - Windows SDK (10.0.19041.0 or later)
     - C++ Clang tools (provides headers like `stdbool.h`)

3. **Claude API Key** (Optional - for AI features in LSP)
   - Sign up at https://console.anthropic.com
   - Generate an API key
   - Set environment variable: `ANTHROPIC_API_KEY=your_key_here`

### For MCP Server (vb6-mcp-server)

1. **Node.js** (18.0+)
   ```powershell
   # Download from https://nodejs.org/ or use a package manager
   winget install OpenJS.NodeJS.LTS
   ```

2. **Visual Studio Build Tools 2022** (same as above - needed for native addon)

## Installation

### Option 1: Rust LSP Server

```bash
# Clone the repository
git clone https://github.com/yourusername/vb6-lsp.git
cd vb6-lsp

# Option A: Use the build script (sets up MSVC environment)
build_with_vs.bat

# Option B: Build directly (if environment is already configured)
cargo build --release

# Binary will be at: target/release/vb6-lsp.exe
```

### Option 2: MCP Server (for Claude Code)

```bash
# From the repository root
cd vb6-mcp-server

# Install dependencies (this also builds the tree-sitter native addon)
npm install

# Build TypeScript
npm run build

# The server is now ready at: dist/index.js
```

#### Configure Claude Code

Add to your MCP configuration file (`~/.mcp.json` or project `.mcp.json`):

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

Now when working with VB6 files, Claude Code will automatically have access to the VB6 analysis tools.

### VSCode Extension Setup

Create a VSCode extension to use the LSP server:

1. **Install the Language Extension**:
   Create `.vscode/settings.json` in your VB6 project:
   ```json
   {
     "vb6.lsp.serverPath": "C:\\path\\to\\vb6-lsp.exe",
     "vb6.lsp.trace.server": "verbose"
   }
   ```

2. **Or use generic LSP client**:
   Install the "Generic LSP Client" extension and configure:
   ```json
   {
     "genericLSP.servers": {
       "vb6": {
         "command": "C:\\path\\to\\vb6-lsp.exe",
         "filetypes": ["vb", "bas", "cls", "frm", "ctl", "pag", "dob"]
       }
     }
   }
   ```

   **Supported File Types:**
   - `.bas` - Standard modules
   - `.cls` - Class modules
   - `.frm` - Forms (with automatic `.frx` binary resource parsing)
   - `.ctl` - UserControls (with automatic `.frx` binary resource parsing)
   - `.pag` - PropertyPages (with automatic `.frx` binary resource parsing)
   - `.dob` - UserDocuments (with automatic `.frx` binary resource parsing)
   - `.vbp` - Project files (parsed for workspace-wide symbol resolution)

### Using with Neovim

```lua
-- In your Neovim config (init.lua or lsp.lua)
local lspconfig = require('lspconfig')
local configs = require('lspconfig.configs')

-- Define VB6 LSP if not already defined
if not configs.vb6_lsp then
  configs.vb6_lsp = {
    default_config = {
      cmd = {'C:\\path\\to\\vb6-lsp.exe'},
      filetypes = {'vb', 'bas', 'cls', 'frm', 'ctl', 'pag', 'dob'},
      root_dir = lspconfig.util.root_pattern('.git', '.vbp'),
      settings = {},
    },
  }
end

-- Start the LSP
lspconfig.vb6_lsp.setup{}
```

## Usage

### Running the LSP Server

The LSP server communicates via stdin/stdout:

```bash
# Run directly
./target/release/vb6-lsp

# With logging
RUST_LOG=vb6_lsp=debug ./target/release/vb6-lsp
```

### CLI Commands for Resource Files ✨ NEW

The vb6-lsp binary provides powerful CLI commands for working with VB6 compiled resource files (.res):

```bash
# Read a .res file (outputs JSON with all Win32 resources)
# Supports: bitmaps, icons, cursors, strings, dialogs, menus, version info, and custom resources
vb6-lsp read-res Game2048.RES

# Example output:
# {
#   "resources": [
#     {
#       "resource_type": "Icon",
#       "name": {"type": "Id", "value": 101},
#       "language_id": 1033,
#       "data_size": 2216,
#       "data_base64": "AAABAAEAEBAAAAEAIABoBA..."
#     }
#   ]
# }

# Write resources from JSON to a .res file
vb6-lsp write-res input.json output.res

# Parse string table from a .res file
vb6-lsp parse-string-table Game2048.RES 1
```

These commands enable:
- **Resource extraction** - Extract icons, bitmaps, strings from compiled .res files
- **Resource modification** - Edit resources via JSON and write back to .res
- **MCP integration** - Used by vb6-mcp-server for Claude AI to work with VB6 resources
- **Build automation** - Programmatic resource file manipulation in build scripts

See [docs/RESOURCE_FILES.md](docs/RESOURCE_FILES.md) for detailed documentation on resource file formats and usage examples.

### Configuration

Set these environment variables:

| Variable | Description | Default |
|----------|-------------|---------|
| `ANTHROPIC_API_KEY` | Claude API key for AI features | None (AI disabled) |
| `RUST_LOG` | Logging level | `vb6_lsp=info` |

### Example VB6 Files

See the `examples/` directory:
- [sample.bas](examples/sample.bas) - Module with functions, subs, types, enums
- [DatabaseConnection.cls](examples/DatabaseConnection.cls) - Class example with properties and events

## Development

### Project Structure

```
vb6-lsp/
├── src/                         # Rust LSP Server
│   ├── main.rs                  # Entry point + CLI commands
│   ├── lib.rs                   # Library exports for external use
│   ├── lsp/                     # LSP server implementation
│   │   ├── mod.rs               # Main LSP handlers
│   │   ├── capabilities.rs      # LSP capabilities
│   │   ├── document.rs          # Document management
│   │   └── handlers.rs          # Request handlers
│   ├── parser/                  # VB6 parser
│   │   ├── mod.rs               # Parser logic & tree-sitter integration
│   │   ├── ast.rs               # AST definitions
│   │   ├── tree_sitter.rs       # Tree-sitter parser wrapper
│   │   ├── converter.rs         # Tree-sitter to AST conversion
│   │   └── lexer.rs             # Legacy tokenizer
│   ├── analysis/                # Code analysis & symbol table
│   │   ├── mod.rs               # Analyzer with symbol table support
│   │   ├── symbol_table.rs      # Symbol table with query methods
│   │   ├── builder.rs           # Builds symbol table from parse tree
│   │   ├── symbol.rs            # Symbol, SymbolKind, TypeInfo, FormControl
│   │   ├── scope.rs             # Scope hierarchy management
│   │   └── position.rs          # Source positions and ranges
│   ├── controls/                # Form control support ✨ NEW
│   │   ├── mod.rs               # Control parsing and management
│   │   ├── properties.rs        # VB6 control property definitions
│   │   ├── colors.rs            # VB6 color constants and conversions
│   │   └── frx.rs               # FRX binary format parser
│   ├── workspace/               # Workspace & project management
│   │   ├── mod.rs               # Workspace manager
│   │   ├── project.rs           # Project structure definitions
│   │   ├── vbp_parser.rs        # VB6 project file (.vbp) parser
│   │   ├── frx_parser.rs        # Binary form resource (.frx) parser
│   │   └── res_parser.rs        # Compiled resource file (.res) parser ✨ NEW
│   ├── utils/                   # Utility functions ✨ NEW
│   │   ├── mod.rs               # Utility module exports
│   │   └── encoding.rs          # VB6 text encoding (ANSI/Unicode)
│   └── claude/                  # Claude AI integration
│       └── mod.rs               # API client
│
├── vb6-mcp-server/              # TypeScript MCP Server
│   ├── src/
│   │   ├── index.ts             # MCP server entry point
│   │   └── analysis/            # Analysis code (ported from Rust)
│   │       ├── types.ts         # Symbol, SymbolKind, TypeInfo
│   │       ├── position.ts      # Source positions and ranges
│   │       ├── scope.ts         # Scope hierarchy
│   │       ├── symbolTable.ts   # Symbol table with queries
│   │       └── builder.ts       # Tree-sitter walker
│   ├── package.json
│   └── tsconfig.json
│
├── tree-sitter-vb6/             # Tree-sitter grammar for VB6
│   ├── grammar.js               # Grammar definition
│   ├── binding.gyp              # Node.js native addon build config
│   ├── bindings/
│   │   ├── node/                # Node.js bindings
│   │   ├── rust/                # Rust bindings
│   │   └── c/                   # C bindings
│   ├── src/                     # Generated parser (C)
│   │   ├── parser.c
│   │   └── scanner.c            # External scanner
│   └── test/                    # Grammar tests
│
├── examples/                    # Example VB6 files
├── tests/                       # Integration tests ✨ NEW
│   ├── integration_vbp_parsing.rs  # VBP project file parsing tests
│   ├── test_real_res.rs         # RES resource file tests
│   └── fixtures/                # Test data files
├── docs/                        # Documentation ✨ NEW
│   └── RESOURCE_FILES.md        # Resource file format documentation
├── Cargo.toml                   # Rust dependencies
└── README.md
```

### Building and Testing

#### Rust LSP Server

```bash
# Debug build
cargo build

# Release build
cargo build --release

# Run tests
cargo test

# Run with debug logging
RUST_LOG=debug cargo run

# Format code
cargo fmt

# Lint
cargo clippy
```

#### MCP Server (TypeScript)

```bash
cd vb6-mcp-server

# Install dependencies
npm install

# Build
npm run build

# Watch mode (rebuild on changes)
npm run dev

# Test the server starts correctly
timeout 2 node dist/index.js
# Should print "VB6 MCP Server started" then timeout
```

#### Tree-sitter Grammar

```bash
cd tree-sitter-vb6

# Regenerate parser from grammar.js
npx tree-sitter generate

# Run grammar tests
npx tree-sitter test

# Parse a sample file
npx tree-sitter parse ../examples/sample.bas
```

### Adding New Features

1. **Parser Enhancement**: Modify `src/parser/mod.rs` and `src/parser/ast.rs`
2. **New LSP Capability**: Add to `src/lsp/mod.rs` initialization
3. **Analysis**: Extend `src/analysis/mod.rs`
4. **Claude Integration**: Add methods to `src/claude/mod.rs`

### Using the FRX Parser

The FRX parser is available for direct use in your code:

```rust
use vb6_lsp::workspace::{resource_file_resolver, list_resolver};

// Example 1: Load an icon from a form's FRX file
let icon_data = resource_file_resolver("MyForm.frx", 0x0C)?;
println!("Icon is {} bytes", icon_data.len());

// Example 2: Extract combo box items from a UserControl's FRX file
let list_data = resource_file_resolver("MyControl.frx", 0x18)?;
let items = list_resolver(&list_data);
for item in items {
    println!("List item: {}", item);
}

// Example 3: Load a picture from a PropertyPage
let picture_data = resource_file_resolver("MyPage.frx", 0x24)?;
// picture_data can be saved or processed as needed

// Example 4: Handle empty resources (removed icons)
match resource_file_resolver("MyForm.frx", 0x00) {
    Ok(data) if data.is_empty() => println!("Icon was removed"),
    Ok(data) => println!("Got {} bytes", data.len()),
    Err(e) => eprintln!("Error: {}", e),
}
```

**Works universally for:**
- Forms (`.frm` + `.frx`)
- UserControls (`.ctl` + `.frx`)
- PropertyPages (`.pag` + `.frx`)
- UserDocuments (`.dob` + `.frx`)

## Supported VB6 Language Features

### Currently Supported (via Tree-Sitter Grammar)
- ✅ Variable declarations (Dim, Private, Public, Global, Static)
- ✅ Constant declarations
- ✅ User-defined Types (with members)
- ✅ Enumerations (with members)
- ✅ Sub procedures
- ✅ Function procedures
- ✅ Property procedures (Get, Let, Set)
- ✅ Parameters (ByVal, ByRef, Optional, ParamArray)
- ✅ Comments and line continuations
- ✅ Option statements
- ✅ Attribute statements (.cls/.frm files)
- ✅ With blocks
- ✅ Control flow (If/Then/Else, Select Case, For/Next, Do/Loop, While/Wend)
- ✅ Events
- ✅ Implements
- ✅ Declare statements (API declarations)
- ✅ Labels and GoTo
- ✅ On Error handling
- ✅ ReDim and Erase
- ✅ Member access (dot notation)
- ✅ Array subscripts
- ✅ All operators and expressions

### Symbol Table Features
- ✅ Hierarchical scope tracking (Module → Procedure → Block)
- ✅ Case-insensitive symbol lookup
- ✅ Precise position-based queries (O(1) line lookup)
- ✅ Parameter tracking with types
- ✅ Type member resolution
- ✅ Enum member resolution

### VBP Project File Parsing ✅

The LSP includes comprehensive VB6 project file (.vbp) parsing with full workspace support:

**Project Structure:**
- Project types: Standard EXE, ActiveX DLL, ActiveX EXE, ActiveX Control
- All source file types: Modules (`.bas`), Classes (`.cls`), Forms (`.frm`), User Controls (`.ctl`), Property Pages (`.pag`), User Documents (`.dob`), Designers (`.dsr`)
- Automatic `.frx` binary resource parsing for all visual designer files (`.frm`, `.ctl`, `.pag`, `.dob`)
- Related documents and resources

**References & Dependencies:**
- Type library references with UUID validation
- SubProject references (*\A format)
- ActiveX/OCX component references
- External library tracking (ADO, MSComCtl, etc.)

**Compilation Settings:**
- Compilation type (P-Code vs Native Code)
- Optimization settings (speed vs size)
- Debug options (bounds check, overflow check, floating point check)
- Conditional compilation arguments

**Version Information:**
- Major/Minor/Revision versioning
- Company name, copyright, product name
- Auto-increment support

**Threading & Compatibility:**
- Threading model (Single-threaded vs Apartment-threaded)
- Start mode (Stand-alone vs Automation)
- Compatibility mode (None, Project, Binary)
- Compatible EXE tracking

**Workspace Management:**
- Automatic VBP discovery in workspace folders
- Multi-project support
- Cross-project symbol resolution
- Public symbol indexing across entire workspace
- File-to-project mapping for navigation

### FRX Binary Resource File Parsing ✅

The LSP includes comprehensive binary resource file (.frx) parsing with universal support for all VB6 visual designers:

**Universal Format Support:**
- **Forms (.frm → .frx)** - Form binary resources
- **UserControls (.ctl → .frx)** - User control binary resources
- **Property Pages (.pag → .frx)** - Property page binary resources
- **UserDocuments (.dob → .frx)** - User document binary resources

The FRX binary format is **identical** across all these file types - VB6 uses the same encoding scheme regardless of the parent file type.

**Resource Types Parsed:**
- **Images & Icons** - Binary image data (12-byte header with `lt\0\0` signature)
- **String Resources** - Text properties stored in binary format (16-bit records)
- **List Items** - ComboBox/ListBox items (list records with 0x03/0x07 signature)
- **Font Data** - Binary font descriptors
- **Control Properties** - Binary-serialized control data
- **OLE Objects** - Embedded object data

**Record Format Support:**
- **12-byte header records** - Large binary blobs (images, OLE objects)
- **16-bit records** - Short string/binary data (marked with 0xFF)
- **List records** - List box/combo box items with item count header
- **4-byte header records** - Medium-sized text/binary data
- **8-bit records** - Very small resources (fallback type)

**Features:**
- Offset-based resource access (no file header required)
- Automatic record type detection based on signatures
- VB6 off-by-one error handling (size correction)
- Empty record detection (removed icons/images)
- List item extraction with UTF-8 conversion
- Comprehensive error handling with detailed messages

**API Functions:**
```rust
// Load any resource from any FRX file type
resource_file_resolver(file_path: &str, offset: usize) -> Result<Vec<u8>, io::Error>

// Parse list items from FRX binary data
list_resolver(buffer: &[u8]) -> Vec<String>
```

### Completed Features ✅
- ✅ Project-wide symbol resolution (.vbp parsing)
- ✅ Binary resource file parsing (.frx for all visual designers)
- ✅ Compiled resource file support (.res read/write for Win32 resources)
- ✅ Form control support with comprehensive property parsing
- ✅ CLI commands for resource file operations (read-res, write-res, parse-string-table)
- ✅ MCP server resource file tools for Claude AI integration
- ✅ Workspace management with multi-project support
- ✅ Cross-file symbol lookup and navigation
- ✅ Reference tracking across project files
- ✅ VB6 text encoding utilities (ANSI/Unicode conversion)
- ✅ Integration test suite for VBP and RES parsing

### Planned Features
- ⏳ Full cross-file rename refactoring
- ⏳ Workspace-wide code actions
- ⏳ Type library (.tlb/.olb) parsing for external references

## Performance

- **Fast startup**: < 100ms
- **Incremental parsing**: Only re-parses changed sections
- **Memory efficient**: Ropey data structure for large files
- **Async operations**: Non-blocking Claude API calls

## Limitations

- **External libraries**: Type library references are parsed but not yet resolved for IntelliSense
- **ActiveX controls**: OCX references are tracked but type definitions not yet loaded
- **Form designer**: Visual layout and control positioning not yet integrated into LSP features

## Roadmap

### Phase 1: Foundation ✅
- [x] Basic LSP server with tower-lsp
- [x] Core completions and diagnostics
- [x] Claude API integration

### Phase 2: Enhanced Parsing ✅
- [x] Full tree-sitter grammar for VB6 (97% validated against ANTLR reference)
- [x] Complete AST with all language constructs
- [x] Symbol table with hierarchical scopes
- [x] Position-based symbol lookup
- [x] VBP project file parsing (comprehensive support)
- [x] FRX binary resource file parsing (universal support for all designers)
- [x] RES compiled resource file support (read/write all Win32 resource types)
- [x] Multi-file symbol resolution
- [x] Workspace management with VBP discovery

### Phase 3: Advanced Features ✅
- [x] Semantic highlighting
- [x] Go-to-definition (single file + cross-file for public symbols)
- [x] Find all references (single file + cross-project)
- [x] Document symbols (outline)
- [x] Cross-file symbol lookup (project-wide)
- [x] Public symbol indexing across workspace
- [x] Binary resource parsing for Forms, UserControls, PropertyPages, UserDocuments
- [x] Form control support with property parsing
- [x] CLI commands for resource file operations
- [x] Compiled resource file (.res) support (read/write all Win32 resource types)
- [x] MCP server resource file tools (vb6_read_res_file, vb6_write_res_file, vb6_get_string_table)
- [x] Encoding utilities for VB6 text formats
- [x] Integration test suite for VBP and RES parsing
- [ ] Full cross-file rename refactoring
- [ ] Code actions (quick fixes)
- [ ] Type library (.tlb/.olb) resolution for external references

## Contributing

Contributions welcome! Please:

1. Fork the repository
2. Create a feature branch
3. Make your changes
4. Add tests if applicable
5. Run `cargo fmt` and `cargo clippy`
6. Submit a pull request

## License

MIT License - see LICENSE file for details

## Acknowledgments

- **tower-lsp**: Excellent LSP framework for Rust
- **tree-sitter**: Fast incremental parsing with error recovery
- **Model Context Protocol (MCP)**: For enabling AI tool integration
- **ANTLR4 VBA Grammar**: Reference grammar used for validation
- **Anthropic**: Claude AI for intelligent assistance
- **VB6 Community**: For keeping the legacy alive

## Support

- Issues: https://github.com/yourusername/vb6-lsp/issues
- Discussions: https://github.com/yourusername/vb6-lsp/discussions

## Resources

- [LSP Specification](https://microsoft.github.io/language-server-protocol/)
- [Model Context Protocol (MCP)](https://modelcontextprotocol.io/)
- [VB6 Language Reference](https://learn.microsoft.com/en-us/previous-versions/visualstudio/visual-basic-6/visual-basic-6.0-documentation)
- [Claude API Documentation](https://docs.anthropic.com/)
- [Tree-sitter Documentation](https://tree-sitter.github.io/tree-sitter/)
