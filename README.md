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
- **Go to Definition** - Navigate to symbol definitions
- **Find References** - Find all references to a symbol
- **Document Symbols** - Outline view of file structure
- **Diagnostics** - Real-time syntax and semantic error checking
- **Code Formatting** - Automatic code indentation and formatting
- **Rename Refactoring** - Safe symbol renaming across files

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
- **vb6_get_symbols** - List all symbols (variables, functions, types) in a file
- **vb6_find_definition** - Go to where a symbol is defined
- **vb6_find_references** - Find all usages of a symbol
- **vb6_get_hover** - Get type/signature info at a position
- **vb6_get_completions** - Get code completion suggestions
- **vb6_get_diagnostics** - Get parse errors and warnings

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
         "filetypes": ["vb", "bas", "cls", "frm", "ctl"]
       }
     }
   }
   ```

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
      filetypes = {'vb', 'bas', 'cls', 'frm', 'ctl'},
      root_dir = lspconfig.util.root_pattern('.git', '.vbp'),
      settings = {},
    },
  }
end

-- Start the LSP
lspconfig.vb6_lsp.setup{}
```

## Usage

### Running the Server

The LSP server communicates via stdin/stdout:

```bash
# Run directly
./target/release/vb6-lsp

# With logging
RUST_LOG=vb6_lsp=debug ./target/release/vb6-lsp
```

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
│   ├── main.rs                  # Entry point
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
│   │   ├── symbol.rs            # Symbol, SymbolKind, TypeInfo
│   │   ├── scope.rs             # Scope hierarchy management
│   │   └── position.rs          # Source positions and ranges
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

### Planned Features
- ⏳ Form designer integration (.frx parsing)
- ⏳ Project-wide symbol resolution (.vbp parsing)
- ⏳ Cross-file go-to-definition
- ⏳ Reference tracking and rename refactoring

## Performance

- **Fast startup**: < 100ms
- **Incremental parsing**: Only re-parses changed sections
- **Memory efficient**: Ropey data structure for large files
- **Async operations**: Non-blocking Claude API calls

## Limitations

- **Single-file analysis**: Symbol resolution is currently per-file only
- **No project files**: Doesn't parse .vbp project files yet
- **Forms**: Limited support for .frm visual designer syntax (.frx not parsed)
- **Reference tracking**: References are collected but not yet fully integrated

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
- [ ] VBP project file parsing
- [ ] Multi-file symbol resolution

### Phase 3: Advanced Features (Current)
- [x] Semantic highlighting
- [x] Go-to-definition (single file)
- [x] Find all references (single file)
- [x] Document symbols (outline)
- [ ] Cross-file go-to-definition
- [ ] Cross-file find references
- [ ] Rename refactoring
- [ ] Code actions (quick fixes)

### Phase 4: AI-Powered
- [ ] Claude-powered smart completions
- [ ] Migration suggestions (VB6 → modern languages)
- [ ] Code quality analysis
- [ ] Legacy code explanation

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
