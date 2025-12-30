# Getting Started with VB6 Language Server

This guide will help you set up and start using the VB6 Language Server with your favorite editor.

## Quick Start (Windows)

### 1. Install Rust

Open PowerShell and run:

```powershell
# Download rustup installer
Invoke-WebRequest -Uri https://win.rustup.rs/x86_64 -OutFile rustup-init.exe

# Run installer (follow prompts, use default settings)
.\rustup-init.exe

# Restart your terminal, then verify installation
rustc --version
cargo --version
```

### 2. Build VB6 LSP

```powershell
# Navigate to project directory
cd C:\projects\VB6_lsp

# Build release version (optimized, smaller binary)
cargo build --release

# The executable will be at: target\release\vb6-lsp.exe
```

### 3. Test the Server

```powershell
# Run the server manually (it will wait for LSP client input)
.\target\release\vb6-lsp.exe

# You should see log output. Press Ctrl+C to exit.
```

## Editor Setup

### Visual Studio Code

#### Option 1: Using Custom LSP Client Extension

1. **Install Generic LSP Client**:
   - Open VSCode
   - Press `Ctrl+P`
   - Type: `ext install mattn.vscode-lsp`

2. **Configure Settings**:
   - Press `Ctrl+,` to open settings
   - Click "Open Settings (JSON)"
   - Add this configuration:

   ```json
   {
     "genericLSP.servers": {
       "vb6": {
         "command": "C:\\projects\\VB6_lsp\\target\\release\\vb6-lsp.exe",
         "filetypes": ["vb", "bas", "cls", "frm", "ctl"],
         "options": {
           "env": {
             "RUST_LOG": "vb6_lsp=info"
           }
         }
       }
     }
   }
   ```

3. **Associate File Extensions**:
   Add to `settings.json`:
   ```json
   {
     "files.associations": {
       "*.bas": "vb",
       "*.cls": "vb",
       "*.frm": "vb",
       "*.ctl": "vb"
     }
   }
   ```

#### Option 2: Manual Launch (for testing)

1. Open terminal in VSCode
2. Run: `.\target\release\vb6-lsp.exe`
3. Open a VB6 file

### Neovim

Add to your `init.lua`:

```lua
local lspconfig = require('lspconfig')
local configs = require('lspconfig.configs')

-- Register VB6 LSP
if not configs.vb6_lsp then
  configs.vb6_lsp = {
    default_config = {
      cmd = {'C:\\projects\\VB6_lsp\\target\\release\\vb6-lsp.exe'},
      filetypes = {'vb', 'bas', 'cls', 'frm', 'ctl'},
      root_dir = function(fname)
        return lspconfig.util.root_pattern('.git', '.vbp')(fname) or vim.fn.getcwd()
      end,
      settings = {},
    },
  }
end

-- Attach to VB6 files
lspconfig.vb6_lsp.setup{
  on_attach = function(client, bufnr)
    -- Enable completion
    vim.api.nvim_buf_set_option(bufnr, 'omnifunc', 'v:lua.vim.lsp.omnifunc')

    -- Keybindings
    local bufopts = { noremap=true, silent=true, buffer=bufnr }
    vim.keymap.set('n', 'gD', vim.lsp.buf.declaration, bufopts)
    vim.keymap.set('n', 'gd', vim.lsp.buf.definition, bufopts)
    vim.keymap.set('n', 'K', vim.lsp.buf.hover, bufopts)
    vim.keymap.set('n', 'gi', vim.lsp.buf.implementation, bufopts)
    vim.keymap.set('n', '<space>rn', vim.lsp.buf.rename, bufopts)
    vim.keymap.set('n', 'gr', vim.lsp.buf.references, bufopts)
  end,
}

-- File type detection for VB6
vim.filetype.add({
  extension = {
    bas = 'vb',
    cls = 'vb',
    frm = 'vb',
    ctl = 'vb',
  },
})
```

### Sublime Text

Install LSP package, then add to `LSP.sublime-settings`:

```json
{
  "clients": {
    "vb6-lsp": {
      "enabled": true,
      "command": ["C:\\projects\\VB6_lsp\\target\\release\\vb6-lsp.exe"],
      "selector": "source.vb",
      "env": {
        "RUST_LOG": "vb6_lsp=info"
      }
    }
  }
}
```

## Enabling Claude AI Features

### 1. Get an API Key

1. Visit https://console.anthropic.com
2. Sign up or log in
3. Navigate to API Keys
4. Create a new key
5. Copy the key (starts with `sk-ant-...`)

### 2. Set Environment Variable

**Windows (PowerShell):**
```powershell
# Temporary (current session only)
$env:ANTHROPIC_API_KEY = "sk-ant-api03-your-key-here"

# Permanent (system-wide)
[System.Environment]::SetEnvironmentVariable('ANTHROPIC_API_KEY', 'sk-ant-api03-your-key-here', 'User')
```

**Windows (Command Prompt):**
```cmd
setx ANTHROPIC_API_KEY "sk-ant-api03-your-key-here"
```

**Linux/Mac:**
```bash
# Add to ~/.bashrc or ~/.zshrc
export ANTHROPIC_API_KEY="sk-ant-api03-your-key-here"

# Reload
source ~/.bashrc
```

### 3. Verify

When you start the LSP, you should see:
```
[INFO] Claude AI integration enabled
```

Instead of:
```
[INFO] Claude AI integration disabled (no ANTHROPIC_API_KEY)
```

## Testing the Setup

### 1. Open Example File

```powershell
# Copy example to test location
copy examples\sample.bas test.bas

# Open in VSCode
code test.bas
```

### 2. Test Features

Try these features:

#### Hover Information
- Move cursor over a function name
- Should show function signature

#### Code Completion
- Type `Dim x As ` and press Ctrl+Space
- Should show type suggestions (Integer, String, etc.)

#### Go to Definition
- Right-click on a function call
- Select "Go to Definition"
- Should jump to function declaration

#### Diagnostics
- Add a line: `Public Function Test()`
- Don't add `End Function`
- Should show error about missing End statement

#### Formatting
- Right-click in file
- Select "Format Document"
- Code should be properly indented

### 3. Check Logs

View LSP output:

**VSCode:**
- Open Output panel: `Ctrl+Shift+U`
- Select "VB6" from dropdown

**Command Line:**
```powershell
# Run with verbose logging
$env:RUST_LOG = "vb6_lsp=debug"
.\target\release\vb6-lsp.exe
```

## Common Issues

### Issue: "Command not found"

**Solution**: Make sure Rust is installed and in PATH
```powershell
# Check if rustc is available
rustc --version

# If not, restart terminal or add to PATH:
$env:PATH += ";$env:USERPROFILE\.cargo\bin"
```

### Issue: "Cannot find vb6-lsp.exe"

**Solution**: Build the project first
```powershell
cargo build --release
```

The executable location:
- Windows: `target\release\vb6-lsp.exe`
- Linux/Mac: `target/release/vb6-lsp`

### Issue: LSP not responding

**Solution**: Check if process is running
```powershell
# Windows
Get-Process | Where-Object {$_.ProcessName -like "*vb6-lsp*"}

# If hung, kill it
Stop-Process -Name vb6-lsp
```

### Issue: Claude features not working

**Solutions**:
1. Verify API key is set:
   ```powershell
   $env:ANTHROPIC_API_KEY
   ```

2. Check logs for Claude errors:
   ```
   RUST_LOG=vb6_lsp=debug cargo run
   ```

3. Verify API key is valid at https://console.anthropic.com

### Issue: No syntax highlighting

**Solution**: VSCode needs a language definition. Install or create:

`.vscode/extensions/vb6/syntaxes/vb6.tmLanguage.json`:
```json
{
  "scopeName": "source.vb",
  "patterns": [
    {
      "name": "keyword.control.vb",
      "match": "\\b(If|Then|Else|ElseIf|End If|For|Next|Do|Loop|While|Select Case)\\b"
    },
    {
      "name": "keyword.other.vb",
      "match": "\\b(Dim|Private|Public|Function|Sub|As|Integer|String|Long)\\b"
    },
    {
      "name": "comment.line.vb",
      "match": "'.*$"
    }
  ]
}
```

## Next Steps

1. **Try the Examples**: Open files in `examples/` directory
2. **Read Documentation**: Check [README.md](README.md) for full feature list
3. **Explore AI Features**: Use Claude for code explanations
4. **Report Issues**: https://github.com/yourusername/vb6-lsp/issues
5. **Contribute**: See [CONTRIBUTING.md](CONTRIBUTING.md)

## Debugging

### Enable Debug Logging

```powershell
$env:RUST_LOG = "vb6_lsp=debug,tower_lsp=debug"
.\target\release\vb6-lsp.exe > lsp.log 2>&1
```

### View LSP Communication

**VSCode**: Set in `settings.json`:
```json
{
  "vb6.lsp.trace.server": "verbose"
}
```

**Neovim**: Use `:LspLog` command

### Test Parser Directly

```rust
// Create a test file: tests/manual_test.rs
use vb6_lsp::parser::Vb6Parser;

fn main() {
    let source = r#"
        Public Function Add(a As Integer, b As Integer) As Integer
            Add = a + b
        End Function
    "#;

    let parser = Vb6Parser::new();
    match parser.parse(source) {
        Ok(ast) => println!("AST: {:#?}", ast),
        Err(errors) => println!("Errors: {:#?}", errors),
    }
}
```

Run with: `cargo run --bin manual_test`

## Performance Tuning

### Optimize Binary Size

```toml
# Add to Cargo.toml
[profile.release]
opt-level = 'z'     # Optimize for size
lto = true          # Link-time optimization
codegen-units = 1   # Better optimization
strip = true        # Remove debug symbols
```

Rebuild:
```powershell
cargo build --release
```

### Reduce Memory Usage

Limit Claude API response size in `src/claude/mod.rs`:
```rust
let request = ClaudeRequest {
    model: self.model.clone(),
    max_tokens: 512,  // Reduce from 1024
    // ...
};
```

## Getting Help

- **Documentation**: [README.md](README.md)
- **Issues**: https://github.com/yourusername/vb6-lsp/issues
- **Discussions**: https://github.com/yourusername/vb6-lsp/discussions
- **Email**: your-email@example.com

---

Happy coding with VB6! 🚀
