# Contributing to VB6 Language Server

Thank you for your interest in contributing! This document provides guidelines and instructions for contributing to the project.

## Getting Started

### Prerequisites

1. **Rust 1.70+** - [Install Rust](https://rustup.rs/)
2. **Git** - For version control
3. **VSCode or Neovim** - For testing the LSP

### Setup Development Environment

```bash
# Clone the repository
git clone https://github.com/yourusername/vb6-lsp.git
cd vb6-lsp

# Build the project
cargo build

# Run tests
cargo test

# Run with debug logging
RUST_LOG=debug cargo run
```

## Development Workflow

### 1. Create a Branch

```bash
git checkout -b feature/your-feature-name
```

### 2. Make Changes

Follow these coding standards:

#### Code Style
- Use `cargo fmt` to format code
- Follow Rust naming conventions
- Add documentation comments (`///`) for public APIs
- Keep functions focused and under 50 lines when possible

#### Example:
```rust
/// Parses a VB6 function declaration
///
/// # Arguments
/// * `line` - The source code line to parse
/// * `line_num` - Line number for error reporting
///
/// # Returns
/// A `Procedure` AST node or a `ParseError`
pub fn parse_function(&self, line: &str, line_num: usize) -> Result<Procedure, ParseError> {
    // Implementation
}
```

### 3. Write Tests

Add tests for new functionality:

```rust
#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn test_parse_function() {
        let parser = Vb6Parser::new();
        let result = parser.parse("Public Function Test() As String");
        assert!(result.is_ok());
    }
}
```

### 4. Run Quality Checks

```bash
# Format code
cargo fmt

# Check for common mistakes
cargo clippy

# Run all tests
cargo test

# Check documentation
cargo doc --no-deps --open
```

### 5. Commit Changes

Use clear, descriptive commit messages:

```bash
git add .
git commit -m "Add support for VB6 property procedures

- Implement Property Get/Let/Set parsing
- Add tests for property declarations
- Update documentation"
```

### 6. Push and Create PR

```bash
git push origin feature/your-feature-name
```

Then create a Pull Request on GitHub.

## Areas for Contribution

### 🌟 High Priority

1. **Improved Parser**
   - Tree-sitter grammar for VB6
   - Better error recovery
   - Multi-line statement support

2. **Project File Support**
   - Parse .vbp files
   - Multi-file symbol resolution

3. **Enhanced Diagnostics**
   - Type checking
   - Dead code detection
   - Unused variable warnings

### 🎯 Medium Priority

4. **Code Actions**
   - Quick fixes for common errors
   - Refactoring suggestions
   - Import organization

5. **Testing**
   - Expand test coverage
   - Integration tests
   - Performance benchmarks

6. **Documentation**
   - More examples
   - Architecture documentation
   - Tutorial videos

### 💡 Future Ideas

7. **Form Designer Support**
   - Parse .frm files completely
   - Visual control information

8. **Migration Tools**
   - VB6 to VB.NET converter
   - VB6 to C# converter

9. **IDE Extensions**
   - VSCode extension with UI
   - IntelliJ plugin
   - Sublime Text support

## Code Organization

### Directory Structure

```
src/
├── main.rs              # Application entry point
├── lsp/                 # LSP server implementation
│   ├── mod.rs           # Main handlers (textDocument/*, initialize, etc.)
│   ├── capabilities.rs  # LSP capability definitions
│   ├── document.rs      # Document management utilities
│   └── handlers.rs      # Specific request handlers
├── parser/              # VB6 parsing logic
│   ├── mod.rs           # Main parser
│   ├── ast.rs           # AST node definitions
│   └── lexer.rs         # Tokenization
├── analysis/            # Semantic analysis
│   └── mod.rs           # Diagnostics, completions, hover, etc.
└── claude/              # Claude AI integration
    └── mod.rs           # API client and helper functions
```

### Module Guidelines

#### Parser (`src/parser/`)
- Keep parsing logic separate from LSP logic
- Maintain an immutable AST
- Return detailed error information

#### LSP (`src/lsp/`)
- Implement LSP spec faithfully
- Use async/await for I/O operations
- Log important events

#### Analysis (`src/analysis/`)
- Work with AST, not raw text
- Cache analysis results when possible
- Provide actionable diagnostics

#### Claude (`src/claude/`)
- Handle API errors gracefully
- Respect rate limits
- Cache responses when appropriate

## Testing Guidelines

### Unit Tests

Place tests in the same file as the code:

```rust
#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn test_feature() {
        // Test code
    }
}
```

### Integration Tests

Place in `tests/` directory:

```rust
// tests/integration_test.rs
use vb6_lsp::parser::Vb6Parser;

#[test]
fn test_parse_complete_file() {
    let source = include_str!("fixtures/sample.bas");
    let parser = Vb6Parser::new();
    let ast = parser.parse(source).unwrap();
    assert_eq!(ast.procedures.len(), 5);
}
```

### Test VB6 Files

Add test files to `examples/` or `tests/fixtures/`:

```vb
' tests/fixtures/test_case.bas
Public Function Add(a As Integer, b As Integer) As Integer
    Add = a + b
End Function
```

## Documentation

### Code Documentation

Use doc comments for public APIs:

```rust
/// Parses Visual Basic 6 source code into an Abstract Syntax Tree.
///
/// # Examples
///
/// ```
/// use vb6_lsp::parser::Vb6Parser;
///
/// let parser = Vb6Parser::new();
/// let ast = parser.parse("Public Sub Main()\nEnd Sub").unwrap();
/// ```
pub fn parse(&self, source: &str) -> Result<Vb6Ast, Vec<ParseError>> {
    // Implementation
}
```

### README Updates

Update README.md when adding features:
- Add to feature list
- Update roadmap
- Add usage examples

## Pull Request Process

1. **Update Documentation**: README, code comments, CHANGELOG
2. **Add Tests**: Ensure new code is tested
3. **Pass CI**: All checks must pass
4. **Get Review**: Wait for maintainer feedback
5. **Address Comments**: Make requested changes
6. **Merge**: Maintainer will merge when ready

### PR Template

```markdown
## Description
Brief description of changes

## Type of Change
- [ ] Bug fix
- [ ] New feature
- [ ] Breaking change
- [ ] Documentation update

## Testing
- [ ] Unit tests added/updated
- [ ] Integration tests added/updated
- [ ] Manual testing completed

## Checklist
- [ ] Code follows project style guidelines
- [ ] Self-review completed
- [ ] Comments added for complex code
- [ ] Documentation updated
- [ ] No new warnings generated
```

## Community

### Getting Help

- **GitHub Issues**: Bug reports and feature requests
- **Discussions**: Questions and general discussion
- **Discord**: [Join our community](#) (coming soon)

### Code of Conduct

- Be respectful and inclusive
- Focus on constructive feedback
- Help newcomers
- Give credit where due

## License

By contributing, you agree that your contributions will be licensed under the MIT License.

## Questions?

Open an issue or start a discussion on GitHub!

---

Thank you for contributing to VB6 LSP! 🎉
