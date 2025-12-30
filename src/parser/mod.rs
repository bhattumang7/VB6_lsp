//! VB6 Parser Module
//!
//! Parses Visual Basic 6 source code into an AST.
//! Handles .bas, .cls, .frm, and .ctl files.
//!
//! Uses tree-sitter for incremental parsing with error recovery.

mod lexer;
mod ast;
mod tree_sitter;
mod converter;

pub use ast::*;
pub use tree_sitter::{TreeSitterVb6Parser, VB6QueryRunner};
pub use converter::ParseErrorInfo;

use tower_lsp::lsp_types::{Position, Range, TextEdit};

/// Parse error with location information
#[derive(Debug, Clone)]
pub struct ParseError {
    pub message: String,
    pub range: Range,
}

/// VB6 Parser using tree-sitter for incremental parsing
pub struct Vb6Parser {
    ts_parser: TreeSitterVb6Parser,
    /// Stored tree for incremental parsing
    last_tree: Option<::tree_sitter::Tree>,
}

impl Vb6Parser {
    pub fn new() -> Self {
        Self {
            ts_parser: TreeSitterVb6Parser::new().expect("Failed to create tree-sitter parser"),
            last_tree: None,
        }
    }

    /// Parse VB6 source code into an AST using tree-sitter
    pub fn parse(&mut self, source: &str) -> std::result::Result<Vb6Ast, Vec<ParseError>> {
        // Use incremental parsing if we have a previous tree
        let tree = self.ts_parser.parse(source, self.last_tree.as_ref());

        match tree {
            Some(tree) => {
                // Convert tree-sitter tree to our AST
                let ast = converter::convert_tree(&tree, source);

                // Extract any parse errors
                let error_infos = converter::extract_errors(&tree, source);
                let errors: Vec<ParseError> = error_infos.into_iter().map(|e| ParseError {
                    message: e.message,
                    range: Range {
                        start: Position {
                            line: e.line as u32,
                            character: e.column as u32,
                        },
                        end: Position {
                            line: e.end_line as u32,
                            character: e.end_column as u32,
                        },
                    },
                }).collect();

                // Store tree for incremental parsing
                self.last_tree = Some(tree);

                // Tree-sitter provides partial AST even with errors
                if errors.is_empty() {
                    Ok(ast)
                } else {
                    // Return AST anyway for error-tolerant parsing
                    // The LSP can still use the partial AST while showing errors
                    Ok(ast)
                }
            }
            None => Err(vec![ParseError {
                message: "Failed to parse source".to_string(),
                range: Range {
                    start: Position { line: 0, character: 0 },
                    end: Position { line: 0, character: 0 },
                },
            }]),
        }
    }

    /// Get parse errors without failing the entire parse
    pub fn get_errors(&mut self, source: &str) -> Vec<ParseError> {
        if let Some(tree) = self.ts_parser.parse(source, self.last_tree.as_ref()) {
            let error_infos = converter::extract_errors(&tree, source);
            error_infos.into_iter().map(|e| ParseError {
                message: e.message,
                range: Range {
                    start: Position {
                        line: e.line as u32,
                        character: e.column as u32,
                    },
                    end: Position {
                        line: e.end_line as u32,
                        character: e.end_column as u32,
                    },
                },
            }).collect()
        } else {
            vec![]
        }
    }

    /// Clear the cached tree (useful when document is closed)
    pub fn clear_cache(&mut self) {
        self.last_tree = None;
    }

    /// Get a reference to the last parsed tree-sitter tree
    pub fn get_tree(&self) -> Option<&::tree_sitter::Tree> {
        self.last_tree.as_ref()
    }

    /// Parse VB6 source code using the legacy line-based parser
    /// Kept for compatibility but tree-sitter is preferred
    pub fn parse_legacy(&self, source: &str) -> std::result::Result<Vb6Ast, Vec<ParseError>> {
        let mut ast = Vb6Ast::new();
        let mut errors = Vec::new();

        let lines: Vec<&str> = source.lines().collect();

        for (line_num, line) in lines.iter().enumerate() {
            let trimmed = line.trim();

            // Skip empty lines and comments
            if trimmed.is_empty() {
                continue;
            }

            // Parse the line
            if let Err(e) = self.parse_line(trimmed, line_num, &mut ast) {
                errors.push(e);
            }
        }

        if errors.is_empty() {
            Ok(ast)
        } else {
            Err(errors)
        }
    }

    /// Parse a single line of VB6 code
    fn parse_line(
        &self,
        line: &str,
        line_num: usize,
        ast: &mut Vb6Ast,
    ) -> std::result::Result<(), ParseError> {
        let upper = line.to_uppercase();

        // Comment
        if upper.starts_with("'") || upper.starts_with("REM ") {
            ast.add_comment(line_num, line);
            return Ok(());
        }

        // Option statements
        if upper.starts_with("OPTION ") {
            ast.add_option(line_num, line);
            return Ok(());
        }

        // Attribute statements (in .cls/.frm files)
        if upper.starts_with("ATTRIBUTE ") {
            ast.add_attribute(line_num, line);
            return Ok(());
        }

        // Variable declarations
        if upper.starts_with("DIM ")
            || upper.starts_with("PRIVATE ")
            || upper.starts_with("PUBLIC ")
            || upper.starts_with("GLOBAL ")
            || upper.starts_with("STATIC ")
        {
            return self.parse_declaration(line, line_num, ast);
        }

        // Const declarations
        if upper.starts_with("CONST ") || upper.contains(" CONST ") {
            return self.parse_const(line, line_num, ast);
        }

        // Type declarations
        if upper.starts_with("TYPE ") || upper.starts_with("PRIVATE TYPE ") || upper.starts_with("PUBLIC TYPE ") {
            return self.parse_type(line, line_num, ast);
        }

        // Enum declarations
        if upper.starts_with("ENUM ") || upper.starts_with("PRIVATE ENUM ") || upper.starts_with("PUBLIC ENUM ") {
            return self.parse_enum(line, line_num, ast);
        }

        // Sub/Function/Property declarations
        if upper.contains("SUB ") || upper.contains("FUNCTION ") || upper.contains("PROPERTY ") {
            return self.parse_procedure(line, line_num, ast);
        }

        // Other statements (assignments, calls, etc.)
        ast.add_statement(line_num, line);

        Ok(())
    }

    /// Parse a variable declaration
    fn parse_declaration(
        &self,
        line: &str,
        line_num: usize,
        ast: &mut Vb6Ast,
    ) -> std::result::Result<(), ParseError> {
        // Basic parsing - extract visibility and variable info
        let upper = line.to_uppercase();
        let visibility = if upper.starts_with("PRIVATE") {
            Visibility::Private
        } else if upper.starts_with("PUBLIC") || upper.starts_with("GLOBAL") {
            Visibility::Public
        } else {
            Visibility::Private // Dim defaults to private
        };

        // Extract variable name and type (simplified)
        // Format: [Visibility] Dim|Static VarName [As Type]
        let parts: Vec<&str> = line.split_whitespace().collect();
        if parts.len() >= 2 {
            let name_part = if parts[0].to_uppercase() == "DIM" {
                parts.get(1)
            } else {
                parts.get(2)
            };

            if let Some(name) = name_part {
                let var_name = name.trim_end_matches(',');
                let var_type = self.extract_type(line);

                ast.add_variable(Variable {
                    name: var_name.to_string(),
                    var_type,
                    visibility,
                    line: line_num,
                    is_array: line.contains("("),
                });
            }
        }

        Ok(())
    }

    /// Parse a constant declaration
    fn parse_const(
        &self,
        line: &str,
        line_num: usize,
        ast: &mut Vb6Ast,
    ) -> std::result::Result<(), ParseError> {
        // Format: [Visibility] Const NAME = VALUE
        let upper = line.to_uppercase();
        let visibility = if upper.starts_with("PRIVATE") {
            Visibility::Private
        } else if upper.starts_with("PUBLIC") {
            Visibility::Public
        } else {
            Visibility::Private
        };

        if let Some(eq_pos) = line.find('=') {
            let before_eq = &line[..eq_pos].trim();
            let parts: Vec<&str> = before_eq.split_whitespace().collect();

            if let Some(name) = parts.last() {
                ast.add_constant(Constant {
                    name: name.to_string(),
                    value: line[eq_pos + 1..].trim().to_string(),
                    visibility,
                    line: line_num,
                });
            }
        }

        Ok(())
    }

    /// Parse a Type declaration
    fn parse_type(
        &self,
        line: &str,
        line_num: usize,
        ast: &mut Vb6Ast,
    ) -> std::result::Result<(), ParseError> {
        let upper = line.to_uppercase();
        let visibility = if upper.starts_with("PRIVATE") {
            Visibility::Private
        } else {
            Visibility::Public
        };

        let parts: Vec<&str> = line.split_whitespace().collect();
        if let Some(pos) = parts.iter().position(|p| p.to_uppercase() == "TYPE") {
            if let Some(name) = parts.get(pos + 1) {
                ast.add_user_type(UserType {
                    name: name.to_string(),
                    visibility,
                    line: line_num,
                    members: Vec::new(), // Will be filled when parsing following lines
                });
            }
        }

        Ok(())
    }

    /// Parse an Enum declaration
    fn parse_enum(
        &self,
        line: &str,
        line_num: usize,
        ast: &mut Vb6Ast,
    ) -> std::result::Result<(), ParseError> {
        let upper = line.to_uppercase();
        let visibility = if upper.starts_with("PRIVATE") {
            Visibility::Private
        } else {
            Visibility::Public
        };

        let parts: Vec<&str> = line.split_whitespace().collect();
        if let Some(pos) = parts.iter().position(|p| p.to_uppercase() == "ENUM") {
            if let Some(name) = parts.get(pos + 1) {
                ast.add_enum(Enumeration {
                    name: name.to_string(),
                    visibility,
                    line: line_num,
                    members: Vec::new(),
                });
            }
        }

        Ok(())
    }

    /// Parse a Sub/Function/Property declaration
    fn parse_procedure(
        &self,
        line: &str,
        line_num: usize,
        ast: &mut Vb6Ast,
    ) -> std::result::Result<(), ParseError> {
        let upper = line.to_uppercase();

        // Determine visibility
        let visibility = if upper.starts_with("PRIVATE") {
            Visibility::Private
        } else if upper.starts_with("PUBLIC") {
            Visibility::Public
        } else if upper.starts_with("FRIEND") {
            Visibility::Friend
        } else {
            Visibility::Public // Default for procedures
        };

        // Determine procedure type
        let proc_type = if upper.contains(" SUB ") || upper.starts_with("SUB ") {
            ProcedureType::Sub
        } else if upper.contains(" FUNCTION ") || upper.starts_with("FUNCTION ") {
            ProcedureType::Function
        } else if upper.contains(" PROPERTY GET ") || upper.starts_with("PROPERTY GET ") {
            ProcedureType::PropertyGet
        } else if upper.contains(" PROPERTY LET ") || upper.starts_with("PROPERTY LET ") {
            ProcedureType::PropertyLet
        } else if upper.contains(" PROPERTY SET ") || upper.starts_with("PROPERTY SET ") {
            ProcedureType::PropertySet
        } else {
            ProcedureType::Sub
        };

        // Extract name and parameters
        if let Some(paren_start) = line.find('(') {
            let before_paren = &line[..paren_start];
            let parts: Vec<&str> = before_paren.split_whitespace().collect();

            if let Some(name) = parts.last() {
                let params = self.extract_parameters(line);
                let return_type = self.extract_return_type(line);

                ast.add_procedure(Procedure {
                    name: name.to_string(),
                    proc_type,
                    visibility,
                    line: line_num,
                    parameters: params,
                    return_type,
                    end_line: None, // Will be set when End Sub/Function is found
                });
            }
        }

        Ok(())
    }

    /// Extract type from "As Type" clause
    fn extract_type(&self, line: &str) -> Option<String> {
        let upper = line.to_uppercase();
        if let Some(pos) = upper.find(" AS ") {
            let after_as = &line[pos + 4..];
            let type_name: String = after_as
                .split(|c: char| !c.is_alphanumeric())
                .next()
                .unwrap_or("")
                .trim()
                .to_string();

            if !type_name.is_empty() {
                return Some(type_name);
            }
        }
        None
    }

    /// Extract parameters from a procedure declaration
    fn extract_parameters(&self, line: &str) -> Vec<Parameter> {
        let mut params = Vec::new();

        if let Some(start) = line.find('(') {
            if let Some(end) = line.rfind(')') {
                let param_str = &line[start + 1..end];

                for param in param_str.split(',') {
                    let param = param.trim();
                    if param.is_empty() {
                        continue;
                    }

                    let upper = param.to_uppercase();
                    let by_ref = !upper.starts_with("BYVAL");
                    let optional = upper.contains("OPTIONAL");

                    let parts: Vec<&str> = param.split_whitespace().collect();
                    let name = parts
                        .iter()
                        .find(|p| {
                            let u = p.to_uppercase();
                            u != "BYVAL" && u != "BYREF" && u != "OPTIONAL" && u != "AS"
                        })
                        .map(|s| s.to_string())
                        .unwrap_or_default();

                    let param_type = self.extract_type(param);

                    if !name.is_empty() {
                        params.push(Parameter {
                            name,
                            param_type,
                            by_ref,
                            optional,
                        });
                    }
                }
            }
        }

        params
    }

    /// Extract return type from a function declaration
    fn extract_return_type(&self, line: &str) -> Option<String> {
        // Look for "As Type" after the closing parenthesis
        if let Some(paren_end) = line.rfind(')') {
            let after_paren = &line[paren_end + 1..];
            return self.extract_type(&format!(" {}", after_paren));
        }
        None
    }

    /// Format VB6 source code
    pub fn format(&self, source: &str) -> Option<Vec<TextEdit>> {
        let mut edits = Vec::new();
        let lines: Vec<&str> = source.lines().collect();
        let mut indent_level: usize = 0;

        for (line_num, line) in lines.iter().enumerate() {
            let trimmed = line.trim();
            let upper = trimmed.to_uppercase();

            // Decrease indent before these keywords
            if upper.starts_with("END ")
                || upper == "END"
                || upper.starts_with("ELSE")
                || upper.starts_with("ELSEIF")
                || upper.starts_with("CASE ")
                || upper.starts_with("LOOP")
                || upper.starts_with("NEXT")
                || upper.starts_with("WEND")
            {
                indent_level = indent_level.saturating_sub(1);
            }

            // Calculate expected indentation
            let expected_indent = "    ".repeat(indent_level);
            let expected_line = format!("{}{}", expected_indent, trimmed);

            // Create edit if line differs
            if *line != expected_line && !trimmed.is_empty() {
                edits.push(TextEdit {
                    range: Range {
                        start: Position {
                            line: line_num as u32,
                            character: 0,
                        },
                        end: Position {
                            line: line_num as u32,
                            character: line.len() as u32,
                        },
                    },
                    new_text: expected_line,
                });
            }

            // Increase indent after these keywords
            if upper.starts_with("IF ") && upper.contains(" THEN") && !upper.contains(" THEN ")
                || upper.starts_with("FOR ")
                || upper.starts_with("DO ")
                || upper.starts_with("DO")
                || upper.starts_with("WHILE ")
                || upper.starts_with("SELECT CASE")
                || upper.starts_with("WITH ")
                || upper.starts_with("SUB ")
                || upper.starts_with("FUNCTION ")
                || upper.starts_with("PROPERTY ")
                || upper.starts_with("TYPE ")
                || upper.starts_with("ENUM ")
                || upper.starts_with("PRIVATE SUB ")
                || upper.starts_with("PRIVATE FUNCTION ")
                || upper.starts_with("PUBLIC SUB ")
                || upper.starts_with("PUBLIC FUNCTION ")
                || upper.starts_with("ELSE")
                || upper.starts_with("ELSEIF")
                || upper.starts_with("CASE ")
            {
                indent_level += 1;
            }
        }

        if edits.is_empty() {
            None
        } else {
            Some(edits)
        }
    }
}

impl Default for Vb6Parser {
    fn default() -> Self {
        Self::new()
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn test_tree_sitter_parse() {
        let mut parser = Vb6Parser::new();
        let source = r#"
Option Explicit

Dim x As Integer
Private y As String

Sub Main()
    x = 10
    y = "Hello"
End Sub

Function Add(a As Integer, b As Integer) As Integer
    Add = a + b
End Function
"#;
        let result = parser.parse(source);
        assert!(result.is_ok());

        let ast = result.unwrap();
        assert_eq!(ast.options.len(), 1);
        assert_eq!(ast.variables.len(), 2);
        assert_eq!(ast.procedures.len(), 2);
    }

    #[test]
    fn test_incremental_parse() {
        let mut parser = Vb6Parser::new();

        // First parse
        let source1 = "Dim x As Integer";
        let _ = parser.parse(source1);

        // Second parse should use incremental parsing
        let source2 = "Dim x As Integer\nDim y As String";
        let result = parser.parse(source2);

        assert!(result.is_ok());
        let ast = result.unwrap();
        assert_eq!(ast.variables.len(), 2);
    }
}
