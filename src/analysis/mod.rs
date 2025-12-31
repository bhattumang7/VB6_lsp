//! Code Analysis Module
//!
//! Provides semantic analysis, diagnostics, and code intelligence.
//! Includes a symbol table for precise position-based lookups.

mod builder;
mod position;
mod scope;
mod symbol;
mod symbol_table;

// Re-export symbol table types
pub use builder::build_symbol_table;
pub use position::{SourcePosition, SourceRange};
pub use scope::{Scope, ScopeId, ScopeKind};
pub use symbol::{ParameterInfo, Symbol, SymbolId, SymbolKind, TypeInfo, Visibility};
pub use symbol_table::{SymbolReference, SymbolTable};

use std::collections::HashMap;

use tower_lsp::lsp_types::*;

use crate::parser::{Procedure, ProcedureType, Vb6Ast, Visibility as AstVisibility};

/// Code analyzer with symbol table support
pub struct Analyzer {
    // Analysis state (reserved for future use)
}

impl Analyzer {
    pub fn new() -> Self {
        Self {}
    }

    // ==========================================
    // Legacy AST-based methods (for compatibility)
    // ==========================================

    /// Analyze AST and produce diagnostics
    pub fn analyze(&self, ast: &Vb6Ast) -> Vec<Diagnostic> {
        let mut diagnostics = Vec::new();

        // Check for duplicate declarations
        let mut var_names: HashMap<String, usize> = HashMap::new();
        for var in &ast.variables {
            if let Some(&first_line) = var_names.get(&var.name) {
                diagnostics.push(Diagnostic {
                    range: Range {
                        start: Position {
                            line: var.line as u32,
                            character: 0,
                        },
                        end: Position {
                            line: var.line as u32,
                            character: var.name.len() as u32,
                        },
                    },
                    severity: Some(DiagnosticSeverity::WARNING),
                    message: format!(
                        "Variable '{}' already declared at line {}",
                        var.name,
                        first_line + 1
                    ),
                    source: Some("vb6-lsp".to_string()),
                    ..Default::default()
                });
            } else {
                var_names.insert(var.name.clone(), var.line);
            }
        }

        // Check for procedures without End Sub/Function
        for proc in &ast.procedures {
            if proc.end_line.is_none() {
                diagnostics.push(Diagnostic {
                    range: Range {
                        start: Position {
                            line: proc.line as u32,
                            character: 0,
                        },
                        end: Position {
                            line: proc.line as u32,
                            character: 50,
                        },
                    },
                    severity: Some(DiagnosticSeverity::ERROR),
                    message: format!(
                        "{} '{}' is missing End statement",
                        match proc.proc_type {
                            ProcedureType::Sub => "Sub",
                            ProcedureType::Function => "Function",
                            _ => "Property",
                        },
                        proc.name
                    ),
                    source: Some("vb6-lsp".to_string()),
                    ..Default::default()
                });
            }
        }

        // Warn about Option Explicit
        if !ast
            .options
            .iter()
            .any(|o| o.to_uppercase().contains("EXPLICIT"))
        {
            diagnostics.push(Diagnostic {
                range: Range {
                    start: Position {
                        line: 0,
                        character: 0,
                    },
                    end: Position {
                        line: 0,
                        character: 0,
                    },
                },
                severity: Some(DiagnosticSeverity::INFORMATION),
                message: "Consider adding 'Option Explicit' to require variable declarations"
                    .to_string(),
                source: Some("vb6-lsp".to_string()),
                ..Default::default()
            });
        }

        diagnostics
    }

    /// Get code completions at a position (legacy)
    pub fn get_completions(&self, ast: &Vb6Ast, _position: Position) -> Vec<CompletionItem> {
        let mut items = Vec::new();

        // Add variables
        for var in &ast.variables {
            items.push(CompletionItem {
                label: var.name.clone(),
                kind: Some(CompletionItemKind::VARIABLE),
                detail: var.var_type.clone(),
                documentation: Some(Documentation::String(format!(
                    "{} variable",
                    match var.visibility {
                        AstVisibility::Public => "Public",
                        AstVisibility::Private => "Private",
                        AstVisibility::Friend => "Friend",
                    }
                ))),
                ..Default::default()
            });
        }

        // Add constants
        for constant in &ast.constants {
            items.push(CompletionItem {
                label: constant.name.clone(),
                kind: Some(CompletionItemKind::CONSTANT),
                detail: Some(constant.value.clone()),
                ..Default::default()
            });
        }

        // Add procedures
        for proc in &ast.procedures {
            let kind = match proc.proc_type {
                ProcedureType::Function => CompletionItemKind::FUNCTION,
                ProcedureType::Sub => CompletionItemKind::FUNCTION,
                _ => CompletionItemKind::PROPERTY,
            };

            let params: Vec<String> = proc
                .parameters
                .iter()
                .map(|p| {
                    let mut s = p.name.clone();
                    if let Some(ref t) = p.param_type {
                        s.push_str(&format!(" As {}", t));
                    }
                    s
                })
                .collect();

            let signature = format!("{}({})", proc.name, params.join(", "));

            items.push(CompletionItem {
                label: proc.name.clone(),
                kind: Some(kind),
                detail: Some(signature),
                insert_text: Some(format!("{}($1)", proc.name)),
                insert_text_format: Some(InsertTextFormat::SNIPPET),
                ..Default::default()
            });
        }

        // Add keywords
        items.extend(self.get_keyword_completions());

        items
    }

    /// Get hover information at a position (legacy)
    pub fn get_hover(&self, ast: &Vb6Ast, position: Position) -> Option<Hover> {
        let line = position.line as usize;

        // Check procedures
        for proc in &ast.procedures {
            if proc.line == line {
                let params: Vec<String> = proc
                    .parameters
                    .iter()
                    .map(|p| {
                        let mut s = format!(
                            "{} {}",
                            if p.by_ref { "ByRef" } else { "ByVal" },
                            p.name
                        );
                        if let Some(ref t) = p.param_type {
                            s.push_str(&format!(" As {}", t));
                        }
                        s
                    })
                    .collect();

                let mut signature = format!(
                    "{} {}{}({})",
                    match proc.visibility {
                        AstVisibility::Public => "Public",
                        AstVisibility::Private => "Private",
                        AstVisibility::Friend => "Friend",
                    },
                    match proc.proc_type {
                        ProcedureType::Sub => "Sub",
                        ProcedureType::Function => "Function",
                        ProcedureType::PropertyGet => "Property Get",
                        ProcedureType::PropertyLet => "Property Let",
                        ProcedureType::PropertySet => "Property Set",
                    },
                    proc.name,
                    params.join(", ")
                );

                if let Some(ref return_type) = proc.return_type {
                    signature.push_str(&format!(" As {}", return_type));
                }

                return Some(Hover {
                    contents: HoverContents::Markup(MarkupContent {
                        kind: MarkupKind::Markdown,
                        value: format!("```vb\n{}\n```", signature),
                    }),
                    range: None,
                });
            }
        }

        // Check variables
        for var in &ast.variables {
            if var.line == line {
                let mut info = format!(
                    "{} {} As {}",
                    match var.visibility {
                        AstVisibility::Public => "Public",
                        AstVisibility::Private => "Private",
                        AstVisibility::Friend => "Friend",
                    },
                    var.name,
                    var.var_type.as_ref().unwrap_or(&"Variant".to_string())
                );

                if var.is_array {
                    info.push_str(" (Array)");
                }

                return Some(Hover {
                    contents: HoverContents::Markup(MarkupContent {
                        kind: MarkupKind::Markdown,
                        value: format!("```vb\n{}\n```", info),
                    }),
                    range: None,
                });
            }
        }

        None
    }

    /// Get definition location (legacy)
    pub fn get_definition(
        &self,
        ast: &Vb6Ast,
        position: Position,
        uri: &Url,
    ) -> Option<GotoDefinitionResponse> {
        let line = position.line as usize;

        // Check procedures
        for proc in &ast.procedures {
            if proc.line == line {
                return Some(GotoDefinitionResponse::Scalar(Location {
                    uri: uri.clone(),
                    range: Range {
                        start: Position {
                            line: proc.line as u32,
                            character: 0,
                        },
                        end: Position {
                            line: proc.line as u32,
                            character: proc.name.len() as u32,
                        },
                    },
                }));
            }
        }

        None
    }

    /// Get references to a symbol (legacy - stub)
    pub fn get_references(
        &self,
        _ast: &Vb6Ast,
        _position: Position,
        _uri: &Url,
    ) -> Vec<Location> {
        Vec::new()
    }

    /// Get document symbols (legacy)
    pub fn get_document_symbols(&self, ast: &Vb6Ast) -> Vec<DocumentSymbol> {
        let mut symbols = Vec::new();

        // Add variables
        for var in &ast.variables {
            #[allow(deprecated)]
            symbols.push(DocumentSymbol {
                name: var.name.clone(),
                detail: var.var_type.clone(),
                kind: tower_lsp::lsp_types::SymbolKind::VARIABLE,
                range: Range {
                    start: Position {
                        line: var.line as u32,
                        character: 0,
                    },
                    end: Position {
                        line: var.line as u32,
                        character: 100,
                    },
                },
                selection_range: Range {
                    start: Position {
                        line: var.line as u32,
                        character: 0,
                    },
                    end: Position {
                        line: var.line as u32,
                        character: var.name.len() as u32,
                    },
                },
                children: None,
                tags: None,
                deprecated: None,
            });
        }

        // Add procedures
        for proc in &ast.procedures {
            let kind = match proc.proc_type {
                ProcedureType::Function => tower_lsp::lsp_types::SymbolKind::FUNCTION,
                ProcedureType::Sub => tower_lsp::lsp_types::SymbolKind::FUNCTION,
                _ => tower_lsp::lsp_types::SymbolKind::PROPERTY,
            };

            #[allow(deprecated)]
            symbols.push(DocumentSymbol {
                name: proc.name.clone(),
                detail: proc.return_type.clone(),
                kind,
                range: Range {
                    start: Position {
                        line: proc.line as u32,
                        character: 0,
                    },
                    end: Position {
                        line: proc.end_line.unwrap_or(proc.line) as u32,
                        character: 0,
                    },
                },
                selection_range: Range {
                    start: Position {
                        line: proc.line as u32,
                        character: 0,
                    },
                    end: Position {
                        line: proc.line as u32,
                        character: proc.name.len() as u32,
                    },
                },
                children: None,
                tags: None,
                deprecated: None,
            });
        }

        symbols
    }

    /// Get code actions (legacy - stub)
    pub fn get_code_actions(
        &self,
        _ast: &Vb6Ast,
        _range: Range,
        _context: &CodeActionContext,
    ) -> Vec<CodeActionOrCommand> {
        Vec::new()
    }

    /// Rename a symbol (legacy - stub)
    pub fn rename(
        &self,
        _ast: &Vb6Ast,
        _position: Position,
        _new_name: &str,
        _uri: &Url,
    ) -> Option<WorkspaceEdit> {
        None
    }

    // ==========================================
    // Symbol Table-based methods (enhanced)
    // ==========================================

    /// Get hover information using symbol table
    pub fn get_hover_with_symbols(
        &self,
        table: &SymbolTable,
        position: Position,
    ) -> Option<Hover> {
        let pos = SourcePosition::from_lsp(position);

        // Find symbol at position
        let symbol = table.symbol_at_position(pos)?;

        // Build hover content
        let signature = symbol.format_signature();

        Some(Hover {
            contents: HoverContents::Markup(MarkupContent {
                kind: MarkupKind::Markdown,
                value: format!("```vb\n{}\n```", signature),
            }),
            range: Some(symbol.name_range.to_lsp()),
        })
    }

    /// Get definition location using symbol table
    pub fn get_definition_with_symbols(
        &self,
        table: &SymbolTable,
        source: &str,
        position: Position,
    ) -> Option<GotoDefinitionResponse> {
        let pos = SourcePosition::from_lsp(position);

        // Try to find symbol at cursor position
        if let Some(symbol) = table.symbol_at_position(pos) {
            return Some(GotoDefinitionResponse::Scalar(Location {
                uri: table.uri.clone(),
                range: symbol.name_range.to_lsp(),
            }));
        }

        // Try to find word at position and look it up
        let word = self.word_at_position(source, position)?;
        let symbol = table.lookup_at_position(&word, pos)?;

        Some(GotoDefinitionResponse::Scalar(Location {
            uri: table.uri.clone(),
            range: symbol.name_range.to_lsp(),
        }))
    }

    /// Get references using symbol table
    pub fn get_references_with_symbols(
        &self,
        table: &SymbolTable,
        position: Position,
    ) -> Vec<Location> {
        let pos = SourcePosition::from_lsp(position);

        table
            .find_all_references(pos)
            .into_iter()
            .map(|range| Location {
                uri: table.uri.clone(),
                range: range.to_lsp(),
            })
            .collect()
    }

    /// Get completions using symbol table
    pub fn get_completions_with_symbols(
        &self,
        table: &SymbolTable,
        position: Position,
        source: &str,
    ) -> Vec<CompletionItem> {
        let pos = SourcePosition::from_lsp(position);
        let mut items = Vec::new();

        // Check if we're completing after a dot (member access)
        if let Some(member_completions) = self.get_member_completions(table, position, source) {
            return member_completions;
        }

        // Get visible symbols at this position
        for symbol in table.visible_symbols(pos) {
            items.push(self.symbol_to_completion_item(symbol));
        }

        // Add keywords
        items.extend(self.get_keyword_completions());

        items
    }

    /// Get member completions (e.g., after typing "txtName.")
    fn get_member_completions(
        &self,
        table: &SymbolTable,
        position: Position,
        source: &str,
    ) -> Option<Vec<CompletionItem>> {
        use tower_lsp::lsp_types::CompletionItemKind;

        // Get the line up to cursor position
        let line_idx = position.line as usize;
        let char_idx = position.character as usize;

        let lines: Vec<&str> = source.lines().collect();
        if line_idx >= lines.len() {
            return None;
        }

        let line = lines[line_idx];
        if char_idx > line.len() {
            return None;
        }

        let before_cursor = &line[..char_idx];

        // Check if we just typed a dot
        if !before_cursor.ends_with('.') {
            return None;
        }

        // Get the identifier before the dot
        let before_dot = before_cursor.trim_end_matches('.');
        let last_word = before_dot
            .split(|c: char| !c.is_alphanumeric() && c != '_')
            .last()?;

        if last_word.is_empty() {
            return None;
        }

        // Look up the symbol
        let pos = SourcePosition::from_lsp(position);
        let symbol = table.lookup_at_position(last_word, pos)?;

        // Check if it's a form control
        if symbol.kind == SymbolKind::FormControl {
            let type_name = symbol.type_info.as_ref()?.name.clone();
            let control = crate::controls::get_control(&type_name)?;

            let mut completions = Vec::new();

            // Add properties
            for prop in control.properties {
                let mut item = CompletionItem {
                    label: prop.name.to_string(),
                    kind: Some(CompletionItemKind::PROPERTY),
                    detail: Some(prop.description.to_string()),
                    documentation: Some(Documentation::String(format!(
                        "**Type:** {}\n\n{}\n\n**Default:** {}",
                        prop.property_type.vb6_type(),
                        prop.description,
                        prop.default_value.unwrap_or("(none)")
                    ))),
                    insert_text: None,
                    insert_text_format: None,
                    ..Default::default()
                };

                // For enum properties, show valid values
                if !prop.valid_values.is_empty() {
                    let mut doc = format!(
                        "**Type:** {}\n\n{}\n\n**Valid Values:**\n",
                        prop.property_type.vb6_type(),
                        prop.description
                    );
                    for value in prop.valid_values.iter().take(10) {
                        doc.push_str(&format!("\n- `{}` ({}): {}", value.value, value.name, value.description));
                    }
                    if prop.valid_values.len() > 10 {
                        doc.push_str(&format!("\n- ... and {} more values", prop.valid_values.len() - 10));
                    }
                    item.documentation = Some(Documentation::String(doc));
                }

                completions.push(item);
            }

            // Add methods
            for method in control.methods {
                completions.push(CompletionItem {
                    label: method.name.to_string(),
                    kind: Some(CompletionItemKind::METHOD),
                    detail: Some(method.description.to_string()),
                    documentation: Some(Documentation::String(format!(
                        "{}\n\n**Signature:** `{}`",
                        method.description,
                        method.signature
                    ))),
                    insert_text: Some(format!("{}($1)", method.name)),
                    insert_text_format: Some(InsertTextFormat::SNIPPET),
                    ..Default::default()
                });
            }

            return Some(completions);
        }

        None
    }

    /// Get document symbols using symbol table
    pub fn get_document_symbols_with_symbols(&self, table: &SymbolTable) -> Vec<DocumentSymbol> {
        let mut symbols = Vec::new();

        for symbol in table.module_symbols() {
            // Skip form controls from document outline - they're for go-to-definition only
            if symbol.kind == SymbolKind::FormControl {
                continue;
            }

            #[allow(deprecated)]
            symbols.push(DocumentSymbol {
                name: symbol.name.clone(),
                detail: symbol.type_info.as_ref().map(|t| t.display()),
                kind: symbol.kind.to_lsp(),
                range: symbol.definition_range.to_lsp(),
                selection_range: symbol.name_range.to_lsp(),
                children: self.get_child_symbols(table, symbol),
                tags: None,
                deprecated: None,
            });
        }

        symbols
    }

    // ==========================================
    // Helper methods
    // ==========================================

    fn get_child_symbols(&self, table: &SymbolTable, parent: &Symbol) -> Option<Vec<DocumentSymbol>> {
        if parent.members.is_empty() {
            return None;
        }

        let children: Vec<DocumentSymbol> = parent
            .members
            .iter()
            .filter_map(|&id| table.get_symbol(id))
            .map(|symbol| {
                #[allow(deprecated)]
                DocumentSymbol {
                    name: symbol.name.clone(),
                    detail: symbol.type_info.as_ref().map(|t| t.display()),
                    kind: symbol.kind.to_lsp(),
                    range: symbol.definition_range.to_lsp(),
                    selection_range: symbol.name_range.to_lsp(),
                    children: None,
                    tags: None,
                    deprecated: None,
                }
            })
            .collect();

        if children.is_empty() {
            None
        } else {
            Some(children)
        }
    }

    fn symbol_to_completion_item(&self, symbol: &Symbol) -> CompletionItem {
        let detail = symbol.type_info.as_ref().map(|t| t.display());

        CompletionItem {
            label: symbol.name.clone(),
            kind: Some(symbol.kind.to_completion_kind()),
            detail,
            documentation: symbol
                .documentation
                .as_ref()
                .map(|d| Documentation::String(d.clone())),
            insert_text: if symbol.kind.is_callable() {
                Some(format!("{}($1)", symbol.name))
            } else {
                None
            },
            insert_text_format: if symbol.kind.is_callable() {
                Some(InsertTextFormat::SNIPPET)
            } else {
                None
            },
            ..Default::default()
        }
    }

    fn get_keyword_completions(&self) -> Vec<CompletionItem> {
        let keywords = [
            "If",
            "Then",
            "Else",
            "ElseIf",
            "End If",
            "For",
            "Next",
            "Do",
            "Loop",
            "While",
            "Wend",
            "Select Case",
            "Case",
            "End Select",
            "With",
            "End With",
            "Sub",
            "End Sub",
            "Function",
            "End Function",
            "Dim",
            "Private",
            "Public",
            "As",
            "Integer",
            "Long",
            "String",
            "Boolean",
            "Variant",
            "Object",
            "Nothing",
            "True",
            "False",
            "And",
            "Or",
            "Not",
            "Exit",
            "GoTo",
            "On Error",
            "Resume",
            "Set",
            "Let",
            "Call",
            "ReDim",
            "Type",
            "End Type",
            "Enum",
            "End Enum",
        ];

        keywords
            .iter()
            .map(|&kw| CompletionItem {
                label: kw.to_string(),
                kind: Some(CompletionItemKind::KEYWORD),
                ..Default::default()
            })
            .collect()
    }

    /// Extract word at position from source
    fn word_at_position(&self, source: &str, position: Position) -> Option<String> {
        let lines: Vec<&str> = source.lines().collect();
        let line = lines.get(position.line as usize)?;
        let col = position.character as usize;

        if col > line.len() {
            return None;
        }

        // Find word boundaries
        let chars: Vec<char> = line.chars().collect();

        let mut start = col;
        while start > 0 && is_identifier_char(chars[start - 1]) {
            start -= 1;
        }

        let mut end = col;
        while end < chars.len() && is_identifier_char(chars[end]) {
            end += 1;
        }

        if start == end {
            None
        } else {
            Some(chars[start..end].iter().collect())
        }
    }
}

impl Default for Analyzer {
    fn default() -> Self {
        Self::new()
    }
}

/// Check if a character is valid in a VB6 identifier
fn is_identifier_char(c: char) -> bool {
    c.is_alphanumeric() || c == '_'
}
