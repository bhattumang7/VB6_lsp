//! LSP Server Implementation
//!
//! Implements the Language Server Protocol handlers for VB6.

mod capabilities;
mod document;
mod handlers;

use std::sync::{Arc, RwLock};

use dashmap::DashMap;
use ropey::Rope;
use tower_lsp::jsonrpc::Result;
use tower_lsp::lsp_types::*;
use tower_lsp::{Client, LanguageServer};

use crate::analysis::{build_symbol_table, Analyzer, SymbolTable};
use crate::claude::ClaudeClient;
use crate::parser::Vb6Parser;
use crate::utils::Encoding;
use crate::workspace::WorkspaceManager;

/// Document information stored in memory
pub struct Document {
    /// The document content as a rope (efficient for edits)
    pub content: Rope,
    /// The document version
    pub version: i32,
    /// Detected encoding (UTF-8 or Windows-1252)
    pub encoding: Encoding,
    /// Parsed AST (if available)
    pub ast: Option<crate::parser::Vb6Ast>,
    /// Tree-sitter tree for incremental parsing
    pub tree: Option<tree_sitter::Tree>,
    /// Symbol table (if available)
    pub symbol_table: Option<SymbolTable>,
}

impl std::fmt::Debug for Document {
    fn fmt(&self, f: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        f.debug_struct("Document")
            .field("content", &self.content)
            .field("version", &self.version)
            .field("ast", &self.ast)
            .field("tree", &self.tree.as_ref().map(|_| "..."))
            .field("symbol_table", &self.symbol_table.as_ref().map(|t| format!("{} symbols", t.symbol_count())))
            .finish()
    }
}

/// The VB6 Language Server
pub struct Vb6LanguageServer {
    /// LSP client for sending notifications
    client: Client,
    /// Open documents
    documents: DashMap<Url, Document>,
    /// VB6 Parser (uses RwLock for incremental parsing support)
    parser: Arc<RwLock<Vb6Parser>>,
    /// Code analyzer
    analyzer: Arc<Analyzer>,
    /// Claude AI client (optional)
    claude: Option<Arc<ClaudeClient>>,
    /// Workspace manager for multi-project support
    workspace: Arc<RwLock<WorkspaceManager>>,
}

impl Vb6LanguageServer {
    pub fn new(client: Client) -> Self {
        // Try to create Claude client if API key is available
        let claude = std::env::var("ANTHROPIC_API_KEY")
            .ok()
            .map(|key| Arc::new(ClaudeClient::new(key)));

        if claude.is_some() {
            tracing::info!("Claude AI integration enabled");
        } else {
            tracing::info!("Claude AI integration disabled (no ANTHROPIC_API_KEY)");
        }

        Self {
            client,
            documents: DashMap::new(),
            parser: Arc::new(RwLock::new(Vb6Parser::new())),
            analyzer: Arc::new(Analyzer::new()),
            claude,
            workspace: Arc::new(RwLock::new(WorkspaceManager::new())),
        }
    }

    /// Get document content by URI
    pub fn get_document(&self, uri: &Url) -> Option<dashmap::mapref::one::Ref<'_, Url, Document>> {
        self.documents.get(uri)
    }

    /// Parse a document and update diagnostics
    async fn parse_and_diagnose(&self, uri: &Url) {
        if let Some(mut doc) = self.documents.get_mut(uri) {
            let content = doc.content.to_string();

            // Parse the document using tree-sitter
            let (parse_result, tree) = {
                let mut parser = self.parser.write().unwrap();
                let result = parser.parse(&content);
                // Get the tree for symbol table building
                let tree = parser.get_tree().cloned();
                (result, tree)
            };

            match parse_result {
                Ok(ast) => {
                    // Get any parse errors for diagnostics
                    let parse_errors = {
                        let mut parser = self.parser.write().unwrap();
                        parser.get_errors(&content)
                    };

                    // Run analysis
                    let mut diagnostics = self.analyzer.analyze(&ast);

                    // Add parse errors as diagnostics
                    for error in parse_errors {
                        diagnostics.push(Diagnostic {
                            range: error.range,
                            severity: Some(DiagnosticSeverity::ERROR),
                            message: error.message,
                            source: Some("vb6-lsp".to_string()),
                            ..Default::default()
                        });
                    }

                    doc.ast = Some(ast);

                    // Build symbol table from tree-sitter tree
                    if let Some(ref ts_tree) = tree {
                        let symbol_table = build_symbol_table(uri.clone(), &content, ts_tree);
                        tracing::debug!(
                            "Built symbol table with {} symbols, {} scopes",
                            symbol_table.symbol_count(),
                            symbol_table.scope_count()
                        );

                        // Register with workspace manager for cross-project navigation
                        if let Ok(file_path) = uri.to_file_path() {
                            let mut workspace = self.workspace.write().unwrap();
                            // Clone the symbol table for workspace (document keeps its own copy)
                            workspace.set_symbol_table(&file_path, symbol_table.clone());
                        }

                        doc.symbol_table = Some(symbol_table);
                    }

                    // Publish diagnostics
                    self.client
                        .publish_diagnostics(uri.clone(), diagnostics, Some(doc.version))
                        .await;
                }
                Err(errors) => {
                    // Convert parse errors to diagnostics
                    let diagnostics: Vec<Diagnostic> = errors
                        .into_iter()
                        .map(|e| Diagnostic {
                            range: e.range,
                            severity: Some(DiagnosticSeverity::ERROR),
                            message: e.message,
                            source: Some("vb6-lsp".to_string()),
                            ..Default::default()
                        })
                        .collect();

                    self.client
                        .publish_diagnostics(uri.clone(), diagnostics, Some(doc.version))
                        .await;
                }
            }
        }
    }

    /// Get tree-sitter tree for a document (for external use)
    #[allow(dead_code)]
    fn get_tree_for_uri(&self, _uri: &Url) -> Option<tree_sitter::Tree> {
        let parser = self.parser.read().unwrap();
        parser.get_tree().cloned()
    }

    /// Extract word at position from source
    fn get_word_at_position(&self, source: &str, position: Position) -> Option<String> {
        let lines: Vec<&str> = source.lines().collect();
        let line = lines.get(position.line as usize)?;
        let col = position.character as usize;

        if col > line.len() {
            return None;
        }

        let chars: Vec<char> = line.chars().collect();

        // Find word boundaries
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

/// Check if a character is valid in a VB6 identifier
fn is_identifier_char(c: char) -> bool {
    c.is_alphanumeric() || c == '_'
}

/// Hover content for a designer resource reference in a `.frm`/`.ctl`/`.pag`/`.dob`.
///
/// When the cursor is on a line like `Icon = "Form1.frx":0000`, this reads the
/// referenced companion blob and summarises what it decodes to.
fn frx_reference_hover(uri: &Url, content: &str, position: Position) -> Option<Hover> {
    use crate::controls::frx;

    let path = uri.to_file_path().ok()?;
    let ext = path.extension()?.to_string_lossy().to_ascii_lowercase();
    if !matches!(ext.as_str(), "frm" | "ctl" | "pag" | "dob") {
        return None;
    }

    let line = content.lines().nth(position.line as usize)?;
    let (name, value) = line.split_once('=')?;
    let prop = name.trim();
    let fref = frx::parse_frx_reference(value.trim())?;

    let dir = path.parent()?;
    let companion = dir.join(&fref.file);
    let kind = frx::kind_for_property(prop, fref.dollar);

    let head = format!("**{}** → `{}`:0x{:04X}", prop, fref.file, fref.offset);
    let body = match std::fs::read(&companion) {
        Ok(bytes) => match frx::decode(&bytes, fref.offset as usize, kind) {
            Ok(value) => format_frx_value(&value),
            Err(e) => format!("⚠️ could not decode: {}", e),
        },
        Err(e) if e.kind() == std::io::ErrorKind::NotFound => {
            format!("⚠️ companion file `{}` not found next to the form", fref.file)
        }
        Err(e) => format!("⚠️ could not read companion: {}", e),
    };

    Some(Hover {
        contents: HoverContents::Markup(MarkupContent {
            kind: MarkupKind::Markdown,
            value: format!("{}\n\n{}", head, body),
        }),
        range: None,
    })
}

/// One-line summary of a decoded FRX value for hover display.
fn format_frx_value(value: &crate::controls::frx::FrxValue) -> String {
    use crate::controls::frx::FrxValue;
    match value {
        FrxValue::Picture { format, data, .. } => {
            format!("🖼 {:?} image · {} bytes", format, data.len())
        }
        FrxValue::Font(f) => {
            let mut style = String::new();
            if f.bold {
                style.push_str(" bold");
            }
            if f.italic {
                style.push_str(" italic");
            }
            if f.underline {
                style.push_str(" underline");
            }
            format!("🔤 Font `{}` · {}pt{}", f.name, f.size_pt, style)
        }
        FrxValue::Text(s) => {
            let preview: String = s.chars().take(120).collect();
            let ellipsis = if s.chars().count() > 120 { "…" } else { "" };
            format!("📝 \"{}{}\"", preview, ellipsis)
        }
        FrxValue::List { items, .. } => {
            let preview: Vec<String> = items.iter().take(5).cloned().collect();
            format!("📋 {} list item(s): {}", items.len(), preview.join(", "))
        }
        FrxValue::ItemData { items, .. } => format!("🔢 {} ItemData long(s)", items.len()),
        FrxValue::PropertyPages(p) => format!("📑 PropertyPages: {}", p.join(", ")),
        FrxValue::OcxBag { clsid, data } => {
            let id = clsid
                .as_ref()
                .map(|g| format_guid(*g))
                .unwrap_or_else(|| "unknown".to_string());
            format!("📦 proprietary control bag · {} bytes · CLSID {}", data.len(), id)
        }
        FrxValue::DecodedBag { properties, .. } => {
            format!("🧩 decoded control bag · {} propert{}", properties.len(),
                if properties.len() == 1 { "y" } else { "ies" })
        }
        FrxValue::Empty => "∅ empty resource".to_string(),
    }
}

/// Format a 16-byte COM GUID (Data1/2/3 little-endian, Data4 big-endian).
fn format_guid(g: [u8; 16]) -> String {
    let d1 = u32::from_le_bytes([g[0], g[1], g[2], g[3]]);
    let d2 = u16::from_le_bytes([g[4], g[5]]);
    let d3 = u16::from_le_bytes([g[6], g[7]]);
    format!(
        "{{{:08X}-{:04X}-{:04X}-{:02X}{:02X}-{:02X}{:02X}{:02X}{:02X}{:02X}{:02X}}}",
        d1, d2, d3, g[8], g[9], g[10], g[11], g[12], g[13], g[14], g[15]
    )
}

#[tower_lsp::async_trait]
impl LanguageServer for Vb6LanguageServer {
    async fn initialize(&self, params: InitializeParams) -> Result<InitializeResult> {
        tracing::info!("Initializing VB6 Language Server");

        // Scan workspace folders for VBP projects
        if let Some(workspace_folders) = params.workspace_folders {
            let mut workspace = self.workspace.write().unwrap();
            for folder in workspace_folders {
                if let Ok(path) = folder.uri.to_file_path() {
                    let discovered = workspace.add_root(path);
                    tracing::info!("Discovered {} VBP projects in {}", discovered.len(), folder.uri);
                }
            }
        } else if let Some(root_uri) = params.root_uri {
            // Fallback to root_uri if workspace_folders not provided
            if let Ok(path) = root_uri.to_file_path() {
                let mut workspace = self.workspace.write().unwrap();
                let discovered = workspace.add_root(path);
                tracing::info!("Discovered {} VBP projects in root", discovered.len());
            }
        }

        Ok(InitializeResult {
            capabilities: ServerCapabilities {
                // Text document sync
                text_document_sync: Some(TextDocumentSyncCapability::Kind(
                    TextDocumentSyncKind::INCREMENTAL,
                )),

                // Completion
                completion_provider: Some(CompletionOptions {
                    trigger_characters: Some(vec![".".to_string(), " ".to_string()]),
                    resolve_provider: Some(true),
                    ..Default::default()
                }),

                // Hover
                hover_provider: Some(HoverProviderCapability::Simple(true)),

                // Signature help
                signature_help_provider: Some(SignatureHelpOptions {
                    trigger_characters: Some(vec!["(".to_string(), ",".to_string()]),
                    retrigger_characters: None,
                    work_done_progress_options: Default::default(),
                }),

                // Go to definition
                definition_provider: Some(OneOf::Left(true)),

                // Find references
                references_provider: Some(OneOf::Left(true)),

                // Document symbols (outline)
                document_symbol_provider: Some(OneOf::Left(true)),

                // Workspace symbols
                workspace_symbol_provider: Some(OneOf::Left(true)),

                // Code actions (quick fixes, refactoring)
                code_action_provider: Some(CodeActionProviderCapability::Simple(true)),

                // Formatting
                document_formatting_provider: Some(OneOf::Left(true)),

                // Rename
                rename_provider: Some(OneOf::Right(RenameOptions {
                    prepare_provider: Some(true),
                    work_done_progress_options: Default::default(),
                })),

                // Semantic tokens for syntax highlighting
                semantic_tokens_provider: Some(
                    SemanticTokensServerCapabilities::SemanticTokensOptions(
                        SemanticTokensOptions {
                            legend: SemanticTokensLegend {
                                token_types: vec![
                                    SemanticTokenType::KEYWORD,
                                    SemanticTokenType::FUNCTION,
                                    SemanticTokenType::VARIABLE,
                                    SemanticTokenType::STRING,
                                    SemanticTokenType::NUMBER,
                                    SemanticTokenType::COMMENT,
                                    SemanticTokenType::TYPE,
                                    SemanticTokenType::CLASS,
                                    SemanticTokenType::PROPERTY,
                                    SemanticTokenType::PARAMETER,
                                ],
                                token_modifiers: vec![
                                    SemanticTokenModifier::DECLARATION,
                                    SemanticTokenModifier::DEFINITION,
                                    SemanticTokenModifier::READONLY,
                                ],
                            },
                            full: Some(SemanticTokensFullOptions::Bool(true)),
                            range: Some(true),
                            ..Default::default()
                        },
                    ),
                ),

                ..Default::default()
            },
            server_info: Some(ServerInfo {
                name: "vb6-lsp".to_string(),
                version: Some(env!("CARGO_PKG_VERSION").to_string()),
            }),
        })
    }

    async fn initialized(&self, _params: InitializedParams) {
        tracing::info!("VB6 Language Server initialized");
        self.client
            .log_message(MessageType::INFO, "VB6 Language Server ready!")
            .await;
    }

    async fn shutdown(&self) -> Result<()> {
        tracing::info!("Shutting down VB6 Language Server");
        Ok(())
    }

    // Document synchronization
    async fn did_open(&self, params: DidOpenTextDocumentParams) {
        let uri = params.text_document.uri;
        let content = params.text_document.text;
        let version = params.text_document.version;

        tracing::debug!("Document opened: {}", uri);

        // Documents opened via LSP are already decoded by the client as UTF-8
        self.documents.insert(
            uri.clone(),
            Document {
                content: Rope::from_str(&content),
                version,
                encoding: Encoding::Utf8, // LSP protocol uses UTF-8
                ast: None,
                tree: None,
                symbol_table: None,
            },
        );

        self.parse_and_diagnose(&uri).await;
    }

    async fn did_change(&self, params: DidChangeTextDocumentParams) {
        let uri = params.text_document.uri;

        if let Some(mut doc) = self.documents.get_mut(&uri) {
            doc.version = params.text_document.version;

            for change in params.content_changes {
                if let Some(range) = change.range {
                    // Incremental update
                    let start_line = range.start.line as usize;
                    let start_char = range.start.character as usize;
                    let end_line = range.end.line as usize;
                    let end_char = range.end.character as usize;

                    let start_idx = doc.content.line_to_char(start_line) + start_char;
                    let end_idx = doc.content.line_to_char(end_line) + end_char;

                    doc.content.remove(start_idx..end_idx);
                    doc.content.insert(start_idx, &change.text);
                } else {
                    // Full replacement
                    doc.content = Rope::from_str(&change.text);
                }
            }
        }

        self.parse_and_diagnose(&uri).await;
    }

    async fn did_close(&self, params: DidCloseTextDocumentParams) {
        let uri = params.text_document.uri;
        tracing::debug!("Document closed: {}", uri);

        // Remove from workspace manager
        if let Ok(file_path) = uri.to_file_path() {
            let mut workspace = self.workspace.write().unwrap();
            workspace.remove_symbol_table(&file_path);
        }

        self.documents.remove(&uri);
    }

    async fn did_save(&self, params: DidSaveTextDocumentParams) {
        let uri = params.text_document.uri;
        tracing::debug!("Document saved: {}", uri);
        // Re-analyze on save
        self.parse_and_diagnose(&uri).await;
    }

    // Completion
    async fn completion(&self, params: CompletionParams) -> Result<Option<CompletionResponse>> {
        let uri = &params.text_document_position.text_document.uri;
        let position = params.text_document_position.position;

        tracing::debug!("Completion requested at {:?}", position);

        // Get completions from analyzer
        if let Some(doc) = self.documents.get(uri) {
            let content = doc.content.to_string();
            // Prefer symbol table for context-aware completions
            if let Some(ref table) = doc.symbol_table {
                let items = self.analyzer.get_completions_with_symbols(table, position, &content);
                return Ok(Some(CompletionResponse::Array(items)));
            }
            // Fall back to AST-based completions
            if let Some(ref ast) = doc.ast {
                let items = self.analyzer.get_completions(ast, position);
                return Ok(Some(CompletionResponse::Array(items)));
            }
        }

        Ok(Some(CompletionResponse::Array(vec![])))
    }

    // Hover
    async fn hover(&self, params: HoverParams) -> Result<Option<Hover>> {
        let uri = &params.text_document_position_params.text_document.uri;
        let position = params.text_document_position_params.position;

        if let Some(doc) = self.documents.get(uri) {
            // Designer resource references (.frm/.ctl): decode the companion blob.
            let content = doc.content.to_string();
            if let Some(h) = frx_reference_hover(uri, &content, position) {
                return Ok(Some(h));
            }
            // Prefer symbol table for precise hover
            if let Some(ref table) = doc.symbol_table {
                return Ok(self.analyzer.get_hover_with_symbols(table, position));
            }
            // Fall back to AST-based hover
            if let Some(ref ast) = doc.ast {
                return Ok(self.analyzer.get_hover(ast, position));
            }
        }

        Ok(None)
    }

    // Go to definition
    async fn goto_definition(
        &self,
        params: GotoDefinitionParams,
    ) -> Result<Option<GotoDefinitionResponse>> {
        let uri = &params.text_document_position_params.text_document.uri;
        let position = params.text_document_position_params.position;

        if let Some(doc) = self.documents.get(uri) {
            let content = doc.content.to_string();
            // Prefer symbol table for precise definition lookup
            if let Some(ref table) = doc.symbol_table {
                // Try local lookup first
                if let Some(result) = self.analyzer.get_definition_with_symbols(table, &content, position) {
                    return Ok(Some(result));
                }

                // If local lookup failed, try workspace-wide lookup
                let word = self.get_word_at_position(&content, position);
                if let Some(word) = word {
                    if let Ok(file_path) = uri.to_file_path() {
                        let workspace = self.workspace.read().unwrap();
                        if let Some(location) = workspace.resolve_symbol(&word, &file_path) {
                            return Ok(Some(GotoDefinitionResponse::Scalar(location)));
                        }
                    }
                }
            }
            // Fall back to AST-based definition
            if let Some(ref ast) = doc.ast {
                return Ok(self.analyzer.get_definition(ast, position, uri));
            }
        }

        Ok(None)
    }

    // Find references
    async fn references(&self, params: ReferenceParams) -> Result<Option<Vec<Location>>> {
        let uri = &params.text_document_position.text_document.uri;
        let position = params.text_document_position.position;

        if let Some(doc) = self.documents.get(uri) {
            // Prefer symbol table for precise references
            if let Some(ref table) = doc.symbol_table {
                return Ok(Some(self.analyzer.get_references_with_symbols(table, position)));
            }
            // Fall back to AST-based references
            if let Some(ref ast) = doc.ast {
                return Ok(Some(self.analyzer.get_references(ast, position, uri)));
            }
        }

        Ok(None)
    }

    // Document symbols
    async fn document_symbol(
        &self,
        params: DocumentSymbolParams,
    ) -> Result<Option<DocumentSymbolResponse>> {
        let uri = &params.text_document.uri;

        if let Some(doc) = self.documents.get(uri) {
            // Prefer symbol table for precise document symbols
            if let Some(ref table) = doc.symbol_table {
                let symbols = self.analyzer.get_document_symbols_with_symbols(table);
                return Ok(Some(DocumentSymbolResponse::Nested(symbols)));
            }
            // Fall back to AST-based symbols
            if let Some(ref ast) = doc.ast {
                let symbols = self.analyzer.get_document_symbols(ast);
                return Ok(Some(DocumentSymbolResponse::Nested(symbols)));
            }
        }

        Ok(None)
    }

    // Code actions
    async fn code_action(&self, params: CodeActionParams) -> Result<Option<CodeActionResponse>> {
        let uri = &params.text_document.uri;
        let range = params.range;

        if let Some(doc) = self.documents.get(uri) {
            if let Some(ref ast) = doc.ast {
                let actions = self.analyzer.get_code_actions(ast, range, &params.context);

                // If Claude is available, add AI-powered actions
                if let Some(ref _claude) = self.claude {
                    // TODO: Add Claude-powered code actions
                }

                return Ok(Some(actions));
            }
        }

        Ok(None)
    }

    // Formatting
    async fn formatting(&self, params: DocumentFormattingParams) -> Result<Option<Vec<TextEdit>>> {
        let uri = &params.text_document.uri;

        if let Some(doc) = self.documents.get(uri) {
            let content = doc.content.to_string();
            let parser = self.parser.read().unwrap();
            return Ok(parser.format(&content));
        }

        Ok(None)
    }

    // Rename
    async fn rename(&self, params: RenameParams) -> Result<Option<WorkspaceEdit>> {
        let uri = &params.text_document_position.text_document.uri;
        let position = params.text_document_position.position;
        let new_name = params.new_name;

        if let Some(doc) = self.documents.get(uri) {
            if let Some(ref ast) = doc.ast {
                return Ok(self.analyzer.rename(ast, position, &new_name, uri));
            }
        }

        Ok(None)
    }
}
