//! LSP Server Implementation
//!
//! Implements the Language Server Protocol handlers for VB6.

mod capabilities;
mod handlers;

use std::sync::{Arc, RwLock};

use dashmap::DashMap;
use ropey::Rope;
use tower_lsp::jsonrpc::Result;
use tower_lsp::lsp_types::*;
use tower_lsp::{Client, LanguageServer};

use crate::workspace::WorkspaceManager;

/// Document information stored in memory
pub struct Document {
    /// The document content as a rope (efficient for edits)
    pub content: Rope,
    /// The document version
    pub version: i32,
}

impl std::fmt::Debug for Document {
    fn fmt(&self, f: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        f.debug_struct("Document")
            .field("content", &self.content)
            .field("version", &self.version)
            .finish()
    }
}

/// The VB6 Language Server
pub struct Vb6LanguageServer {
    /// LSP client for sending notifications
    client: Client,
    /// Open documents
    documents: DashMap<Url, Document>,
    /// Workspace manager for multi-project support (VBP discovery)
    workspace: Arc<RwLock<WorkspaceManager>>,
    /// VB6 analysis engine. Holds the project's parsed/bound state
    /// and answers navigation/diagnostics queries.
    engine: Arc<RwLock<vb6_engine::session::Session>>,
}

impl Vb6LanguageServer {
    pub fn new(client: Client) -> Self {
        Self {
            client,
            documents: DashMap::new(),
            workspace: Arc::new(RwLock::new(WorkspaceManager::new())),
            engine: Arc::new(RwLock::new(vb6_engine::session::Session::from_sources(Vec::new()))),
        }
    }

    /// Update the engine with the document's current content and publish diagnostics.
    async fn parse_and_diagnose(&self, uri: &Url) {
        let (content, version) = match self.documents.get(uri) {
            Some(doc) => (doc.content.to_string(), doc.version),
            None => return,
        };
        let key = crate::engine_glue::doc_key(uri);

        {
            let mut engine = self.engine.write().unwrap();
            engine.update_file(&key, crate::engine_glue::to_cp1252(&content));
        }
        let diagnostics = {
            let engine = self.engine.read().unwrap();
            crate::engine_glue::diagnostics_for(&engine, &key)
        };

        self.client
            .publish_diagnostics(uri.clone(), diagnostics, Some(version))
            .await;
    }
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
            format!(
                "🧩 decoded control bag · {} propert{}",
                properties.len(),
                if properties.len() == 1 { "y" } else { "ies" }
            )
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

        if let Some(workspace_folders) = params.workspace_folders {
            let mut workspace = self.workspace.write().unwrap();
            for folder in workspace_folders {
                if let Ok(path) = folder.uri.to_file_path() {
                    let discovered = workspace.add_root(path);
                    tracing::info!("Discovered {} VBP projects in {}", discovered.len(), folder.uri);
                }
            }
        } else if let Some(root_uri) = params.root_uri {
            if let Ok(path) = root_uri.to_file_path() {
                let mut workspace = self.workspace.write().unwrap();
                let discovered = workspace.add_root(path);
                let stats = workspace.stats();
                tracing::info!(
                    "Workspace: {} roots, {} projects, {} source files ({} VBP files discovered)",
                    stats.root_count,
                    stats.project_count,
                    stats.total_source_files,
                    discovered.len()
                );
                for project in workspace.projects() {
                    tracing::debug!(
                        "  Project '{}' at {}",
                        project.name(),
                        project.vbp_path().display()
                    );
                }
            }
        }

        Ok(InitializeResult {
            capabilities: ServerCapabilities {
                text_document_sync: Some(TextDocumentSyncCapability::Kind(
                    TextDocumentSyncKind::INCREMENTAL,
                )),

                hover_provider: Some(HoverProviderCapability::Simple(true)),

                signature_help_provider: Some(SignatureHelpOptions {
                    trigger_characters: Some(vec!["(".to_string(), ",".to_string()]),
                    retrigger_characters: Some(vec![",".to_string()]),
                    work_done_progress_options: Default::default(),
                }),

                definition_provider: Some(OneOf::Left(true)),

                references_provider: Some(OneOf::Left(true)),

                document_highlight_provider: Some(OneOf::Left(true)),

                document_symbol_provider: Some(OneOf::Left(true)),

                workspace_symbol_provider: Some(OneOf::Left(true)),

                completion_provider: Some(CompletionOptions {
                    trigger_characters: Some(vec![
                        ".".to_string(),
                        " ".to_string(),
                    ]),
                    resolve_provider: Some(false),
                    ..Default::default()
                }),

                code_action_provider: Some(CodeActionProviderCapability::Simple(true)),

                document_formatting_provider: Some(OneOf::Left(true)),

                folding_range_provider: Some(FoldingRangeProviderCapability::Simple(true)),

                rename_provider: Some(OneOf::Left(true)),

                call_hierarchy_provider: Some(CallHierarchyServerCapability::Simple(true)),

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
                            range: None,
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

    async fn did_open(&self, params: DidOpenTextDocumentParams) {
        let uri = params.text_document.uri;
        let content = params.text_document.text;
        let version = params.text_document.version;

        tracing::debug!("Document opened: {}", uri);

        self.documents.insert(
            uri.clone(),
            Document {
                content: Rope::from_str(&content),
                version,
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
                    let start_line = range.start.line as usize;
                    let start_char = range.start.character as usize;
                    let end_line = range.end.line as usize;
                    let end_char = range.end.character as usize;

                    let start_idx = doc.content.line_to_char(start_line) + start_char;
                    let end_idx = doc.content.line_to_char(end_line) + end_char;

                    doc.content.remove(start_idx..end_idx);
                    doc.content.insert(start_idx, &change.text);
                } else {
                    doc.content = Rope::from_str(&change.text);
                }
            }
        }

        self.parse_and_diagnose(&uri).await;
    }

    async fn did_close(&self, params: DidCloseTextDocumentParams) {
        let uri = params.text_document.uri;
        tracing::debug!("Document closed: {}", uri);

        {
            let mut engine = self.engine.write().unwrap();
            engine.remove_file(&crate::engine_glue::doc_key(&uri));
        }

        self.documents.remove(&uri);
    }

    async fn did_save(&self, params: DidSaveTextDocumentParams) {
        let uri = params.text_document.uri;
        tracing::debug!("Document saved: {}", uri);
        self.parse_and_diagnose(&uri).await;
    }

    // Hover
    async fn hover(&self, params: HoverParams) -> Result<Option<Hover>> {
        let uri = &params.text_document_position_params.text_document.uri;
        let position = params.text_document_position_params.position;

        if let Some(doc) = self.documents.get(uri) {
            let content = doc.content.to_string();
            if let Some(h) = frx_reference_hover(uri, &content, position) {
                return Ok(Some(h));
            }
        }

        let key = crate::engine_glue::doc_key(uri);
        let g = self.engine.read().unwrap();
        let Some(m) = g.module_of(&key) else { return Ok(None) };
        let Some(off) = crate::engine_glue::offset_at(&g, m, position) else { return Ok(None) };
        Ok(g.hover(m, off).map(|h| crate::engine_glue::hover(&g, m, h)))
    }

    // Go to definition
    async fn goto_definition(
        &self,
        params: GotoDefinitionParams,
    ) -> Result<Option<GotoDefinitionResponse>> {
        let uri = &params.text_document_position_params.text_document.uri;
        let position = params.text_document_position_params.position;

        let key = crate::engine_glue::doc_key(uri);
        let g = self.engine.read().unwrap();
        let Some(m) = g.module_of(&key) else { return Ok(None) };
        let Some(off) = crate::engine_glue::offset_at(&g, m, position) else { return Ok(None) };
        Ok(g.definition(m, off)
            .and_then(|loc| crate::engine_glue::location(&g, loc))
            .map(GotoDefinitionResponse::Scalar))
    }

    // Find references
    async fn references(&self, params: ReferenceParams) -> Result<Option<Vec<Location>>> {
        let uri = &params.text_document_position.text_document.uri;
        let position = params.text_document_position.position;
        let include_decl = params.context.include_declaration;

        let key = crate::engine_glue::doc_key(uri);
        let g = self.engine.read().unwrap();
        let Some(m) = g.module_of(&key) else { return Ok(None) };
        let Some(off) = crate::engine_glue::offset_at(&g, m, position) else { return Ok(None) };
        let locations = g
            .references(m, off, include_decl)
            .into_iter()
            .filter_map(|l| crate::engine_glue::location(&g, l))
            .collect();
        Ok(Some(locations))
    }

    // Document symbols
    async fn document_symbol(
        &self,
        params: DocumentSymbolParams,
    ) -> Result<Option<DocumentSymbolResponse>> {
        let uri = &params.text_document.uri;
        let key = crate::engine_glue::doc_key(uri);
        let g = self.engine.read().unwrap();
        let Some(m) = g.module_of(&key) else { return Ok(None) };
        Ok(Some(DocumentSymbolResponse::Nested(
            crate::engine_glue::document_symbols(&g, m),
        )))
    }

    // Workspace symbols
    async fn symbol(
        &self,
        params: WorkspaceSymbolParams,
    ) -> Result<Option<Vec<SymbolInformation>>> {
        let g = self.engine.read().unwrap();
        Ok(Some(crate::engine_glue::workspace_symbols(&g, &params.query)))
    }

    // Semantic tokens (syntax highlighting)
    async fn semantic_tokens_full(
        &self,
        params: SemanticTokensParams,
    ) -> Result<Option<SemanticTokensResult>> {
        let uri = &params.text_document.uri;
        let key = crate::engine_glue::doc_key(uri);
        let g = self.engine.read().unwrap();
        let Some(m) = g.module_of(&key) else { return Ok(None) };
        Ok(Some(SemanticTokensResult::Tokens(
            crate::engine_glue::semantic_tokens(&g, m),
        )))
    }

    // Code actions (quick-fixes and refactors, computed by the engine)
    async fn code_action(&self, params: CodeActionParams) -> Result<Option<CodeActionResponse>> {
        let uri = &params.text_document.uri;
        let range = params.range;

        let key = crate::engine_glue::doc_key(uri);
        let g = self.engine.read().unwrap();
        let Some(m) = g.module_of(&key) else { return Ok(None) };
        let actions = crate::engine_glue::code_actions(&g, m, range);
        if actions.is_empty() {
            return Ok(None);
        }
        Ok(Some(actions))
    }

    // Formatting (engine-driven: indentation, trailing-whitespace, keyword case)
    async fn formatting(&self, params: DocumentFormattingParams) -> Result<Option<Vec<TextEdit>>> {
        let uri = &params.text_document.uri;

        let key = crate::engine_glue::doc_key(uri);
        let g = self.engine.read().unwrap();
        let Some(m) = g.module_of(&key) else { return Ok(None) };
        let edits = crate::engine_glue::formatting(&g, m);
        if edits.is_empty() {
            return Ok(None);
        }
        Ok(Some(edits))
    }

    // Rename
    async fn rename(&self, params: RenameParams) -> Result<Option<WorkspaceEdit>> {
        let uri = &params.text_document_position.text_document.uri;
        let position = params.text_document_position.position;
        let new_name = params.new_name;

        let key = crate::engine_glue::doc_key(uri);
        let g = self.engine.read().unwrap();
        let Some(m) = g.module_of(&key) else { return Ok(None) };
        let Some(off) = crate::engine_glue::offset_at(&g, m, position) else { return Ok(None) };
        let edits = g.rename(m, off, &new_name);
        if edits.is_empty() {
            return Ok(None);
        }
        Ok(Some(crate::engine_glue::workspace_edit(&g, edits)))
    }

    // Signature help (shown when typing inside a call's argument list)
    async fn signature_help(&self, params: SignatureHelpParams) -> Result<Option<SignatureHelp>> {
        let uri = &params.text_document_position_params.text_document.uri;
        let position = params.text_document_position_params.position;

        let key = crate::engine_glue::doc_key(uri);
        let g = self.engine.read().unwrap();
        let Some(m) = g.module_of(&key) else { return Ok(None) };
        let Some(off) = crate::engine_glue::offset_at(&g, m, position) else { return Ok(None) };
        Ok(crate::engine_glue::sig_help(&g, m, off))
    }

    // Document highlights (all occurrences of the symbol under the cursor in this file)
    async fn document_highlight(
        &self,
        params: DocumentHighlightParams,
    ) -> Result<Option<Vec<DocumentHighlight>>> {
        let uri = &params.text_document_position_params.text_document.uri;
        let position = params.text_document_position_params.position;

        let key = crate::engine_glue::doc_key(uri);
        let g = self.engine.read().unwrap();
        let Some(m) = g.module_of(&key) else { return Ok(None) };
        let Some(off) = crate::engine_glue::offset_at(&g, m, position) else { return Ok(None) };
        let spans = g.document_highlights(m, off);
        if spans.is_empty() {
            return Ok(None);
        }
        Ok(Some(crate::engine_glue::document_highlights(&g, m, spans)))
    }

    // Folding ranges (Sub/Function/If/For/…/End blocks)
    async fn folding_range(
        &self,
        params: FoldingRangeParams,
    ) -> Result<Option<Vec<FoldingRange>>> {
        let uri = &params.text_document.uri;
        let key = crate::engine_glue::doc_key(uri);
        let g = self.engine.read().unwrap();
        let Some(m) = g.module_of(&key) else { return Ok(None) };
        let ranges = crate::engine_glue::folding_ranges(&g, m);
        if ranges.is_empty() {
            Ok(None)
        } else {
            Ok(Some(ranges))
        }
    }

    // Completion (identifier suggestions at the cursor position)
    async fn completion(&self, params: CompletionParams) -> Result<Option<CompletionResponse>> {
        let uri = &params.text_document_position.text_document.uri;
        let position = params.text_document_position.position;

        let key = crate::engine_glue::doc_key(uri);
        let g = self.engine.read().unwrap();
        let Some(m) = g.module_of(&key) else { return Ok(None) };
        let Some(off) = crate::engine_glue::offset_at(&g, m, position) else { return Ok(None) };
        let items = crate::engine_glue::completion_items(&g, m, off);
        Ok(Some(CompletionResponse::Array(items)))
    }

    // Call hierarchy — resolve the procedure under the cursor
    async fn prepare_call_hierarchy(
        &self,
        params: CallHierarchyPrepareParams,
    ) -> Result<Option<Vec<CallHierarchyItem>>> {
        let uri = &params.text_document_position_params.text_document.uri;
        let position = params.text_document_position_params.position;

        let key = crate::engine_glue::doc_key(uri);
        let g = self.engine.read().unwrap();
        let Some(m) = g.module_of(&key) else { return Ok(None) };
        let Some(off) = crate::engine_glue::offset_at(&g, m, position) else { return Ok(None) };
        let items = crate::engine_glue::prepare_call_hierarchy(&g, m, off);
        if items.is_empty() {
            Ok(None)
        } else {
            Ok(Some(items))
        }
    }

    // Call hierarchy — who calls the given procedure?
    async fn incoming_calls(
        &self,
        params: CallHierarchyIncomingCallsParams,
    ) -> Result<Option<Vec<CallHierarchyIncomingCall>>> {
        let g = self.engine.read().unwrap();
        let calls = crate::engine_glue::incoming_calls(&g, &params.item);
        Ok(Some(calls))
    }

    // Call hierarchy — what does the given procedure call?
    async fn outgoing_calls(
        &self,
        params: CallHierarchyOutgoingCallsParams,
    ) -> Result<Option<Vec<CallHierarchyOutgoingCall>>> {
        let g = self.engine.read().unwrap();
        let calls = crate::engine_glue::outgoing_calls(&g, &params.item);
        Ok(Some(calls))
    }
}
