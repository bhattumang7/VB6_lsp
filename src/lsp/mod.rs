//! LSP Server Implementation
//!
//! Implements the Language Server Protocol handlers for VB6.

mod capabilities;
mod handlers;
mod projects;

pub use projects::{ProjectStatus, ProjectStatusNotification};

use std::path::PathBuf;
use std::sync::{Arc, Mutex, RwLock};

use dashmap::DashMap;
use ropey::Rope;
use tower_lsp::jsonrpc::Result;
use tower_lsp::lsp_types::*;
use tower_lsp::{Client, LanguageServer};


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

/// Open a `.vbp` or `.vbg` as the project; argument: its path or `file:` URI.
pub const COMMAND_OPEN_PROJECT: &str = "vb6.openProject";
/// Stop analysing the opened project.
pub const COMMAND_CLOSE_PROJECT: &str = "vb6.closeProject";
/// Describe the opened project (members, references, problems), or null.
pub const COMMAND_PROJECT_INFO: &str = "vb6.projectInfo";
/// Describe a `.vbp` or `.vbg` without opening it; argument: its path or `file:` URI.
pub const COMMAND_PROJECT_PREVIEW: &str = "vb6.projectPreview";

/// The VB6 Language Server
pub struct Vb6LanguageServer {
    /// LSP client for sending notifications
    client: Client,
    /// Open documents
    documents: DashMap<Url, Document>,
    /// Workspace manager for multi-project support (VBP discovery)
    /// Analysis of the open documents that no opened project lists.
    loose: Arc<RwLock<vb6_engine::session::Session>>,
    /// The opened project or group, one analysis per project.
    selection: Arc<RwLock<Option<vb6_project::Selection>>>,
    /// The diagnostics the editor currently has, per file.
    published: DashMap<Url, Vec<Diagnostic>>,
    /// A project to open as soon as the handshake is complete.
    initial_project: Mutex<Option<PathBuf>>,
    /// Opening or closing a project is one at a time.
    load_lock: tokio::sync::Mutex<()>,
}

impl Vb6LanguageServer {
    pub fn new(client: Client) -> Self {
        Self {
            client,
            documents: DashMap::new(),
            loose: Arc::new(RwLock::new(vb6_engine::session::Session::from_sources(Vec::new()))),
            selection: Arc::new(RwLock::new(None)),
            published: DashMap::new(),
            initial_project: Mutex::new(None),
            load_lock: tokio::sync::Mutex::new(()),
        }
    }
}

/// The text of line `line` (from 0) of `content`, without its break. Lines are counted as the
/// editor counts them: a line feed, a carriage return, or both together end a line.
fn line_text(content: &str, line: u32) -> Option<&str> {
    let physical = vb6_engine::frontend::line_reader::physical_lines(content.as_bytes()).nth(line as usize)?;
    Some(&content[physical.start..physical.end])
}

/// Hover content for a designer resource reference in a `.frm`/`.ctl`/`.pag`/`.dob`.
///
/// When the cursor is on a line like `Icon = "Form1.frx":0000`, this reads the
/// referenced companion blob and summarises what it decodes to.
fn frx_reference_hover(uri: &Url, content: &str, position: Position) -> Option<Hover> {
    use vb6_formats::frx;

    let path = uri.to_file_path().ok()?;
    let ext = path.extension()?.to_string_lossy().to_ascii_lowercase();
    if !matches!(ext.as_str(), "frm" | "ctl" | "pag" | "dob") {
        return None;
    }

    let line = line_text(content, position.line)?;
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
fn format_frx_value(value: &vb6_formats::frx::FrxValue) -> String {
    use vb6_formats::controls::export::format_guid;
    use vb6_formats::frx::FrxValue;
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
            if f.strikethrough {
                style.push_str(" strikethrough");
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

#[tower_lsp::async_trait]
impl LanguageServer for Vb6LanguageServer {
    async fn initialize(&self, params: InitializeParams) -> Result<InitializeResult> {
        tracing::info!("Initializing VB6 Language Server");

        if let Some(path) = params
            .initialization_options
            .as_ref()
            .and_then(|o| o.get("project"))
            .and_then(projects::path_argument)
        {
            *self.initial_project.lock().unwrap() = Some(path);
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

                execute_command_provider: Some(ExecuteCommandOptions {
                    commands: vec![
                        COMMAND_OPEN_PROJECT.to_string(),
                        COMMAND_CLOSE_PROJECT.to_string(),
                        COMMAND_PROJECT_INFO.to_string(),
                        COMMAND_PROJECT_PREVIEW.to_string(),
                    ],
                    work_done_progress_options: Default::default(),
                }),

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

        let initial = self.initial_project.lock().unwrap().take();
        if let Some(path) = initial {
            let _ = self.open_selection(path).await;
        }
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

        self.sync_document(&uri).await;
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

        self.sync_document(&uri).await;
    }

    async fn did_close(&self, params: DidCloseTextDocumentParams) {
        let uri = params.text_document.uri;
        tracing::debug!("Document closed: {}", uri);
        self.close_document(&uri).await;
    }

    async fn did_change_watched_files(&self, params: DidChangeWatchedFilesParams) {
        let changes = params
            .changes
            .into_iter()
            .map(|c| {
                let change = if c.typ == FileChangeType::DELETED {
                    projects::DiskChange::Deleted
                } else {
                    projects::DiskChange::Written
                };
                (c.uri, change)
            })
            .collect();
        self.apply_disk_changes(changes).await;
    }

    async fn execute_command(
        &self,
        params: ExecuteCommandParams,
    ) -> Result<Option<serde_json::Value>> {
        match params.command.as_str() {
            COMMAND_OPEN_PROJECT => {
                let Some(path) = params.arguments.first().and_then(projects::path_argument) else {
                    return Err(tower_lsp::jsonrpc::Error::invalid_params(
                        "vb6.openProject takes the path of a .vbp or .vbg file",
                    ));
                };
                match self.open_selection(path).await {
                    Ok(info) => Ok(serde_json::to_value(info).ok()),
                    Err(error) => {
                        let mut e = tower_lsp::jsonrpc::Error::invalid_params(error.to_string());
                        e.code = tower_lsp::jsonrpc::ErrorCode::ServerError(1);
                        Err(e)
                    }
                }
            }
            COMMAND_CLOSE_PROJECT => {
                self.close_selection().await;
                Ok(None)
            }
            COMMAND_PROJECT_INFO => Ok(self.project_info().and_then(|i| serde_json::to_value(i).ok())),
            COMMAND_PROJECT_PREVIEW => {
                let Some(path) = params.arguments.first().and_then(projects::path_argument) else {
                    return Err(tower_lsp::jsonrpc::Error::invalid_params(
                        "vb6.projectPreview takes the path of a .vbp or .vbg file",
                    ));
                };
                match vb6_project::Selection::preview(&path) {
                    Ok(info) => Ok(serde_json::to_value(info).ok()),
                    Err(error) => {
                        let mut e = tower_lsp::jsonrpc::Error::invalid_params(error.to_string());
                        e.code = tower_lsp::jsonrpc::ErrorCode::ServerError(1);
                        Err(e)
                    }
                }
            }
            other => Err(tower_lsp::jsonrpc::Error::invalid_params(format!("unknown command {other}"))),
        }
    }

    async fn did_save(&self, params: DidSaveTextDocumentParams) {
        let uri = params.text_document.uri;
        tracing::debug!("Document saved: {}", uri);
        self.sync_document(&uri).await;
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

        let route = self.route(uri);
        let g = route.session.read().unwrap();
        let Some(m) = g.module_of(&route.key) else { return Ok(None) };
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

        let route = self.route(uri);
        let g = route.session.read().unwrap();
        let Some(m) = g.module_of(&route.key) else { return Ok(None) };
        let Some(off) = crate::engine_glue::offset_at(&g, m, position) else { return Ok(None) };
        if let Some(loc) = g.definition(m, off) {
            return Ok(crate::engine_glue::location(&g, loc).map(GotoDefinitionResponse::Scalar));
        }
        // A name from a referenced project is declared in that project's own files.
        Ok(g.external_target_at(m, off)
            .and_then(|target| self.external_definition(&route, &target))
            .map(GotoDefinitionResponse::Scalar))
    }

    // Find references
    async fn references(&self, params: ReferenceParams) -> Result<Option<Vec<Location>>> {
        let uri = &params.text_document_position.text_document.uri;
        let position = params.text_document_position.position;
        let include_decl = params.context.include_declaration;

        let route = self.route(uri);
        let g = route.session.read().unwrap();
        let Some(m) = g.module_of(&route.key) else { return Ok(None) };
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
        let route = self.route(uri);
        let g = route.session.read().unwrap();
        let Some(m) = g.module_of(&route.key) else { return Ok(None) };
        Ok(Some(DocumentSymbolResponse::Nested(
            crate::engine_glue::document_symbols(&g, m),
        )))
    }

    // Workspace symbols
    async fn symbol(
        &self,
        params: WorkspaceSymbolParams,
    ) -> Result<Option<Vec<SymbolInformation>>> {
        let mut symbols = {
            let g = self.loose.read().unwrap();
            crate::engine_glue::workspace_symbols(&g, &params.query)
        };
        if let Some(selection) = self.selection.read().unwrap().as_ref() {
            for unit in selection.units.iter().filter(|u| u.listed) {
                let g = unit.session.read().unwrap();
                symbols.extend(crate::engine_glue::workspace_symbols(&g, &params.query));
            }
        }
        Ok(Some(symbols))
    }

    // Semantic tokens (syntax highlighting)
    async fn semantic_tokens_full(
        &self,
        params: SemanticTokensParams,
    ) -> Result<Option<SemanticTokensResult>> {
        let uri = &params.text_document.uri;
        let route = self.route(uri);
        let g = route.session.read().unwrap();
        let Some(m) = g.module_of(&route.key) else { return Ok(None) };
        Ok(Some(SemanticTokensResult::Tokens(
            crate::engine_glue::semantic_tokens(&g, m),
        )))
    }

    // Code actions (quick-fixes and refactors, computed by the engine)
    async fn code_action(&self, params: CodeActionParams) -> Result<Option<CodeActionResponse>> {
        let uri = &params.text_document.uri;
        let range = params.range;

        let route = self.route(uri);
        let g = route.session.read().unwrap();
        let Some(m) = g.module_of(&route.key) else { return Ok(None) };
        let actions = crate::engine_glue::code_actions(&g, m, range);
        if actions.is_empty() {
            return Ok(None);
        }
        Ok(Some(actions))
    }

    // Formatting (engine-driven: indentation, trailing-whitespace, keyword case)
    async fn formatting(&self, params: DocumentFormattingParams) -> Result<Option<Vec<TextEdit>>> {
        let uri = &params.text_document.uri;

        let route = self.route(uri);
        let g = route.session.read().unwrap();
        let Some(m) = g.module_of(&route.key) else { return Ok(None) };
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

        let route = self.route(uri);
        let g = route.session.read().unwrap();
        let Some(m) = g.module_of(&route.key) else { return Ok(None) };
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

        let route = self.route(uri);
        let g = route.session.read().unwrap();
        let Some(m) = g.module_of(&route.key) else { return Ok(None) };
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

        let route = self.route(uri);
        let g = route.session.read().unwrap();
        let Some(m) = g.module_of(&route.key) else { return Ok(None) };
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
        let route = self.route(uri);
        let g = route.session.read().unwrap();
        let Some(m) = g.module_of(&route.key) else { return Ok(None) };
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

        let route = self.route(uri);
        let g = route.session.read().unwrap();
        let Some(m) = g.module_of(&route.key) else { return Ok(None) };
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

        let route = self.route(uri);
        let g = route.session.read().unwrap();
        let Some(m) = g.module_of(&route.key) else { return Ok(None) };
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
        let route = self.route(&params.item.uri);
        let g = route.session.read().unwrap();
        let calls = crate::engine_glue::incoming_calls(&g, &params.item);
        Ok(Some(calls))
    }

    // Call hierarchy — what does the given procedure call?
    async fn outgoing_calls(
        &self,
        params: CallHierarchyOutgoingCallsParams,
    ) -> Result<Option<Vec<CallHierarchyOutgoingCall>>> {
        let route = self.route(&params.item.uri);
        let g = route.session.read().unwrap();
        let calls = crate::engine_glue::outgoing_calls(&g, &params.item);
        Ok(Some(calls))
    }
}

#[cfg(test)]
#[path = "tests/line_text.rs"]
mod tests;
