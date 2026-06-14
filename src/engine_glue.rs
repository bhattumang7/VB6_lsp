//! Glue between the `vb6-engine` analysis Session and LSP types.
//!
//! The engine works in Windows-1252 bytes with byte-offset spans; LSP positions
//! are UTF-16 over the document. Because VB6 source is single-byte Windows-1252
//! (every char is one byte and one BMP UTF-16 unit), the engine's `LineIndex`
//! converts byte offsets ↔ `(line, character)` consistently with the editor.

use tower_lsp::lsp_types as lsp;
use vb6_engine::session::Session;

/// Stable key for a document in the engine (the URI string).
pub fn doc_key(uri: &lsp::Url) -> String {
    uri.to_string()
}

/// Encode editor text (UTF-8) to the Windows-1252 bytes the engine scans.
pub fn to_cp1252(text: &str) -> Vec<u8> {
    let (bytes, _, _) = encoding_rs::WINDOWS_1252.encode(text);
    bytes.into_owned()
}

fn pos(p: vb6_engine::session::Position) -> lsp::Position {
    lsp::Position { line: p.line, character: p.character }
}

/// Convert an engine byte [`Span`](vb6_engine::frontend::ast::Span) within
/// `module` to an LSP range.
pub fn span_range(session: &Session, module: usize, span: vb6_engine::frontend::ast::Span) -> lsp::Range {
    match session.line_index(module) {
        Some(li) => {
            let r = li.range(span);
            lsp::Range { start: pos(r.start), end: pos(r.end) }
        }
        None => lsp::Range::default(),
    }
}

/// Convert an LSP position in `module` to an engine byte offset.
pub fn offset_at(session: &Session, module: usize, position: lsp::Position) -> Option<u32> {
    let li = session.line_index(module)?;
    Some(li.offset(vb6_engine::session::Position {
        line: position.line,
        character: position.character,
    }))
}

/// Convert an engine `Location` to an LSP `Location` (maps module → URI).
pub fn location(session: &Session, loc: vb6_engine::session::Location) -> Option<lsp::Location> {
    let key = session.module_path(loc.module)?;
    let uri = lsp::Url::parse(key).ok()?;
    Some(lsp::Location { uri, range: span_range(session, loc.module, loc.span) })
}

fn symbol_kind(k: vb6_engine::session::SymbolKind) -> lsp::SymbolKind {
    use vb6_engine::session::SymbolKind as E;
    match k {
        E::Sub | E::Function => lsp::SymbolKind::FUNCTION,
        E::PropertyGet | E::PropertyLet | E::PropertySet => lsp::SymbolKind::PROPERTY,
        E::Variable | E::Parameter | E::Local => lsp::SymbolKind::VARIABLE,
        E::Constant => lsp::SymbolKind::CONSTANT,
        E::Type => lsp::SymbolKind::STRUCT,
        E::Enum => lsp::SymbolKind::ENUM,
        E::EnumMember => lsp::SymbolKind::ENUM_MEMBER,
    }
}

/// Token-type index into the semantic-tokens legend declared in `initialize`:
/// [KEYWORD, FUNCTION, VARIABLE, STRING, NUMBER, COMMENT, TYPE, CLASS, PROPERTY, PARAMETER].
fn token_type_index(k: vb6_engine::session::SemTokenKind) -> u32 {
    use vb6_engine::session::SemTokenKind as K;
    match k {
        K::Keyword => 0,
        K::Function => 1,
        K::Variable => 2,
        K::String => 3,
        K::Number => 4,
        K::Comment => 5,
        K::Type => 6,
        K::EnumMember => 8, // PROPERTY (no enumMember in the legend)
        K::Parameter => 9,
    }
}

/// Engine hover → LSP hover (VB6 signature in a code fence).
pub fn hover(session: &Session, module: usize, h: vb6_engine::session::Hover) -> lsp::Hover {
    lsp::Hover {
        contents: lsp::HoverContents::Markup(lsp::MarkupContent {
            kind: lsp::MarkupKind::Markdown,
            value: format!("```vb\n{}\n```", h.text),
        }),
        range: Some(span_range(session, module, h.span)),
    }
}

/// Engine document symbols → LSP nested document symbols (flat, one level).
pub fn document_symbols(session: &Session, module: usize) -> Vec<lsp::DocumentSymbol> {
    session
        .document_symbols(module)
        .into_iter()
        .map(|s| {
            let range = span_range(session, s.location.module, s.location.span);
            #[allow(deprecated)]
            lsp::DocumentSymbol {
                name: s.name,
                detail: None,
                kind: symbol_kind(s.kind),
                tags: None,
                deprecated: None,
                range,
                selection_range: range,
                children: None,
            }
        })
        .collect()
}

/// Engine workspace symbols → LSP `SymbolInformation`.
pub fn workspace_symbols(session: &Session, query: &str) -> Vec<lsp::SymbolInformation> {
    session
        .workspace_symbols(query)
        .into_iter()
        .filter_map(|s| {
            let loc = location(session, s.location)?;
            #[allow(deprecated)]
            Some(lsp::SymbolInformation {
                name: s.name,
                kind: symbol_kind(s.kind),
                tags: None,
                deprecated: None,
                location: loc,
                container_name: None,
            })
        })
        .collect()
}

/// Engine rename edits → LSP `WorkspaceEdit`, grouped by document URI.
pub fn workspace_edit(session: &Session, edits: Vec<vb6_engine::session::TextEdit>) -> lsp::WorkspaceEdit {
    use std::collections::HashMap;
    let mut changes: HashMap<lsp::Url, Vec<lsp::TextEdit>> = HashMap::new();
    for e in edits {
        let Some(key) = session.module_path(e.module) else { continue };
        let Ok(uri) = lsp::Url::parse(key) else { continue };
        let range = span_range(session, e.module, e.span);
        changes.entry(uri).or_default().push(lsp::TextEdit { range, new_text: e.new_text });
    }
    lsp::WorkspaceEdit { changes: Some(changes), document_changes: None, change_annotations: None }
}

/// Engine formatting edits for `module` → LSP `TextEdit`s (single document).
pub fn formatting(session: &Session, module: usize) -> Vec<lsp::TextEdit> {
    session
        .format(module)
        .into_iter()
        .map(|e| lsp::TextEdit { range: span_range(session, e.module, e.span), new_text: e.new_text })
        .collect()
}

/// Engine code actions for the byte range in `module` → LSP `CodeAction`s,
/// each carrying a `WorkspaceEdit` keyed by the document URI.
pub fn code_actions(
    session: &Session,
    module: usize,
    range: lsp::Range,
) -> Vec<lsp::CodeActionOrCommand> {
    let Some(start) = offset_at(session, module, range.start) else {
        return Vec::new();
    };
    let end = offset_at(session, module, range.end).unwrap_or(start);
    session
        .code_actions(module, start, end)
        .into_iter()
        .map(|a| {
            let kind = match a.kind {
                vb6_engine::session::CodeActionKind::QuickFix => lsp::CodeActionKind::QUICKFIX,
                vb6_engine::session::CodeActionKind::RefactorRewrite => {
                    lsp::CodeActionKind::REFACTOR_REWRITE
                }
            };
            lsp::CodeActionOrCommand::CodeAction(lsp::CodeAction {
                title: a.title,
                kind: Some(kind),
                edit: Some(workspace_edit(session, a.edits)),
                ..Default::default()
            })
        })
        .collect()
}

/// Engine semantic tokens → LSP delta-encoded `SemanticTokens`.
pub fn semantic_tokens(session: &Session, module: usize) -> lsp::SemanticTokens {
    let mut data = Vec::new();
    let mut prev_line = 0u32;
    let mut prev_start = 0u32;
    for t in session.semantic_tokens(module) {
        let r = span_range(session, module, t.span);
        let (line, start) = (r.start.line, r.start.character);
        let (delta_line, delta_start) = if line == prev_line {
            (0, start.saturating_sub(prev_start))
        } else {
            (line - prev_line, start)
        };
        data.push(lsp::SemanticToken {
            delta_line,
            delta_start,
            length: t.span.len,
            token_type: token_type_index(t.kind),
            token_modifiers_bitset: 0,
        });
        prev_line = line;
        prev_start = start;
    }
    lsp::SemanticTokens { result_id: None, data }
}

/// Engine signature help → LSP `SignatureHelp`.
pub fn sig_help(session: &Session, module: usize, offset: u32) -> Option<lsp::SignatureHelp> {
    let sh = session.signature_help(module, offset)?;
    let params: Vec<lsp::ParameterInformation> = sh
        .params
        .iter()
        .map(|&(start, end)| lsp::ParameterInformation {
            label: lsp::ParameterLabel::LabelOffsets([start, end]),
            documentation: None,
        })
        .collect();
    let active = sh.active_param as u32;
    Some(lsp::SignatureHelp {
        signatures: vec![lsp::SignatureInformation {
            label: sh.label,
            documentation: None,
            parameters: Some(params),
            active_parameter: Some(active),
        }],
        active_signature: Some(0),
        active_parameter: Some(active),
    })
}

/// Engine document highlights → LSP `DocumentHighlight`s (all text, no read/write distinction).
pub fn document_highlights(
    session: &Session,
    module: usize,
    spans: Vec<vb6_engine::frontend::ast::Span>,
) -> Vec<lsp::DocumentHighlight> {
    spans
        .into_iter()
        .map(|sp| lsp::DocumentHighlight {
            range: span_range(session, module, sp),
            kind: None,
        })
        .collect()
}

/// Engine folding ranges → LSP `FoldingRange`s.
pub fn folding_ranges(session: &Session, module: usize) -> Vec<lsp::FoldingRange> {
    session
        .folding_ranges(module)
        .into_iter()
        .map(|r| lsp::FoldingRange {
            start_line: r.start_line,
            start_character: None,
            end_line: r.end_line,
            end_character: None,
            kind: None,
            collapsed_text: None,
        })
        .collect()
}

fn completion_item_kind(k: vb6_engine::session::CompletionKind) -> lsp::CompletionItemKind {
    use vb6_engine::session::CompletionKind as K;
    match k {
        K::Variable | K::Parameter => lsp::CompletionItemKind::VARIABLE,
        K::Constant => lsp::CompletionItemKind::CONSTANT,
        K::Function | K::Sub | K::Builtin => lsp::CompletionItemKind::FUNCTION,
        K::Property => lsp::CompletionItemKind::PROPERTY,
        K::Keyword => lsp::CompletionItemKind::KEYWORD,
        K::EnumMember => lsp::CompletionItemKind::ENUM_MEMBER,
        K::Type => lsp::CompletionItemKind::STRUCT,
        K::Enum => lsp::CompletionItemKind::ENUM,
    }
}

/// Engine completions → LSP `CompletionItem`s.
pub fn completion_items(session: &Session, module: usize, offset: u32) -> Vec<lsp::CompletionItem> {
    session
        .completions(module, offset)
        .into_iter()
        .map(|e| lsp::CompletionItem {
            label: e.name,
            kind: Some(completion_item_kind(e.kind)),
            detail: e.detail,
            ..Default::default()
        })
        .collect()
}

/// Build an LSP `CallHierarchyItem` from an engine `CallHierarchyDecl`.
fn call_hierarchy_item(
    session: &Session,
    decl: vb6_engine::session::CallHierarchyDecl,
) -> Option<lsp::CallHierarchyItem> {
    let key = session.module_path(decl.location.module)?;
    let uri = lsp::Url::parse(key).ok()?;
    let range = span_range(session, decl.location.module, decl.location.span);
    Some(lsp::CallHierarchyItem {
        name: decl.name.clone(),
        kind: lsp::SymbolKind::FUNCTION,
        tags: None,
        detail: None,
        uri,
        range,
        selection_range: range,
        data: Some(serde_json::Value::String(decl.name)),
    })
}

/// `textDocument/prepareCallHierarchy` → engine lookup → LSP items.
pub fn prepare_call_hierarchy(
    session: &Session,
    module: usize,
    offset: u32,
) -> Vec<lsp::CallHierarchyItem> {
    match session.prepare_call_hierarchy(module, offset) {
        Some(decl) => call_hierarchy_item(session, decl).into_iter().collect(),
        None => Vec::new(),
    }
}

/// `callHierarchy/incomingCalls` — extract the proc name from the item's `data`
/// field, then ask the engine for callers and convert to LSP.
pub fn incoming_calls(
    session: &Session,
    item: &lsp::CallHierarchyItem,
) -> Vec<lsp::CallHierarchyIncomingCall> {
    let name = match &item.data {
        Some(serde_json::Value::String(s)) => s.as_str(),
        _ => item.name.as_str(),
    };
    session
        .incoming_calls(name)
        .into_iter()
        .filter_map(|ic| {
            let from = call_hierarchy_item(session, ic.caller)?;
            let from_ranges = ic
                .call_sites
                .into_iter()
                .map(|loc| span_range(session, loc.module, loc.span))
                .collect();
            Some(lsp::CallHierarchyIncomingCall { from, from_ranges })
        })
        .collect()
}

/// `callHierarchy/outgoingCalls` — same round-trip as incoming, reversed.
pub fn outgoing_calls(
    session: &Session,
    item: &lsp::CallHierarchyItem,
) -> Vec<lsp::CallHierarchyOutgoingCall> {
    let name = match &item.data {
        Some(serde_json::Value::String(s)) => s.as_str(),
        _ => item.name.as_str(),
    };
    session
        .outgoing_calls(name)
        .into_iter()
        .filter_map(|oc| {
            let to = call_hierarchy_item(session, oc.callee)?;
            let from_ranges = oc
                .call_sites
                .into_iter()
                .map(|loc| span_range(session, loc.module, loc.span))
                .collect();
            Some(lsp::CallHierarchyOutgoingCall { to, from_ranges })
        })
        .collect()
}

/// All diagnostics for the document keyed by `key`, mapped to LSP.
pub fn diagnostics_for(session: &Session, key: &str) -> Vec<lsp::Diagnostic> {
    let Some(m) = session.module_of(key) else {
        return Vec::new();
    };
    session
        .diagnostics(m)
        .into_iter()
        .map(|d| lsp::Diagnostic {
            range: span_range(session, m, d.span),
            severity: Some(lsp::DiagnosticSeverity::ERROR),
            message: d.message()
                .unwrap_or_else(|| format!("VB6 error {:#06x}", d.code)),
            source: Some("vb6-lsp".to_string()),
            ..Default::default()
        })
        .collect()
}
