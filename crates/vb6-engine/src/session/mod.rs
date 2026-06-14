//! Host-facing session/driver layer.
//!
//! Ties the frontend (scan → parse) and `sema` (bind) together and exposes
//! LSP-shaped queries over a project of VB6 source files. It is the single
//! public boundary an editor shell (e.g. the existing VB6_lsp tower-lsp server)
//! calls into; nothing here speaks the Language Server Protocol itself.
//!
//! Queries take a `(module, byte_offset)` cursor and return byte-offset
//! [`Span`]s / strings; the shell converts to LSP line/character positions with
//! [`LineIndex`].
//!
//! On construction a session:
//!   1. scans + parses + binds every file (single-module),
//!   2. builds a project-wide [`ModuleIndex`] of public declarations,
//!   3. upgrades each module's `Unresolved` cross-module references to
//!      [`External`](crate::sema::NameResolution::External),
//!   4. builds a reverse [`ReferenceIndex`] (declaration → use sites),
//!   5. emits project-scoped diagnostics (undeclared variable under
//!      `Option Explicit`, now that cross-module names are resolved).

mod actions;
mod call_hierarchy;
mod completion;
mod folding;
mod format;
mod signature;
pub mod forms;
pub mod hover;
pub mod line_index;
pub mod module_index;
pub mod reverse_index;
pub mod tokens;

pub use call_hierarchy::{CallHierarchyDecl, IncomingCall, OutgoingCall};
pub use completion::{CompletionEntry, CompletionKind};
pub use folding::FoldRange;
pub use forms::{describe_frx_reference, form_controls, FormControl};
pub use line_index::{LineIndex, Position, Range};
pub use module_index::{ExternalRef, ModuleIndex};
pub use reverse_index::{DeclId, ReferenceIndex, RefSite};
pub use signature::SigHelp;
pub use tokens::{SemToken, SemTokenKind};

use std::collections::{HashMap, HashSet};

use crate::frm::parse_frm;
use crate::frontend::ast::{ExprArena, ExprNode, NodeId, NodeSpans, ProcKind, Span};
use crate::frontend::diagnostics::Diagnostic;
use crate::frontend::parser::Parser;
use crate::frontend::scanner::{Scanner, ScannerContext};
use crate::frontend::token::{Kw, Token, TokenKind};
use crate::sema::binder::{ERR_SUB_OR_FUNCTION_NOT_DEFINED, ERR_VARIABLE_NOT_DEFINED};
use crate::sema::builtins::is_builtin;
use crate::sema::{bind, BoundEnumDecl, BoundModule, BoundProc, NameResolution};

use reverse_index::decl_id_of;

// ── Public query result types ──────────────────────────────────────────────────

/// A source location: a byte [`Span`] within a given module.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct Location {
    pub module: usize,
    pub span: Span,
}

/// The kind of a symbol, for document/workspace symbol queries.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum SymbolKind {
    Sub,
    Function,
    PropertyGet,
    PropertyLet,
    PropertySet,
    Variable,
    Constant,
    Type,
    Enum,
    EnumMember,
    Parameter,
    Local,
}

/// A named symbol with its kind and definition location.
#[derive(Debug, Clone)]
pub struct SymbolInfo {
    pub name: String,
    pub kind: SymbolKind,
    pub location: Location,
}

/// Hover information: rendered text plus the span of the symbol under the cursor.
#[derive(Debug, Clone)]
pub struct Hover {
    pub text: String,
    pub span: Span,
}

/// A single text edit: replace `span` in `module` with `new_text`. A
/// zero-length `span` is an insertion at `span.start`. Used by rename,
/// formatting, and code actions.
#[derive(Debug, Clone)]
pub struct TextEdit {
    pub module: usize,
    pub span: Span,
    pub new_text: String,
}

/// The category of a [`CodeAction`], mapped by the host to an LSP code-action
/// kind (quick-fix vs refactor).
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum CodeActionKind {
    QuickFix,
    RefactorRewrite,
}

/// An offered code action: a human-readable title plus the edits that apply it.
#[derive(Debug, Clone)]
pub struct CodeAction {
    pub title: String,
    pub kind: CodeActionKind,
    pub edits: Vec<TextEdit>,
}

// ── Per-module analyzed state ──────────────────────────────────────────────────

struct ModuleData {
    path: String,
    source: Vec<u8>,
    ctx: ScannerContext,
    arena: ExprArena,
    spans: NodeSpans,
    line_index: LineIndex,
    /// Single-module bind result, never mutated after construction. `relink`
    /// resets the working `bound` from this so project passes (cross-module
    /// resolution, undeclared diagnostics) are recomputed cleanly on each update.
    raw_bound: BoundModule,
    /// Working copy of `raw_bound` with project-level passes applied. Queries
    /// read this.
    bound: BoundModule,
    parse_diags: Vec<Diagnostic>,
    /// `NameRef` nodes used as a call target (callee of a `Call`, or a bare-name
    /// statement = implicit sub call). Excluded from the undeclared-*variable*
    /// check, since an undefined call is "Sub or Function not defined" (a
    /// different error), not "Variable not defined".
    callee_nodes: HashSet<u32>,
}

/// Cursor resolution target.
enum CursorTarget {
    /// Cursor sits on a `NameRef` use site.
    Use { node: u32, res: NameResolution },
    /// Cursor sits on a declaration's own name.
    Def(DeclId),
}

// ── Session ────────────────────────────────────────────────────────────────────

/// An analyzed project: a set of modules plus cross-module indices.
pub struct Session {
    modules: Vec<ModuleData>,
    path_to_module: HashMap<String, usize>,
    index: ModuleIndex,
    refs: ReferenceIndex,
}

impl Session {
    /// Build a session from in-memory files: `(path, raw_bytes)`.
    ///
    /// `raw_bytes` is the file's on-disk content (Windows-1252); the scanner is
    /// byte-oriented, so no transcoding is required.
    pub fn from_sources(files: Vec<(String, Vec<u8>)>) -> Self {
        let modules: Vec<ModuleData> = files
            .into_iter()
            .map(|(path, source)| analyze_module(path, source))
            .collect();

        let mut s = Session {
            modules,
            path_to_module: HashMap::new(),
            index: ModuleIndex::new(),
            refs: ReferenceIndex::new(),
        };
        s.rebuild_path_index();
        s.relink();
        s
    }

    /// Add or replace a file's content and re-analyze the project.
    ///
    /// Only the named module is re-parsed/re-bound; the project-level indexes
    /// and cross-module diagnostics are then recomputed from every module's raw
    /// snapshot. Use for editor open/change (and for didClose, by reloading the
    /// on-disk content). `source` is the file's Windows-1252 bytes.
    pub fn update_file(&mut self, path: &str, source: Vec<u8>) {
        let md = analyze_module(path.to_string(), source);
        match self.path_to_module.get(path).copied() {
            Some(i) => self.modules[i] = md,
            None => self.modules.push(md),
        }
        self.rebuild_path_index();
        self.relink();
    }

    /// Remove a file from the project and re-analyze. Returns whether it existed.
    pub fn remove_file(&mut self, path: &str) -> bool {
        let before = self.modules.len();
        self.modules.retain(|m| m.path != path);
        if self.modules.len() == before {
            return false;
        }
        self.rebuild_path_index();
        self.relink();
        true
    }

    // ── Construction passes ─────────────────────────────────────────────────────

    fn rebuild_path_index(&mut self) {
        self.path_to_module.clear();
        for (i, m) in self.modules.iter().enumerate() {
            self.path_to_module.insert(m.path.clone(), i);
        }
    }

    /// Recompute all project-level state from each module's raw snapshot:
    /// reset working bounds, rebuild the public-symbol index, re-resolve
    /// cross-module references, rebuild the reverse index, and re-emit
    /// project-scoped diagnostics. Idempotent — safe to call after any edit.
    fn relink(&mut self) {
        for m in &mut self.modules {
            m.bound = m.raw_bound.clone();
        }
        self.build_index();
        self.resolve_cross_module();
        self.build_references();
        self.check_undeclared();
    }

    fn build_index(&mut self) {
        let mut index = ModuleIndex::new();
        for (i, m) in self.modules.iter().enumerate() {
            index.add_module(i, &m.bound, |sym| {
                m.ctx.symbol(sym as usize).name.to_ascii_lowercase()
            });
        }
        self.index = index;
    }

    /// Upgrade each module's `Unresolved` `NameRef`s to `External` when the name
    /// matches a public declaration in another module.
    fn resolve_cross_module(&mut self) {
        let index = &self.index;
        for m in &mut self.modules {
            let upgrades: Vec<(u32, NameResolution)> = m
                .bound
                .resolutions
                .iter()
                .filter(|(_, r)| **r == NameResolution::Unresolved)
                .filter_map(|(&node, _)| {
                    let ExprNode::NameRef { sym, .. } = m.arena.get(NodeId(node)) else {
                        return None;
                    };
                    let name = m.ctx.symbol(*sym as usize).name.to_ascii_lowercase();
                    index
                        .lookup(&name)
                        .map(|r| (node, NameResolution::External { module: r.module, decl: r.decl }))
                })
                .collect();
            for (node, res) in upgrades {
                m.bound.resolutions.insert(node, res);
            }
        }
    }

    fn build_references(&mut self) {
        let mut refs = ReferenceIndex::new();
        for (i, m) in self.modules.iter().enumerate() {
            refs.add_module(i, &m.bound.resolutions);
        }
        self.refs = refs;
    }

    /// Emit diagnostics for names still unresolved after cross-module resolution.
    ///
    /// VB6 distinguishes two cases:
    ///   * a name in **call** position → "Sub or Function not defined" (`0x23`),
    ///     emitted regardless of `Option Explicit` (an undefined call is always
    ///     an error);
    ///   * a name in **value** position → "Variable not defined" (`0x9caf`), only
    ///     under `Option Explicit` (otherwise VB6 implicitly declares it).
    ///
    /// Builtins are excluded. **Caveat:** the builtin registry is not yet
    /// exhaustive (`sema::builtins`), so a call to an unrecognised builtin can
    /// be falsely flagged; this closes when the registry is completed.
    fn check_undeclared(&mut self) {
        for m in &mut self.modules {
            Self::check_undeclared_module(m);
        }
    }

    /// Flag unresolved names in one module: calls become "Sub or Function not
    /// defined"; other names become "Variable not defined" under Option Explicit.
    fn check_undeclared_module(m: &mut ModuleData) {
        let option_explicit = m.bound.option_explicit;
        let mut undefined_calls: Vec<Span> = Vec::new();
        let mut undefined_vars: Vec<Span> = Vec::new();

        for (&node, res) in m.bound.resolutions.iter() {
            let Some(span) = Self::undeclared_name_span(m, node, res) else {
                continue;
            };
            if m.callee_nodes.contains(&node) {
                undefined_calls.push(span);
            } else if option_explicit {
                undefined_vars.push(span);
            }
        }

        for span in undefined_calls {
            m.bound.diagnostics.push(ERR_SUB_OR_FUNCTION_NOT_DEFINED as u32, span);
        }
        for span in undefined_vars {
            m.bound.diagnostics.push(ERR_VARIABLE_NOT_DEFINED as u32, span);
        }
    }

    /// The span of an unresolved, non-builtin `NameRef` at `node`, or `None` if it
    /// is resolved, not a name reference, or a recognised builtin.
    fn undeclared_name_span(m: &ModuleData, node: u32, res: &NameResolution) -> Option<Span> {
        if *res != NameResolution::Unresolved {
            return None;
        }
        let ExprNode::NameRef { sym, .. } = m.arena.get(NodeId(node)) else {
            return None;
        };
        let name = m.ctx.symbol(*sym as usize).name.to_ascii_lowercase();
        if name.is_empty() || is_builtin(&name) {
            return None;
        }
        Some(m.spans.get(NodeId(node)))
    }

    // ── Project introspection ───────────────────────────────────────────────────

    pub fn module_count(&self) -> usize {
        self.modules.len()
    }

    pub fn module_of(&self, path: &str) -> Option<usize> {
        self.path_to_module.get(path).copied()
    }

    pub fn module_path(&self, module: usize) -> Option<&str> {
        self.modules.get(module).map(|m| m.path.as_str())
    }

    /// The line index for a module (byte-offset ↔ LSP position conversion).
    pub fn line_index(&self, module: usize) -> Option<&LineIndex> {
        self.modules.get(module).map(|m| &m.line_index)
    }

    // ── Queries ─────────────────────────────────────────────────────────────────

    /// Classified tokens for syntax highlighting, in source order.
    ///
    /// Keywords, string/number literals, and comments are classified from the
    /// token stream; identifiers are classified from the bound model (function /
    /// variable / parameter / type / enum member). Unresolved identifiers and
    /// member names are omitted (left for the client's base grammar).
    pub fn semantic_tokens(&self, module: usize) -> Vec<SemToken> {
        let Some(m) = self.modules.get(module) else {
            return Vec::new();
        };

        // Offset → kind for every classified identifier (uses + declarations).
        let by_offset = Self::classified_identifier_offsets(m);

        // Re-scan the source (a fresh context — we only need spans + kinds) and
        // emit one classified token at a time, in order.
        let mut ctx = ScannerContext::new(1, 1, 0x0409);
        ctx.intern_keywords();
        let mut sc = Scanner::new(&mut ctx, &m.source);
        let mut out = Vec::new();
        loop {
            let tok = sc.next_token();
            if tok.kind == TokenKind::Eof {
                break;
            }
            if let Some(kind) = Self::classify_sem_token(&tok, &by_offset, &m.source) {
                out.push(SemToken { span: tok.span, kind });
            }
        }
        out
    }

    /// Build the offset → kind map for every classified identifier (resolved
    /// name uses plus all declaration-name sites).
    fn classified_identifier_offsets(m: &ModuleData) -> HashMap<u32, SemTokenKind> {
        let mut by_offset: HashMap<u32, SemTokenKind> = HashMap::new();
        for (&node, res) in m.bound.resolutions.iter() {
            if let Some(k) = tokens::kind_of_resolution(res) {
                let sp = m.spans.get(NodeId(node));
                if sp.len > 0 {
                    by_offset.insert(sp.start, k);
                }
            }
        }
        Self::add_declaration_offsets(m, &mut by_offset);
        by_offset
    }

    /// Add the declaration-name sites (procs, params, locals, module vars, types,
    /// enums, and their members) to `by_offset`, not overwriting use-site kinds.
    fn add_declaration_offsets(m: &ModuleData, by_offset: &mut HashMap<u32, SemTokenKind>) {
        let mut put = |span: Span, k: SemTokenKind| {
            if span.len > 0 {
                by_offset.entry(span.start).or_insert(k);
            }
        };
        for p in &m.bound.procs {
            put(p.name_span, SemTokenKind::Function);
            for prm in &p.params {
                put(prm.name_span, SemTokenKind::Parameter);
            }
            for loc in &p.locals {
                put(loc.name_span, SemTokenKind::Variable);
            }
        }
        for v in &m.bound.module_vars {
            put(v.name_span, SemTokenKind::Variable);
        }
        for t in &m.bound.type_decls {
            put(t.name_span, SemTokenKind::Type);
            for mem in &t.members {
                put(mem.name_span, SemTokenKind::Variable);
            }
        }
        for e in &m.bound.enum_decls {
            put(e.name_span, SemTokenKind::Type);
            for mem in &e.members {
                put(mem.name_span, SemTokenKind::EnumMember);
            }
        }
    }

    /// Classify one scanned token for highlighting, or `None` to skip it
    /// (`Eol`/`Error`, unclassified identifiers, and punctuation).
    fn classify_sem_token(
        tok: &Token,
        by_offset: &HashMap<u32, SemTokenKind>,
        source: &[u8],
    ) -> Option<SemTokenKind> {
        match tok.kind {
            TokenKind::Eol | TokenKind::Error | TokenKind::Eof => None,
            TokenKind::StrLit => Some(SemTokenKind::String),
            TokenKind::IntLit
            | TokenKind::LongLit
            | TokenKind::SngLit
            | TokenKind::DblLit
            | TokenKind::CurLit
            | TokenKind::DateLit => Some(SemTokenKind::Number),
            TokenKind::Kw(Kw::Apos) => Some(SemTokenKind::Comment),
            TokenKind::Ident => by_offset.get(&tok.span.start).copied(),
            // A keyword token that is actually a classified identifier (e.g. a
            // bracketed name) keeps its identifier kind; otherwise word keywords
            // highlight, punctuation/operators do not.
            TokenKind::Kw(_) => by_offset.get(&tok.span.start).copied().or_else(|| {
                let alpha = matches!(source.get(tok.span.start as usize), Some(b) if b.is_ascii_alphabetic());
                alpha.then_some(SemTokenKind::Keyword)
            }),
        }
    }

    /// All diagnostics for a module: parse errors plus semantic diagnostics.
    pub fn diagnostics(&self, module: usize) -> Vec<Diagnostic> {
        let Some(m) = self.modules.get(module) else {
            return Vec::new();
        };
        let mut out = m.parse_diags.clone();
        out.extend(m.bound.diagnostics.items().iter().cloned());
        out
    }

    /// Go to the definition of the symbol at `offset`.
    pub fn definition(&self, module: usize, offset: u32) -> Option<Location> {
        match self.target_at(module, offset)? {
            CursorTarget::Use { res, .. } => self.decl_location(decl_id_of(module, &res)?),
            CursorTarget::Def(id) => self.decl_location(id),
        }
    }

    /// All references to the symbol at `offset`. With `include_decl`, the
    /// declaration's own name is included.
    pub fn references(&self, module: usize, offset: u32, include_decl: bool) -> Vec<Location> {
        let Some(id) = self.decl_at_cursor(module, offset) else {
            return Vec::new();
        };
        let mut out: Vec<Location> = self
            .refs
            .references(id)
            .into_iter()
            .filter_map(|s| {
                self.modules.get(s.module).map(|m| Location {
                    module: s.module,
                    span: m.spans.get(NodeId(s.node)),
                })
            })
            .collect();
        if include_decl {
            if let Some(loc) = self.decl_location(id) {
                if !out.contains(&loc) {
                    out.push(loc);
                }
            }
        }
        out.sort_by_key(|l| (l.module, l.span.start));
        out
    }

    /// Hover text for the symbol at `offset`.
    pub fn hover(&self, module: usize, offset: u32) -> Option<Hover> {
        match self.target_at(module, offset)? {
            CursorTarget::Use { node, res } => {
                let span = self.modules.get(module)?.spans.get(NodeId(node));
                let text = match decl_id_of(module, &res) {
                    Some(id) => self.decl_hover_text(id)?,
                    None if res == NameResolution::Builtin => {
                        let m = self.modules.get(module)?;
                        if let ExprNode::NameRef { sym, .. } = m.arena.get(NodeId(node)) {
                            format!("VBA builtin: {}", m.ctx.symbol(*sym as usize).name)
                        } else {
                            return None;
                        }
                    }
                    None => return None,
                };
                Some(Hover { text, span })
            }
            CursorTarget::Def(id) => Some(Hover {
                text: self.decl_hover_text(id)?,
                span: self.decl_location(id)?.span,
            }),
        }
    }

    /// All declared symbols in a module (outline).
    pub fn document_symbols(&self, module: usize) -> Vec<SymbolInfo> {
        let Some(m) = self.modules.get(module) else {
            return Vec::new();
        };
        let ctx = &m.ctx;
        let name = |sym: u32| ctx.symbol(sym as usize).name.clone();
        let mut out = Vec::new();

        for p in &m.bound.procs {
            out.push(SymbolInfo {
                name: name(p.sym_id),
                kind: proc_symbol_kind(p.kind),
                location: Location { module, span: p.name_span },
            });
        }
        for v in &m.bound.module_vars {
            out.push(SymbolInfo {
                name: name(v.sym_id),
                kind: if v.is_const { SymbolKind::Constant } else { SymbolKind::Variable },
                location: Location { module, span: v.name_span },
            });
        }
        for t in &m.bound.type_decls {
            out.push(SymbolInfo {
                name: name(t.sym_id),
                kind: SymbolKind::Type,
                location: Location { module, span: t.name_span },
            });
        }
        for e in &m.bound.enum_decls {
            out.push(SymbolInfo {
                name: name(e.sym_id),
                kind: SymbolKind::Enum,
                location: Location { module, span: e.name_span },
            });
            for member in &e.members {
                out.push(SymbolInfo {
                    name: name(member.sym_id),
                    kind: SymbolKind::EnumMember,
                    location: Location { module, span: member.name_span },
                });
            }
        }
        out
    }

    /// All symbols across the project whose name contains `query` (case-insensitive).
    pub fn workspace_symbols(&self, query: &str) -> Vec<SymbolInfo> {
        let q = query.to_ascii_lowercase();
        let mut out = Vec::new();
        for module in 0..self.modules.len() {
            for sym in self.document_symbols(module) {
                if q.is_empty() || sym.name.to_ascii_lowercase().contains(&q) {
                    out.push(sym);
                }
            }
        }
        out
    }

    /// All spans of the symbol at `offset` within the same module (document
    /// highlights). Includes the declaration's name span.
    pub fn document_highlights(&self, module: usize, offset: u32) -> Vec<Span> {
        let Some(id) = self.decl_at_cursor(module, offset) else {
            return Vec::new();
        };
        let m = &self.modules[module];
        let mut spans: Vec<Span> = self
            .refs
            .references(id)
            .into_iter()
            .filter(|s| s.module == module)
            .map(|s| m.spans.get(NodeId(s.node)))
            .filter(|sp| sp.len > 0)
            .collect();
        // Include the declaration if it lives in this module.
        if let Some(loc) = self.decl_location(id) {
            if loc.module == module && !spans.contains(&loc.span) {
                spans.push(loc.span);
            }
        }
        spans.sort_by_key(|s| s.start);
        spans
    }

    /// Rename the symbol at `offset` to `new_name`, returning edits across all
    /// modules (including the declaration).
    pub fn rename(&self, module: usize, offset: u32, new_name: &str) -> Vec<TextEdit> {
        self.references(module, offset, true)
            .into_iter()
            .map(|loc| TextEdit {
                module: loc.module,
                span: loc.span,
                new_text: new_name.to_string(),
            })
            .collect()
    }

    // ── Cursor / declaration resolution helpers ─────────────────────────────────

    fn target_at(&self, module: usize, offset: u32) -> Option<CursorTarget> {
        let m = self.modules.get(module)?;
        // 1) a NameRef use whose span contains the cursor
        for (&node, res) in m.bound.resolutions.iter() {
            if span_contains(m.spans.get(NodeId(node)), offset) {
                return Some(CursorTarget::Use { node, res: res.clone() });
            }
        }
        // 2) a declaration name whose span contains the cursor
        self.decl_at(module, offset).map(CursorTarget::Def)
    }

    /// The `DeclId` the cursor refers to, whether it sits on a use or a decl.
    fn decl_at_cursor(&self, module: usize, offset: u32) -> Option<DeclId> {
        match self.target_at(module, offset)? {
            CursorTarget::Use { res, .. } => decl_id_of(module, &res),
            CursorTarget::Def(id) => Some(id),
        }
    }

    /// Find a declaration whose name identifier covers `offset`.
    fn decl_at(&self, module: usize, offset: u32) -> Option<DeclId> {
        let m = self.modules.get(module)?;
        for (i, p) in m.bound.procs.iter().enumerate() {
            if let Some(id) = proc_decl_at(module, i, p, offset) {
                return Some(id);
            }
        }
        for (i, v) in m.bound.module_vars.iter().enumerate() {
            if span_contains(v.name_span, offset) {
                return Some(DeclId::ModuleVar { module, idx: i });
            }
        }
        for (i, t) in m.bound.type_decls.iter().enumerate() {
            if span_contains(t.name_span, offset) {
                return Some(DeclId::Type { module, idx: i });
            }
        }
        for (i, e) in m.bound.enum_decls.iter().enumerate() {
            if let Some(id) = enum_decl_at(module, i, e, offset) {
                return Some(id);
            }
        }
        None
    }

    fn decl_location(&self, id: DeclId) -> Option<Location> {
        let module = decl_module(id);
        let m = self.modules.get(module)?;
        let span = match id {
            DeclId::Proc { idx, .. } => m.bound.procs.get(idx)?.name_span,
            DeclId::ModuleVar { idx, .. } => m.bound.module_vars.get(idx)?.name_span,
            DeclId::Local { proc, idx, .. } => m.bound.procs.get(proc)?.locals.get(idx)?.name_span,
            DeclId::Param { proc, idx, .. } => m.bound.procs.get(proc)?.params.get(idx)?.name_span,
            DeclId::Type { idx, .. } => m.bound.type_decls.get(idx)?.name_span,
            DeclId::Enum { idx, .. } => m.bound.enum_decls.get(idx)?.name_span,
            DeclId::EnumMember { enum_idx, member_idx, .. } => {
                m.bound.enum_decls.get(enum_idx)?.members.get(member_idx)?.name_span
            }
        };
        Some(Location { module, span })
    }

    fn decl_hover_text(&self, id: DeclId) -> Option<String> {
        let module = decl_module(id);
        let m = self.modules.get(module)?;
        let ctx = &m.ctx;
        let src = m.source.as_slice();
        Some(match id {
            DeclId::Proc { idx, .. } => hover::proc_signature(ctx, src, m.bound.procs.get(idx)?),
            DeclId::ModuleVar { idx, .. } => {
                hover::var_signature(ctx, src, m.bound.module_vars.get(idx)?)
            }
            DeclId::Local { proc, idx, .. } => {
                hover::var_signature(ctx, src, m.bound.procs.get(proc)?.locals.get(idx)?)
            }
            DeclId::Param { proc, idx, .. } => {
                let p = m.bound.procs.get(proc)?.params.get(idx)?;
                hover::param_signature(ctx, src, p.name_span, &p.vba_type)
            }
            DeclId::Type { idx, .. } => hover::type_decl_signature(src, m.bound.type_decls.get(idx)?),
            DeclId::Enum { idx, .. } => hover::enum_decl_signature(src, m.bound.enum_decls.get(idx)?),
            DeclId::EnumMember { enum_idx, member_idx, .. } => {
                let e = m.bound.enum_decls.get(enum_idx)?;
                hover::enum_member_signature(src, e, e.members.get(member_idx)?)
            }
        })
    }
}

// ── Free helpers ────────────────────────────────────────────────────────────────

/// The `DeclId` for a proc, parameter, or local whose name span covers `offset`.
fn proc_decl_at(module: usize, proc_idx: usize, p: &BoundProc, offset: u32) -> Option<DeclId> {
    if span_contains(p.name_span, offset) {
        return Some(DeclId::Proc { module, idx: proc_idx });
    }
    for (pi, param) in p.params.iter().enumerate() {
        if span_contains(param.name_span, offset) {
            return Some(DeclId::Param { module, proc: proc_idx, idx: pi });
        }
    }
    for (li, local) in p.locals.iter().enumerate() {
        if span_contains(local.name_span, offset) {
            return Some(DeclId::Local { module, proc: proc_idx, idx: li });
        }
    }
    None
}

/// The `DeclId` for an enum or one of its members whose name span covers `offset`.
fn enum_decl_at(module: usize, enum_idx: usize, e: &BoundEnumDecl, offset: u32) -> Option<DeclId> {
    if span_contains(e.name_span, offset) {
        return Some(DeclId::Enum { module, idx: enum_idx });
    }
    for (mi, member) in e.members.iter().enumerate() {
        if span_contains(member.name_span, offset) {
            return Some(DeclId::EnumMember { module, enum_idx, member_idx: mi });
        }
    }
    None
}

pub(super) fn span_contains(span: Span, offset: u32) -> bool {
    span.len > 0 && offset >= span.start && offset < span.start + span.len
}

fn decl_module(id: DeclId) -> usize {
    match id {
        DeclId::Proc { module, .. }
        | DeclId::ModuleVar { module, .. }
        | DeclId::Local { module, .. }
        | DeclId::Param { module, .. }
        | DeclId::Type { module, .. }
        | DeclId::Enum { module, .. }
        | DeclId::EnumMember { module, .. } => module,
    }
}

fn proc_symbol_kind(kind: ProcKind) -> SymbolKind {
    match kind {
        ProcKind::Sub => SymbolKind::Sub,
        ProcKind::Function => SymbolKind::Function,
        ProcKind::PropGet => SymbolKind::PropertyGet,
        ProcKind::PropLet => SymbolKind::PropertyLet,
        ProcKind::PropSet => SymbolKind::PropertySet,
    }
}

/// For .frm / .cls / .ctl / .dob / .pag files, replace the designer-section
/// bytes (VERSION + Object lines + root Begin/End block) with spaces so the
/// VB parser sees only whitespace there. Newlines are preserved so that all
/// diagnostic byte offsets and line numbers remain correct.
fn strip_designer_header(path: &str, mut source: Vec<u8>) -> Vec<u8> {
    let ext = path.rsplit('.').next().unwrap_or("").to_ascii_lowercase();
    if !matches!(ext.as_str(), "frm" | "ctl" | "dob" | "pag" | "cls") {
        return source;
    }
    let src_str = String::from_utf8_lossy(&source);
    let Ok(frm) = parse_frm(&src_str) else {
        return source;
    };
    if frm.designer_lines == 0 {
        return source;
    }
    let mut lines_seen = 0usize;
    for byte in source.iter_mut() {
        if lines_seen >= frm.designer_lines {
            break;
        }
        match *byte {
            b'\n' => lines_seen += 1,
            b'\r' => {}
            _ => *byte = b' ',
        }
    }
    source
}

/// Scan + parse + bind one module into [`ModuleData`].
fn analyze_module(path: String, source: Vec<u8>) -> ModuleData {
    let source = strip_designer_header(&path, source);
    let mut ctx = ScannerContext::new(1, 1, 0x0409);
    ctx.intern_keywords();
    let mut arena = ExprArena::new();
    let mut parser = Parser::new(&mut ctx, &source);
    let _top = parser.parse_module(&mut arena);
    let spans = std::mem::take(&mut parser.node_spans);
    let visibility = std::mem::take(&mut parser.decl_public);
    let parse_diags = std::mem::take(&mut parser.diagnostics).into_items();
    drop(parser);
    let raw_bound = bind(&ctx, &arena, &_top, &spans, &visibility);
    let bound = raw_bound.clone(); // overwritten by relink once the project is known
    let line_index = LineIndex::new(&source);
    let callee_nodes = collect_callees(&arena);

    ModuleData {
        path,
        source,
        ctx,
        arena,
        spans,
        line_index,
        raw_bound,
        bound,
        parse_diags,
        callee_nodes,
    }
}

/// Collect `NameRef` nodes that occupy a call-target position.
fn collect_callees(arena: &ExprArena) -> HashSet<u32> {
    let mut callees = HashSet::new();
    for i in 0..arena.len() as u32 {
        collect_callees_of(arena, NodeId(i), &mut callees);
    }
    callees
}

/// Record the `NameRef` call targets reachable from one arena node.
fn collect_callees_of(arena: &ExprArena, node: NodeId, callees: &mut HashSet<u32>) {
    match arena.get(node) {
        ExprNode::Call { func, .. } => insert_if_nameref(arena, *func, callees),
        ExprNode::CallStmt { callee, .. } => insert_if_nameref(arena, *callee, callees),
        ExprNode::Block { stmts } => {
            for &st in stmts {
                insert_if_nameref(arena, st, callees);
            }
        }
        _ => {}
    }
}

/// Insert `node`'s id into `callees` when it is a `NameRef`.
fn insert_if_nameref(arena: &ExprArena, node: NodeId, callees: &mut HashSet<u32>) {
    if matches!(arena.get(node), ExprNode::NameRef { .. }) {
        callees.insert(node.0);
    }
}
