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
mod library;
mod project_model;
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
pub use forms::{describe_frx_reference, form_controls, form_controls_from_bytes, FormControl};
pub use library::{
    Library, LibraryMember, LibraryMemberKind, LibraryType, LibraryTypeKind, SessionConfig,
};
pub use line_index::{LineIndex, Position, Range};
pub use module_index::{ExternalRef, ModuleIndex};
pub use reverse_index::{DeclId, LibraryKey, ReferenceIndex, RefSite};
pub use signature::SigHelp;
pub use tokens::{SemToken, SemTokenKind};

use std::collections::{HashMap, HashSet};

use crate::frm::parse_frm;
use crate::frontend::ast::{ExprArena, ExprNode, NodeId, NodeSpans, ProcKind, Span};
use crate::frontend::diagnostics::Diagnostic;
use crate::frontend::parser::{ModuleKind, Parser};
use crate::frontend::scanner::{Scanner, ScannerContext};
use crate::frontend::token::{Kw, Token, TokenKind};
use crate::sema::binder::{ERR_SUB_OR_FUNCTION_NOT_DEFINED, ERR_VARIABLE_NOT_DEFINED};
use crate::sema::builtins::is_builtin;
use crate::sema::project::MemberStatus;
use crate::sema::symbol::TypeName;
use crate::sema::{
    bind, bind_in_project, unaccounted_member_accesses, BoundEnumDecl, BoundModule, BoundProc,
    NameResolution,
};

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

/// How the member accesses of a project resolved, by outcome. The sum of all
/// fields is the number of member accesses (and `ReDim` of member arrays).
#[derive(Debug, Clone, Copy, Default, PartialEq, Eq)]
pub struct MemberAccessSummary {
    pub resolved: usize,
    /// Base is `Variant`, `Object` or untyped.
    pub late_bound: usize,
    /// `base!key`.
    pub bang: usize,
    /// Base is a value type or an array.
    pub not_an_object: usize,
    /// Base's declared type is not a type of the project and not qualified.
    pub unknown_type: usize,
    /// Base's type, or the member's, belongs to a library that is not loaded.
    pub library_not_loaded: usize,
    /// Base's class is known and has no reachable member of that name.
    pub no_such_member: usize,
    /// The member exists, not as the accessor the use needs.
    pub missing_accessor: usize,
    /// Base expression itself did not resolve.
    pub base_unresolved: usize,
}

/// A single text edit: replace `span` in `module` with `new_text`. A
/// zero-length `span` is an insertion at `span.start`. Used by rename,
/// formatting, and code actions.
#[derive(Debug, Clone, PartialEq, Eq)]
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
    /// Operands of `AddressOf`: a procedure named there is passed by address,
    /// not called.
    address_of_nodes: HashSet<u32>,
    /// Top-level AST nodes and the parser's visibility table, kept so the
    /// module can be bound again once the project is known.
    top: Vec<NodeId>,
    visibility: HashMap<u32, bool>,
    /// The module's name: `Attribute VB_Name`, else the file's base name.
    module_name: String,
    /// A `.bas` standard module (a namespace), not a class or designer.
    standard: bool,
    /// The module's name is also a default instance (forms, and classes marked
    /// `VB_PredeclaredId`).
    predeclared: bool,
    /// For a designer, the library type it extends (`VB.Form`).
    base_type: Option<TypeName>,
    /// Controls of a designer, grouped (a control array is one control).
    controls: Vec<ControlDecl>,
    /// Span of the `VB_Name` attribute's value: where the module name is
    /// declared.
    vb_name_span: Option<Span>,
}

/// A control declared in a designer, with one span per element.
struct ControlDecl {
    name: String,
    control_type: String,
    is_array: bool,
    /// Name span of each declaration of the control (several for an array).
    spans: Vec<Span>,
    /// Span of the control type on each declaration's `Begin` line.
    type_spans: Vec<Span>,
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
    refs: ReferenceIndex,
    /// Uses that are not AST nodes: the control types written on the `Begin`
    /// lines of a designer, by the declaration they name.
    designer_uses: HashMap<DeclId, Vec<Location>>,
    /// The project as the binder saw it at the last relink; answers questions
    /// about library members.
    model: Option<project_model::ProjectModel>,
    config: SessionConfig,
}

impl Session {
    /// Build a session from in-memory files: `(path, raw_bytes)`.
    ///
    /// `raw_bytes` is the file's on-disk content (Windows-1252); the scanner is
    /// byte-oriented, so no transcoding is required.
    pub fn from_sources(files: Vec<(String, Vec<u8>)>) -> Self {
        Self::with_config(files, SessionConfig::default())
    }

    /// Like [`from_sources`](Self::from_sources), with what the host knows
    /// about the project beyond its files: the project's name and the external
    /// libraries it references (see [`Library`]).
    pub fn with_config(files: Vec<(String, Vec<u8>)>, config: SessionConfig) -> Self {
        let modules: Vec<ModuleData> = files
            .into_iter()
            .map(|(path, source)| analyze_module(path, source))
            .collect();

        let mut s = Session {
            modules,
            path_to_module: HashMap::new(),
            refs: ReferenceIndex::new(),
            designer_uses: HashMap::new(),
            model: None,
            config,
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

    /// Recompute all project-level state from each module's parsed form: build
    /// the project model (every module's declarations plus the external
    /// libraries), bind every module against it, rebuild the reverse index, and
    /// re-emit project-scoped diagnostics. Idempotent — safe to call after any
    /// edit.
    fn relink(&mut self) {
        let model = project_model::ProjectModel::build(
            &self.modules,
            self.config.project_name.as_deref(),
            &self.config.libraries,
        );
        for (i, m) in self.modules.iter_mut().enumerate() {
            m.bound = bind_in_project(&m.ctx, &m.arena, &m.top, &m.spans, &m.visibility, i, &model);
        }
        self.build_references();
        self.designer_uses = designer_type_uses(&self.modules, &model);
        self.model = Some(model);
        self.check_undeclared();
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

    // ── Member-access accounting ────────────────────────────────────────────────

    /// How every member access in the project resolved, by outcome.
    pub fn member_access_summary(&self) -> MemberAccessSummary {
        let mut out = MemberAccessSummary::default();
        for m in &self.modules {
            for status in m.bound.member_status.values() {
                match status {
                    MemberStatus::Resolved => out.resolved += 1,
                    MemberStatus::LateBound => out.late_bound += 1,
                    MemberStatus::Bang => out.bang += 1,
                    MemberStatus::NotAnObject => out.not_an_object += 1,
                    MemberStatus::UnknownType(_) => out.unknown_type += 1,
                    MemberStatus::LibraryNotLoaded(_) => out.library_not_loaded += 1,
                    MemberStatus::NoSuchMember => out.no_such_member += 1,
                    MemberStatus::MissingAccessor(_) => out.missing_accessor += 1,
                    MemberStatus::BaseUnresolved => out.base_unresolved += 1,
                }
            }
        }
        out
    }

    /// Every member access of `module` with its status, as the span of the
    /// member name and why it did or did not resolve, in source order.
    pub fn member_accesses(&self, module: usize) -> Vec<(Span, MemberStatus)> {
        let Some(m) = self.modules.get(module) else {
            return Vec::new();
        };
        let mut out: Vec<(Span, MemberStatus)> = m
            .bound
            .member_status
            .iter()
            .map(|(node, status)| (m.spans.get(NodeId(*node)), status.clone()))
            .collect();
        out.sort_by_key(|(span, _)| span.start);
        out
    }

    /// The status of the member access whose member name covers `offset`.
    pub fn member_status_at(&self, module: usize, offset: u32) -> Option<MemberStatus> {
        let m = self.modules.get(module)?;
        m.bound
            .member_status
            .iter()
            .find(|(node, _)| span_contains(m.spans.get(NodeId(**node)), offset))
            .map(|(_, status)| status.clone())
    }

    /// Spans of member accesses in `module` that have no status. Always empty
    /// for a well-formed module: every member access either resolves or is
    /// reported with a reason.
    pub fn unaccounted_member_accesses(&self, module: usize) -> Vec<Span> {
        let Some(m) = self.modules.get(module) else {
            return Vec::new();
        };
        unaccounted_member_accesses(&m.arena, &m.bound.member_status)
            .into_iter()
            .map(|n| m.spans.get(n))
            .collect()
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
    ///
    /// The accessors of a property (`Get` / `Let` / `Set`) are one symbol: a
    /// reference to any of them is a reference to the property. A member of an
    /// external library has no declaration in the project, so its sites are
    /// those with the same `(library, type, member)`.
    pub fn references(&self, module: usize, offset: u32, include_decl: bool) -> Vec<Location> {
        if let Some(key) = self.library_key_at(module, offset) {
            return self.site_locations(self.refs.library_references(key));
        }
        let Some(id) = self.decl_at_cursor(module, offset) else {
            return Vec::new();
        };
        let group = self.decl_group(id);
        let mut sites: Vec<RefSite> =
            group.iter().flat_map(|g| self.refs.references(*g)).collect();
        sites.sort_by_key(|s| (s.module, s.node));
        let mut out = self.site_locations(sites);
        for g in &group {
            out.extend(self.designer_uses.get(g).into_iter().flatten().copied());
        }
        if include_decl {
            for g in &group {
                if let Some(loc) = self.decl_location(*g) {
                    if !out.contains(&loc) {
                        out.push(loc);
                    }
                }
            }
        }
        out.sort_by_key(|l| (l.module, l.span.start));
        out
    }

    fn site_locations(&self, sites: Vec<RefSite>) -> Vec<Location> {
        sites
            .into_iter()
            .filter_map(|s| {
                self.modules.get(s.module).map(|m| Location {
                    module: s.module,
                    span: m.spans.get(NodeId(s.node)),
                })
            })
            .collect()
    }

    /// Every use of member `member` of type `type_name` of the external
    /// library `library`, by name (case-insensitive): the member as declared by
    /// the type or by any interface it exposes. Empty when the library, type or
    /// member is not in the session's libraries.
    pub fn external_member_references(
        &self,
        library: &str,
        type_name: &str,
        member: &str,
    ) -> Vec<Location> {
        let Some(model) = &self.model else {
            return Vec::new();
        };
        let mut sites: Vec<RefSite> = model
            .library_member_keys(library, type_name, member)
            .into_iter()
            .flat_map(|key| self.refs.library_references(key))
            .collect();
        sites.sort_by_key(|s| (s.module, s.node));
        sites.dedup();
        self.site_locations(sites)
    }

    /// The library member the cursor sits on, if it is a use of one.
    fn library_key_at(&self, module: usize, offset: u32) -> Option<LibraryKey> {
        match self.target_at(module, offset)? {
            CursorTarget::Use { res: NameResolution::LibraryMember { lib, ty, member }, .. } => {
                Some((lib, ty, member))
            }
            _ => None,
        }
    }

    /// The declarations that are the same symbol as `id`: all accessors of a
    /// property declared in a module, otherwise `id` alone.
    fn decl_group(&self, id: DeclId) -> Vec<DeclId> {
        let DeclId::Proc { module, idx } = id else {
            return vec![id];
        };
        let Some(m) = self.modules.get(module) else {
            return vec![id];
        };
        let Some(p) = m.bound.procs.get(idx) else {
            return vec![id];
        };
        let is_property = |k: ProcKind| matches!(k, ProcKind::PropGet | ProcKind::PropLet | ProcKind::PropSet);
        if !is_property(p.kind) {
            return vec![id];
        }
        let name = m.ctx.symbol(p.sym_id as usize).name.to_ascii_lowercase();
        m.bound
            .procs
            .iter()
            .enumerate()
            .filter(|(_, q)| {
                is_property(q.kind) && m.ctx.symbol(q.sym_id as usize).name.to_ascii_lowercase() == name
            })
            .map(|(i, _)| DeclId::Proc { module, idx: i })
            .collect()
    }

    /// The declaration the symbol at `offset` refers to: for a use, the
    /// declaration it resolved to; for a declaration's own name, itself.
    /// `None` for builtins, library members, unresolved names and whitespace.
    pub fn declaration_at(&self, module: usize, offset: u32) -> Option<DeclId> {
        self.decl_at_cursor(module, offset)
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
                    None => self.external_hover_text(&res)?,
                };
                Some(Hover { text, span })
            }
            CursorTarget::Def(id) => Some(Hover {
                text: self.decl_hover_text(id)?,
                span: self.decl_location(id)?.span,
            }),
        }
    }

    /// Hover text for a name that stands for the project or a library, or a
    /// member of a library.
    fn external_hover_text(&self, res: &NameResolution) -> Option<String> {
        let libs = &self.config.libraries;
        Some(match res {
            NameResolution::Project => format!("Project {}", self.config.project_name.as_deref()?),
            NameResolution::Library(l) => format!("Library {}", libs.get(*l)?.name),
            NameResolution::LibraryType { lib, ty } => {
                let l = libs.get(*lib)?;
                format!("{}.{}", l.name, l.types.get(*ty)?.name)
            }
            NameResolution::LibraryMember { lib, ty, member } => {
                let l = libs.get(*lib)?;
                let t = l.types.get(*ty)?;
                let m = t.members.get(*member)?;
                let kind = match m.kind {
                    LibraryMemberKind::Method if m.return_type.is_some() => "Function",
                    LibraryMemberKind::Method => "Sub",
                    LibraryMemberKind::PropertyGet => "Property Get",
                    LibraryMemberKind::PropertyLet => "Property Let",
                    LibraryMemberKind::PropertySet => "Property Set",
                    LibraryMemberKind::Field => "Field",
                    LibraryMemberKind::Constant => "Const",
                };
                let ret = m.return_type.as_ref().map(|r| format!(" As {r}")).unwrap_or_default();
                format!("{kind} {}.{}.{}{ret}", l.name, t.name, m.name)
            }
            _ => return None,
        })
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
        let m = &self.modules[module];
        let mut spans: Vec<Span> = Vec::new();
        if let Some(key) = self.library_key_at(module, offset) {
            spans.extend(
                self.refs
                    .library_references(key)
                    .into_iter()
                    .filter(|s| s.module == module)
                    .map(|s| m.spans.get(NodeId(s.node)))
                    .filter(|sp| sp.len > 0),
            );
        } else if let Some(id) = self.decl_at_cursor(module, offset) {
            for g in self.decl_group(id) {
                spans.extend(
                    self.designer_uses
                        .get(&g)
                        .into_iter()
                        .flatten()
                        .filter(|l| l.module == module)
                        .map(|l| l.span),
                );
                spans.extend(
                    self.refs
                        .references(g)
                        .into_iter()
                        .filter(|s| s.module == module)
                        .map(|s| m.spans.get(NodeId(s.node)))
                        .filter(|sp| sp.len > 0),
                );
                // Include the declaration if it lives in this module.
                if let Some(loc) = self.decl_location(g) {
                    if loc.module == module && loc.span.len > 0 && !spans.contains(&loc.span) {
                        spans.push(loc.span);
                    }
                }
            }
        }
        spans.sort_by_key(|s| s.start);
        spans.dedup();
        spans
    }

    /// Rename the symbol at `offset` to `new_name`, returning edits across all
    /// modules (including the declaration).
    ///
    /// Nothing is renamed for a symbol the project does not own: a member of an
    /// external library, a library or project name, or a module (its name is
    /// also its file name and its `.vbp` entry, which are outside the source).
    pub fn rename(&self, module: usize, offset: u32, new_name: &str) -> Vec<TextEdit> {
        match self.decl_at_cursor(module, offset) {
            Some(DeclId::Module { .. }) => return Vec::new(),
            Some(_) => {}
            None => {
                // Not a project declaration: library members and names have
                // nothing to rename, and an unresolved name has no sites.
                if self.library_key_at(module, offset).is_some() {
                    return Vec::new();
                }
            }
        }
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
        for (ti, t) in m.bound.type_decls.iter().enumerate() {
            for (mi, member) in t.members.iter().enumerate() {
                if span_contains(member.name_span, offset) {
                    return Some(DeclId::TypeMember { module, type_idx: ti, member_idx: mi });
                }
            }
        }
        for (ci, c) in m.controls.iter().enumerate() {
            if c.spans.iter().any(|s| span_contains(*s, offset)) {
                return Some(DeclId::Control { module, idx: ci });
            }
        }
        if m.vb_name_span.map_or(false, |s| span_contains(s, offset)) {
            return Some(DeclId::Module { module });
        }
        // A control type written on a designer's `Begin` line.
        self.designer_uses
            .iter()
            .find(|(_, uses)| uses.iter().any(|l| l.module == module && span_contains(l.span, offset)))
            .map(|(id, _)| *id)
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
            DeclId::TypeMember { type_idx, member_idx, .. } => {
                m.bound.type_decls.get(type_idx)?.members.get(member_idx)?.name_span
            }
            DeclId::Control { idx, .. } => *m.controls.get(idx)?.spans.first()?,
            // A module is declared where its name is written; a module without
            // a `VB_Name` attribute is named by its file and has no such place.
            DeclId::Module { .. } => m.vb_name_span?,
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
            DeclId::TypeMember { type_idx, member_idx, .. } => {
                let t = m.bound.type_decls.get(type_idx)?;
                let mem = t.members.get(member_idx)?;
                format!(
                    "{}.{} As {}",
                    hover::name_at(src, t.name_span),
                    hover::name_at(src, mem.name_span),
                    hover::type_str(ctx, &mem.vba_type)
                )
            }
            DeclId::Control { idx, .. } => {
                let c = m.controls.get(idx)?;
                let array = if c.is_array { "()" } else { "" };
                format!("{}{array} As {}", c.name, c.control_type)
            }
            DeclId::Module { .. } => format!("Module {}", m.module_name),
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
        | DeclId::EnumMember { module, .. }
        | DeclId::TypeMember { module, .. }
        | DeclId::Control { module, .. }
        | DeclId::Module { module } => module,
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

/// The module kind a file's extension implies. Only `.bas` is a standard
/// module; class modules and every designer (form, UserControl, PropertyPage,
/// UserDocument, designer `.dsr`) are object modules, where `Implements`,
/// `Event` and `WithEvents` are legal.
pub fn module_kind_for_path(path: &str) -> ModuleKind {
    match path.rsplit('.').next().unwrap_or("").to_ascii_lowercase().as_str() {
        "cls" => ModuleKind::Class,
        "frm" | "ctl" | "pag" | "dob" | "dsr" => ModuleKind::Form,
        _ => ModuleKind::Standard,
    }
}

/// Scan + parse + bind one module into [`ModuleData`].
fn analyze_module(path: String, raw: Vec<u8>) -> ModuleData {
    // The designer section is read from the original bytes; the code parser
    // sees it blanked.
    let designer_controls = if is_designer_path(&path) {
        form_controls_from_bytes(&raw)
    } else {
        Vec::new()
    };
    let source = strip_designer_header(&path, raw);
    let mut ctx = ScannerContext::new(1, 1, 0x0409);
    ctx.intern_keywords();
    let mut arena = ExprArena::new();
    let kind = module_kind_for_path(&path);
    let mut parser = Parser::with_module_kind(&mut ctx, &source, kind);
    let top = parser.parse_module(&mut arena);
    let spans = std::mem::take(&mut parser.node_spans);
    let visibility = std::mem::take(&mut parser.decl_public);
    let parse_diags = std::mem::take(&mut parser.diagnostics).into_items();
    drop(parser);
    let raw_bound = bind(&ctx, &arena, &top, &spans, &visibility);
    let bound = raw_bound.clone(); // overwritten by relink once the project is known
    let line_index = LineIndex::new(&source);
    let callee_nodes = collect_callees(&arena);
    let address_of_nodes: HashSet<u32> = (0..arena.len() as u32)
        .filter_map(|i| match arena.get(NodeId(i)) {
            ExprNode::AddressOf { operand } => Some(operand.0),
            _ => None,
        })
        .collect();

    let module_name = raw_bound
        .vb_name
        .clone()
        .unwrap_or_else(|| file_stem(&path).to_string());
    let base_type = designer_controls.first().and_then(|root| {
        let ty = root.control_type.to_ascii_lowercase();
        match ty.split_once('.') {
            Some((q, n)) => Some(TypeName { qualifier: Some(q.to_string()), name: n.to_string() }),
            None if ty.is_empty() => None,
            None => Some(TypeName { qualifier: None, name: ty }),
        }
    });
    let vb_name_span = (0..arena.len() as u32).find_map(|i| match arena.get(NodeId(i)) {
        ExprNode::Attribute { target: None, name, .. }
            if ctx.symbol(*name as usize).name.eq_ignore_ascii_case("vb_name") =>
        {
            Some(spans.get(NodeId(i))).filter(|s| s.len > 0)
        }
        _ => None,
    });

    ModuleData {
        path,
        source,
        ctx,
        arena,
        spans,
        line_index,
        standard: kind == ModuleKind::Standard,
        predeclared: raw_bound.predeclared_id || kind == ModuleKind::Form && base_type.as_ref().map_or(false, |t| t.name == "form" || t.name == "mdiform"),
        raw_bound,
        bound,
        parse_diags,
        callee_nodes,
        address_of_nodes,
        top,
        visibility,
        module_name,
        base_type,
        controls: group_controls(&designer_controls),
        vb_name_span,
    }
}

/// The control declarations of a designer: the root form is the module itself
/// and is skipped; elements sharing a name (a control array) are one control.
fn group_controls(flat: &[FormControl]) -> Vec<ControlDecl> {
    let mut out: Vec<ControlDecl> = Vec::new();
    for c in flat.iter().filter(|c| c.depth > 0) {
        match out.iter_mut().find(|d| d.name.eq_ignore_ascii_case(&c.name)) {
            Some(d) => {
                d.is_array |= c.index.is_some();
                d.spans.push(c.span);
                d.type_spans.push(c.type_span);
            }
            None => out.push(ControlDecl {
                name: c.name.clone(),
                control_type: c.control_type.clone(),
                is_array: c.index.is_some(),
                spans: vec![c.span],
                type_spans: vec![c.type_span],
            }),
        }
    }
    out
}

/// The uses of a declaration that are written on a designer's `Begin` lines:
/// a control typed by a class or UserControl of the project is a use of it.
fn designer_type_uses(
    modules: &[ModuleData],
    model: &project_model::ProjectModel,
) -> HashMap<DeclId, Vec<Location>> {
    use crate::sema::project::ProjectScope;
    let mut out: HashMap<DeclId, Vec<Location>> = HashMap::new();
    for (module, m) in modules.iter().enumerate() {
        for c in &m.controls {
            let ty = c.control_type.to_ascii_lowercase();
            let tn = match ty.split_once('.') {
                Some((q, n)) => TypeName { qualifier: Some(q.to_string()), name: n.to_string() },
                None => TypeName { qualifier: None, name: ty },
            };
            let res = model.resolve_type_decl(module, &tn);
            let Some(id) = decl_id_of(module, &res) else { continue };
            for span in c.type_spans.iter().filter(|s| s.len > 0) {
                out.entry(id).or_default().push(Location { module, span: *span });
            }
        }
    }
    out
}

/// A path's file name without directory or extension (either separator).
fn file_stem(path: &str) -> &str {
    let name = path.rsplit(['/', '\\']).next().unwrap_or(path);
    name.rsplit_once('.').map_or(name, |(stem, _)| stem)
}

fn is_designer_path(path: &str) -> bool {
    matches!(
        path.rsplit('.').next().unwrap_or("").to_ascii_lowercase().as_str(),
        "frm" | "ctl" | "pag" | "dob" | "dsr"
    )
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
