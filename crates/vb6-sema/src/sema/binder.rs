//! VB6 binder (semantic analysis / name resolution).
//!
//! Takes the parser output (a flat `Vec<NodeId>` of top-level AST nodes) and
//! produces a [`BoundModule`] that:
//!
//! 1. Lists all declared procedures, module-level variables, type declarations,
//!    and enum declarations.
//! 2. Records a [`NameResolution`] for every `NameRef` node in the AST.
//! 3. Annotates every expression node with an inferred [`VbaType`].

use std::collections::HashMap;

// ── Semantic error codes ─────────────────────────────────────────────────────
//
// Each code is also a string-resource ID: the message text is looked up by the
// numeric code, so the code value equals the resource ID for its message.

/// `0x9caf` = 40111 = "Variable not defined"
///
/// Emitted when the RequireDeclaration flag (`Option Explicit`) is set and a
/// name is used without a declaration.
pub const ERR_VARIABLE_NOT_DEFINED: u16 = 0x9caf;

/// `0x9c9f` = 40095 = "Duplicate declaration in current scope"
///
/// Emitted when a name is inserted into a symbol or type table that already
/// holds an active entry for that key.
pub const ERR_DUPLICATE_DECLARATION: u16 = 0x9c9f;

/// `0x23` = 35 = "Sub or Function not defined"
///
/// Emitted for an unresolved name in *call* position — the call-binding sibling
/// of the `0x9caf` variable branch. Unlike `0x9caf`, this is **not** gated on
/// the RequireDeclaration bit: an undefined call is always an error, regardless
/// of `Option Explicit`.
pub const ERR_SUB_OR_FUNCTION_NOT_DEFINED: u16 = 0x23;

/// `461` = "Method or data member not found"
///
/// Emitted for `obj.Member` when the static type of `obj` is a class (or
/// user-defined type) of the project and it has no member of that name
/// reachable from the use site.
pub const ERR_METHOD_OR_DATA_MEMBER_NOT_FOUND: u16 = 461;

use crate::frontend::ast::{ExprArena, ExprNode, FileIoKind, NodeId, NodeSpans, ProcKind, Span};
use crate::frontend::diagnostics::Diagnostics;
use crate::frontend::scanner::ScannerContext;
use crate::sema::builtins::is_builtin;
use crate::sema::project::{
    Access, Arity, MemberHit, MemberLookup, MemberStatus, ProjectScope, Qualifier, SemType,
};
use crate::sema::symbol::{
    AccessorKind, BoundEnumDecl, BoundEnumMember, BoundModule, BoundParam, BoundProc,
    BoundTypeDecl, BoundTypeMember, BoundVar, BuiltinCall, ClassMemberSlot, DeclTypeNames,
    ExternalClass, ExternalDecl, NameResolution, ParamFlags, ResolvedClassMember, RtcArg,
    TypeName, UnaryIntrinsic,
};
use crate::sema::types::VbaType;

// ── Public entry point ────────────────────────────────────────────────────────

/// Bind a parsed module, resolving names and annotating types.
///
/// `ctx` — the scanner context (provides name→sym_id lookup).
/// `arena` — the expression arena produced by the parser.
/// `top_nodes` — the top-level node list from `Parser::parse_module`.
/// `spans` — the parser's node-span side table (declaration-name spans, used to
///   populate `name_span` on the bound declarations for LSP go-to-definition).
/// `visibility` — the parser's explicit-visibility table (`Parser::decl_public`);
///   declarations absent from it take the VB6 default (procedures/types/enums are
///   Public, module variables/constants are Private).
pub fn bind(
    ctx: &ScannerContext,
    arena: &ExprArena,
    top_nodes: &[NodeId],
    spans: &NodeSpans,
    visibility: &HashMap<u32, bool>,
) -> BoundModule {
    let empty = HashMap::new();
    bind_with_classes(ctx, arena, top_nodes, spans, visibility, &empty)
}

/// Like [`bind`], but with cross-module class member resolution: `classes`
/// maps a class's name (lowercase) to its already-bound Public field list
/// (see [`ExternalClass`]) — the caller (a project/fixture-level compiler)
/// binds the class's own module first, then passes its field list in here so
/// `Dim o As New ClassName` / `o.Field` in THIS module resolves `o.Field`'s
/// type instead of falling back to `Variant`. Sema itself has no "class
/// module" or "project" concept beyond this lookup table.
pub fn bind_with_classes(
    ctx: &ScannerContext,
    arena: &ExprArena,
    top_nodes: &[NodeId],
    spans: &NodeSpans,
    visibility: &HashMap<u32, bool>,
    classes: &HashMap<String, ExternalClass>,
) -> BoundModule {
    let mut b = Binder::new(ctx, arena, spans, visibility, classes, None);
    b.bind_top_level(top_nodes);
    b.finish()
}

/// Like [`bind`], but with the project around the module: `module` is this
/// module's index in the host's project and `scope` answers for every other
/// module and any external library. Beyond what [`bind`] records, this
/// resolves project-wide names and member accesses (`obj.Member`) against the
/// declared type of the base, records the static type of every expression in
/// `sem_types`, and accounts for each member access in `member_status`.
pub fn bind_in_project(
    ctx: &ScannerContext,
    arena: &ExprArena,
    top_nodes: &[NodeId],
    spans: &NodeSpans,
    visibility: &HashMap<u32, bool>,
    module: usize,
    scope: &dyn ProjectScope,
) -> BoundModule {
    let empty = HashMap::new();
    let mut b = Binder::new(ctx, arena, spans, visibility, &empty, Some((scope, module)));
    b.bind_top_level(top_nodes);
    b.finish()
}

/// Member-access coverage invariant for [`bind_in_project`]: every
/// `MemberAccess` node (and every `ReDim` of a member array) in `arena` must
/// have a status in `member_status`: resolved, or unresolved for a stated
/// reason. Returns the node ids that have none (empty = invariant holds).
pub fn unaccounted_member_accesses(
    arena: &ExprArena,
    member_status: &HashMap<u32, MemberStatus>,
) -> Vec<NodeId> {
    (0..arena.len() as u32)
        .filter(|i| match arena.get(NodeId(*i)) {
            ExprNode::MemberAccess { .. } => true,
            ExprNode::ReDimItem { base, .. } => base.is_some(),
            _ => false,
        })
        .filter(|i| !member_status.contains_key(i))
        .map(NodeId)
        .collect()
}

/// Resolution-coverage invariant: every `NameRef` node allocated in `arena` must
/// have an entry in `resolutions`. A `NameRef` that does not is a use site the
/// binder never reached — either the parser orphaned it (allocated but not
/// attached to the tree) or a tree-walk missed a child. For well-formed source
/// this must return empty; it is the structural backstop for the whole class of
/// "go-to-definition silently finds nothing" bugs.
///
/// Returns the offending `NameRef` node ids (empty = invariant holds).
pub fn unbound_namerefs(
    arena: &ExprArena,
    resolutions: &HashMap<u32, NameResolution>,
) -> Vec<NodeId> {
    // The qualified tail of a type name (`Proj.Class`) is a type-position name,
    // not a value use: it is not bound as a variable.
    let type_tails: std::collections::HashSet<u32> = (0..arena.len() as u32)
        .filter_map(|i| match arena.get(NodeId(i)) {
            ExprNode::UserType { child: Some(c), .. } => Some(c.0),
            _ => None,
        })
        .collect();
    (0..arena.len() as u32)
        .filter(|i| matches!(arena.get(NodeId(*i)), ExprNode::NameRef { .. }))
        .filter(|i| !type_tails.contains(i))
        .filter(|i| !resolutions.contains_key(i))
        .map(NodeId)
        .collect()
}

/// What a pre-pass over a procedure body found out about its candidate locals.
#[derive(Default)]
struct InferenceScan {
    /// Values `Set` into each local.
    sets: HashMap<usize, Vec<NodeId>>,
    /// Values `Let` into each local.
    lets: HashMap<usize, Vec<NodeId>>,
    /// Locals something else may store into.
    blocked: std::collections::HashSet<usize>,
}

/// What the `Set` statements of a procedure say about a `Variant`/`Object` local.
#[derive(Debug, Clone, PartialEq)]
enum Inference {
    /// Every `Set` assigns an object of this one type.
    Single(SemType),
    /// The `Set`s assign objects of these different types, in order first seen.
    Ambiguous(Vec<SemType>),
}

// ── Internal binder state ─────────────────────────────────────────────────────

struct Binder<'a> {
    ctx: &'a ScannerContext,
    arena: &'a ExprArena,
    spans: &'a NodeSpans,
    visibility: &'a HashMap<u32, bool>,
    /// Known external classes (name lowercase → Public field list), for
    /// cross-module member resolution. See [`bind_with_classes`].
    known_classes: &'a HashMap<String, ExternalClass>,
    /// Type-sym -> matched external class, recorded as `extract_type` resolves
    /// each `UserType` reference against `known_classes` (interior mutability:
    /// `extract_type` is called with `&self` from many non-`&mut` contexts).
    /// Exposed on `BoundModule` so codegen (no scanner/interner of its own)
    /// can recognize a `VbaType::UserDefined(sym)` local as a class instance
    /// by sym_id, without a name-text lookup.
    class_field_info: std::cell::RefCell<HashMap<u32, ExternalClass>>,
    /// Resolved vtable dispatch slots per class-member `MemberAccess` node,
    /// recorded as a side effect of `resolve_member_type` (interior
    /// mutability for the same reason as `class_field_info`). See
    /// `BoundModule::class_member_slots`.
    class_member_slots: std::cell::RefCell<HashMap<u32, ResolvedClassMember>>,

    procs:       Vec<BoundProc>,
    /// Parameter `ArgList` node per proc (aligned with `procs`), kept so the
    /// body pass can bind `Optional` default-value expressions, which live in the
    /// parameter list rather than the body block.
    proc_param_nodes: Vec<Option<NodeId>>,
    /// Return type-spec node per proc (aligned with `procs`), bound with the
    /// parameters so a named return type is a use of that type.
    proc_ret_nodes: Vec<Option<NodeId>>,
    module_vars: Vec<BoundVar>,
    type_decls:  Vec<BoundTypeDecl>,
    enum_decls:  Vec<BoundEnumDecl>,

    resolutions: HashMap<u32, NameResolution>,
    types:       HashMap<u32, VbaType>,
    builtins:    HashMap<u32, BuiltinCall>,

    /// `DefType` letter-range map: first character of an untyped name → inferred VbaType.
    deftype_map: HashMap<char, VbaType>,

    diagnostics:     Diagnostics,
    option_explicit: bool,

    decl_types: DeclTypeNames,
    vb_name: Option<String>,
    predeclared_id: bool,
    global_namespace: bool,
    default_member: Option<String>,

    /// Diagnostics held when the body being bound was started: a second pass
    /// over the body drops what the first one reported.
    proc_diag_mark: usize,
    /// The project and this module's index in it; `None` for a plain [`bind`].
    project: Option<(&'a dyn ProjectScope, usize)>,
    sem_types: HashMap<u32, SemType>,
    member_status: HashMap<u32, MemberStatus>,
    /// Accessor a member use needs, keyed by the `MemberAccess` / `NameRef`
    /// node: set from the enclosing statement before the node is bound.
    access_ctx: HashMap<u32, Access>,
    /// Nodes that resolved to a procedure, with its declared parameter count.
    callable: HashMap<u32, Arity>,
    /// Static types of the `With` objects in force, innermost last.
    with_stack: Vec<SemType>,
    /// Declared type names of each procedure's parameters, aligned with `procs`.
    param_names: Vec<Vec<Option<TypeName>>>,
    /// The inferred type of procedure locals of the procedure being bound,
    /// by local index (project mode, second pass of a procedure only).
    inference: HashMap<usize, Inference>,
    /// Nodes whose static type came from an inferred local, directly or
    /// through members of it.
    inferred_nodes: std::collections::HashSet<u32>,
    /// Reads of an ambiguous local, with the candidate type names.
    ambiguous_nodes: HashMap<u32, Vec<String>>,
    /// Static types of declarations (project mode only).
    var_sem: Vec<SemType>,
    ret_sem: Vec<SemType>,
    param_sem: Vec<Vec<SemType>>,
    local_sem: Vec<Vec<SemType>>,
}

impl<'a> Binder<'a> {
    fn new(
        ctx: &'a ScannerContext,
        arena: &'a ExprArena,
        spans: &'a NodeSpans,
        visibility: &'a HashMap<u32, bool>,
        known_classes: &'a HashMap<String, ExternalClass>,
        project: Option<(&'a dyn ProjectScope, usize)>,
    ) -> Self {
        Self {
            ctx,
            arena,
            spans,
            visibility,
            known_classes,
            class_field_info: std::cell::RefCell::new(HashMap::new()),
            class_member_slots: std::cell::RefCell::new(HashMap::new()),
            procs:       Vec::new(),
            proc_param_nodes: Vec::new(),
            proc_ret_nodes: Vec::new(),
            module_vars: Vec::new(),
            type_decls:  Vec::new(),
            enum_decls:  Vec::new(),
            resolutions: HashMap::new(),
            types:       HashMap::new(),
            builtins:    HashMap::new(),
            deftype_map: HashMap::new(),
            diagnostics:     Diagnostics::new(),
            option_explicit: false,
            decl_types: DeclTypeNames::default(),
            vb_name: None,
            predeclared_id: false,
            global_namespace: false,
            default_member: None,
            proc_diag_mark: 0,
            project,
            sem_types: HashMap::new(),
            member_status: HashMap::new(),
            access_ctx: HashMap::new(),
            callable: HashMap::new(),
            with_stack: Vec::new(),
            param_names: Vec::new(),
            inference: HashMap::new(),
            inferred_nodes: std::collections::HashSet::new(),
            ambiguous_nodes: HashMap::new(),
            var_sem: Vec::new(),
            ret_sem: Vec::new(),
            param_sem: Vec::new(),
            local_sem: Vec::new(),
        }
    }

    fn finish(self) -> BoundModule {
        BoundModule {
            procs:       self.procs,
            module_vars: self.module_vars,
            type_decls:  self.type_decls,
            enum_decls:  self.enum_decls,
            resolutions: self.resolutions,
            types:       self.types,
            builtins:    self.builtins,
            diagnostics:     self.diagnostics,
            option_explicit: self.option_explicit,
            class_field_info: self.class_field_info.into_inner(),
            class_member_slots: self.class_member_slots.into_inner(),
            decl_types: self.decl_types,
            vb_name: self.vb_name,
            predeclared_id: self.predeclared_id,
            global_namespace: self.global_namespace,
            default_member: self.default_member,
            sem_types: self.sem_types,
            member_status: self.member_status,
        }
    }

    // ── Name helpers ──────────────────────────────────────────────────────────

    /// Visibility of the declaration at `id`: the parser's explicit modifier if
    /// present, otherwise `default` (the VB6 default for that declaration kind).
    fn vis(&self, id: NodeId, default: bool) -> bool {
        self.visibility.get(&id.0).copied().unwrap_or(default)
    }

    /// Return the lowercase name for a sym_id, or "" if not found.
    fn name_of(&self, sym_id: u32) -> String {
        if sym_id == 0 {
            return String::new();
        }
        self.ctx
            .symbol(sym_id as usize)
            .name
            .to_ascii_lowercase()
    }

    // ── Type extraction from AST nodes ────────────────────────────────────────

    /// Extract the `VbaType` from a type-spec node (`None` = absent → Variant).
    fn extract_type(&self, type_node: Option<NodeId>) -> VbaType {
        let Some(type_node_id) = type_node else {
            return VbaType::Variant;
        };
        match self.arena.get(type_node_id) {
            ExprNode::BuiltinType { kind } => VbaType::from_kind(*kind),
            ExprNode::StringType { .. }   => VbaType::String,
            ExprNode::AnyType             => VbaType::Any,
            ExprNode::ArrayType { elem } => {
                VbaType::Array(Box::new(self.extract_type(Some(*elem))))
            }
            ExprNode::UserType { name, child } => {
                // A qualified name (`Project.Class`, `Module.Type`) names the
                // type by its last part; the qualifier only selects where to
                // look.
                let name = match child.map(|c| self.arena.get(c)) {
                    Some(ExprNode::NameRef { sym, .. }) => sym,
                    _ => name,
                };
                // If `name` matches a known external class (not a same-module
                // `Type` declaration — those stay resolved via `type_decls`),
                // record its field list keyed by this type sym so codegen
                // (which has no scanner/interner to resolve sym -> name
                // itself) can look the class up directly by sym_id, without
                // redoing the name-text lookup `resolve_member_type` does.
                if let Some(class) = self.known_classes.get(&self.name_of(*name)) {
                    self.class_field_info
                        .borrow_mut()
                        .insert(*name, class.clone());
                }
                VbaType::UserDefined(*name)
            }
            // User-defined-type *references* (`Dim x As MyType`) are emitted by
            // the recursive-descent parser as `UserType` and resolved above. The
            // raw 0xb3 `TypeSpec`/`UdtTypeSpec` nodes (built only by
            // `make_type_spec_node`/`make_udt_type_node`) are the low-level VB6
            // node form; for the UDT flag they carry no name field (`type_kind`
            // is 0, the name lives in the declaration's field-type table, not
            // here), so `Variant` is the correct conservative result rather than
            // a fabricated symbol.
            ExprNode::TypeSpec { type_kind, type_flags, .. } => {
                if type_flags & 0x8000 != 0 {
                    VbaType::Variant
                } else {
                    VbaType::from_kind(*type_kind)
                }
            }
            ExprNode::UdtTypeSpec { .. } => VbaType::Variant,
            _ => VbaType::Variant,
        }
    }

    /// The named (user) type a type-spec node declares, lowercased, or `None`
    /// for a built-in type or no type.
    fn type_name_of(&self, type_node: Option<NodeId>) -> Option<TypeName> {
        let mut node = type_node?;
        while let ExprNode::ArrayType { elem } = self.arena.get(node) {
            node = *elem;
        }
        let ExprNode::UserType { name, child } = self.arena.get(node) else {
            return None;
        };
        let first = self.name_of(*name);
        if first.is_empty() {
            return None;
        }
        match child.map(|c| self.arena.get(c)) {
            Some(ExprNode::NameRef { sym, .. }) => {
                let last = self.name_of(*sym);
                (!last.is_empty()).then(|| TypeName { qualifier: Some(first), name: last })
            }
            _ => Some(TypeName { qualifier: None, name: first }),
        }
    }

    /// The type names of a `Type` block's members, aligned with
    /// [`collect_type_decl`](Self::collect_type_decl)'s member list.
    fn type_member_names(&self, member_nodes: &[NodeId]) -> Vec<Option<TypeName>> {
        member_nodes
            .iter()
            .filter_map(|&id| match self.arena.get(id) {
                ExprNode::DimItem { type_node, .. } => Some(self.type_name_of(*type_node)),
                _ => None,
            })
            .collect()
    }

    // ── Top-level pass ────────────────────────────────────────────────────────

    fn bind_top_level(&mut self, top_nodes: &[NodeId]) {
        // Pass 1: collect all declarations.
        for &id in top_nodes {
            self.collect_top_decl(id);
        }

        if self.project.is_some() {
            self.compute_declaration_types();
        }

        // Pass 2: bind module-level declaration expressions in module scope —
        // array bounds (`Dim a(1 To MAX)`), `Const`/enum values, and type-member
        // bounds. These live outside any procedure body, so the body pass below
        // never reaches them.
        for &id in top_nodes {
            self.bind_module_node(id);
        }

        // Pass 3: bind procedures (resolve names + annotate types).
        // We iterate by index so we can mutate self.resolutions/types.
        let proc_count = self.procs.len();
        for proc_idx in 0..proc_count {
            // Parameter defaults bind in proc scope even for body-less `Declare`s.
            if let Some(params) = self.proc_param_nodes[proc_idx] {
                self.bind_node(Some(proc_idx), params);
            }
            if let Some(ret) = self.proc_ret_nodes[proc_idx] {
                self.bind_node(Some(proc_idx), ret);
            }
            let body = self.procs[proc_idx].body;
            if body != u32::MAX {
                self.bind_proc_body(proc_idx, NodeId(body));
            }
        }

        // Pass 4: module-local semantic diagnostics (locals are now collected).
        self.check_duplicate_declarations();
    }

    /// Bind a top-level node's expressions in module scope, skipping procedures
    /// and `Declare`s (their bodies/params are bound in proc scope by pass 3).
    fn bind_module_node(&mut self, id: NodeId) {
        match self.arena.get(id) {
            ExprNode::ProcDecl { .. } | ExprNode::DeclareDecl { .. } => {}
            ExprNode::Block { stmts } => {
                let stmts = stmts.clone();
                for s in stmts {
                    self.bind_module_node(s);
                }
            }
            _ => self.bind_node(None, id),
        }
    }

    /// Emit `ERR_DUPLICATE_DECLARATION` for names declared twice in one scope.
    ///
    /// Module scope: module-level variables, procedures, types, and enums.
    /// For procedures, Property Get/Let/Set accessors with the same name are
    /// **not** duplicates of each other (VB6 allows one of each kind), but two
    /// `Property Get` with the same name, or a `Sub`+`Property Get` with the
    /// same name, are errors.
    fn check_duplicate_declarations(&mut self) {
        use std::collections::{HashMap, HashSet};

        let mut dups: Vec<Span> = Vec::new();

        // ── Module-level variables ──────────────────────────────────────────
        let mut seen: HashSet<String> = HashSet::new();
        for v in &self.module_vars {
            self.record_dup(&mut seen, v.sym_id, v.name_span, &mut dups);
        }

        // ── Procedures — with Property accessor grouping ────────────────────
        // Group procs by lowercase name. Within each group, at most one of each
        // Property accessor kind (Get/Let/Set) is allowed; any non-Property proc
        // mixed with a Property, or two procs of the same non-Property kind, is
        // a duplicate.
        {
            // seen_accessor[name] = bitfield: bit0=Get, bit1=Let, bit2=Set, bit3=non-property
            let mut accessor_seen: HashMap<String, u8> = HashMap::new();
            for p in &self.procs {
                let name = self.name_of(p.sym_id);
                if name.is_empty() { continue; }
                let bit = match p.kind {
                    ProcKind::PropGet => 0b0001,
                    ProcKind::PropLet => 0b0010,
                    ProcKind::PropSet => 0b0100,
                    _ => 0b1000, // Sub or Function
                };
                let entry = accessor_seen.entry(name).or_insert(0);
                if *entry & bit != 0 {
                    dups.push(p.name_span);
                } else if bit == 0b1000 && (*entry & 0b0111 != 0) {
                    // Non-property proc clashes with a Property accessor
                    dups.push(p.name_span);
                } else if bit != 0b1000 && (*entry & 0b1000 != 0) {
                    // Property accessor clashes with an existing non-property proc
                    dups.push(p.name_span);
                } else {
                    *entry |= bit;
                }
            }
        }

        // ── Type and enum names ─────────────────────────────────────────────
        let mut type_names: HashSet<String> = HashSet::new();
        for t in &self.type_decls {
            self.record_dup(&mut type_names, t.sym_id, t.name_span, &mut dups);
        }
        for e in &self.enum_decls {
            // Enum names share the module type-name space with UDT names.
            let name = self.name_of(e.sym_id);
            if !name.is_empty() && !type_names.insert(name) {
                dups.push(e.name_span);
            }
        }

        // ── Procedure scope: parameters and locals share one scope ──────────
        for p in &self.procs {
            let mut pseen: HashSet<String> = HashSet::new();
            for param in &p.params {
                self.record_dup(&mut pseen, param.sym_id, param.name_span, &mut dups);
            }
            for local in &p.locals {
                self.record_dup(&mut pseen, local.sym_id, local.name_span, &mut dups);
            }
        }

        for span in dups {
            self.diagnostics.push(ERR_DUPLICATE_DECLARATION as u32, span);
        }
    }

    /// Record `name_span` as a duplicate if this scope's `seen` set already holds
    /// the (non-empty) name; otherwise insert it. Shared by module and proc scopes.
    fn record_dup(
        &self,
        seen: &mut std::collections::HashSet<String>,
        sym_id: u32,
        name_span: Span,
        dups: &mut Vec<Span>,
    ) {
        let name = self.name_of(sym_id);
        if name.is_empty() {
            return;
        }
        if !seen.insert(name) {
            dups.push(name_span);
        }
    }

    // ── Declaration collection ────────────────────────────────────────────────

    fn collect_top_decl(&mut self, id: NodeId) {
        match self.arena.get(id) {
            ExprNode::ProcDecl { kind, name, params, ret_type, body } => {
                self.collect_top_proc(id, *kind, *name, *params, *ret_type, *body);
            }
            ExprNode::Block { stmts } => {
                let stmts = stmts.clone();
                for sid in stmts {
                    self.collect_top_decl(sid);
                }
            }
            ExprNode::DimItem { name, is_const, is_new, bounds, type_node, init, .. } => {
                self.collect_top_var(id, *name, *is_const, *is_new, bounds.is_some(), *type_node, *init);
            }
            ExprNode::TypeDecl { name, members } => {
                let sym_id = *name;
                let members_clone = members.clone();
                let name_span = self.spans.get(id);
                let is_public = self.vis(id, true); // types default Public
                let decl = self.collect_type_decl(sym_id, &members_clone, name_span, is_public);
                let names = self.type_member_names(&members_clone);
                self.decl_types.type_members.push(names);
                self.type_decls.push(decl);
            }
            ExprNode::EnumDecl { name, members } => {
                let sym_id = *name;
                let members_clone = members.clone();
                let name_span = self.spans.get(id);
                let is_public = self.vis(id, true); // enums default Public
                let decl = self.collect_enum_decl(sym_id, &members_clone, name_span, is_public);
                self.enum_decls.push(decl);
            }
            ExprNode::DeclareDecl { kind, name, params, ret_type, .. } => {
                // A `Declare` (external API) is a callable procedure with no body.
                // Collecting it lets calls resolve (so they are not falsely
                // flagged "Sub or Function not defined"). `body = u32::MAX` marks
                // the missing body; the body-binding pass skips it.
                self.collect_top_proc(id, *kind, *name, *params, *ret_type, NodeId(u32::MAX));
            }
            ExprNode::OptionExplicit => {
                self.option_explicit = true;
            }
            ExprNode::DefType { type_kw, ranges } => {
                let ty = deftype_kw_to_vbatype(*type_kw);
                let ranges = ranges.clone();
                for (lo_sym, hi_sym) in &ranges {
                    let lo = self.name_of(*lo_sym).chars().next().unwrap_or('a');
                    let hi = self.name_of(*hi_sym).chars().next().unwrap_or(lo);
                    for ch in lo..=hi {
                        self.deftype_map.insert(ch, ty.clone());
                    }
                }
            }
            // Ignored at module level: Implements, EventDecl, empty Block, Generic
            _ => {}
        }
    }

    /// Collect a module-level procedure (`ProcDecl` or bodyless `Declare`).
    /// `body_id` is `NodeId(u32::MAX)` for `Declare`. Procedures, including
    /// `Declare`s, default to Public.
    fn collect_top_proc(
        &mut self,
        id: NodeId,
        kind: ProcKind,
        sym_id: u32,
        params_id: Option<NodeId>,
        ret_type_id: Option<NodeId>,
        body_id: NodeId,
    ) {
        let name_span = self.spans.get(id);
        let is_public = self.vis(id, true);
        let proc =
            self.collect_proc(sym_id, kind, params_id, ret_type_id, body_id, name_span, is_public);
        let ret_name = self.type_name_of(ret_type_id);
        self.decl_types.proc_rets.push(ret_name);
        let names = self.param_type_names(params_id);
        self.decl_types.proc_params.push(names.clone());
        self.param_names.push(names);
        self.procs.push(proc);
        self.proc_param_nodes.push(params_id);
        self.proc_ret_nodes.push(ret_type_id);
    }

    /// Collect a module-level variable (`Dim`/`Const` item). Module `Dim`/`Const`
    /// default to Private; array bounds wrap the element type in `Array`.
    ///
    /// A module-level `Const`'s value is folded here the SAME way a
    /// procedure-level one is (`collect_locals`'s own `const_value`/
    /// `const_lit` derivation) — this used to unconditionally leave both
    /// `None` regardless of `is_const`, silently discarding the `init`
    /// expression entirely (the old signature didn't even take one). The
    /// code generator's `NameResolution::ModuleVar` path had no fallback for
    /// this: it read a module-level `Const` back as an ordinary global
    /// variable (a bogus load through an unused frame slot, wrong bytes) —
    /// oracle-confirmed broken (`oracle_bank/c21_module_const`) before this
    /// fix.
    fn collect_top_var(
        &mut self,
        id: NodeId,
        sym_id: u32,
        is_const: bool,
        is_new: bool,
        has_bounds: bool,
        type_node: Option<NodeId>,
        init: Option<NodeId>,
    ) {
        let const_value = if is_const { init.and_then(|i| self.eval_const_i64(i)) } else { None };
        let const_lit = if is_const && const_value.is_none() {
            init.and_then(|i| self.eval_const_lit(i))
        } else {
            None
        };
        let var = BoundVar {
            sym_id,
            vba_type: self.extract_type(type_node),
            is_const,
            const_value,
            const_lit,
            fixed_string_len: None,
            array_dims: None,
            is_static: false,
            is_public: self.vis(id, false),
            is_new,
            name_span: self.spans.get(id),
        };
        let var = if has_bounds {
            BoundVar { vba_type: as_array(var.vba_type), ..var }
        } else {
            var
        };
        let type_name = self.type_name_of(type_node);
        self.decl_types.module_vars.push(type_name);
        self.module_vars.push(var);
    }

    fn collect_proc(
        &self,
        sym_id: u32,
        kind: ProcKind,
        params_id: Option<NodeId>,
        ret_type_id: Option<NodeId>,
        body_id: NodeId,
        name_span: Span,
        is_public: bool,
    ) -> BoundProc {
        let ret_type = self.extract_type(ret_type_id);
        let params = match params_id {
            Some(p) => self.collect_params(p),
            None => Vec::new(),
        };
        BoundProc {
            sym_id,
            kind,
            params,
            ret_type,
            locals: Vec::new(),  // filled in during body pass
            body: body_id.0,
            is_public,
            name_span,
        }
    }

    fn collect_params(&self, params_node_id: NodeId) -> Vec<BoundParam> {
        match self.arena.get(params_node_id) {
            ExprNode::ArgList { args } => {
                let args = args.clone();
                args.iter().map(|&id| self.collect_param(id)).collect()
            }
            _ => Vec::new(),
        }
    }

    /// The declared type name of each parameter, aligned with
    /// [`collect_params`](Self::collect_params).
    fn param_type_names(&self, params_id: Option<NodeId>) -> Vec<Option<TypeName>> {
        let Some(ExprNode::ArgList { args }) = params_id.map(|p| self.arena.get(p)) else {
            return Vec::new();
        };
        args.iter()
            .map(|&a| match self.arena.get(a) {
                ExprNode::ParamDef { type_node, .. } => self.type_name_of(*type_node),
                _ => None,
            })
            .collect()
    }

    fn collect_param(&self, id: NodeId) -> BoundParam {
        match self.arena.get(id) {
            ExprNode::ParamDef { flags, name, type_node, .. } => BoundParam {
                sym_id:   *name,
                vba_type: {
                    let t = self.extract_type(*type_node);
                    if flags & 0x08 != 0 { VbaType::Array(Box::new(t)) } else { t }
                },
                flags: ParamFlags::from_bits(*flags),
                name_span: self.spans.get(id),
            },
            _ => BoundParam {
                sym_id:   0,
                vba_type: VbaType::Variant,
                flags:    ParamFlags::default(),
                name_span: Span::DUMMY,
            },
        }
    }

    fn collect_type_decl(&self, sym_id: u32, member_nodes: &[NodeId], name_span: Span, is_public: bool) -> BoundTypeDecl {
        let members = member_nodes
            .iter()
            .filter_map(|&id| match self.arena.get(id) {
                ExprNode::DimItem { name, type_node, bounds, .. } => Some(BoundTypeMember {
                    sym_id:   *name,
                    vba_type: {
                        let t = self.extract_type(*type_node);
                        if bounds.is_some() { as_array(t) } else { t }
                    },
                    name_span: self.spans.get(id),
                }),
                _ => None,
            })
            .collect();
        BoundTypeDecl { sym_id, members, is_public, name_span }
    }

    fn collect_enum_decl(&self, sym_id: u32, member_nodes: &[NodeId], name_span: Span, is_public: bool) -> BoundEnumDecl {
        let mut next_val: i64 = 0;
        let members = member_nodes
            .iter()
            .filter_map(|&id| match self.arena.get(id) {
                ExprNode::DimItem { name, init, .. } => {
                    let value = init
                        .as_ref()
                        .and_then(|&init_id| self.eval_const_i64(init_id))
                        .unwrap_or(next_val);
                    next_val = value + 1;
                    Some(BoundEnumMember { sym_id: *name, value, name_span: self.spans.get(id) })
                }
                _ => None,
            })
            .collect();
        BoundEnumDecl { sym_id, members, is_public, name_span }
    }

    /// Attempt to constant-fold an expression to i64.  Returns None if the
    /// expression is not a simple integer/long literal.
    fn eval_const_i64(&self, id: NodeId) -> Option<i64> {
        match self.arena.get(id) {
            ExprNode::Literal { lit } => match lit {
                crate::frontend::ast::AstLit::Int(n)  => Some(*n as i64),
                crate::frontend::ast::AstLit::Long(n) => Some(*n as i64),
                _ => None,
            },
            ExprNode::UnOp { op: crate::frontend::ast::UnOpKind::Neg, operand } => {
                self.eval_const_i64(*operand).map(|v| -v)
            }
            ExprNode::Paren { inner } => self.eval_const_i64(*inner),
            _ => None,
        }
    }

    /// Resolve a const initializer that is a (parenthesised) literal to that
    /// literal, for non-integer constant folding by the code generator.
    fn eval_const_lit(&self, id: NodeId) -> Option<crate::frontend::ast::AstLit> {
        match self.arena.get(id) {
            ExprNode::Literal { lit } => Some(lit.clone()),
            ExprNode::Paren { inner } => self.eval_const_lit(*inner),
            _ => None,
        }
    }

    // ── Procedure body binding ────────────────────────────────────────────────

    fn bind_proc_body(&mut self, proc_idx: usize, body_id: NodeId) {
        // Collect local variables from the body first.
        let mut locals: Vec<BoundVar> = Vec::new();
        // A Function / Property Get exposes its own name as an implicit
        // return-value local at frame slot 0, typed as the return type. Injecting
        // it as local index 0 makes `F = …` / reads of `F` resolve as a normal
        // local (and the code generator allocates the return slot before the
        // user locals, matching the compiler's frame layout).
        let mut local_names: Vec<Option<TypeName>> = Vec::new();
        let (kind, ret_type, sym_id, name_span) = {
            let p = &self.procs[proc_idx];
            (p.kind, p.ret_type.clone(), p.sym_id, p.name_span)
        };
        if matches!(kind, ProcKind::Function | ProcKind::PropGet) {
            locals.push(BoundVar {
                sym_id,
                vba_type: ret_type,
                is_const: false,
                const_value: None,
                const_lit: None,
                fixed_string_len: None,
                array_dims: None,
                is_static: false,
                is_public: false,
                is_new: false,
                name_span,
            });
            local_names.push(self.decl_types.proc_rets.get(proc_idx).cloned().flatten());
        }
        self.proc_diag_mark = self.diagnostics.len();
        self.collect_locals(proc_idx, body_id, &mut locals, &mut local_names);
        if self.project.is_some() {
            let sems = locals
                .iter()
                .zip(&local_names)
                .map(|(l, n)| self.sem_of_declared(&l.vba_type, n.as_ref()))
                .collect();
            self.local_sem[proc_idx] = sems;
        }
        self.procs[proc_idx].locals = locals;

        // Bind all name references in the body. (`Optional` parameter defaults
        // are bound separately by the caller, since they live in the parameter
        // list rather than the body block.)
        self.bind_node(Some(proc_idx), body_id);
        if self.project.is_some() {
            self.infer_and_rebind(proc_idx, body_id);
        }
    }

    /// Procedure-local object-type inference. For each local of type `Variant`
    /// or `Object` that is only ever set by `Set` statements and never stored
    /// into or passed by reference otherwise, the types assigned by those
    /// `Set`s are the local's type: one type is an inference, several are an
    /// ambiguity. The body is bound again with that knowledge, so member
    /// accesses on the local resolve (or are reported ambiguous).
    fn infer_and_rebind(&mut self, proc_idx: usize, body_id: NodeId) {
        let procs_locals = &self.procs[proc_idx].locals;
        // Candidates by lowercase name. The implicit return slot is excluded.
        let proc_name = self.name_of(self.procs[proc_idx].sym_id);
        let mut candidates: HashMap<String, usize> = HashMap::new();
        for (i, l) in procs_locals.iter().enumerate() {
            let n = self.name_of(l.sym_id);
            let loose = matches!(l.vba_type, VbaType::Variant | VbaType::Object);
            if loose && !l.is_const && n != proc_name {
                candidates.entry(n).or_insert(i);
            }
        }
        if candidates.is_empty() {
            return;
        }
        let mut scan = InferenceScan::default();
        self.scan_for_inference(body_id, &candidates, &mut scan);
        let mut table: HashMap<usize, Inference> = HashMap::new();
        for (&local, values) in &scan.sets {
            if scan.blocked.contains(&local) || values.is_empty() {
                continue;
            }
            // A value-typed `Let` is harmless; a `Let` of anything else may
            // copy an object in.
            let lets_unknown = scan.lets.get(&local).map_or(false, |vs| {
                vs.iter().any(|v| !matches!(self.sem_of_node(*v), SemType::Value(_)))
            });
            if lets_unknown {
                continue;
            }
            let mut types: Vec<SemType> = Vec::new();
            let mut known = true;
            for v in values {
                match self.sem_of_node(*v) {
                    t @ (SemType::Class { .. } | SemType::Lib { .. }) => {
                        if !types.contains(&t) {
                            types.push(t);
                        }
                    }
                    _ => known = false,
                }
            }
            if !known || types.is_empty() {
                continue;
            }
            let entry = if types.len() == 1 {
                Inference::Single(types.remove(0))
            } else {
                Inference::Ambiguous(types)
            };
            table.insert(local, entry);
        }
        if table.is_empty() {
            return;
        }
        // Forget everything the first pass recorded for this body and bind it
        // again with the inferred types in force.
        let mut ids: Vec<u32> = Vec::new();
        self.collect_ids(body_id, &mut ids);
        for id in &ids {
            self.resolutions.remove(id);
            self.types.remove(id);
            self.builtins.remove(id);
            self.sem_types.remove(id);
            self.member_status.remove(id);
            self.callable.remove(id);
            self.class_member_slots.borrow_mut().remove(id);
        }
        self.diagnostics.truncate(self.proc_diag_mark);
        self.inference = table;
        self.bind_node(Some(proc_idx), body_id);
        self.inference.clear();
    }

    fn collect_ids(&self, id: NodeId, out: &mut Vec<u32>) {
        out.push(id.0);
        if let ExprNode::ForRange { loop_var, range, step } = self.arena.get(id) {
            for raw in [*loop_var, *range, *step] {
                if raw != 0 {
                    self.collect_ids(NodeId(raw), out);
                }
            }
            return;
        }
        let mut kids: Vec<NodeId> = Vec::new();
        self.arena.get(id).for_each_child(&mut |c| kids.push(c));
        for c in kids {
            self.collect_ids(c, out);
        }
    }

    /// Record, for every candidate local, the values `Set` into it and
    /// whether anything else could change what it holds.
    fn scan_for_inference(
        &self,
        id: NodeId,
        candidates: &HashMap<String, usize>,
        scan: &mut InferenceScan,
    ) {
        let candidate = |n: NodeId| -> Option<usize> {
            match self.arena.get(n) {
                ExprNode::NameRef { sym, .. } => candidates.get(&self.name_of(*sym)).copied(),
                _ => None,
            }
        };
        match self.arena.get(id) {
            ExprNode::SetAssign { target, value } => {
                if let Some(l) = candidate(*target) {
                    // `Set x = Nothing` says nothing about the type.
                    if !matches!(self.arena.get(*value), ExprNode::Nothing) {
                        scan.sets.entry(l).or_default().push(*value);
                    }
                }
            }
            ExprNode::Assign { target, value } => {
                if let Some(l) = candidate(*target) {
                    scan.lets.entry(l).or_default().push(*value);
                }
            }
            ExprNode::ForEach { var, .. } => {
                if let Some(l) = candidate(*var) {
                    scan.blocked.insert(l);
                }
            }
            ExprNode::FileIoStmt { kind, args, .. } => {
                if matches!(kind, FileIoKind::Input | FileIoKind::LineInput | FileIoKind::Get) {
                    for a in args {
                        if let Some(l) = candidate(*a) {
                            scan.blocked.insert(l);
                        }
                    }
                }
            }
            ExprNode::ReDimItem { base: None, name, .. } => {
                if let Some(l) = candidates.get(&self.name_of(*name)) {
                    scan.blocked.insert(*l);
                }
            }
            ExprNode::Call { func, args } | ExprNode::CallStmt { callee: func, args } => {
                self.scan_call_arguments(*func, *args, candidates, scan);
            }
            ExprNode::RaiseEvent { args, .. } => {
                self.scan_call_arguments(id, *args, candidates, scan);
            }
            _ => {}
        }
        let mut kids: Vec<NodeId> = Vec::new();
        self.arena.get(id).for_each_child(&mut |c| kids.push(c));
        for c in kids {
            self.scan_for_inference(c, candidates, scan);
        }
    }

    /// A candidate local handed to a procedure as a bare argument may be
    /// reassigned through the reference, unless the callee is a built-in or a
    /// procedure of this module whose parameter is `ByVal`.
    fn scan_call_arguments(
        &self,
        func: NodeId,
        args: NodeId,
        candidates: &HashMap<String, usize>,
        scan: &mut InferenceScan,
    ) {
        let ExprNode::ArgList { args } = self.arena.get(args) else { return };
        let callee_name = match self.arena.get(func) {
            ExprNode::NameRef { sym, .. } => Some(self.name_of(*sym)),
            _ => None,
        };
        let own = callee_name.as_ref().and_then(|n| {
            self.procs.iter().find(|p| self.name_of(p.sym_id) == *n)
        });
        let builtin = match (&callee_name, own) {
            (Some(n), None) => {
                is_builtin(n)
                    && !self.module_vars.iter().any(|v| self.name_of(v.sym_id) == *n)
            }
            _ => false,
        };
        for (i, a) in args.iter().enumerate() {
            let mut node = *a;
            if let ExprNode::NamedArg { value, .. } = self.arena.get(node) {
                node = *value;
            }
            let ExprNode::NameRef { sym, .. } = self.arena.get(node) else { continue };
            let Some(&l) = candidates.get(&self.name_of(*sym)) else { continue };
            let by_val = own
                .and_then(|p| p.params.get(i))
                .map_or(false, |prm| prm.flags.by_val);
            if !(builtin || by_val) {
                scan.blocked.insert(l);
            }
        }
    }

    /// Whether `name` is already a declaration visible to a `ReDim` in
    /// procedure `proc_idx`: a local collected so far, a parameter, the
    /// procedure's own name, a module-level variable, or (in a project) a
    /// public variable of another module.
    fn is_declared_for_redim(&self, proc_idx: usize, name: u32, collected: &[BoundVar]) -> bool {
        let text = self.name_of(name);
        if text.is_empty() {
            return true;
        }
        let same = |sym: u32| self.name_of(sym) == text;
        let proc = &self.procs[proc_idx];
        if collected.iter().any(|l| same(l.sym_id))
            || proc.params.iter().any(|p| same(p.sym_id))
            || same(proc.sym_id)
            || self.module_vars.iter().any(|v| same(v.sym_id))
        {
            return true;
        }
        match self.project {
            Some((project, module)) => matches!(
                project.lookup_global(module, &text),
                Some(MemberHit {
                    resolution: NameResolution::External { decl: ExternalDecl::Var(_), .. },
                    ..
                })
            ),
            None => false,
        }
    }

    /// Collect every `DimItem` in a procedure body subtree as a local variable.
    ///
    /// VB6 locals are procedure-scoped regardless of nesting (no block scope), so
    /// every `DimItem` anywhere under the body is a local. Descent goes through
    /// [`ExprNode::for_each_child`] — the same single source of truth the binder
    /// uses — so a local declared inside any statement form (even one added
    /// later) is collected, never silently resolved to the wrong scope. A proc
    /// body contains no nested procedure declarations, so a full descent is safe.
    fn collect_locals(
        &self,
        proc_idx: usize,
        id: NodeId,
        out: &mut Vec<BoundVar>,
        names: &mut Vec<Option<TypeName>>,
    ) {
        // `ReDim a(n)` of a name that is not declared at procedure or module
        // level declares a local dynamic array at that statement.
        if let ExprNode::ReDimItem { base: None, name, bounds, type_node, .. } = self.arena.get(id) {
            if !self.is_declared_for_redim(proc_idx, *name, out) {
                let elem = self.extract_type(*type_node);
                let dims = (*bounds).and_then(|b| match self.arena.get(b) {
                    ExprNode::ArgList { args } if !args.is_empty() => Some(args.len() as u16),
                    _ => None,
                });
                out.push(BoundVar {
                    sym_id: *name,
                    vba_type: as_array(elem),
                    is_const: false,
                    const_value: None,
                    const_lit: None,
                    fixed_string_len: None,
                    array_dims: dims,
                    is_static: false,
                    is_public: false,
                    is_new: false,
                    name_span: self.spans.get(id),
                });
                names.push(self.type_name_of(*type_node));
            }
        }
        if let ExprNode::DimItem { name, is_const, is_static, is_new, type_node, bounds, init, .. } = self.arena.get(id) {
            let t = self.extract_type(*type_node);
            let vba_type = if bounds.is_some() { as_array(t) } else { t };
            let const_value = if *is_const {
                (*init).and_then(|i| self.eval_const_i64(i))
            } else {
                None
            };
            // Non-integer const initializers (String/Double/Single/Currency/Date/
            // Boolean literals) are carried as the literal itself for the code
            // generator to fold; integer-valued consts use `const_value`.
            let const_lit = if *is_const && const_value.is_none() {
                (*init).and_then(|i| self.eval_const_lit(i))
            } else {
                None
            };
            let fixed_string_len = (*type_node).and_then(|tn| {
                if let ExprNode::StringType { fixed_len: Some(n) } = self.arena.get(tn) {
                    self.eval_const_i64(*n).map(|v| v as u16)
                } else {
                    None
                }
            });
            let array_dims = (*bounds).and_then(|b| match self.arena.get(b) {
                ExprNode::ArgList { args } if !args.is_empty() => Some(args.len() as u16),
                _ => None,
            });
            out.push(BoundVar {
                sym_id: *name, vba_type,
                is_const: *is_const, const_value, const_lit, fixed_string_len, array_dims,
                is_static: *is_static, is_public: false, is_new: *is_new,
                name_span: self.spans.get(id),
            });
            names.push(self.type_name_of(*type_node));
        }
        let mut kids: Vec<NodeId> = Vec::new();
        self.arena.get(id).for_each_child(&mut |c| kids.push(c));
        for c in kids {
            self.collect_locals(proc_idx, c, out, names);
        }
    }

    // ── Name-binding / type-annotation walker ─────────────────────────────────
    //
    // Traversal and analysis are deliberately separated:
    //
    //   * `bind_node` is the *only* place recursion happens. It descends into a
    //     node's children through [`ExprNode::for_each_child`] — the single,
    //     derive-generated source of truth for child edges — so no child can be
    //     silently skipped (a field added to a node is traversed automatically).
    //   * `bind_local` does only *node-local* work (resolve a name, compute a
    //     type) and never recurses.
    //
    // The descent is post-order: every child is fully bound before its parent's
    // `bind_local` runs, so type/resolution computations that read child results
    // (e.g. `Call` return type, `BinOp` operand types) see them ready.

    fn bind_node(&mut self, scope: Option<usize>, id: NodeId) {
        // A type name is not a value use: its qualified tail is never bound as
        // a variable.
        if let ExprNode::UserType { name, child } = self.arena.get(id) {
            let (name, child) = (*name, *child);
            if self.project.is_some() {
                self.resolve_type_position(id, name, child);
            }
            return;
        }
        // ForRange stores its sub-nodes as raw u32 (not NodeId), so the derive-generated
        // for_each_child doesn't visit them. Descend manually before bind_local.
        if let ExprNode::ForRange { loop_var, range, step } = self.arena.get(id) {
            let (lv, r, s) = (*loop_var, *range, *step);
            self.bind_node(scope, NodeId(lv));
            self.bind_node(scope, NodeId(r));
            if s != 0 { self.bind_node(scope, NodeId(s)); }
            self.bind_local(scope, id);
            return;
        }
        if self.project.is_some() {
            // `With obj`: the object is bound first, then it is the base of
            // every leading-dot member in the body.
            if let ExprNode::With { obj, body } = self.arena.get(id) {
                let (obj, body) = (*obj, *body);
                self.bind_node(scope, obj);
                let sem = self.sem_of_node(obj);
                self.with_stack.push(sem);
                self.bind_node(scope, body);
                self.with_stack.pop();
                self.bind_local(scope, id);
                return;
            }
        }
        self.mark_assignment_targets(id);
        let mut kids: Vec<NodeId> = Vec::new();
        self.arena.get(id).for_each_child(&mut |c| kids.push(c));
        for c in kids {
            self.bind_node(scope, c);
        }
        self.bind_local(scope, id);
    }

    /// Record, before the children of an assigning statement are bound, which
    /// accessor each assigned member needs: `Let` for a value assignment,
    /// `Set` for `Set … =`. Only the outermost member of a target is marked
    /// (`a.b.c = 1` assigns `c`; `a` and `b` are read), and an indexed
    /// property target (`o.Item(1) = x`) marks the member being indexed.
    fn mark_assignment_targets(&mut self, id: NodeId) {
        let mut mark: Vec<(NodeId, Access)> = Vec::new();
        match self.arena.get(id) {
            ExprNode::Assign { target, .. } | ExprNode::RangeAssign { target, .. } => {
                mark.push((*target, Access::Let));
            }
            ExprNode::SetAssign { target, .. } => mark.push((*target, Access::Set)),
            ExprNode::MidAssign { args, .. } => {
                if let ExprNode::ArgList { args } = self.arena.get(*args) {
                    if let Some(first) = args.first() {
                        mark.push((*first, Access::Let));
                    }
                }
            }
            ExprNode::For { var, .. } | ExprNode::ForEach { var, .. } => {
                mark.push((*var, Access::Let));
            }
            ExprNode::FileIoStmt { kind, args, .. } => match kind {
                FileIoKind::Input | FileIoKind::LineInput => {
                    mark.extend(args.iter().map(|a| (*a, Access::Let)));
                }
                FileIoKind::Get => {
                    if let Some(last) = args.last() {
                        mark.push((*last, Access::Let));
                    }
                }
                _ => {}
            },
            _ => {}
        }
        for (target, access) in mark {
            let mut node = target;
            if let ExprNode::Call { func, .. } = self.arena.get(node) {
                node = *func;
            }
            if matches!(
                self.arena.get(node),
                ExprNode::MemberAccess { .. } | ExprNode::NameRef { .. }
            ) {
                self.access_ctx.insert(node.0, access);
            }
        }
    }

    /// Node-local binding for `id`, assuming its children are already bound.
    /// Performs name resolution and type annotation only — no recursion.
    /// `scope` is `None` at module level, `Some(proc_idx)` inside a procedure.
    fn bind_local(&mut self, scope: Option<usize>, id: NodeId) {
        match self.arena.get(id) {
            ExprNode::Literal { lit } => {
                let t = lit_type(lit);
                self.types.insert(id.0, t);
            }
            ExprNode::Me | ExprNode::Nothing => {
                self.types.insert(id.0, VbaType::Object);
                if let (Some((project, module)), ExprNode::Me) = (self.project, self.arena.get(id)) {
                    self.sem_types.insert(id.0, project.self_type(module));
                }
            }
            ExprNode::WithContext => {
                if self.project.is_some() {
                    let sem = self.with_stack.last().cloned().unwrap_or(SemType::Unresolved);
                    self.sem_types.insert(id.0, sem);
                }
            }

            // Name reference — the core of the binder.
            ExprNode::NameRef { sym, suffix } => {
                let sym_id = *sym;
                let dollar = *suffix == crate::frontend::token::TypeSuffix::String;
                let (res, sem) = if self.project.is_some() {
                    let access = self.access_ctx.get(&id.0).copied().unwrap_or(Access::Get);
                    let (res, mut sem, params) =
                        self.resolve_in_project(scope, sym_id, access, dollar);
                    if let Some(arity) = params {
                        self.callable.insert(id.0, arity);
                    }
                    // A `Variant`/`Object` local with an inferred type.
                    if let NameResolution::Local { local_idx, .. } = &res {
                        match self.inference.get(local_idx) {
                            Some(Inference::Single(t)) if access == Access::Get => {
                                sem = t.clone();
                                self.inferred_nodes.insert(id.0);
                            }
                            Some(Inference::Ambiguous(cands)) if access == Access::Get => {
                                let names = cands.iter().map(|c| self.type_name_of_sem(c)).collect();
                                self.ambiguous_nodes.insert(id.0, names);
                            }
                            _ => {}
                        }
                    }
                    (res, Some(sem))
                } else {
                    (self.resolve(scope, sym_id), None)
                };
                let mut ty = self.type_of_resolution(scope, &res);
                if let Some(SemType::Value(v)) = &sem {
                    ty = v.clone();
                }
                // Apply DefType letter-range inference: if the name has no explicit
                // type (resolves to Variant), look up its first letter in the map.
                if ty == VbaType::Variant {
                    let name = self.name_of(sym_id);
                    if let Some(first) = name.chars().next() {
                        if let Some(mapped) = self.deftype_map.get(&first) {
                            ty = mapped.clone();
                        }
                    }
                }
                self.resolutions.insert(id.0, res);
                self.types.insert(id.0, ty);
                if let Some(sem) = sem {
                    self.sem_types.insert(id.0, sem);
                }
            }

            ExprNode::Paren { inner } => {
                let t = self.types.get(&inner.0).cloned().unwrap_or_default();
                self.types.insert(id.0, t);
                if self.project.is_some() {
                    let sem = self.sem_of_node(*inner);
                    self.sem_types.insert(id.0, sem);
                }
            }
            ExprNode::BinOp { op, lhs, rhs } => {
                let t = binop_type(
                    *op,
                    self.types.get(&lhs.0).unwrap_or(&VbaType::Variant),
                    self.types.get(&rhs.0).unwrap_or(&VbaType::Variant),
                );
                self.types.insert(id.0, t);
            }
            ExprNode::UnOp { operand, .. } => {
                let t = self.types.get(&operand.0).cloned().unwrap_or_default();
                self.types.insert(id.0, t);
            }
            ExprNode::Call { func, .. } => {
                // Return type: the resolution of `func` if it names a proc; a
                // type-conversion intrinsic returns (and is classified as) its
                // target type.
                let ret = if let ExprNode::NameRef { sym, .. } = self.arena.get(*func) {
                    let sym = *sym;
                    match self.resolutions.get(&func.0).cloned() {
                        Some(NameResolution::Proc(pi)) => {
                            self.procs.get(pi).map(|p| p.ret_type.clone()).unwrap_or_default()
                        }
                        Some(NameResolution::Builtin) => {
                            let name = self.name_of(sym).to_ascii_lowercase();
                            // Type of this call's single argument (already typed
                            // bottom-up), for intrinsics whose result follows it.
                            let arg_ty = || {
                                if let ExprNode::Call { args, .. } = self.arena.get(id) {
                                    if let ExprNode::ArgList { args } = self.arena.get(*args) {
                                        if let Some(&a) = args.first() {
                                            return self.types.get(&a.0).cloned();
                                        }
                                    }
                                }
                                None
                            };
                            let arg_count = || {
                                if let ExprNode::Call { args, .. } = self.arena.get(id) {
                                    if let ExprNode::ArgList { args } = self.arena.get(*args) {
                                        return args.len();
                                    }
                                }
                                0
                            };
                            if let Some(t) = conversion_intrinsic_type(&name) {
                                self.builtins.insert(id.0, BuiltinCall::Convert(t.clone()));
                                t
                            } else if let Some(k) = unary_intrinsic(&name) {
                                self.builtins.insert(id.0, BuiltinCall::Unary(k));
                                match k {
                                    UnaryIntrinsic::Len => VbaType::Long,
                                    UnaryIntrinsic::Sgn => VbaType::Integer,
                                    // Abs/Int/Fix return their argument type.
                                    _ => arg_ty().unwrap_or(VbaType::Variant),
                                }
                            } else if let Some((a, r)) = rtc_numeric_intrinsic(&name, arg_ty()) {
                                self.builtins.insert(id.0, BuiltinCall::RtcNumeric { arg: a, ret: r.clone() });
                                r
                            } else if let Some(args) = rtc_string_intrinsic(&name, arg_count()) {
                                self.builtins.insert(id.0, BuiltinCall::RtcString { args });
                                VbaType::String
                            } else if matches!(name.as_str(), "instr")
                                && (arg_count() == 2 || arg_count() == 3)
                            {
                                self.builtins
                                    .insert(id.0, BuiltinCall::Instr { three_arg: arg_count() == 3 });
                                VbaType::Long
                            } else {
                                VbaType::Variant
                            }
                        }
                        _ => VbaType::Variant,
                    }
                } else {
                    VbaType::Variant
                };
                if self.project.is_some() {
                    let sem = self.call_sem(id, &ret);
                    self.sem_types.insert(id.0, sem);
                    if let ExprNode::Call { func, .. } = self.arena.get(id) {
                        if self.inferred_nodes.contains(&func.0) {
                            self.inferred_nodes.insert(id.0);
                        }
                    }
                }
                self.types.insert(id.0, ret);
            }
            ExprNode::MemberAccess { base, member, bang } => {
                let base_id = *base;
                let member_sym = *member;
                if self.project.is_some() {
                    let access = self.access_ctx.get(&id.0).copied().unwrap_or(Access::Get);
                    self.resolve_member_site(id, base_id, member_sym, *bang, access);
                } else if !*bang && matches!(self.arena.get(base_id), ExprNode::Me) {
                    let access = self.access_ctx.get(&id.0).copied().unwrap_or(Access::Get);
                    if let Some(res) = self.resolve_me_member(member_sym, access) {
                        self.resolutions.insert(id.0, res);
                    }
                }
                let ty = self.resolve_member_type(id, base_id, member_sym);
                self.types.insert(id.0, ty);
            }
            ExprNode::ReDimItem { base, name, .. } => {
                if self.project.is_some() {
                    match *base {
                        // `ReDim a(n)`: a use of the array variable.
                        None => {
                            let (res, sem, _) =
                                self.resolve_in_project(scope, *name, Access::Let, false);
                            // The `ReDim` that declares the array is the
                            // declaration, not a use of it.
                            let declares = matches!(
                                (&res, scope),
                                (NameResolution::Local { proc_idx, local_idx }, Some(_))
                                    if self.procs[*proc_idx].locals[*local_idx].name_span
                                        == self.spans.get(id)
                            );
                            if res != NameResolution::Unresolved && !declares {
                                self.resolutions.insert(id.0, res);
                            }
                            self.sem_types.insert(id.0, sem);
                        }
                        // `ReDim o.Items(n)`: a member of the base object.
                        Some(b) => {
                            self.resolve_member_site(id, b, *name, false, Access::Let);
                        }
                    }
                }
            }
            // A function pointer is a Long-sized address in VB6.
            ExprNode::AddressOf { .. } => {
                self.types.insert(id.0, VbaType::Long);
            }
            ExprNode::TypeOf { .. } => {
                self.types.insert(id.0, VbaType::Boolean);
            }
            ExprNode::New { type_spec } => {
                let t = self.extract_type(Some(*type_spec));
                self.types.insert(id.0, t);
                if self.project.is_some() {
                    let sem = self.sem_of_type_node(Some(*type_spec));
                    self.sem_types.insert(id.0, sem);
                }
            }
            ExprNode::RangeTo { .. } | ExprNode::Generic { .. } => {
                self.types.insert(id.0, VbaType::Variant);
            }
            ExprNode::Implements { library, name } => {
                if let Some((project, module)) = self.project {
                    let tn = TypeName {
                        qualifier: library.map(|l| self.name_of(l)),
                        name: self.name_of(*name),
                    };
                    let res = project.resolve_type_decl(module, &tn);
                    if res != NameResolution::Unresolved {
                        self.resolutions.insert(id.0, res);
                    }
                }
            }
            ExprNode::Attribute { target, name, values } => {
                let (target, name, first) = (*target, *name, values.first().copied());
                self.record_attribute(target, name, first);
            }

            // Everything else carries no node-local type/resolution work; its
            // children (if any) were already bound by the generic descent.
            _ => {}
        }
    }

    /// Record the module facts a metadata line carries: the module name, the
    /// default-instance and global-namespace flags, and the default member.
    fn record_attribute(&mut self, target: Option<u32>, name: u32, first: Option<NodeId>) {
        let lit = first.and_then(|v| match self.arena.get(v) {
            ExprNode::Literal { lit } => Some(lit.clone()),
            _ => None,
        });
        use crate::frontend::ast::AstLit;
        match (self.name_of(name).as_str(), target, lit) {
            ("vb_name", None, Some(AstLit::Str(s))) => self.vb_name = Some(s.to_string()),
            ("vb_predeclaredid", None, Some(AstLit::Bool(b))) => self.predeclared_id = b,
            ("vb_globalnamespace", None, Some(AstLit::Bool(b))) => self.global_namespace = b,
            ("vb_usermemid" | "vb_varusermemid", Some(t), Some(AstLit::Int(0) | AstLit::Long(0))) => {
                let member = self.name_of(t);
                if !member.is_empty() {
                    self.default_member = Some(member);
                }
            }
            _ => {}
        }
    }

    // ── Name resolution ───────────────────────────────────────────────────────

    /// Resolve `sym_id` in `scope` (`None` = module scope, used for module-level
    /// declaration expressions such as array bounds and `Const`/enum values).
    fn resolve(&self, scope: Option<usize>, sym_id: u32) -> NameResolution {
        let name = self.name_of(sym_id);
        if name.is_empty() {
            return NameResolution::Unresolved;
        }

        // Scope chain, innermost first. Each step returns `Some` on a hit; the
        // first hit wins, matching VB6's lookup order.
        self.resolve_in_proc(scope, &name)
            .or_else(|| self.resolve_module_var(&name))
            .or_else(|| self.resolve_proc(&name))
            .or_else(|| self.resolve_enum_member(&name))
            .or_else(|| self.resolve_builtin(&name))
            .unwrap_or(NameResolution::Unresolved)
    }

    /// Procedure scope: locals, parameters, then the return variable (the proc's
    /// own name). `None` at module scope, which has no enclosing procedure.
    fn resolve_in_proc(&self, scope: Option<usize>, name: &str) -> Option<NameResolution> {
        let proc_idx = scope?;
        let proc = &self.procs[proc_idx];

        // 1. Local variables (innermost scope first)
        for (li, local) in proc.locals.iter().enumerate() {
            if self.name_of(local.sym_id) == name {
                return Some(NameResolution::Local { proc_idx, local_idx: li });
            }
        }

        // 2. Parameters
        for (pi, param) in proc.params.iter().enumerate() {
            if self.name_of(param.sym_id) == name {
                return Some(NameResolution::Param { proc_idx, param_idx: pi });
            }
        }

        // 3. The function/sub return variable (same name as the proc itself)
        if self.name_of(proc.sym_id) == name {
            return Some(NameResolution::Proc(proc_idx));
        }

        None
    }

    /// 4. Module-level variables.
    fn resolve_module_var(&self, name: &str) -> Option<NameResolution> {
        self.module_vars
            .iter()
            .position(|var| self.name_of(var.sym_id) == name)
            .map(NameResolution::ModuleVar)
    }

    /// `Me.member`: a member of this module's own public interface (public or
    /// Friend variables and procedures). Private members are not reachable
    /// through `Me`.
    fn resolve_me_member(&self, member_sym: u32, access: Access) -> Option<NameResolution> {
        let name = self.name_of(member_sym);
        if name.is_empty() {
            return None;
        }
        self.module_vars
            .iter()
            .position(|v| v.is_public && self.name_of(v.sym_id) == name)
            .map(NameResolution::ModuleVar)
            .or_else(|| self.pick_proc(&name, access, true).map(NameResolution::Proc))
    }

    /// The procedure named `name` that serves `access`: among the accessors of
    /// a property, `Get` for a read, `Let` / `Set` for an assignment. A
    /// `Function` or `Sub` serves a read. With no accessor of the wanted kind,
    /// the first declaration of that name is returned, so a bare use still
    /// resolves; `public_only` restricts to public / Friend members.
    fn pick_proc(&self, name: &str, access: Access, public_only: bool) -> Option<usize> {
        let named: Vec<usize> = self
            .procs
            .iter()
            .enumerate()
            .filter(|(_, p)| (!public_only || p.is_public) && self.name_of(p.sym_id) == name)
            .map(|(i, _)| i)
            .collect();
        let wanted = |k: ProcKind| match access {
            Access::Get => matches!(k, ProcKind::Sub | ProcKind::Function | ProcKind::PropGet),
            Access::Let => k == ProcKind::PropLet,
            Access::Set => k == ProcKind::PropSet,
        };
        named
            .iter()
            .copied()
            .find(|&i| wanted(self.procs[i].kind))
            .or_else(|| named.first().copied())
    }

    /// 5. Other procs in this module.
    fn resolve_proc(&self, name: &str) -> Option<NameResolution> {
        self.procs
            .iter()
            .position(|p| self.name_of(p.sym_id) == name)
            .map(NameResolution::Proc)
    }

    /// 6. Enum members.
    fn resolve_enum_member(&self, name: &str) -> Option<NameResolution> {
        for (ei, e) in self.enum_decls.iter().enumerate() {
            for (mi, m) in e.members.iter().enumerate() {
                if self.name_of(m.sym_id) == name {
                    return Some(NameResolution::EnumMember { enum_idx: ei, member_idx: mi });
                }
            }
        }
        None
    }

    /// 7. Known built-ins.
    fn resolve_builtin(&self, name: &str) -> Option<NameResolution> {
        is_builtin(name).then_some(NameResolution::Builtin)
    }

    // ── Project-mode typing and resolution ────────────────────────────────────

    /// Static type of a declared type: `vba_type` is the declared `VbaType`,
    /// `name` the named type behind it (if any).
    fn sem_of_declared(&self, vba_type: &VbaType, name: Option<&TypeName>) -> SemType {
        match vba_type {
            VbaType::Array(inner) => SemType::Array(Box::new(self.sem_of_declared(inner, name))),
            VbaType::Variant | VbaType::Object | VbaType::Any => SemType::LateBound,
            VbaType::UserDefined(sym) => match (self.project, name) {
                (Some((project, module)), Some(tn)) => project.sem_of_type_name(module, tn),
                _ => SemType::Unknown(self.name_of(*sym)),
            },
            other => SemType::Value(other.clone()),
        }
    }

    /// Static type named by a type-spec node (`New T`, `TypeOf … Is T`).
    fn sem_of_type_node(&self, type_node: Option<NodeId>) -> SemType {
        let ty = self.extract_type(type_node);
        let name = self.type_name_of(type_node);
        self.sem_of_declared(&ty, name.as_ref())
    }

    /// Record what a type name in a declaration (`As Foo`, `As Proj.Foo`,
    /// `New Foo`) refers to, so that the type's uses are references: the
    /// `UserType` node (spanning the first name) takes the type, or the
    /// qualifier when qualified, and the qualified tail takes the type.
    fn resolve_type_position(&mut self, id: NodeId, name_sym: u32, child: Option<NodeId>) {
        let Some((project, module)) = self.project else { return };
        let first = self.name_of(name_sym);
        if first.is_empty() {
            return;
        }
        let tail = child.and_then(|c| match self.arena.get(c) {
            ExprNode::NameRef { sym, .. } => Some((c, self.name_of(*sym))),
            _ => None,
        });
        match tail {
            Some((tail_id, last)) => {
                let q = project.resolve_qualifier(module, &first);
                if q != NameResolution::Unresolved {
                    self.resolutions.insert(id.0, q);
                }
                if !last.is_empty() {
                    let tn = TypeName { qualifier: Some(first), name: last };
                    let r = project.resolve_type_decl(module, &tn);
                    if r != NameResolution::Unresolved {
                        self.resolutions.insert(tail_id.0, r);
                    }
                }
            }
            None => {
                let r = project.resolve_type_decl(module, &TypeName { qualifier: None, name: first });
                if r != NameResolution::Unresolved {
                    self.resolutions.insert(id.0, r);
                }
            }
        }
    }

    /// Compute the static type of every module-level declaration, once the
    /// module's declarations are collected.
    fn compute_declaration_types(&mut self) {
        let var_sem = self
            .module_vars
            .iter()
            .enumerate()
            .map(|(i, v)| {
                self.sem_of_declared(&v.vba_type, self.decl_types.module_vars.get(i).and_then(|n| n.as_ref()))
            })
            .collect();
        let ret_sem = self
            .procs
            .iter()
            .enumerate()
            .map(|(i, p)| {
                self.sem_of_declared(&p.ret_type, self.decl_types.proc_rets.get(i).and_then(|n| n.as_ref()))
            })
            .collect();
        let param_sem = self
            .procs
            .iter()
            .enumerate()
            .map(|(i, p)| {
                p.params
                    .iter()
                    .enumerate()
                    .map(|(j, prm)| {
                        let name = self.param_names.get(i).and_then(|n| n.get(j)).and_then(|n| n.as_ref());
                        self.sem_of_declared(&prm.vba_type, name)
                    })
                    .collect()
            })
            .collect();
        self.var_sem = var_sem;
        self.ret_sem = ret_sem;
        self.param_sem = param_sem;
        self.local_sem = vec![Vec::new(); self.procs.len()];
    }

    /// The static type recorded for `id`, falling back to the value type the
    /// untyped table carries (literals, operators).
    fn sem_of_node(&self, id: NodeId) -> SemType {
        if let Some(sem) = self.sem_types.get(&id.0) {
            return sem.clone();
        }
        match self.types.get(&id.0) {
            Some(VbaType::Variant | VbaType::Object | VbaType::Any) | None => SemType::LateBound,
            Some(VbaType::UserDefined(_)) => SemType::Unresolved,
            Some(VbaType::Array(_)) => SemType::Unresolved,
            Some(other) => SemType::Value(other.clone()),
        }
    }

    /// Resolve a bare name in project mode. Returns the resolution, the static
    /// type of the name and, when it is a procedure, its parameter count.
    ///
    /// Scope chain, innermost first: locals, parameters, the procedure's own
    /// name, module variables, procedures, the module's own scope (controls),
    /// enum members, enum names, project-wide names, then built-ins.
    fn resolve_in_project(
        &self,
        scope: Option<usize>,
        sym_id: u32,
        access: Access,
        dollar: bool,
    ) -> (NameResolution, SemType, Option<Arity>) {
        let unresolved = (NameResolution::Unresolved, SemType::Unresolved, None);
        let name = self.name_of(sym_id);
        let Some((project, module)) = self.project else { return unresolved };
        if name.is_empty() {
            return unresolved;
        }
        if let Some(proc_idx) = scope {
            let proc = &self.procs[proc_idx];
            for (li, local) in proc.locals.iter().enumerate() {
                if self.name_of(local.sym_id) == name {
                    let sem = self.local_sem[proc_idx].get(li).cloned().unwrap_or(SemType::LateBound);
                    return (NameResolution::Local { proc_idx, local_idx: li }, sem, None);
                }
            }
            for (pi, param) in proc.params.iter().enumerate() {
                if self.name_of(param.sym_id) == name {
                    let sem = self.param_sem[proc_idx].get(pi).cloned().unwrap_or(SemType::LateBound);
                    return (NameResolution::Param { proc_idx, param_idx: pi }, sem, None);
                }
            }
            if self.name_of(proc.sym_id) == name {
                let arity = Arity::Known(proc.params.len());
                return (NameResolution::Proc(proc_idx), self.ret_sem[proc_idx].clone(), Some(arity));
            }
        }
        if let Some(vi) = self.module_vars.iter().position(|v| self.name_of(v.sym_id) == name) {
            return (NameResolution::ModuleVar(vi), self.var_sem[vi].clone(), None);
        }
        if let Some(pi) = self.pick_proc(&name, access, false) {
            let arity = Arity::Known(self.procs[pi].params.len());
            return (NameResolution::Proc(pi), self.ret_sem[pi].clone(), Some(arity));
        }
        if let Some(hit) = project.lookup_module_scope(module, &name, access) {
            return (hit.resolution, hit.sem, hit.callable);
        }
        if let Some(res) = self.resolve_enum_member(&name) {
            return (res, SemType::Value(VbaType::Long), None);
        }
        if let Some(idx) = self.enum_decls.iter().position(|e| self.name_of(e.sym_id) == name) {
            return (
                NameResolution::External { module, decl: ExternalDecl::Enum(idx) },
                SemType::Namespace(Qualifier::Enum { module, idx }),
                None,
            );
        }
        // `Left$(…)` is the `String` form of `Left`.
        if dollar {
            if let Some(hit) = project.lookup_global(module, &format!("{name}$")) {
                return (hit.resolution, hit.sem, hit.callable);
            }
        }
        if let Some(hit) = project.lookup_global(module, &name) {
            return (hit.resolution, hit.sem, hit.callable);
        }
        if is_builtin(&name) {
            return (NameResolution::Builtin, SemType::LibraryNotLoaded("VBA".to_string()), None);
        }
        unresolved
    }

    /// Resolve `base.member` (or `base!member`, or `ReDim base.member(…)`) at
    /// `site`: record the declaration it names, its static type, and a status
    /// that accounts for the site either way. Reports "Method or data member
    /// not found" when the base's class is known and has no such member.
    fn resolve_member_site(
        &mut self,
        site: NodeId,
        base: NodeId,
        member_sym: u32,
        bang: bool,
        access: Access,
    ) {
        let Some((project, module)) = self.project else { return };
        let name = self.name_of(member_sym);
        let base_sem = self.sem_of_node(base);
        let mut result = SemType::Unresolved;
        let base_inferred = self.inferred_nodes.contains(&base.0);
        let ambiguous = self.ambiguous_nodes.get(&base.0).cloned();
        let status = if let (false, Some(cands)) = (bang, ambiguous) {
            result = SemType::LateBound;
            MemberStatus::Ambiguous(cands)
        } else if bang {
            result = SemType::LateBound;
            MemberStatus::Bang
        } else if name.is_empty() {
            MemberStatus::BaseUnresolved
        } else {
            match base_sem {
                SemType::LateBound => {
                    result = SemType::LateBound;
                    MemberStatus::LateBound
                }
                SemType::Unresolved => MemberStatus::BaseUnresolved,
                // A member of a value whose type is not in the project is a
                // value of unknown type for the same reason, so a chain
                // reports where the knowledge ran out.
                SemType::Unknown(n) => {
                    result = SemType::Unknown(n.clone());
                    MemberStatus::UnknownType(n)
                }
                SemType::LibraryNotLoaded(n) => {
                    result = SemType::LibraryNotLoaded(n.clone());
                    MemberStatus::LibraryNotLoaded(n)
                }
                SemType::Value(_) | SemType::Array(_) => MemberStatus::NotAnObject,
                ref ty @ (SemType::Udt { .. }
                | SemType::Class { .. }
                | SemType::Lib { .. }
                | SemType::Namespace(_)) => match project.lookup_member(module, ty, &name, access) {
                    MemberLookup::Found(MemberHit { resolution, sem, callable }) => {
                        if let Some(arity) = callable {
                            self.callable.insert(site.0, arity);
                        }
                        self.resolutions.insert(site.0, resolution);
                        result = sem;
                        if base_inferred {
                            self.inferred_nodes.insert(site.0);
                            MemberStatus::Inferred
                        } else {
                            MemberStatus::Resolved
                        }
                    }
                    MemberLookup::NoSuchMember => {
                        let span = self.spans.get(site);
                        self.diagnostics.push(ERR_METHOD_OR_DATA_MEMBER_NOT_FOUND as u32, span);
                        MemberStatus::NoSuchMember
                    }
                    MemberLookup::MissingAccessor(a) => MemberStatus::MissingAccessor(a),
                    MemberLookup::LibraryNotLoaded(n) => MemberStatus::LibraryNotLoaded(n),
                    MemberLookup::LateBound => {
                        result = SemType::LateBound;
                        MemberStatus::LateBound
                    }
                },
            }
        };
        self.member_status.insert(site.0, status);
        self.sem_types.insert(site.0, result);
    }

    /// How the project writes the type `ty`.
    fn type_name_of_sem(&self, ty: &SemType) -> String {
        match self.project {
            Some((project, _)) => project.type_name(ty),
            None => String::new(),
        }
    }

    /// Static type of the call or index expression `id`. `builtin_ret` is the
    /// value type already derived for a built-in callee.
    fn call_sem(&self, id: NodeId, builtin_ret: &VbaType) -> SemType {
        let ExprNode::Call { func, args } = self.arena.get(id) else { return SemType::Unresolved };
        let arg_count = match self.arena.get(*args) {
            ExprNode::ArgList { args } => args.len(),
            _ => 0,
        };
        if matches!(self.resolutions.get(&func.0), Some(NameResolution::Builtin)) {
            return match builtin_ret {
                VbaType::Variant | VbaType::Object | VbaType::Any => SemType::LateBound,
                other => SemType::Value(other.clone()),
            };
        }
        let callee = self.sem_of_node(*func);
        let default_member = |ty: &SemType| -> SemType {
            let Some((project, module)) = self.project else { return SemType::Unresolved };
            match ty {
                SemType::Class { .. } | SemType::Lib { .. } => {
                    // `rs(0)` is `rs.Fields(0)`: when the default member takes no arguments, the
                    // arguments index the object it returns, through that object's default member.
                    let mut current = ty.clone();
                    let mut result = SemType::Unresolved;
                    for _ in 0..4 {
                        let Some(hit) = project.default_member(module, &current) else {
                            return result;
                        };
                        result = hit.sem.clone();
                        let takes_no_arguments = matches!(hit.callable, Some(Arity::Known(0)));
                        if arg_count == 0 || !takes_no_arguments
                            || !matches!(result, SemType::Class { .. } | SemType::Lib { .. })
                        {
                            break;
                        }
                        current = result.clone();
                    }
                    result
                }
                // Indexing or calling a value of a type outside the project
                // yields a value of that same unknown kind.
                SemType::LateBound | SemType::Unknown(_) | SemType::LibraryNotLoaded(_) => ty.clone(),
                _ => SemType::Unresolved,
            }
        };
        match self.callable.get(&func.0) {
            // A procedure: the call's value is its declared result. A
            // procedure without parameters that returns an array (or an
            // object with a default member) is indexed by the arguments.
            Some(Arity::Known(0)) if arg_count > 0 => match callee {
                SemType::Array(elem) => *elem,
                other => default_member(&other),
            },
            Some(_) => callee,
            None => match callee {
                SemType::Array(elem) => *elem,
                other => default_member(&other),
            },
        }
    }

    fn type_of_resolution(&self, _scope: Option<usize>, res: &NameResolution) -> VbaType {
        match res {
            NameResolution::Local { proc_idx: pi, local_idx: li } => {
                self.procs.get(*pi)
                    .and_then(|p| p.locals.get(*li))
                    .map(|v| v.vba_type.clone())
                    .unwrap_or_default()
            }
            NameResolution::Param { proc_idx: pi, param_idx: pi2 } => {
                self.procs.get(*pi)
                    .and_then(|p| p.params.get(*pi2))
                    .map(|param| param.vba_type.clone())
                    .unwrap_or_default()
            }
            NameResolution::ModuleVar(vi) => {
                self.module_vars.get(*vi)
                    .map(|v| v.vba_type.clone())
                    .unwrap_or_default()
            }
            NameResolution::Proc(pi) => {
                self.procs.get(*pi)
                    .map(|p| p.ret_type.clone())
                    .unwrap_or_default()
            }
            NameResolution::EnumMember { .. } => VbaType::Long,
            // `External` is produced only by the project-level cross-module pass,
            // never by single-module bind; its type is computed there.
            NameResolution::Builtin
            | NameResolution::External { .. }
            | NameResolution::Module(_)
            | NameResolution::Project
            | NameResolution::Library(_)
            | NameResolution::LibraryType { .. }
            | NameResolution::LibraryMember { .. }
            | NameResolution::Unresolved => VbaType::Variant,
        }
    }

    /// Resolve the type of `base.member_sym` by looking up the UDT whose sym_id
    /// matches `base`'s resolved type. Returns `Variant` when the type or member
    /// is unknown (e.g. the base is `Object`, or the type isn't a local UDT).
    ///
    /// As a side effect, when `base` is a known external class instance, also
    /// resolves this access SITE's (`node_id`'s) vtable dispatch slots into
    /// `class_member_slots` — see `BoundModule::class_member_slots`.
    fn resolve_member_type(&self, node_id: NodeId, base: NodeId, member_sym: u32) -> VbaType {
        let base_ty = self.types.get(&base.0).unwrap_or(&VbaType::Variant);
        let VbaType::UserDefined(type_sym) = base_ty else {
            return VbaType::Variant;
        };
        let member_name = self.name_of(member_sym);
        if member_name.is_empty() {
            return VbaType::Variant;
        }
        // Find the UDT declaration in this module whose name matches.
        for decl in &self.type_decls {
            if decl.sym_id == *type_sym {
                for mem in &decl.members {
                    if self.name_of(mem.sym_id) == member_name {
                        return mem.vba_type.clone();
                    }
                }
                return VbaType::Variant;
            }
        }
        // Not a same-module UDT — try a known external class (Dim o As New
        // ClassName; o.Field), matched by name text (cross-module, so sym_ids
        // aren't comparable — see `bind_with_classes`).
        let type_name = self.name_of(*type_sym);
        if let Some(class) = self.known_classes.get(&type_name) {
            if let Some(resolved) = resolve_class_member_slots(class, &member_name) {
                self.class_member_slots
                    .borrow_mut()
                    .insert(node_id.0, resolved);
            }
            for member in &class.members {
                match member {
                    ClassMemberSlot::Field { name, vba_type, .. }
                        if name.eq_ignore_ascii_case(&member_name) =>
                    {
                        return vba_type.clone();
                    }
                    ClassMemberSlot::PropertyAccessor { name, vba_type, .. }
                        if name.eq_ignore_ascii_case(&member_name) =>
                    {
                        return vba_type.clone();
                    }
                    ClassMemberSlot::Method { name, ret_type, .. }
                        if name.eq_ignore_ascii_case(&member_name) =>
                    {
                        return ret_type.clone();
                    }
                    _ => {}
                }
            }
        }
        VbaType::Variant
    }
}

/// The class vtable's fixed IDispatch prefix width: 7 standard COM methods
/// (QueryInterface, AddRef, Release, GetTypeInfoCount, GetTypeInfo,
/// GetIDsOfNames, Invoke) × 4 bytes. A class's first member's own dispatch
/// slots start here. This is the observed base of the first member's slots in
/// every class-member fixture, confirmed byte-for-byte against the real VB6
/// compiler's output.
const IDISPATCH_PREFIX: u16 = 0x1c;

/// Resolve `member_name`'s vtable dispatch slots within `class`, by walking
/// `class.members` in declaration order and assigning each entry its slot
/// width at stride 4 from the base `0x1c` (see `ExternalClass::members`'s
/// ordering invariant): a scalar field takes a Get and a Let slot, an
/// object-typed field additionally a Set slot, a single property accessor or
/// a `Sub`/`Function` one slot each; non-exposed (`Private`) members are not
/// in the list and consume no slot. This closed form reproduces the real VB6
/// compiler's emitted `0x0d <slot>` operands byte-for-byte across every
/// class-member fixture — scalar fields of mixed types, an object field
/// interleaved between scalars, `Get`-only / `Let`-only / `Get`+`Let`+`Set`
/// properties, interleaved `Sub`/`Function`/property members, a skipped
/// `Private` field, and declaration-order permutations (access order does not
/// affect the value). Returns `None` if no member of that name exists.
fn resolve_class_member_slots(
    class: &ExternalClass,
    member_name: &str,
) -> Option<ResolvedClassMember> {
    let mut next = IDISPATCH_PREFIX;
    let mut get_slot = None;
    let mut let_slot = None;
    let mut set_slot = None;
    let mut method_slot = None;
    let mut method_ret_type = None;
    let mut method_params = Vec::new();
    let mut is_property = false;
    let mut found = false;
    for member in &class.members {
        match member {
            ClassMemberSlot::Field {
                name, is_object, ..
            } => {
                let g = next;
                next += 4;
                let l = next;
                next += 4;
                let s = if *is_object {
                    let v = next;
                    next += 4;
                    Some(v)
                } else {
                    None
                };
                if name.eq_ignore_ascii_case(member_name) {
                    get_slot = Some(g);
                    let_slot = Some(l);
                    set_slot = s;
                    is_property = false;
                    found = true;
                }
            }
            ClassMemberSlot::PropertyAccessor { name, kind, .. } => {
                let slot = next;
                next += 4;
                if name.eq_ignore_ascii_case(member_name) {
                    match kind {
                        AccessorKind::Get => get_slot = Some(slot),
                        AccessorKind::Let => let_slot = Some(slot),
                        AccessorKind::Set => set_slot = Some(slot),
                    }
                    is_property = true;
                    found = true;
                }
            }
            ClassMemberSlot::Method { name, ret_type, params } => {
                let slot = next;
                next += 4;
                if name.eq_ignore_ascii_case(member_name) {
                    method_slot = Some(slot);
                    method_ret_type = Some(ret_type.clone());
                    method_params = params.clone();
                    found = true;
                }
            }
        }
    }
    found.then_some(ResolvedClassMember {
        get_slot,
        let_slot,
        set_slot,
        method_slot,
        method_ret_type,
        method_params,
        is_property,
    })
}

// ── Type inference helpers ────────────────────────────────────────────────────

/// `t` as the type of a declaration with array bounds: an array of `t`, unless
/// the declared type is already an array (`Dim a(3) As Long()` is one array).
fn as_array(t: VbaType) -> VbaType {
    match t {
        VbaType::Array(_) => t,
        other => VbaType::Array(Box::new(other)),
    }
}

fn lit_type(lit: &crate::frontend::ast::AstLit) -> VbaType {
    use crate::frontend::ast::AstLit::*;
    match lit {
        Int(_)      => VbaType::Integer,
        Long(_)     => VbaType::Long,
        Single(_)   => VbaType::Single,
        Double(_)   => VbaType::Double,
        Currency(_) => VbaType::Currency,
        Str(_)      => VbaType::String,
        Date(_)     => VbaType::Date,
        Bool(_)     => VbaType::Boolean,
        Empty       => VbaType::Variant,
        Null        => VbaType::Variant,
    }
}

/// Target type of a type-conversion intrinsic that reuses the assignment-
/// conversion opcodes. `CBool`/`CDate`/`CVar` use special conversions and are
/// intentionally excluded here.
fn conversion_intrinsic_type(name_lower: &str) -> Option<VbaType> {
    Some(match name_lower {
        "cint" => VbaType::Integer,
        "clng" => VbaType::Long,
        "cbyte" => VbaType::Byte,
        "csng" => VbaType::Single,
        "cdbl" => VbaType::Double,
        "ccur" => VbaType::Currency,
        "cstr" => VbaType::String,
        _ => return None,
    })
}

/// Classify a single-argument numeric-result runtime-library intrinsic by name,
/// returning `(argument type, result type)`. `Asc` takes a String → Integer; `Sqr`
/// a Double → Double; `Val` a String → Double. Other runtime intrinsics (string-
/// returning / multi-argument) are not classified here and are gated.
fn rtc_numeric_intrinsic(name_lower: &str, arg: Option<VbaType>) -> Option<(VbaType, VbaType)> {
    match name_lower {
        "asc" => Some((VbaType::String, VbaType::Integer)),
        "sqr" => Some((arg.unwrap_or(VbaType::Double), VbaType::Double)),
        "val" => Some((VbaType::String, VbaType::Double)),
        _ => None,
    }
}

/// Classify a String-returning runtime intrinsic by name and argument count,
/// returning the passing mode of each parameter (in source order). The result is
/// always a String. Each parameter is either pushed by value ([`RtcArg::ByVal`])
/// or boxed into a hidden 16-byte temp ([`RtcArg::Boxed`]).
///
/// `Mid` with an omitted Optional length (the 2-argument form) needs the
/// Missing-Variant marshalling form and is left unclassified (gated).
fn rtc_string_intrinsic(name_lower: &str, arg_count: usize) -> Option<Vec<RtcArg>> {
    use RtcArg::{ByVal, Boxed, MissingVariant};
    Some(match name_lower {
        // Numeric count → String: the count is pushed by value as a Long.
        "chr" | "chr$" | "space" | "space$" => vec![ByVal(VbaType::Long)],
        // String → String: the argument is boxed.
        "ucase" | "ucase$" | "lcase" | "lcase$" | "trim" | "trim$"
        | "ltrim" | "ltrim$" | "rtrim" | "rtrim$" => vec![Boxed],
        // Number → String, boxed as a Variant argument.
        "str" | "str$" | "hex" | "hex$" | "oct" | "oct$" => vec![Boxed],
        // (String, length As Long): the string is boxed, the length by value.
        "left" | "left$" | "right" | "right$" if arg_count == 2 => {
            vec![Boxed, ByVal(VbaType::Long)]
        }
        // (number As Long, character As String).
        "string" | "string$" if arg_count == 2 => vec![ByVal(VbaType::Long), Boxed],
        // Mid(String, start As Long, length As Variant). With the optional length
        // omitted (2-argument form) a Missing variant is materialised for it.
        "mid" | "mid$" if arg_count == 3 => vec![Boxed, ByVal(VbaType::Long), Boxed],
        "mid" | "mid$" if arg_count == 2 => {
            vec![Boxed, ByVal(VbaType::Long), MissingVariant]
        }
        _ => return None,
    })
}

/// Classify a single-argument dedicated-opcode intrinsic by name.
fn unary_intrinsic(name_lower: &str) -> Option<UnaryIntrinsic> {
    Some(match name_lower {
        "len" => UnaryIntrinsic::Len,
        "abs" => UnaryIntrinsic::Abs,
        "sgn" => UnaryIntrinsic::Sgn,
        "int" => UnaryIntrinsic::Int,
        "fix" => UnaryIntrinsic::Fix,
        _ => return None,
    })
}

fn binop_type(
    op: crate::frontend::ast::BinOpKind,
    lhs: &VbaType,
    rhs: &VbaType,
) -> VbaType {
    use crate::frontend::ast::BinOpKind as B;
    match op {
        // Comparison ops always return Boolean.
        B::Eq | B::Ne | B::Lt | B::Gt | B::Le | B::Ge
        | B::Like | B::Is
            => VbaType::Boolean,
        // And/Or/Xor/Eqv/Imp are bitwise: the back-end (EbEmitBinaryOperation2
        // @0fab2e1e) selects their opcode from the bound node's *own* type tag on
        // the same arithmetic dispatch path as Add/Sub/Mul (RT_DISPATCH_FLAG bit
        // 0x10 clear → RT_TYPE_OFFSET[node.type_tag]). The node therefore carries
        // the operand-promoted type, not Boolean (Long Xor Long → Long).
        B::And | B::Or | B::Xor | B::Eqv | B::Imp => numeric_promote(lhs, rhs),
        // String concatenation always returns String
        B::Cat => VbaType::String,
        // Integer division and modulo round each operand to an integer type and
        // return that promoted integer type: Byte\Byte → Byte, Integer\Integer →
        // Integer, Long\Long → Long. Floating-point / Currency / Date / Variant
        // operands are rounded to Long.
        B::IDiv | B::Mod => match numeric_promote(lhs, rhs) {
            t @ (VbaType::Byte | VbaType::Integer | VbaType::Long | VbaType::Boolean) => t,
            _ => VbaType::Long,
        },
        // Exponentiation always computes and returns Double, regardless of operand
        // type (unlike the other arithmetic ops, which promote to the operands'
        // common type).
        B::Pow => VbaType::Double,
        // Arithmetic: numeric promotion. Date operands are computed in Double (the
        // OLE serial is widened; the Date result type is restored on store), so an
        // arithmetic expression never has type Date.
        B::Add | B::Sub | B::Mul | B::Div => {
            match numeric_promote(lhs, rhs) {
                VbaType::Date => VbaType::Double,
                // Floating division (`/`) of Currency is performed in Double; the
                // result is Double (other Currency arithmetic stays Currency).
                VbaType::Currency if matches!(op, B::Div) => VbaType::Double,
                other => other,
            }
        }
        // Member access: Variant (resolved at runtime)
        B::Dot | B::Bang => VbaType::Variant,
    }
}

/// Map a `DefType` keyword value (the numeric `Kw` discriminant) to its `VbaType`.
/// Kw values: DefBool=49, DefByte=50, DefCur=51, DefDate=52, DefDec=53, DefDbl=54,
///            DefInt=55, DefLng=56, DefObj=57, DefSng=58, DefStr=59, DefVar=60.
fn deftype_kw_to_vbatype(kw: u16) -> VbaType {
    match kw {
        49 => VbaType::Boolean,
        50 => VbaType::Byte,
        51 => VbaType::Currency,
        52 => VbaType::Date,
        53 => VbaType::Decimal,
        54 => VbaType::Double,
        55 => VbaType::Integer,
        56 => VbaType::Long,
        57 => VbaType::Object,
        58 => VbaType::Single,
        59 => VbaType::String,
        _  => VbaType::Variant, // DefVar (60) and unknown
    }
}

fn numeric_promote(a: &VbaType, b: &VbaType) -> VbaType {
    // VB6 numeric promotion order (lowest to highest):
    // Boolean < Integer < Long < Single < Double < Currency < Variant
    let rank = |t: &VbaType| -> u8 {
        match t {
            VbaType::Boolean  => 0,
            VbaType::Byte     => 1,
            VbaType::Integer  => 2,
            VbaType::Long     => 3,
            VbaType::Single   => 4,
            VbaType::Double   => 5,
            VbaType::Currency => 6,
            VbaType::Date     => 5,
            _                 => 7,
        }
    };
    if rank(a) >= rank(b) { a.clone() } else { b.clone() }
}
