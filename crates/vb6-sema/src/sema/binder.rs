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

use crate::frontend::ast::{ExprArena, ExprNode, NodeId, NodeSpans, ProcKind, Span};
use crate::frontend::diagnostics::Diagnostics;
use crate::frontend::scanner::ScannerContext;
use crate::sema::builtins::is_builtin;
use crate::sema::symbol::{
    BoundEnumDecl, BoundEnumMember, BoundModule, BoundParam, BoundProc, BoundTypeDecl,
    BoundTypeMember, BoundVar, NameResolution, ParamFlags,
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
    let mut b = Binder::new(ctx, arena, spans, visibility);
    b.bind_top_level(top_nodes);
    b.finish()
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
    (0..arena.len() as u32)
        .filter(|i| matches!(arena.get(NodeId(*i)), ExprNode::NameRef { .. }))
        .filter(|i| !resolutions.contains_key(i))
        .map(NodeId)
        .collect()
}

// ── Internal binder state ─────────────────────────────────────────────────────

struct Binder<'a> {
    ctx: &'a ScannerContext,
    arena: &'a ExprArena,
    spans: &'a NodeSpans,
    visibility: &'a HashMap<u32, bool>,

    procs:       Vec<BoundProc>,
    /// Parameter `ArgList` node per proc (aligned with `procs`), kept so the
    /// body pass can bind `Optional` default-value expressions, which live in the
    /// parameter list rather than the body block.
    proc_param_nodes: Vec<Option<NodeId>>,
    module_vars: Vec<BoundVar>,
    type_decls:  Vec<BoundTypeDecl>,
    enum_decls:  Vec<BoundEnumDecl>,

    resolutions: HashMap<u32, NameResolution>,
    types:       HashMap<u32, VbaType>,

    /// `DefType` letter-range map: first character of an untyped name → inferred VbaType.
    deftype_map: HashMap<char, VbaType>,

    diagnostics:     Diagnostics,
    option_explicit: bool,
}

impl<'a> Binder<'a> {
    fn new(
        ctx: &'a ScannerContext,
        arena: &'a ExprArena,
        spans: &'a NodeSpans,
        visibility: &'a HashMap<u32, bool>,
    ) -> Self {
        Self {
            ctx,
            arena,
            spans,
            visibility,
            procs:       Vec::new(),
            proc_param_nodes: Vec::new(),
            module_vars: Vec::new(),
            type_decls:  Vec::new(),
            enum_decls:  Vec::new(),
            resolutions: HashMap::new(),
            types:       HashMap::new(),
            deftype_map: HashMap::new(),
            diagnostics:     Diagnostics::new(),
            option_explicit: false,
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
            diagnostics:     self.diagnostics,
            option_explicit: self.option_explicit,
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
            ExprNode::UserType { name, .. } => VbaType::UserDefined(*name),
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

    // ── Top-level pass ────────────────────────────────────────────────────────

    fn bind_top_level(&mut self, top_nodes: &[NodeId]) {
        // Pass 1: collect all declarations.
        for &id in top_nodes {
            self.collect_top_decl(id);
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
            ExprNode::DimItem { name, is_const, bounds, type_node, .. } => {
                self.collect_top_var(id, *name, *is_const, bounds.is_some(), *type_node);
            }
            ExprNode::TypeDecl { name, members } => {
                let sym_id = *name;
                let members_clone = members.clone();
                let name_span = self.spans.get(id);
                let is_public = self.vis(id, true); // types default Public
                let decl = self.collect_type_decl(sym_id, &members_clone, name_span, is_public);
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
        self.procs.push(proc);
        self.proc_param_nodes.push(params_id);
    }

    /// Collect a module-level variable (`Dim`/`Const` item). Module `Dim`/`Const`
    /// default to Private; array bounds wrap the element type in `Array`.
    fn collect_top_var(
        &mut self,
        id: NodeId,
        sym_id: u32,
        is_const: bool,
        has_bounds: bool,
        type_node: Option<NodeId>,
    ) {
        let var = BoundVar {
            sym_id,
            vba_type: self.extract_type(type_node),
            is_const,
            const_value: None,
            const_lit: None,
            fixed_string_len: None,
            array_dims: None,
            is_static: false,
            is_public: self.vis(id, false),
            name_span: self.spans.get(id),
        };
        let var = if has_bounds {
            BoundVar { vba_type: VbaType::Array(Box::new(var.vba_type)), ..var }
        } else {
            var
        };
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
                        if bounds.is_some() { VbaType::Array(Box::new(t)) } else { t }
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
        self.collect_locals(body_id, &mut locals);
        self.procs[proc_idx].locals = locals;

        // Bind all name references in the body. (`Optional` parameter defaults
        // are bound separately by the caller, since they live in the parameter
        // list rather than the body block.)
        self.bind_node(Some(proc_idx), body_id);
    }

    /// Collect every `DimItem` in a procedure body subtree as a local variable.
    ///
    /// VB6 locals are procedure-scoped regardless of nesting (no block scope), so
    /// every `DimItem` anywhere under the body is a local. Descent goes through
    /// [`ExprNode::for_each_child`] — the same single source of truth the binder
    /// uses — so a local declared inside any statement form (even one added
    /// later) is collected, never silently resolved to the wrong scope. A proc
    /// body contains no nested procedure declarations, so a full descent is safe.
    fn collect_locals(&self, id: NodeId, out: &mut Vec<BoundVar>) {
        if let ExprNode::DimItem { name, is_const, type_node, bounds, init } = self.arena.get(id) {
            let t = self.extract_type(*type_node);
            let vba_type = if bounds.is_some() { VbaType::Array(Box::new(t)) } else { t };
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
                is_static: false, is_public: false,
                name_span: self.spans.get(id),
            });
        }
        let mut kids: Vec<NodeId> = Vec::new();
        self.arena.get(id).for_each_child(&mut |c| kids.push(c));
        for c in kids {
            self.collect_locals(c, out);
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
        let mut kids: Vec<NodeId> = Vec::new();
        self.arena.get(id).for_each_child(&mut |c| kids.push(c));
        for c in kids {
            self.bind_node(scope, c);
        }
        self.bind_local(scope, id);
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
            }

            // Name reference — the core of the binder.
            ExprNode::NameRef { sym, .. } => {
                let sym_id = *sym;
                let res = self.resolve(scope, sym_id);
                let mut ty = self.type_of_resolution(scope, &res);
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
            }

            ExprNode::Paren { inner } => {
                let t = self.types.get(&inner.0).cloned().unwrap_or_default();
                self.types.insert(id.0, t);
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
                // Return type: the resolution of `func` if it names a proc.
                let ret = if matches!(self.arena.get(*func), ExprNode::NameRef { .. }) {
                    match self.resolutions.get(&func.0).cloned() {
                        Some(NameResolution::Proc(pi)) => {
                            self.procs.get(pi).map(|p| p.ret_type.clone()).unwrap_or_default()
                        }
                        _ => VbaType::Variant,
                    }
                } else {
                    VbaType::Variant
                };
                self.types.insert(id.0, ret);
            }
            ExprNode::MemberAccess { base, member, .. } => {
                let base_id = *base;
                let member_sym = *member;
                let ty = self.resolve_member_type(base_id, member_sym);
                self.types.insert(id.0, ty);
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
            }
            ExprNode::RangeTo { .. } | ExprNode::Generic { .. } => {
                self.types.insert(id.0, VbaType::Variant);
            }

            // Everything else carries no node-local type/resolution work; its
            // children (if any) were already bound by the generic descent.
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
            | NameResolution::Unresolved => VbaType::Variant,
        }
    }

    /// Resolve the type of `base.member_sym` by looking up the UDT whose sym_id
    /// matches `base`'s resolved type. Returns `Variant` when the type or member
    /// is unknown (e.g. the base is `Object`, or the type isn't a local UDT).
    fn resolve_member_type(&self, base: NodeId, member_sym: u32) -> VbaType {
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
        VbaType::Variant
    }
}

// ── Type inference helpers ────────────────────────────────────────────────────

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

fn binop_type(
    op: crate::frontend::ast::BinOpKind,
    lhs: &VbaType,
    rhs: &VbaType,
) -> VbaType {
    use crate::frontend::ast::BinOpKind as B;
    match op {
        // Comparison ops always return Boolean.
        B::Eq | B::Ne | B::Lt | B::Gt | B::Le | B::Ge
        | B::Like | B::Is | B::IsNot
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
        // Arithmetic: numeric promotion. Date operands are computed in Double (the
        // OLE serial is widened; the Date result type is restored on store), so an
        // arithmetic expression never has type Date.
        B::Add | B::Sub | B::Mul | B::Div | B::Pow => {
            match numeric_promote(lhs, rhs) {
                VbaType::Date => VbaType::Double,
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
