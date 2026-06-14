//! Compiler context (`CompilerContext`) — the central state object for
//! one compilation unit.
//!
//! It is threaded through every compiler pass.  Here it is an owned Rust struct
//! that holds the three node arenas and the predefined-module nodes.

use std::collections::HashMap;

use crate::frontend::ast::{
    alloc_scope_block, ast_node_create, ast_node_link, DeclArena, NodeId, ScopeArena,
    ScopeBlockArena, ScopeNodeKind,
};
use crate::frontend::keyword_table::keyword_string;

/// Token id of the keyword whose string names the second predefined scope.
/// VB6's `build_predefined_nodes` interns `keyword_string(0xba)` (= "Unknown")
/// with symbol kind 4 as that scope's name.
const PREDEFINED_SCOPE_KEYWORD: u16 = 0xba;

/// A name interned via [`CompilerContext::intern_string`]. The interned symbol
/// heads the declaration chain for that name (the `ScopeBlock` passed to
/// `alloc_decl_node`).
struct InternedSymbol {
    /// Canonical (original-case) name text.
    name: String,
    /// Symbol kind flag (2 = module name, 4 = predefined-scope name).
    kind: u32,
    /// The `ScopeBlock` that heads this symbol's declaration chain.
    block: NodeId,
}

/// The per-compilation-unit compiler context.
///
/// Owns the three node arenas and records the predefined built-in module nodes
/// that are created during initialisation.
pub struct CompilerContext {
    /// Arena for scope-container nodes (`ScopeNode`, created by `ast_node_create`).
    pub scope_nodes: ScopeArena,
    /// Arena for declaration/symbol nodes (`DeclNode`).
    pub decls: DeclArena,
    /// Arena for scope-block list heads (`ScopeBlock`).
    pub scopes: ScopeBlockArena,

    /// First predefined scope node.
    pub first_predefined_scope: Option<NodeId>,
    /// ModuleList child of the first predefined scope.
    pub first_predefined_child: Option<NodeId>,
    /// ModuleList child of the second predefined scope.
    pub second_predefined_child: Option<NodeId>,

    /// Name interner: the symbols created for module and predefined-scope
    /// names. Indexed by case-folded name via `interner_by_name`.
    interner: Vec<InternedSymbol>,
    /// Case-folded name → index into `interner`.
    interner_by_name: HashMap<String, usize>,
}

impl CompilerContext {
    /// Create an empty compiler context.
    pub fn new() -> Self {
        Self {
            scope_nodes: ScopeArena::new(),
            decls: DeclArena::new(),
            scopes: ScopeBlockArena::new(),
            first_predefined_scope: None,
            first_predefined_child: None,
            second_predefined_child: None,
            interner: Vec::new(),
            interner_by_name: HashMap::new(),
        }
    }

    /// Intern `name` and return the `ScopeBlock` that heads its declaration
    /// chain (a case-insensitive lookup-or-insert). Re-interning the same name
    /// returns the existing symbol; `kind` is recorded on first insert.
    pub fn intern_string(&mut self, name: &str, kind: u32) -> NodeId {
        let key = name.to_ascii_lowercase();
        if let Some(&idx) = self.interner_by_name.get(&key) {
            return self.interner[idx].block;
        }
        let block = alloc_scope_block(&mut self.scopes);
        let idx = self.interner.len();
        self.interner.push(InternedSymbol { name: name.to_string(), kind, block });
        self.interner_by_name.insert(key, idx);
        block
    }

    /// The `ScopeBlock` heading the declaration chain for an already-interned
    /// `name` (case-insensitive), or `None` if it was never interned.
    pub fn interned_symbol(&self, name: &str) -> Option<NodeId> {
        self.interner_by_name
            .get(&name.to_ascii_lowercase())
            .map(|&idx| self.interner[idx].block)
    }

    /// The symbol-kind flag recorded when `name` was first interned (2 = module
    /// name, 4 = predefined-scope name), or `None` if it was never interned.
    pub fn interned_kind(&self, name: &str) -> Option<u32> {
        self.interner_by_name
            .get(&name.to_ascii_lowercase())
            .map(|&idx| self.interner[idx].kind)
    }

    /// The canonical (original-case) name of the interned symbol whose
    /// declaration chain is headed by `block`, or `None` if `block` is not an
    /// interned symbol.
    pub fn interned_name(&self, block: NodeId) -> Option<&str> {
        self.interner.iter().find(|s| s.block == block).map(|s| s.name.as_str())
    }

    /// Build the predefined module nodes.
    ///
    /// Creates two predefined module-scope nodes and links declaration nodes
    /// to them.  Called once during context initialisation.
    ///
    /// Parameters:
    /// * `outer_scope_list` — scope-block to use for both DeclNode allocations
    ///   (`None` when the outer block is null).
    /// * `parent_decl` — the enclosing standard-module DeclNode.
    ///
    /// Writes to context fields:
    /// * `first_predefined_scope` ← first scope id
    /// * `first_predefined_child` ← first scope's ModuleList child id
    /// * `second_predefined_child` ← second scope's ModuleList child id
    pub fn build_predefined_nodes(
        &mut self,
        outer_scope_list: Option<NodeId>,
        parent_decl: NodeId,
    ) {
        let parent_ext_flags = self.decls.get(parent_decl).ext_flags;

        // --- First predefined scope (unnamed) ---
        let scope1 = ast_node_create(&mut self.scope_nodes, ScopeNodeKind::Module);
        self.scope_nodes.get_mut(scope1).flags_b6 |= 4;
        self.first_predefined_scope = Some(scope1);

        let decl1 = ast_node_link(
            &mut self.scope_nodes,
            &mut self.decls,
            &mut self.scopes,
            scope1,
            outer_scope_list,
            Some(parent_decl),
            2,
        );
        self.decls.get_mut(decl1).sec_flags |= 8;
        self.decls.get_mut(decl1).type_info = 5;

        self.scope_nodes.get_mut(scope1).flags_b6 |= 0x80;
        // first_predefined_child = scope1.child (ModuleList child NodeId)
        self.first_predefined_child = self.scope_nodes.get(scope1).child;

        if parent_ext_flags & 1 != 0 {
            self.decls.get_mut(decl1).flags_10 |= 8;
        }

        // --- Second predefined scope (named; keyword 0xba interned as kind 4) ---
        // VB6 names this scope by interning `keyword_string(0xba)` ("Unknown")
        // as a kind-4 symbol and passing that symbol as the decl's name/scope.
        let scope2 = ast_node_create(&mut self.scope_nodes, ScopeNodeKind::Module);
        self.scope_nodes.get_mut(scope2).flags_b6 |= 8;

        let name_sym = self.intern_string(keyword_string(PREDEFINED_SCOPE_KEYWORD), 4);
        let decl2 = ast_node_link(
            &mut self.scope_nodes,
            &mut self.decls,
            &mut self.scopes,
            scope2,
            Some(name_sym),
            Some(parent_decl),
            2,
        );
        self.decls.get_mut(decl2).sec_flags |= 8;
        self.decls.get_mut(decl2).type_info = 5;
        // second_predefined_child = scope2.child (ModuleList child NodeId)
        self.second_predefined_child = self.scope_nodes.get(scope2).child;

        if parent_ext_flags & 1 != 0 {
            self.decls.get_mut(decl2).flags_10 |= 8;
        }
    }
}

impl Default for CompilerContext {
    fn default() -> Self {
        Self::new()
    }
}
