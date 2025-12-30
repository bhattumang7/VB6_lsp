//! Symbol Table
//!
//! The main symbol table that stores all symbols and scopes for a document.

use std::collections::HashMap;

use tower_lsp::lsp_types::Url;

use super::position::{SourcePosition, SourceRange};
use super::scope::{Scope, ScopeId, ScopeKind};
use super::symbol::{ParameterInfo, Symbol, SymbolId, SymbolKind, TypeInfo, Visibility};

/// A reference to a symbol (usage site)
#[derive(Debug, Clone)]
pub struct SymbolReference {
    /// The symbol being referenced
    pub symbol_id: SymbolId,
    /// The range of this reference
    pub range: SourceRange,
    /// The scope where this reference occurs
    pub scope_id: ScopeId,
    /// Whether this is an assignment target (LHS of =)
    pub is_assignment: bool,
    /// For member access chains, the qualifying reference (e.g., obj in obj.member)
    pub qualifying_reference: Option<Box<SymbolReference>>,
}

/// The complete symbol table for a document
#[derive(Debug)]
pub struct SymbolTable {
    /// Document URI
    pub uri: Url,

    /// All symbols, indexed by ID
    symbols: Vec<Symbol>,

    /// All scopes, indexed by ID
    scopes: Vec<Scope>,

    /// The module-level (root) scope
    pub module_scope: ScopeId,

    /// All references to symbols
    references: Vec<SymbolReference>,

    /// Spatial index: map from line number to symbols defined on that line
    symbols_by_line: HashMap<u32, Vec<SymbolId>>,

    /// Spatial index: map from line number to scopes that contain that line
    scopes_by_line: HashMap<u32, Vec<ScopeId>>,

    /// Next symbol ID to allocate
    next_symbol_id: u32,

    /// Next scope ID to allocate
    next_scope_id: u32,
}

impl SymbolTable {
    /// Create a new empty symbol table for a document
    pub fn new(uri: Url) -> Self {
        let mut table = Self {
            uri,
            symbols: Vec::new(),
            scopes: Vec::new(),
            module_scope: ScopeId(0),
            references: Vec::new(),
            symbols_by_line: HashMap::new(),
            scopes_by_line: HashMap::new(),
            next_symbol_id: 0,
            next_scope_id: 0,
        };

        // Create the module scope (covers entire file)
        let module_range = SourceRange::new(
            SourcePosition::new(0, 0),
            SourcePosition::new(u32::MAX - 1, 0),
        );
        table.module_scope = table.create_scope(ScopeKind::Module, None, module_range);

        table
    }

    // ==========================================
    // Symbol Management
    // ==========================================

    /// Create a new symbol and add it to the table
    pub fn create_symbol(
        &mut self,
        name: String,
        kind: SymbolKind,
        visibility: Visibility,
        definition_range: SourceRange,
        name_range: SourceRange,
        scope_id: ScopeId,
    ) -> SymbolId {
        let id = SymbolId(self.next_symbol_id);
        self.next_symbol_id += 1;

        let symbol = Symbol::new(
            id,
            name.clone(),
            kind,
            visibility,
            definition_range,
            name_range,
            scope_id,
        );

        // Add to spatial index (index by name_range lines for precise lookup)
        for line in name_range.start.line..=name_range.end.line {
            self.symbols_by_line.entry(line).or_default().push(id);
        }

        // Add to scope
        if let Some(scope) = self.scopes.get_mut(scope_id.0 as usize) {
            scope.add_symbol(&name, id);
        }

        self.symbols.push(symbol);
        id
    }

    /// Get a symbol by ID
    pub fn get_symbol(&self, id: SymbolId) -> Option<&Symbol> {
        self.symbols.get(id.0 as usize)
    }

    /// Get a mutable symbol by ID
    pub fn get_symbol_mut(&mut self, id: SymbolId) -> Option<&mut Symbol> {
        self.symbols.get_mut(id.0 as usize)
    }

    /// Set type info for a symbol
    pub fn set_type_info(&mut self, id: SymbolId, type_info: TypeInfo) {
        if let Some(symbol) = self.get_symbol_mut(id) {
            symbol.type_info = Some(type_info);
        }
    }

    /// Set value for a symbol (constants, enum members)
    pub fn set_value(&mut self, id: SymbolId, value: String) {
        if let Some(symbol) = self.get_symbol_mut(id) {
            symbol.value = Some(value);
        }
    }

    /// Add parameters to a procedure symbol
    pub fn set_parameters(&mut self, id: SymbolId, parameters: Vec<ParameterInfo>) {
        if let Some(symbol) = self.get_symbol_mut(id) {
            symbol.parameters = parameters;
        }
    }

    /// Add a member to a type/enum symbol
    pub fn add_member(&mut self, parent_id: SymbolId, member_id: SymbolId) {
        if let Some(symbol) = self.get_symbol_mut(parent_id) {
            symbol.members.push(member_id);
        }
    }

    /// Set documentation for a symbol
    pub fn set_documentation(&mut self, id: SymbolId, doc: String) {
        if let Some(symbol) = self.get_symbol_mut(id) {
            symbol.documentation = Some(doc);
        }
    }

    // ==========================================
    // Scope Management
    // ==========================================

    /// Create a new scope
    pub fn create_scope(
        &mut self,
        kind: ScopeKind,
        parent: Option<ScopeId>,
        range: SourceRange,
    ) -> ScopeId {
        let id = ScopeId(self.next_scope_id);
        self.next_scope_id += 1;

        let scope = Scope::new(id, kind, parent, range);

        // Add to spatial index (limit to reasonable range to avoid memory issues)
        let end_line = range.end.line.min(range.start.line + 10000);
        for line in range.start.line..=end_line {
            self.scopes_by_line.entry(line).or_default().push(id);
        }

        // Add as child to parent
        if let Some(parent_id) = parent {
            if let Some(parent_scope) = self.scopes.get_mut(parent_id.0 as usize) {
                parent_scope.add_child(id);
            }
        }

        self.scopes.push(scope);
        id
    }

    /// Get a scope by ID
    pub fn get_scope(&self, id: ScopeId) -> Option<&Scope> {
        self.scopes.get(id.0 as usize)
    }

    /// Get a mutable scope by ID
    pub fn get_scope_mut(&mut self, id: ScopeId) -> Option<&mut Scope> {
        self.scopes.get_mut(id.0 as usize)
    }

    /// Link a procedure symbol to its scope
    pub fn link_procedure_scope(&mut self, symbol_id: SymbolId, scope_id: ScopeId) {
        if let Some(scope) = self.get_scope_mut(scope_id) {
            scope.defining_symbol = Some(symbol_id);
        }
    }

    /// Set with object for a with block scope
    pub fn set_with_object(&mut self, scope_id: ScopeId, object: String) {
        if let Some(scope) = self.get_scope_mut(scope_id) {
            scope.with_object = Some(object);
        }
    }

    // ==========================================
    // Reference Tracking
    // ==========================================

    /// Add a reference to a symbol
    pub fn add_reference(
        &mut self,
        symbol_id: SymbolId,
        range: SourceRange,
        scope_id: ScopeId,
        is_assignment: bool,
    ) {
        self.references.push(SymbolReference {
            symbol_id,
            range,
            scope_id,
            is_assignment,
            qualifying_reference: None,
        });
    }

    /// Add a qualified reference (member access)
    pub fn add_qualified_reference(
        &mut self,
        symbol_id: SymbolId,
        range: SourceRange,
        scope_id: ScopeId,
        is_assignment: bool,
        qualifying: SymbolReference,
    ) {
        self.references.push(SymbolReference {
            symbol_id,
            range,
            scope_id,
            is_assignment,
            qualifying_reference: Some(Box::new(qualifying)),
        });
    }

    /// Get all references to a symbol
    pub fn get_references(&self, symbol_id: SymbolId) -> Vec<&SymbolReference> {
        self.references
            .iter()
            .filter(|r| r.symbol_id == symbol_id)
            .collect()
    }

    // ==========================================
    // Query Methods
    // ==========================================

    /// Find the innermost scope containing a position
    pub fn scope_at_position(&self, pos: SourcePosition) -> ScopeId {
        // Get candidate scopes from spatial index
        if let Some(scope_ids) = self.scopes_by_line.get(&pos.line) {
            // Find the innermost (smallest) scope that contains the position
            let mut best = self.module_scope;
            let mut best_size = u64::MAX;

            for &scope_id in scope_ids {
                if let Some(scope) = self.get_scope(scope_id) {
                    if scope.range.contains(pos) {
                        let size = scope.range.size();
                        if size < best_size {
                            best = scope_id;
                            best_size = size;
                        }
                    }
                }
            }

            best
        } else {
            self.module_scope
        }
    }

    /// Look up a symbol by name, searching from the given scope up to the root
    pub fn lookup_symbol(&self, name: &str, from_scope: ScopeId) -> Option<&Symbol> {
        let mut current = Some(from_scope);

        while let Some(scope_id) = current {
            if let Some(scope) = self.get_scope(scope_id) {
                if let Some(symbol_id) = scope.lookup_local(name) {
                    return self.get_symbol(symbol_id);
                }
                current = scope.parent;
            } else {
                break;
            }
        }

        None
    }

    /// Look up a symbol at a specific position
    pub fn lookup_at_position(&self, name: &str, pos: SourcePosition) -> Option<&Symbol> {
        let scope = self.scope_at_position(pos);
        self.lookup_symbol(name, scope)
    }

    /// Find the symbol whose name_range contains the given position
    pub fn symbol_at_position(&self, pos: SourcePosition) -> Option<&Symbol> {
        // First check if we're on a symbol definition
        if let Some(symbol_ids) = self.symbols_by_line.get(&pos.line) {
            for &symbol_id in symbol_ids {
                if let Some(symbol) = self.get_symbol(symbol_id) {
                    if symbol.name_range.contains(pos) {
                        return Some(symbol);
                    }
                }
            }
        }

        // Then check references
        for reference in &self.references {
            if reference.range.contains(pos) {
                return self.get_symbol(reference.symbol_id);
            }
        }

        None
    }

    /// Find the reference at a specific position (if any)
    pub fn reference_at_position(&self, pos: SourcePosition) -> Option<&SymbolReference> {
        self.references.iter().find(|r| r.range.contains(pos))
    }

    /// Get all symbols in a scope (not recursive)
    pub fn symbols_in_scope(&self, scope_id: ScopeId) -> Vec<&Symbol> {
        if let Some(scope) = self.get_scope(scope_id) {
            scope
                .symbols()
                .filter_map(|id| self.get_symbol(id))
                .collect()
        } else {
            Vec::new()
        }
    }

    /// Get all module-level symbols
    pub fn module_symbols(&self) -> Vec<&Symbol> {
        self.symbols_in_scope(self.module_scope)
    }

    /// Get all symbols (for document outline)
    pub fn all_symbols(&self) -> impl Iterator<Item = &Symbol> {
        self.symbols.iter()
    }

    /// Get all symbols of a specific kind
    pub fn symbols_of_kind(&self, kind: SymbolKind) -> impl Iterator<Item = &Symbol> {
        self.symbols.iter().filter(move |s| s.kind == kind)
    }

    /// Find definition of a symbol by name and position
    pub fn find_definition(&self, name: &str, pos: SourcePosition) -> Option<&Symbol> {
        self.lookup_at_position(name, pos)
    }

    /// Find all references to a symbol at position (including the definition)
    pub fn find_all_references(&self, pos: SourcePosition) -> Vec<SourceRange> {
        // First find what symbol is at this position
        let symbol = match self.symbol_at_position(pos) {
            Some(s) => s,
            None => return Vec::new(),
        };

        let mut ranges = vec![symbol.name_range];

        // Add all reference locations
        for reference in self.get_references(symbol.id) {
            ranges.push(reference.range);
        }

        ranges
    }

    /// Get visible symbols at a position (for completion)
    pub fn visible_symbols(&self, pos: SourcePosition) -> Vec<&Symbol> {
        let scope_id = self.scope_at_position(pos);
        let mut visible = Vec::new();
        let mut current = Some(scope_id);
        let mut seen_names = std::collections::HashSet::new();

        // Walk up the scope chain
        while let Some(id) = current {
            if let Some(scope) = self.get_scope(id) {
                for symbol_id in scope.symbols() {
                    if let Some(symbol) = self.get_symbol(symbol_id) {
                        let lower_name = symbol.name.to_lowercase();
                        if !seen_names.contains(&lower_name) {
                            visible.push(symbol);
                            seen_names.insert(lower_name);
                        }
                    }
                }
                current = scope.parent;
            } else {
                break;
            }
        }

        visible
    }

    /// Get procedures (for document outline)
    pub fn procedures(&self) -> impl Iterator<Item = &Symbol> {
        self.symbols.iter().filter(|s| s.kind.is_procedure())
    }

    /// Get all scopes
    pub fn all_scopes(&self) -> impl Iterator<Item = &super::scope::Scope> {
        self.scopes.iter()
    }

    /// Get the count of symbols
    pub fn symbol_count(&self) -> usize {
        self.symbols.len()
    }

    /// Get the count of scopes
    pub fn scope_count(&self) -> usize {
        self.scopes.len()
    }

    /// Get the count of references
    pub fn reference_count(&self) -> usize {
        self.references.len()
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    fn create_test_table() -> SymbolTable {
        SymbolTable::new(Url::parse("file:///test.bas").unwrap())
    }

    #[test]
    fn test_create_symbol() {
        let mut table = create_test_table();

        let id = table.create_symbol(
            "MyVar".to_string(),
            SymbolKind::Variable,
            Visibility::Public,
            SourceRange::new(SourcePosition::new(5, 0), SourcePosition::new(5, 20)),
            SourceRange::new(SourcePosition::new(5, 7), SourcePosition::new(5, 12)),
            table.module_scope,
        );

        let symbol = table.get_symbol(id).unwrap();
        assert_eq!(symbol.name, "MyVar");
        assert_eq!(symbol.kind, SymbolKind::Variable);
    }

    #[test]
    fn test_scope_lookup() {
        let mut table = create_test_table();

        // Create module-level variable
        table.create_symbol(
            "GlobalVar".to_string(),
            SymbolKind::Variable,
            Visibility::Public,
            SourceRange::new(SourcePosition::new(1, 0), SourcePosition::new(1, 25)),
            SourceRange::new(SourcePosition::new(1, 7), SourcePosition::new(1, 16)),
            table.module_scope,
        );

        // Create procedure scope
        let proc_range = SourceRange::new(SourcePosition::new(5, 0), SourcePosition::new(10, 0));
        let proc_scope =
            table.create_scope(ScopeKind::Procedure, Some(table.module_scope), proc_range);

        // Create local variable in procedure
        table.create_symbol(
            "LocalVar".to_string(),
            SymbolKind::LocalVariable,
            Visibility::Private,
            SourceRange::new(SourcePosition::new(6, 4), SourcePosition::new(6, 20)),
            SourceRange::new(SourcePosition::new(6, 8), SourcePosition::new(6, 16)),
            proc_scope,
        );

        // Lookup from procedure scope should find local
        let local = table.lookup_symbol("LocalVar", proc_scope);
        assert!(local.is_some());
        assert_eq!(local.unwrap().name, "LocalVar");

        // Lookup from procedure scope should also find global
        let global = table.lookup_symbol("GlobalVar", proc_scope);
        assert!(global.is_some());
        assert_eq!(global.unwrap().name, "GlobalVar");

        // Lookup from module scope should NOT find local
        let local_from_module = table.lookup_symbol("LocalVar", table.module_scope);
        assert!(local_from_module.is_none());
    }

    #[test]
    fn test_scope_at_position() {
        let mut table = create_test_table();

        // Create procedure scope
        let proc_range = SourceRange::new(SourcePosition::new(5, 0), SourcePosition::new(10, 0));
        let proc_scope =
            table.create_scope(ScopeKind::Procedure, Some(table.module_scope), proc_range);

        // Position inside procedure
        let scope = table.scope_at_position(SourcePosition::new(7, 5));
        assert_eq!(scope, proc_scope);

        // Position outside procedure (should be module scope)
        let scope = table.scope_at_position(SourcePosition::new(3, 0));
        assert_eq!(scope, table.module_scope);
    }

    #[test]
    fn test_case_insensitive_lookup() {
        let mut table = create_test_table();

        table.create_symbol(
            "MyVariable".to_string(),
            SymbolKind::Variable,
            Visibility::Public,
            SourceRange::new(SourcePosition::new(1, 0), SourcePosition::new(1, 20)),
            SourceRange::new(SourcePosition::new(1, 7), SourcePosition::new(1, 17)),
            table.module_scope,
        );

        // All case variations should find the symbol
        assert!(table
            .lookup_symbol("MyVariable", table.module_scope)
            .is_some());
        assert!(table
            .lookup_symbol("myvariable", table.module_scope)
            .is_some());
        assert!(table
            .lookup_symbol("MYVARIABLE", table.module_scope)
            .is_some());
    }
}
