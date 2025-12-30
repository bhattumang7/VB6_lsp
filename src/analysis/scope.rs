//! Scope Definitions
//!
//! Defines scope hierarchy for the symbol table.

use std::collections::HashMap;

use super::position::SourceRange;
use super::symbol::SymbolId;

/// Unique identifier for a scope
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash, Default)]
pub struct ScopeId(pub u32);

/// The type of scope
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
pub enum ScopeKind {
    /// Module-level scope (file scope)
    Module,
    /// Procedure scope (Sub, Function, Property)
    Procedure,
    /// With block scope (implicit object reference)
    WithBlock,
    /// For loop scope (loop variable)
    ForLoop,
    /// For Each loop scope
    ForEachLoop,
}

impl ScopeKind {
    /// Check if this scope kind introduces a new variable scope
    /// (where local variables can be declared)
    pub fn is_variable_scope(&self) -> bool {
        matches!(self, ScopeKind::Module | ScopeKind::Procedure)
    }

    /// Get a display name for the scope kind
    pub fn display_name(&self) -> &'static str {
        match self {
            ScopeKind::Module => "Module",
            ScopeKind::Procedure => "Procedure",
            ScopeKind::WithBlock => "With Block",
            ScopeKind::ForLoop => "For Loop",
            ScopeKind::ForEachLoop => "For Each Loop",
        }
    }
}

/// A scope in the scope hierarchy
#[derive(Debug, Clone)]
pub struct Scope {
    /// Unique identifier
    pub id: ScopeId,
    /// Scope kind
    pub kind: ScopeKind,
    /// Parent scope (None for module scope)
    pub parent: Option<ScopeId>,
    /// Range covered by this scope
    pub range: SourceRange,
    /// Symbols declared in this scope (lowercase name -> symbol id)
    symbols: HashMap<String, SymbolId>,
    /// Child scopes (in document order)
    pub children: Vec<ScopeId>,
    /// For WithBlock: the object expression being referenced
    pub with_object: Option<String>,
    /// The symbol that created this scope (for procedure scopes)
    pub defining_symbol: Option<SymbolId>,
}

impl Scope {
    /// Create a new scope
    pub fn new(id: ScopeId, kind: ScopeKind, parent: Option<ScopeId>, range: SourceRange) -> Self {
        Self {
            id,
            kind,
            parent,
            range,
            symbols: HashMap::new(),
            children: Vec::new(),
            with_object: None,
            defining_symbol: None,
        }
    }

    /// Add a symbol to this scope
    pub fn add_symbol(&mut self, name: &str, symbol_id: SymbolId) {
        // Use lowercase for case-insensitive lookup
        self.symbols.insert(name.to_lowercase(), symbol_id);
    }

    /// Look up a symbol by name in this scope only (case-insensitive)
    pub fn lookup_local(&self, name: &str) -> Option<SymbolId> {
        self.symbols.get(&name.to_lowercase()).copied()
    }

    /// Get all symbols declared in this scope
    pub fn symbols(&self) -> impl Iterator<Item = SymbolId> + '_ {
        self.symbols.values().copied()
    }

    /// Get the number of symbols in this scope
    pub fn symbol_count(&self) -> usize {
        self.symbols.len()
    }

    /// Check if a symbol name exists in this scope
    pub fn has_symbol(&self, name: &str) -> bool {
        self.symbols.contains_key(&name.to_lowercase())
    }

    /// Add a child scope
    pub fn add_child(&mut self, child_id: ScopeId) {
        self.children.push(child_id);
    }
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::analysis::position::SourcePosition;

    #[test]
    fn test_scope_case_insensitive_lookup() {
        let mut scope = Scope::new(
            ScopeId(0),
            ScopeKind::Module,
            None,
            SourceRange::new(SourcePosition::new(0, 0), SourcePosition::new(100, 0)),
        );

        scope.add_symbol("MyVariable", SymbolId(1));

        assert_eq!(scope.lookup_local("MyVariable"), Some(SymbolId(1)));
        assert_eq!(scope.lookup_local("myvariable"), Some(SymbolId(1)));
        assert_eq!(scope.lookup_local("MYVARIABLE"), Some(SymbolId(1)));
        assert_eq!(scope.lookup_local("NotFound"), None);
    }

    #[test]
    fn test_scope_has_symbol() {
        let mut scope = Scope::new(
            ScopeId(0),
            ScopeKind::Procedure,
            Some(ScopeId(1)),
            SourceRange::default(),
        );

        scope.add_symbol("x", SymbolId(0));

        assert!(scope.has_symbol("x"));
        assert!(scope.has_symbol("X"));
        assert!(!scope.has_symbol("y"));
    }
}
