//! Project-wide index of public declarations across modules.
//!
//! Single-module [`bind`](crate::sema::bind) resolves only names declared in the
//! same module; anything else is left
//! [`Unresolved`](crate::sema::NameResolution::Unresolved). VB6 projects are
//! multi-file, so the session builds a `ModuleIndex` over every bound module and
//! uses it to upgrade those `Unresolved` references to
//! [`External`](crate::sema::NameResolution::External).
//!
//! This mirrors VB6's project-scope identifier lookup: names are matched by
//! their text, since each module interns symbols in its own scanner context.
//!
//! **Visibility:** the binder receives the parser's `decl_public` table, which
//! captures explicit `Public`/`Private`/`Friend` modifiers. Declarations absent
//! from that table take VB6 defaults (procedures/types/enums = Public, module
//! variables/constants = Private). Only `is_public = true` declarations are
//! indexed here.
//!
//! **Ambiguity:** if the same public name is declared in two modules, the first
//! one indexed wins. Project-scope ambiguity resolution is a later refinement.

use std::collections::HashMap;

use crate::sema::symbol::{BoundEnumDecl, BoundModule, ExternalDecl};

/// A resolved location of a public declaration: which module, and which decl.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct ExternalRef {
    pub module: usize,
    pub decl: ExternalDecl,
}

/// Project-wide map from lowercased declaration name to its location.
#[derive(Debug, Default, Clone)]
pub struct ModuleIndex {
    map: HashMap<String, ExternalRef>,
}

impl ModuleIndex {
    pub fn new() -> Self {
        Self { map: HashMap::new() }
    }

    /// Add a module's public declarations to the index.
    ///
    /// `module_idx` identifies the module within the project's module list.
    /// `name_of` resolves a `sym_id` to its source name (the caller supplies the
    /// module's scanner context). Names are stored lowercased (VB6 is
    /// case-insensitive). The first declaration of a given name wins.
    ///
    /// `members_global` says whether the module's procedures and variables are
    /// project-wide names. They are for a standard module; a class's members are
    /// reached through an instance, so they are indexed only for a class marked
    /// `VB_GlobalNameSpace`. Types and enums are project-wide either way.
    pub fn add_module(
        &mut self,
        module_idx: usize,
        module: &BoundModule,
        name_of: impl Fn(u32) -> String,
        members_global: bool,
    ) {
        // Only public declarations are visible to other modules in the project.
        if members_global {
            self.add_global_members(module_idx, module, &name_of);
        }
        for (i, t) in module.type_decls.iter().enumerate() {
            if t.is_public {
                self.insert(name_of(t.sym_id), module_idx, ExternalDecl::Type(i));
            }
        }
        for (i, e) in module.enum_decls.iter().enumerate() {
            if e.is_public {
                self.add_enum(module_idx, i, e, &name_of);
            }
        }
    }

    /// Index the public procedures and variables of a module whose members
    /// are project-wide names.
    fn add_global_members(
        &mut self,
        module_idx: usize,
        module: &BoundModule,
        name_of: &impl Fn(u32) -> String,
    ) {
        for (i, p) in module.procs.iter().enumerate() {
            if p.is_public {
                self.insert(name_of(p.sym_id), module_idx, ExternalDecl::Proc(i));
            }
        }
        for (i, v) in module.module_vars.iter().enumerate() {
            if v.is_public {
                self.insert(name_of(v.sym_id), module_idx, ExternalDecl::Var(i));
            }
        }
    }

    /// Index a public enum and its members.
    fn add_enum(
        &mut self,
        module_idx: usize,
        enum_idx: usize,
        e: &BoundEnumDecl,
        name_of: &impl Fn(u32) -> String,
    ) {
        self.insert(name_of(e.sym_id), module_idx, ExternalDecl::Enum(enum_idx));
        // Enum members are project-scoped constants accessible unqualified
        // (when the enum itself is public).
        for (mi, member) in e.members.iter().enumerate() {
            self.insert(
                name_of(member.sym_id),
                module_idx,
                ExternalDecl::EnumMember { enum_idx, member_idx: mi },
            );
        }
    }

    fn insert(&mut self, name: String, module: usize, decl: ExternalDecl) {
        if name.is_empty() {
            return;
        }
        self.map
            .entry(name)
            .or_insert(ExternalRef { module, decl });
    }

    /// Look up a name (any case) in the project index.
    pub fn lookup(&self, name: &str) -> Option<ExternalRef> {
        self.map.get(&name.to_ascii_lowercase()).copied()
    }

    /// Number of distinct names indexed.
    pub fn len(&self) -> usize {
        self.map.len()
    }

    pub fn is_empty(&self) -> bool {
        self.map.is_empty()
    }
}

#[cfg(test)]
#[path = "tests/module_index.rs"]
mod tests;
