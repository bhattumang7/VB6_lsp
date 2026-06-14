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

use crate::sema::symbol::{BoundModule, ExternalDecl};

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
    pub fn add_module(
        &mut self,
        module_idx: usize,
        module: &BoundModule,
        name_of: impl Fn(u32) -> String,
    ) {
        // Only public declarations are visible to other modules in the project.
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
        for (i, t) in module.type_decls.iter().enumerate() {
            if t.is_public {
                self.insert(name_of(t.sym_id), module_idx, ExternalDecl::Type(i));
            }
        }
        for (i, e) in module.enum_decls.iter().enumerate() {
            if !e.is_public {
                continue;
            }
            self.insert(name_of(e.sym_id), module_idx, ExternalDecl::Enum(i));
            // Enum members are project-scoped constants accessible unqualified
            // (when the enum itself is public).
            for (mi, member) in e.members.iter().enumerate() {
                self.insert(
                    name_of(member.sym_id),
                    module_idx,
                    ExternalDecl::EnumMember { enum_idx: i, member_idx: mi },
                );
            }
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
mod tests {
    use super::*;
    use crate::frontend::ast::ExprArena;
    use crate::frontend::parser::Parser;
    use crate::frontend::scanner::ScannerContext;
    use crate::sema::bind;

    /// Parse+bind one module, returning its bound form plus a name resolver.
    fn bind_module(src: &str) -> (BoundModule, ScannerContext) {
        let mut ctx = ScannerContext::new(1, 1, 0x0409);
        ctx.intern_keywords();
        let mut arena = ExprArena::new();
        let mut parser = Parser::new(&mut ctx, src.as_bytes());
        let top = parser.parse_module(&mut arena);
        let spans = std::mem::take(&mut parser.node_spans);
        let vis = std::mem::take(&mut parser.decl_public);
        drop(parser);
        let m = bind(&ctx, &arena, &top, &spans, &vis);
        (m, ctx)
    }

    #[test]
    fn indexes_public_procs_and_vars() {
        let (m, ctx) = bind_module(
            "Public gConfig As Long\nPublic Sub DoThing()\nEnd Sub\nFunction Calc() As Long\nEnd Function\n",
        );
        let mut idx = ModuleIndex::new();
        idx.add_module(0, &m, |s| ctx.symbol(s as usize).name.to_ascii_lowercase());

        // Case-insensitive lookup.
        assert_eq!(idx.lookup("dothing").map(|r| r.decl), Some(ExternalDecl::Proc(0)));
        assert_eq!(idx.lookup("CALC").map(|r| r.decl), Some(ExternalDecl::Proc(1)));
        assert_eq!(idx.lookup("gconfig").map(|r| r.decl), Some(ExternalDecl::Var(0)));
        assert!(idx.lookup("missing").is_none());
    }

    #[test]
    fn indexes_enum_members() {
        let (m, ctx) = bind_module(
            "Public Enum Color\n    Red\n    Green\n    Blue\nEnd Enum\n",
        );
        let mut idx = ModuleIndex::new();
        idx.add_module(2, &m, |s| ctx.symbol(s as usize).name.to_ascii_lowercase());

        assert_eq!(idx.lookup("color").map(|r| r.decl), Some(ExternalDecl::Enum(0)));
        assert_eq!(
            idx.lookup("green").map(|r| r.decl),
            Some(ExternalDecl::EnumMember { enum_idx: 0, member_idx: 1 }),
        );
        // The referenced module index is propagated.
        assert_eq!(idx.lookup("red").map(|r| r.module), Some(2));
    }

    #[test]
    fn first_module_wins_on_clash() {
        let (m0, c0) = bind_module("Public Sub Shared()\nEnd Sub\n");
        let (m1, c1) = bind_module("Public Sub Shared()\nEnd Sub\n");
        let mut idx = ModuleIndex::new();
        idx.add_module(0, &m0, |s| c0.symbol(s as usize).name.to_ascii_lowercase());
        idx.add_module(1, &m1, |s| c1.symbol(s as usize).name.to_ascii_lowercase());
        assert_eq!(idx.lookup("shared").map(|r| r.module), Some(0));
    }
}
