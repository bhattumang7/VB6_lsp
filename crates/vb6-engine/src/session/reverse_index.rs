//! Reverse reference index: declaration → all use sites.
//!
//! Single-module bind produces a *forward* map (`resolutions`: each `NameRef`
//! node → the declaration it resolves to). Find-references and rename need the
//! *inverse*: given a declaration, every `NameRef` that points at it, across all
//! modules in the project.
//!
//! This mirrors VB6 marking a symbol's use sites.
//!
//! Use sites only: the index holds `NameRef` occurrences (uses). The
//! declaration's own name span lives on the `Bound*` struct and is added by the
//! caller when the client requests `includeDeclaration`.

use std::collections::HashMap;

use crate::sema::symbol::{ExternalDecl, NameResolution};

/// A project-global identity for a declaration, used as the reverse-index key.
///
/// Indices are into the owning module's corresponding `BoundModule` vectors.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
pub enum DeclId {
    Proc { module: usize, idx: usize },
    ModuleVar { module: usize, idx: usize },
    Local { module: usize, proc: usize, idx: usize },
    Param { module: usize, proc: usize, idx: usize },
    Type { module: usize, idx: usize },
    Enum { module: usize, idx: usize },
    EnumMember { module: usize, enum_idx: usize, member_idx: usize },
}

/// A single use site: the `NameRef` node `node` (a `NodeId.0`) in module `module`.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct RefSite {
    pub module: usize,
    pub node: u32,
}

/// Map the declaration a resolution points at to a project-global [`DeclId`].
///
/// `cur_module` is the module the *use* lives in — it owns the target for every
/// single-module resolution; [`External`](NameResolution::External) names the
/// target module itself. Returns `None` for builtins and unresolved names
/// (no declaration to point at).
pub fn decl_id_of(cur_module: usize, res: &NameResolution) -> Option<DeclId> {
    Some(match res {
        NameResolution::Local { proc_idx, local_idx } => {
            DeclId::Local { module: cur_module, proc: *proc_idx, idx: *local_idx }
        }
        NameResolution::Param { proc_idx, param_idx } => {
            DeclId::Param { module: cur_module, proc: *proc_idx, idx: *param_idx }
        }
        NameResolution::ModuleVar(i) => DeclId::ModuleVar { module: cur_module, idx: *i },
        NameResolution::Proc(i) => DeclId::Proc { module: cur_module, idx: *i },
        NameResolution::EnumMember { enum_idx, member_idx } => {
            DeclId::EnumMember { module: cur_module, enum_idx: *enum_idx, member_idx: *member_idx }
        }
        NameResolution::External { module, decl } => match decl {
            ExternalDecl::Proc(i) => DeclId::Proc { module: *module, idx: *i },
            ExternalDecl::Var(i) => DeclId::ModuleVar { module: *module, idx: *i },
            ExternalDecl::Type(i) => DeclId::Type { module: *module, idx: *i },
            ExternalDecl::Enum(i) => DeclId::Enum { module: *module, idx: *i },
            ExternalDecl::EnumMember { enum_idx, member_idx } => {
                DeclId::EnumMember { module: *module, enum_idx: *enum_idx, member_idx: *member_idx }
            }
        },
        NameResolution::Builtin | NameResolution::Unresolved => return None,
    })
}

/// Inverse of all modules' `resolutions`: declaration → use sites.
#[derive(Debug, Default, Clone)]
pub struct ReferenceIndex {
    map: HashMap<DeclId, Vec<RefSite>>,
}

impl ReferenceIndex {
    pub fn new() -> Self {
        Self { map: HashMap::new() }
    }

    /// Fold one module's forward resolution map into the reverse index.
    ///
    /// `module_idx` is the module the references live in. `resolutions` is that
    /// module's `BoundModule.resolutions` (after any cross-module upgrade).
    pub fn add_module(
        &mut self,
        module_idx: usize,
        resolutions: &HashMap<u32, NameResolution>,
    ) {
        for (&node, res) in resolutions {
            if let Some(id) = decl_id_of(module_idx, res) {
                self.map.entry(id).or_default().push(RefSite { module: module_idx, node });
            }
        }
    }

    /// All recorded use sites for a declaration, in deterministic order
    /// (by module, then node id).
    pub fn references(&self, id: DeclId) -> Vec<RefSite> {
        let mut sites = self.map.get(&id).cloned().unwrap_or_default();
        sites.sort_by_key(|s| (s.module, s.node));
        sites
    }
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::frontend::ast::{ExprArena, NodeSpans, Span};
    use crate::frontend::parser::Parser;
    use crate::frontend::scanner::ScannerContext;
    use crate::sema::{bind, BoundModule};

    fn analyze(src: &str) -> (BoundModule, NodeSpans, String) {
        let mut ctx = ScannerContext::new(1, 1, 0x0409);
        ctx.intern_keywords();
        let mut arena = ExprArena::new();
        let mut parser = Parser::new(&mut ctx, src.as_bytes());
        let top = parser.parse_module(&mut arena);
        let spans = std::mem::take(&mut parser.node_spans);
        let vis = std::mem::take(&mut parser.decl_public);
        drop(parser);
        let m = bind(&ctx, &arena, &top, &spans, &vis);
        (m, spans, src.to_string())
    }

    fn text(src: &str, span: Span) -> &str {
        let lo = span.start as usize;
        &src[lo..lo + span.len as usize]
    }

    #[test]
    fn collects_local_uses() {
        let src = "Sub Foo()\n    Dim y As Long\n    y = y + 1\n    y = y * 2\nEnd Sub\n";
        let (m, spans, src) = analyze(src);
        let mut idx = ReferenceIndex::new();
        idx.add_module(0, &m.resolutions);

        // `y` is local 0 of proc 0.
        let sites = idx.references(DeclId::Local { module: 0, proc: 0, idx: 0 });
        // Four uses: target+rhs on two assignment lines.
        assert_eq!(sites.len(), 4);
        for s in &sites {
            assert_eq!(s.module, 0);
            assert_eq!(text(&src, spans.get(crate::frontend::ast::NodeId(s.node))), "y");
        }
    }

    #[test]
    fn collects_module_var_uses_and_skips_builtins() {
        let src = "Public gN As Long\nSub Foo()\n    gN = Len(\"hi\")\n    gN = gN + 1\nEnd Sub\n";
        let (m, _spans, _src) = analyze(src);
        let mut idx = ReferenceIndex::new();
        idx.add_module(0, &m.resolutions);

        // gN used three times (two targets + one rhs); Len is a builtin (no decl).
        let sites = idx.references(DeclId::ModuleVar { module: 0, idx: 0 });
        assert_eq!(sites.len(), 3);
    }

    #[test]
    fn external_resolution_maps_to_target_module() {
        // Build a reverse index from a hand-made external resolution to confirm
        // decl_id_of routes the use to the *target* module, not the use's module.
        let mut resolutions = HashMap::new();
        resolutions.insert(
            7u32,
            NameResolution::External {
                module: 3,
                decl: ExternalDecl::Proc(2),
            },
        );
        let mut idx = ReferenceIndex::new();
        idx.add_module(0, &resolutions); // use lives in module 0

        let sites = idx.references(DeclId::Proc { module: 3, idx: 2 });
        assert_eq!(sites, vec![RefSite { module: 0, node: 7 }]);
    }
}
