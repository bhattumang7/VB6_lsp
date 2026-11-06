//! The name each `Attribute VB_Name` line gives a module has to be one the project can take:
//! not the project's name, a referenced library's name or another module's, and, for a
//! standard module, a legal name. The checks need the whole project, so they run when the
//! project is linked.

use super::*;
use crate::frontend::module_name::{ModuleNameSetting, NameScope};

impl Session {
    /// Checks the name every module is given.
    pub(super) fn check_module_names(&mut self) {
        for idx in 0..self.modules.len() {
            self.check_module_name(idx);
        }
    }

    /// Checks the names module `idx` is given, in the order its `VB_Name` lines give them, and
    /// reports the failure they leave.
    pub(super) fn check_module_name(&mut self, idx: usize) {
        let module = &self.modules[idx];
        let others: Vec<&str> = self
            .modules
            .iter()
            .enumerate()
            .filter(|(i, _)| *i != idx)
            .map(|(_, m)| m.module_name.as_str())
            .collect();
        let libraries: Vec<&str> = self.config.libraries.iter().map(|l| l.name.as_str()).collect();
        let scope = NameScope {
            kind: module.kind,
            project: self.config.project_name.as_deref(),
            libraries: &libraries,
            modules: &others,
        };
        let mut setting = ModuleNameSetting::new();
        for i in 0..module.arena.len() as u32 {
            let ExprNode::Attribute { target: None, name, values } = module.arena.get(NodeId(i)) else { continue };
            if !module.ctx.symbol(*name as usize).name.eq_ignore_ascii_case("vb_name") {
                continue;
            }
            let Some(ExprNode::Literal { lit: crate::frontend::ast::AstLit::Str(text) }) =
                values.first().map(|v| module.arena.get(*v))
            else {
                continue;
            };
            setting.set(&scope, text, module.spans.get(NodeId(i)));
        }
        if let Some(diagnostic) = setting.diagnostic() {
            self.modules[idx].bound.diagnostics.push_diagnostic(diagnostic);
        }
    }
}
