//! Completion item generation from the bound module state.

use crate::frontend::ast::{ExprNode, NodeId, ProcKind};
use crate::sema::VbaType;
use super::hover;
use super::{span_contains, ModuleData, Session};

/// The semantic kind of a completion item.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum CompletionKind {
    Variable,
    Constant,
    Parameter,
    Function,
    Sub,
    Property,
    Keyword,
    Builtin,
    EnumMember,
    Type,
    Enum,
}

/// A single completion candidate.
#[derive(Debug, Clone)]
pub struct CompletionEntry {
    pub name: String,
    pub kind: CompletionKind,
    /// Rendered declaration string for the detail/documentation pane.
    pub detail: Option<String>,
}

impl Session {
    /// All completion candidates visible at `offset` in `module`.
    ///
    /// When the cursor is after a `.` in a member-access expression (e.g. `obj.|`),
    /// returns only the UDT members of the base object's type. Otherwise returns
    /// identifiers in scope (locals, params, module vars, procs, types, enums,
    /// cross-module public names), known built-ins, and VB6 keywords.
    /// Results are deduplicated case-insensitively, with richer entries winning.
    pub fn completions(&self, module: usize, offset: u32) -> Vec<CompletionEntry> {
        let Some(m) = self.modules.get(module) else { return Vec::new() };

        // Dot-completion: if the cursor is inside a member-access expression, return
        // only the UDT members of the base type.
        if let Some(members) = self.dot_completions(m, offset) {
            return members;
        }

        let mut out = Vec::new();
        push_local_entries(m, offset, &mut out);
        push_module_var_entries(m, &mut out);
        push_proc_entries(m, &mut out);
        push_type_and_enum_entries(m, &mut out);
        for (i, other) in self.modules.iter().enumerate() {
            if i != module {
                push_public_entries(other, &mut out);
            }
        }
        push_builtin_and_keyword_entries(&mut out);
        dedup(out)
    }

    /// If `offset` falls inside a `MemberAccess` expression, return completion
    /// entries for the UDT members of the base object. Returns `None` if we're
    /// not in a dot context, so the caller can fall back to normal completions.
    fn dot_completions(&self, m: &ModuleData, offset: u32) -> Option<Vec<CompletionEntry>> {
        let n = m.arena.len();
        // Walk all arena nodes looking for a MemberAccess whose span contains offset.
        for i in 0..n {
            let id = NodeId(i as u32);
            let ExprNode::MemberAccess { base, .. } = m.arena.get(id) else { continue };
            let span = m.spans.get(id);
            if !span_contains(span, offset) { continue; }

            // Base type must be a local UDT.
            let base_ty = m.bound.types.get(&base.0)?;
            let VbaType::UserDefined(type_sym) = base_ty else { return None };

            // Find the UDT declaration in this module or across modules.
            let members = self.udt_members_for_sym(m, *type_sym)?;
            if members.is_empty() { return None; }
            return Some(members);
        }
        None
    }

    /// Return `CompletionEntry` items for each member of the UDT identified by
    /// `type_sym` in module `m`. Searches the current module first, then others.
    fn udt_members_for_sym(&self, m: &ModuleData, type_sym: u32) -> Option<Vec<CompletionEntry>> {
        // Search current module's type_decls first.
        for decl in &m.bound.type_decls {
            if decl.sym_id == type_sym {
                return Some(members_to_completions(&m.ctx, &m.source, &decl.members));
            }
        }
        // Then search other modules for a public UDT with the same name text.
        let type_name = m.ctx.symbol(type_sym as usize).name.to_ascii_lowercase();
        for other in &self.modules {
            for decl in &other.bound.type_decls {
                if !decl.is_public { continue; }
                let other_name = other.ctx.symbol(decl.sym_id as usize).name.to_ascii_lowercase();
                if other_name == type_name {
                    return Some(members_to_completions(&other.ctx, &other.source, &decl.members));
                }
            }
        }
        None
    }
}

/// Constant or variable kind for a declared variable.
fn var_completion_kind(is_const: bool) -> CompletionKind {
    if is_const { CompletionKind::Constant } else { CompletionKind::Variable }
}

/// Locals and parameters of the procedure containing `offset` (innermost
/// scope first).
fn push_local_entries(m: &ModuleData, offset: u32, out: &mut Vec<CompletionEntry>) {
    let Some(proc_idx) = proc_containing(m, offset) else { return };
    let p = &m.bound.procs[proc_idx];
    for param in &p.params {
        let name = hover::name_at(&m.source, param.name_span);
        if name.is_empty() { continue; }
        out.push(CompletionEntry {
            detail: Some(format!(
                "{} As {}",
                name,
                hover::type_str(&m.ctx, &param.vba_type)
            )),
            name,
            kind: CompletionKind::Parameter,
        });
    }
    for local in &p.locals {
        let name = hover::name_at(&m.source, local.name_span);
        if name.is_empty() { continue; }
        out.push(CompletionEntry {
            detail: Some(hover::var_signature(&m.ctx, &m.source, local)),
            name,
            kind: var_completion_kind(local.is_const),
        });
    }
}

/// Module-level variables and constants.
fn push_module_var_entries(m: &ModuleData, out: &mut Vec<CompletionEntry>) {
    for var in &m.bound.module_vars {
        let name = hover::name_at(&m.source, var.name_span);
        if name.is_empty() { continue; }
        out.push(CompletionEntry {
            detail: Some(hover::var_signature(&m.ctx, &m.source, var)),
            name,
            kind: var_completion_kind(var.is_const),
        });
    }
}

/// Procedures declared in the module.
fn push_proc_entries(m: &ModuleData, out: &mut Vec<CompletionEntry>) {
    for p in &m.bound.procs {
        let name = hover::name_at(&m.source, p.name_span);
        if name.is_empty() { continue; }
        out.push(CompletionEntry {
            detail: Some(hover::proc_signature(&m.ctx, &m.source, p)),
            name,
            kind: proc_completion_kind(p.kind),
        });
    }
}

/// User-defined types, enums and enum members declared in the module.
fn push_type_and_enum_entries(m: &ModuleData, out: &mut Vec<CompletionEntry>) {
    for t in &m.bound.type_decls {
        let name = hover::name_at(&m.source, t.name_span);
        if name.is_empty() { continue; }
        out.push(CompletionEntry { name, kind: CompletionKind::Type, detail: None });
    }
    for e in &m.bound.enum_decls {
        let ename = hover::name_at(&m.source, e.name_span);
        if !ename.is_empty() {
            out.push(CompletionEntry { name: ename.clone(), kind: CompletionKind::Enum, detail: None });
        }
        push_enum_member_entries(m, e, &ename, out);
    }
}

/// Members of one enum declaration.
fn push_enum_member_entries(
    m: &ModuleData,
    e: &crate::sema::symbol::BoundEnumDecl,
    ename: &str,
    out: &mut Vec<CompletionEntry>,
) {
    for mem in &e.members {
        let mname = hover::name_at(&m.source, mem.name_span);
        if mname.is_empty() { continue; }
        out.push(CompletionEntry {
            detail: Some(format!("{}.{} = {}", ename, mname, mem.value)),
            name: mname,
            kind: CompletionKind::EnumMember,
        });
    }
}

/// Public procedures, variables and enum members of another module.
fn push_public_entries(other: &ModuleData, out: &mut Vec<CompletionEntry>) {
    for p in &other.bound.procs {
        if !p.is_public { continue; }
        let name = hover::name_at(&other.source, p.name_span);
        if name.is_empty() { continue; }
        out.push(CompletionEntry {
            detail: Some(hover::proc_signature(&other.ctx, &other.source, p)),
            name,
            kind: proc_completion_kind(p.kind),
        });
    }
    for v in &other.bound.module_vars {
        if !v.is_public { continue; }
        let name = hover::name_at(&other.source, v.name_span);
        if name.is_empty() { continue; }
        out.push(CompletionEntry {
            detail: Some(hover::var_signature(&other.ctx, &other.source, v)),
            name,
            kind: var_completion_kind(v.is_const),
        });
    }
    for e in other.bound.enum_decls.iter().filter(|e| e.is_public) {
        let ename = hover::name_at(&other.source, e.name_span);
        push_enum_member_entries(other, e, &ename, out);
    }
}

/// Built-in functions/statements followed by the VB6 keywords.
fn push_builtin_and_keyword_entries(out: &mut Vec<CompletionEntry>) {
    for &name in crate::sema::builtins::builtin_names() {
        out.push(CompletionEntry {
            name: name.to_string(),
            kind: CompletionKind::Builtin,
            detail: None,
        });
    }
    for &kw in VB6_KEYWORDS {
        out.push(CompletionEntry {
            name: kw.to_string(),
            kind: CompletionKind::Keyword,
            detail: None,
        });
    }
}

/// Convert UDT member declarations to completion entries.
fn members_to_completions(
    ctx: &crate::frontend::scanner::ScannerContext,
    src: &[u8],
    members: &[crate::sema::symbol::BoundTypeMember],
) -> Vec<CompletionEntry> {
    members.iter().filter_map(|mem| {
        let name = hover::name_at(src, mem.name_span);
        if name.is_empty() { return None; }
        Some(CompletionEntry {
            detail: Some(format!("{} As {}", name, hover::type_str(ctx, &mem.vba_type))),
            name,
            kind: CompletionKind::Variable,
        })
    }).collect()
}

/// Find the index of the proc whose body spans `offset`. Procs are assumed to
/// be in source order; each proc extends from its name_span.start to just
/// before the next proc's name_span.start (or EOF).
pub(super) fn proc_containing(m: &ModuleData, offset: u32) -> Option<usize> {
    let procs = &m.bound.procs;
    for i in 0..procs.len() {
        let start = procs[i].name_span.start;
        let end = procs.get(i + 1)
            .map(|next| next.name_span.start)
            .unwrap_or(m.source.len() as u32);
        if offset >= start && offset < end {
            return Some(i);
        }
    }
    None
}

fn proc_completion_kind(kind: ProcKind) -> CompletionKind {
    match kind {
        ProcKind::Sub => CompletionKind::Sub,
        ProcKind::Function => CompletionKind::Function,
        _ => CompletionKind::Property,
    }
}

/// Deduplicate by name (case-insensitive). For duplicates, keep the entry with
/// detail (richer information wins).
fn dedup(mut entries: Vec<CompletionEntry>) -> Vec<CompletionEntry> {
    entries.sort_by(|a, b| {
        a.name.to_ascii_lowercase().cmp(&b.name.to_ascii_lowercase())
            .then(b.detail.is_some().cmp(&a.detail.is_some()))
    });
    entries.dedup_by(|a, b| a.name.eq_ignore_ascii_case(&b.name));
    entries
}

/// VB6 statement and control-flow keywords offered as completion items.
static VB6_KEYWORDS: &[&str] = &[
    "And", "As", "Boolean", "ByRef", "ByVal", "Call", "Case", "Close",
    "Const", "Date", "Declare", "Dim", "Do", "Double", "Each", "Else",
    "ElseIf", "End", "Enum", "Error", "Exit", "False", "For", "Friend",
    "Function", "Get", "GoSub", "GoTo", "If", "Implements", "In",
    "Integer", "Is", "Let", "Like", "Long", "Loop", "Me", "Mod", "New",
    "Next", "Not", "Nothing", "Null", "Object", "On", "Open", "Option",
    "Or", "ParamArray", "Private", "Property", "Public", "RaiseEvent",
    "ReDim", "Resume", "Select", "Set", "Single", "Static", "Step",
    "Stop", "String", "Sub", "Then", "To", "True", "Type", "TypeOf",
    "Until", "Variant", "Wend", "While", "With", "Xor",
];
