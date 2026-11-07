//! Names that stand for a type or member of a referenced library, and finding where a library
//! built from source declares them.
//!
//! A project reference makes another project's public classes a library of this session. The
//! library describes types and members but not where they are written, so a host that has the
//! referenced project's own session asks it for the declaration by name.

use super::library::LibraryMemberKind;
use super::{CursorTarget, Location, NameResolution, Session};
use crate::frontend::ast::{ProcKind, Span};

/// A library type or member a name in the session resolves to.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct ExternalTarget {
    /// The library's name: the name of the project it was built from.
    pub library: String,
    pub type_name: String,
    /// The member, with its kind, when the name is a member rather than the type.
    pub member: Option<(String, LibraryMemberKind)>,
}

impl Session {
    /// The library type or member the name at `offset` resolves to, if it resolves to one.
    pub fn external_target_at(&self, module: usize, offset: u32) -> Option<ExternalTarget> {
        let CursorTarget::Use { res, .. } = self.target_at(module, offset)? else { return None };
        let libraries = &self.config.libraries;
        match res {
            NameResolution::LibraryType { lib, ty } => {
                let l = libraries.get(lib)?;
                Some(ExternalTarget { library: l.name.clone(), type_name: l.types.get(ty)?.name.clone(), member: None })
            }
            NameResolution::LibraryMember { lib, ty, member } => {
                let l = libraries.get(lib)?;
                let t = l.types.get(ty)?;
                let m = t.members.get(member)?;
                Some(ExternalTarget {
                    library: l.name.clone(),
                    type_name: t.name.clone(),
                    member: Some((m.name.clone(), m.kind)),
                })
            }
            _ => None,
        }
    }

    /// Where this session's source declares the public type `type_name` or, with `member`, its
    /// member. A class is declared by its module (the start of the file); an enumeration or
    /// record by its name. Names match case-insensitively. A property's accessors share a name:
    /// the one of the requested kind is preferred.
    pub fn declaration_of(&self, type_name: &str, member: Option<(&str, LibraryMemberKind)>) -> Option<Location> {
        for (module, m) in self.modules.iter().enumerate() {
            let name_of = |sym: u32| m.ctx.symbol(sym as usize).name.as_str();
            let is_class = matches!(m.path.rsplit('.').next().map(str::to_ascii_lowercase).as_deref(), Some("cls" | "ctl" | "dob"));
            if is_class && m.module_name.eq_ignore_ascii_case(type_name) {
                let Some((wanted, kind)) = member else {
                    return Some(Location { module, span: Span { start: 0, len: 0 } });
                };
                let procs: Vec<_> = m
                    .bound
                    .procs
                    .iter()
                    .filter(|p| p.is_public && name_of(p.sym_id).eq_ignore_ascii_case(wanted))
                    .collect();
                if let Some(p) = procs.iter().find(|p| proc_kind(p.kind) == kind).or(procs.first()) {
                    return Some(Location { module, span: p.name_span });
                }
                if let Some(v) = m.bound.module_vars.iter().find(|v| v.is_public && name_of(v.sym_id).eq_ignore_ascii_case(wanted)) {
                    return Some(Location { module, span: v.name_span });
                }
                continue;
            }
            if let Some(e) = m.bound.enum_decls.iter().find(|e| e.is_public && name_of(e.sym_id).eq_ignore_ascii_case(type_name)) {
                return Some(match member {
                    None => Location { module, span: e.name_span },
                    Some((wanted, _)) => {
                        let found = e.members.iter().find(|em| name_of(em.sym_id).eq_ignore_ascii_case(wanted))?;
                        Location { module, span: found.name_span }
                    }
                });
            }
            if let Some(t) = m.bound.type_decls.iter().find(|t| t.is_public && name_of(t.sym_id).eq_ignore_ascii_case(type_name)) {
                return Some(match member {
                    None => Location { module, span: t.name_span },
                    Some((wanted, _)) => {
                        let found = t.members.iter().find(|tm| name_of(tm.sym_id).eq_ignore_ascii_case(wanted))?;
                        Location { module, span: found.name_span }
                    }
                });
            }
        }
        None
    }
}

fn proc_kind(kind: ProcKind) -> LibraryMemberKind {
    match kind {
        ProcKind::Sub | ProcKind::Function => LibraryMemberKind::Method,
        ProcKind::PropGet => LibraryMemberKind::PropertyGet,
        ProcKind::PropLet => LibraryMemberKind::PropertyLet,
        ProcKind::PropSet => LibraryMemberKind::PropertySet,
    }
}
