//! Call hierarchy: incoming and outgoing call chains for a procedure.

use std::collections::HashMap;

use super::hover;
use super::reverse_index::{decl_id_of, DeclId, RefSite};
use super::{Location, Session};
use crate::frontend::ast::NodeId;

/// A procedure node in the call hierarchy tree.
#[derive(Debug, Clone)]
pub struct CallHierarchyDecl {
    pub name: String,
    pub location: Location,
}

/// Incoming call edge: a caller and the use-site spans within it.
#[derive(Debug, Clone)]
pub struct IncomingCall {
    pub caller: CallHierarchyDecl,
    pub call_sites: Vec<Location>,
}

/// Outgoing call edge: a callee and the call-site spans within the queried proc.
#[derive(Debug, Clone)]
pub struct OutgoingCall {
    pub callee: CallHierarchyDecl,
    pub call_sites: Vec<Location>,
}

impl Session {
    /// Resolve the procedure whose name span contains `offset` in `module`.
    /// Returns `None` when the cursor is not on a proc name.
    pub fn prepare_call_hierarchy(
        &self,
        module: usize,
        offset: u32,
    ) -> Option<CallHierarchyDecl> {
        let m = self.modules.get(module)?;
        for (idx, p) in m.bound.procs.iter().enumerate() {
            if super::span_contains(p.name_span, offset) {
                let name = hover::name_at(&m.source, p.name_span);
                let location = self.decl_location(DeclId::Proc { module, idx })?;
                return Some(CallHierarchyDecl { name, location });
            }
        }
        None
    }

    /// The `(module, procedure index)` whose name span contains `offset`.
    fn proc_at(&self, module: usize, offset: u32) -> Option<usize> {
        let m = self.modules.get(module)?;
        m.bound.procs.iter().position(|p| super::span_contains(p.name_span, offset))
    }

    /// Callers of the procedure whose name is at `offset` in `module` — the
    /// declaration itself, so procedures sharing a name in different modules
    /// are not confused.
    pub fn incoming_calls_at(&self, module: usize, offset: u32) -> Vec<IncomingCall> {
        match self.proc_at(module, offset) {
            Some(idx) => self.incoming_calls_to(module, idx),
            None => Vec::new(),
        }
    }

    /// Callees of the procedure whose name is at `offset` in `module`.
    pub fn outgoing_calls_at(&self, module: usize, offset: u32) -> Vec<OutgoingCall> {
        match self.proc_at(module, offset) {
            Some(idx) => self.outgoing_calls_from(module, idx),
            None => Vec::new(),
        }
    }

    /// All callers of the named procedure (the first procedure of that name in
    /// project order; use [`incoming_calls_to`](Self::incoming_calls_to) when
    /// several share a name).
    pub fn incoming_calls(&self, proc_name: &str) -> Vec<IncomingCall> {
        match self.find_proc_by_name(proc_name) {
            Some((module, idx)) => self.incoming_calls_to(module, idx),
            None => Vec::new(),
        }
    }

    /// All callers of procedure `proc_idx` of `module`: every use of it, other
    /// than as the operand of `AddressOf`, grouped by the procedure that
    /// contains it. Uses outside any procedure are not calls.
    pub fn incoming_calls_to(&self, module: usize, proc_idx: usize) -> Vec<IncomingCall> {
        if self.modules.get(module).and_then(|m| m.bound.procs.get(proc_idx)).is_none() {
            return Vec::new();
        }
        let id = DeclId::Proc { module, idx: proc_idx };
        let mut by_caller: HashMap<(usize, usize), Vec<Location>> = HashMap::new();

        for site in self.refs.references(id) {
            let m = &self.modules[site.module];
            if m.address_of_nodes.contains(&site.node) {
                continue;
            }
            let Some(site_span) = self.ref_site_location(&site) else { continue };
            if let Some(cidx) = proc_containing_offset(&m.bound.procs, site_span.span.start) {
                by_caller.entry((site.module, cidx)).or_default().push(site_span);
            }
        }

        let mut callers: Vec<_> = by_caller.into_iter().collect();
        callers.sort_by_key(|(k, _)| *k);
        callers
            .into_iter()
            .filter_map(|((mod_idx, cidx), mut call_sites)| {
                call_sites.sort_by_key(|l| l.span.start);
                let p = &self.modules[mod_idx].bound.procs[cidx];
                let name = hover::name_at(&self.modules[mod_idx].source, p.name_span);
                let location = self.decl_location(DeclId::Proc { module: mod_idx, idx: cidx })?;
                Some(IncomingCall {
                    caller: CallHierarchyDecl { name, location },
                    call_sites,
                })
            })
            .collect()
    }

    /// All callees of the named procedure (see
    /// [`outgoing_calls_from`](Self::outgoing_calls_from)).
    pub fn outgoing_calls(&self, proc_name: &str) -> Vec<OutgoingCall> {
        match self.find_proc_by_name(proc_name) {
            Some((module, idx)) => self.outgoing_calls_from(module, idx),
            None => Vec::new(),
        }
    }

    /// All callees of procedure `proc_idx` of `module`: every use inside its
    /// body that resolves to a procedure of the project, other than as the
    /// operand of `AddressOf`. That covers calls by bare name, through an
    /// object (`o.Save`), and property reads and writes (`x = o.Title`).
    pub fn outgoing_calls_from(&self, module: usize, proc_idx: usize) -> Vec<OutgoingCall> {
        let Some(m) = self.modules.get(module) else {
            return Vec::new();
        };
        if m.bound.procs.get(proc_idx).is_none() {
            return Vec::new();
        }

        let mut by_callee: HashMap<DeclId, Vec<Location>> = HashMap::new();
        for (&node_id, res) in &m.bound.resolutions {
            let Some(callee_id @ DeclId::Proc { .. }) = decl_id_of(module, res) else {
                continue;
            };
            if m.address_of_nodes.contains(&node_id) {
                continue;
            }
            let span = m.spans.get(NodeId(node_id));
            if span.len == 0 || proc_containing_offset(&m.bound.procs, span.start) != Some(proc_idx) {
                continue;
            }
            by_callee.entry(callee_id).or_default().push(Location { module, span });
        }

        let mut callees: Vec<_> = by_callee.into_iter().collect();
        callees.sort_by_key(|(id, _)| match id {
            DeclId::Proc { module, idx } => (*module, *idx),
            _ => (usize::MAX, usize::MAX),
        });
        callees
            .into_iter()
            .filter_map(|(callee_id, mut call_sites)| {
                let DeclId::Proc { module: tmod, idx: tidx } = callee_id else {
                    return None;
                };
                call_sites.sort_by_key(|l| l.span.start);
                let tp = &self.modules[tmod].bound.procs[tidx];
                let name = hover::name_at(&self.modules[tmod].source, tp.name_span);
                let location = self.decl_location(callee_id)?;
                Some(OutgoingCall {
                    callee: CallHierarchyDecl { name, location },
                    call_sites,
                })
            })
            .collect()
    }

    // ── helpers ─────────────────────────────────────────────────────────────────

    fn find_proc_by_name(&self, name: &str) -> Option<(usize, usize)> {
        for (mi, m) in self.modules.iter().enumerate() {
            for (pi, p) in m.bound.procs.iter().enumerate() {
                if hover::name_at(&m.source, p.name_span).eq_ignore_ascii_case(name) {
                    return Some((mi, pi));
                }
            }
        }
        None
    }

    fn ref_site_location(&self, site: &RefSite) -> Option<Location> {
        let m = self.modules.get(site.module)?;
        let span = m.spans.get(NodeId(site.node));
        Some(Location { module: site.module, span })
    }
}

fn proc_containing_offset(
    procs: &[crate::sema::symbol::BoundProc],
    offset: u32,
) -> Option<usize> {
    for i in 0..procs.len() {
        let start = procs[i].name_span.start;
        let end = procs.get(i + 1).map(|n| n.name_span.start).unwrap_or(u32::MAX);
        if offset >= start && offset < end {
            return Some(i);
        }
    }
    None
}
