//! Call hierarchy: incoming and outgoing call chains for a procedure.

use std::collections::HashMap;

use super::hover;
use super::reverse_index::{DeclId, RefSite};
use super::{Location, Session};
use crate::frontend::ast::NodeId;
use crate::sema::symbol::{ExternalDecl, NameResolution};

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

    /// All callers of the named procedure.
    pub fn incoming_calls(&self, proc_name: &str) -> Vec<IncomingCall> {
        let Some((target_module, target_idx)) = self.find_proc_by_name(proc_name) else {
            return Vec::new();
        };
        let id = DeclId::Proc { module: target_module, idx: target_idx };
        let sites = self.refs.references(id);

        // Group call-site locations by (caller_module, caller_proc_idx).
        let mut by_caller: HashMap<(usize, usize), Vec<Location>> = HashMap::new();

        for site in sites {
            let site_span = match self.ref_site_location(&site) {
                Some(loc) => loc,
                None => continue,
            };
            let m = &self.modules[site.module];
            let site_offset = m.spans.get(NodeId(site.node)).start;
            if let Some(cidx) = proc_containing_offset(&m.bound.procs, site_offset) {
                by_caller.entry((site.module, cidx)).or_default().push(site_span);
            }
        }

        by_caller
            .into_iter()
            .filter_map(|((mod_idx, cidx), call_sites)| {
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

    /// All callees of the named procedure.
    pub fn outgoing_calls(&self, proc_name: &str) -> Vec<OutgoingCall> {
        let Some((mod_idx, proc_idx)) = self.find_proc_by_name(proc_name) else {
            return Vec::new();
        };
        let m = &self.modules[mod_idx];
        let procs = &m.bound.procs;
        let proc_start = procs[proc_idx].name_span.start;
        let proc_end = procs.get(proc_idx + 1)
            .map(|next| next.name_span.start)
            .unwrap_or(m.source.len() as u32);

        // Collect call-site locations, grouped by callee DeclId.
        let mut by_callee: HashMap<DeclId, Vec<Location>> = HashMap::new();

        for (&node_id, res) in &m.bound.resolutions {
            if !m.callee_nodes.contains(&node_id) { continue; }
            let span = m.spans.get(NodeId(node_id));
            if span.start < proc_start || span.start >= proc_end { continue; }

            let callee_id = match res {
                NameResolution::Proc(i) => DeclId::Proc { module: mod_idx, idx: *i },
                NameResolution::External { module, decl: ExternalDecl::Proc(i) } => {
                    DeclId::Proc { module: *module, idx: *i }
                }
                _ => continue,
            };
            let site_loc = Location { module: mod_idx, span };
            by_callee.entry(callee_id).or_default().push(site_loc);
        }

        by_callee
            .into_iter()
            .filter_map(|(callee_id, call_sites)| {
                let DeclId::Proc { module: tmod, idx: tidx } = callee_id else {
                    return None;
                };
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
