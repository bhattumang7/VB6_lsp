//! Workspace Management
//!
//! Handles multi-project workspaces with VBP discovery and cross-project navigation.

mod project;
mod res_parser;
mod vbp_parser;

use vbp_parser::TypeLibReference;
#[allow(unused_imports)]
pub use vbp_parser::VbpParseError;
pub use res_parser::{
    parse_string_table, read_res_file, write_res_file, ResourceEntry, ResourceId, ResourceType,
};
#[allow(unused_imports)]
pub use vbp_parser::{
    CompatibilityMode, CompilationType, ObjectReference, OptimizationType, ProjectMember,
    ProjectType, StartMode, ThreadingModel, TypeLibReference as VbpReference, VbpFile,
};
#[allow(unused_imports)]
pub use project::{ProjectStats, Vb6Project};

use std::path::{Path, PathBuf};

use walkdir::WalkDir;

/// Statistics for the whole workspace
#[derive(Debug, Default)]
pub struct WorkspaceStats {
    pub root_count: usize,
    pub project_count: usize,
    pub total_source_files: usize,
}

/// Manages all VB6 projects in a workspace
#[derive(Debug, Default)]
pub struct WorkspaceManager {
    roots: Vec<PathBuf>,
    projects: Vec<Vb6Project>,
}

impl WorkspaceManager {
    /// Create a new empty workspace manager
    pub fn new() -> Self {
        Self::default()
    }

    /// Add a workspace root and scan for VBP files
    pub fn add_root(&mut self, root: PathBuf) -> Vec<PathBuf> {
        let discovered = self.scan_for_vbp_files(&root);

        for vbp_path in &discovered {
            match Vb6Project::from_vbp(vbp_path) {
                Ok(project) => {
                    self.log_project(&project);
                    self.projects.push(project);
                }
                Err(e) => {
                    tracing::warn!("Failed to load VBP {}: {}", vbp_path.display(), e);
                }
            }
        }

        self.roots.push(root);
        discovered
    }

    /// Iterate over all loaded projects
    pub fn projects(&self) -> impl Iterator<Item = &Vb6Project> {
        self.projects.iter()
    }

    /// Workspace-wide statistics
    pub fn stats(&self) -> WorkspaceStats {
        WorkspaceStats {
            root_count: self.roots.len(),
            project_count: self.projects.len(),
            total_source_files: self.projects.iter().map(|p| p.stats().total_files()).sum(),
        }
    }

    /// Scan a directory recursively for .vbp files
    fn scan_for_vbp_files(&self, root: &Path) -> Vec<PathBuf> {
        let mut vbp_files = Vec::new();

        for entry in WalkDir::new(root)
            .follow_links(true)
            .into_iter()
            .filter_map(|e| e.ok())
        {
            let path = entry.path();
            if path
                .extension()
                .map_or(false, |ext| ext.eq_ignore_ascii_case("vbp"))
            {
                vbp_files.push(path.to_path_buf());
            }
        }

        tracing::info!(
            "Discovered {} VBP files in {}",
            vbp_files.len(),
            root.display()
        );

        vbp_files
    }

    /// Log project details at load time
    fn log_project(&self, project: &Vb6Project) {
        let stats = project.stats();
        let file_count = project
            .source_files()
            .filter(|m| m.absolute_path.exists())
            .count();

        tracing::info!(
            "Loaded '{}' ({}): {}/{} source files on disk, {} references, {} objects",
            project.name(),
            project.vbp_path().display(),
            file_count,
            stats.total_files(),
            project.vbp.references.len(),
            project.vbp.objects.len(),
        );

        for r in project.vbp.get_compiled_references() {
            if let TypeLibReference::Compiled { path, major, minor, lcid, .. } = r {
                tracing::debug!(
                    "  TypeLib: {} ({}) guid={} ver={}.{} lcid={}",
                    r.description(),
                    path.display(),
                    r.uuid().unwrap_or(""),
                    major,
                    minor,
                    lcid
                );
            }
        }

        for r in project.vbp.get_subproject_references() {
            tracing::debug!("  Subproject: {}", r.description());
        }

        for obj in &project.vbp.objects {
            tracing::debug!("  Object: {} v{} ({})", obj.filename, obj.version, obj.guid);
        }

        for member in project.source_files() {
            tracing::trace!(
                "  {} -> {} ({})",
                member.name,
                member.relative_path.display(),
                if member.absolute_path.exists() { "found" } else { "missing" }
            );
        }

        for section in project.vbp.custom_sections.keys() {
            let props = project.vbp.get_custom_section(section);
            tracing::debug!(
                "  Section [{}]: {} properties",
                section,
                props.map_or(0, |m| m.len())
            );
        }
    }
}
