//! Workspace Management
//!
//! Handles multi-project workspaces with VBP discovery and cross-project navigation.

mod frx_parser;
mod project;
mod res_parser;
mod vbp_parser;

pub use frx_parser::{list_resolver, resource_file_resolver};
pub use project::{ProjectStats, Vb6Project};
pub use res_parser::{
    create_string_table, parse_string_table, read_res_file, write_res_file, MemoryFlags,
    ResHeader, ResourceEntry, ResourceId, ResourceType, StringTableEntry,
};
pub use vbp_parser::{
    ObjectReference, ProjectMember, ProjectType, TypeLibReference, VbpFile, VbpParseError,
};

use std::collections::HashMap;
use std::path::{Path, PathBuf};

use tower_lsp::lsp_types::{Location, Url};
use walkdir::WalkDir;

use crate::analysis::{SymbolKind, SymbolTable};

/// Manages all VB6 projects in a workspace
#[derive(Debug)]
pub struct WorkspaceManager {
    /// Workspace root folders
    roots: Vec<PathBuf>,

    /// All discovered VB6 projects, keyed by VBP path
    projects: HashMap<PathBuf, Vb6Project>,

    /// Reverse index: source file path -> VBP path
    file_to_project: HashMap<PathBuf, PathBuf>,

    /// Files that don't belong to any VBP (orphans)
    orphan_files: HashMap<PathBuf, SymbolTable>,
}

impl WorkspaceManager {
    /// Create a new empty workspace manager
    pub fn new() -> Self {
        Self {
            roots: Vec::new(),
            projects: HashMap::new(),
            file_to_project: HashMap::new(),
            orphan_files: HashMap::new(),
        }
    }

    /// Add a workspace root and scan for VBP files
    pub fn add_root(&mut self, root: PathBuf) -> Vec<PathBuf> {
        let discovered = self.scan_for_vbp_files(&root);

        for vbp_path in &discovered {
            if let Err(e) = self.load_project(vbp_path) {
                tracing::warn!("Failed to load VBP {}: {}", vbp_path.display(), e);
            }
        }

        self.roots.push(root);
        discovered
    }

    /// Remove a workspace root
    pub fn remove_root(&mut self, root: &Path) {
        self.roots.retain(|r| r != root);

        // Remove projects under this root
        let to_remove: Vec<PathBuf> = self
            .projects
            .keys()
            .filter(|p| p.starts_with(root))
            .cloned()
            .collect();

        for vbp_path in to_remove {
            self.unload_project(&vbp_path);
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
            if path.extension().map_or(false, |ext| {
                ext.eq_ignore_ascii_case("vbp")
            }) {
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

    /// Load a VBP project
    pub fn load_project(&mut self, vbp_path: &Path) -> Result<(), VbpParseError> {
        let project = Vb6Project::from_vbp(vbp_path)?;

        // Build file-to-project index
        for member in project.source_files() {
            let normalized = normalize_path(&member.absolute_path);
            self.file_to_project
                .insert(normalized, vbp_path.to_path_buf());
        }

        tracing::info!(
            "Loaded project '{}' with {} source files",
            project.name(),
            project.source_files().count()
        );

        self.projects.insert(vbp_path.to_path_buf(), project);
        Ok(())
    }

    /// Unload a VBP project
    pub fn unload_project(&mut self, vbp_path: &Path) {
        if let Some(project) = self.projects.remove(vbp_path) {
            // Remove from file-to-project index
            for member in project.source_files() {
                let normalized = normalize_path(&member.absolute_path);
                self.file_to_project.remove(&normalized);
            }

            tracing::info!("Unloaded project '{}'", project.name());
        }
    }

    /// Get the project that contains a file
    pub fn project_for_file(&self, file_path: &Path) -> Option<&Vb6Project> {
        let normalized = normalize_path(file_path);
        self.file_to_project
            .get(&normalized)
            .and_then(|vbp_path| self.projects.get(vbp_path))
    }

    /// Get a mutable reference to the project that contains a file
    pub fn project_for_file_mut(&mut self, file_path: &Path) -> Option<&mut Vb6Project> {
        let normalized = normalize_path(file_path);
        if let Some(vbp_path) = self.file_to_project.get(&normalized).cloned() {
            self.projects.get_mut(&vbp_path)
        } else {
            None
        }
    }

    /// Check if a file belongs to any project
    pub fn is_file_in_project(&self, file_path: &Path) -> bool {
        let normalized = normalize_path(file_path);
        self.file_to_project.contains_key(&normalized)
    }

    /// Store a symbol table for a file
    pub fn set_symbol_table(&mut self, file_path: &Path, table: SymbolTable) {
        let normalized = normalize_path(file_path);

        if let Some(project) = self.project_for_file_mut(file_path) {
            project.set_symbol_table(normalized, table);
        } else {
            // Orphan file
            self.orphan_files.insert(normalized, table);
        }
    }

    /// Get a symbol table for a file
    pub fn get_symbol_table(&self, file_path: &Path) -> Option<&SymbolTable> {
        let normalized = normalize_path(file_path);

        if let Some(project) = self.project_for_file(file_path) {
            project.get_symbol_table(&normalized)
        } else {
            self.orphan_files.get(&normalized)
        }
    }

    /// Remove a symbol table
    pub fn remove_symbol_table(&mut self, file_path: &Path) {
        let normalized = normalize_path(file_path);

        if let Some(project) = self.project_for_file_mut(file_path) {
            project.remove_symbol_table(&normalized);
        } else {
            self.orphan_files.remove(&normalized);
        }
    }

    /// Resolve a symbol across the workspace
    /// First checks the current file's project, then other projects
    pub fn resolve_symbol(&self, name: &str, from_file: &Path) -> Option<Location> {
        // 1. Check current file's project first
        if let Some(project) = self.project_for_file(from_file) {
            // Check if it's a module/class name reference
            if let Some(location) = project.resolve_module_reference(name) {
                return Some(location);
            }

            // Check public symbols in the same project
            if let Some(location) = project.find_public_symbol(name) {
                return Some(location);
            }
        }

        // 2. Check other projects for public symbols
        for project in self.projects.values() {
            if let Some(location) = project.find_public_symbol(name) {
                return Some(location);
            }
        }

        // 3. Check orphan files
        let name_lower = name.to_lowercase();
        for table in self.orphan_files.values() {
            if let Some(symbol) = table.lookup_symbol(&name_lower, table.module_scope) {
                if symbol.visibility == crate::analysis::Visibility::Public {
                    let range = symbol.name_range.to_lsp();
                    return Some(Location {
                        uri: table.uri.clone(),
                        range,
                    });
                }
            }
        }

        None
    }

    /// Find all public symbols matching a prefix (for workspace-wide completion)
    pub fn find_symbols_with_prefix(&self, prefix: &str) -> Vec<(String, PathBuf, SymbolKind)> {
        let mut results = Vec::new();

        for project in self.projects.values() {
            for (name, path, kind) in project.find_public_symbols_with_prefix(prefix) {
                results.push((name.to_string(), path.to_path_buf(), kind));
            }
        }

        results
    }

    /// Get all projects
    pub fn projects(&self) -> impl Iterator<Item = &Vb6Project> {
        self.projects.values()
    }

    /// Get a project by its VBP path
    pub fn get_project(&self, vbp_path: &Path) -> Option<&Vb6Project> {
        self.projects.get(vbp_path)
    }

    /// Get workspace statistics
    pub fn stats(&self) -> WorkspaceStats {
        let mut total_files = 0;
        let mut loaded_tables = 0;

        for project in self.projects.values() {
            let stats = project.stats();
            total_files += stats.module_count + stats.class_count + stats.form_count + stats.user_control_count;
            loaded_tables += stats.loaded_symbol_tables;
        }

        WorkspaceStats {
            root_count: self.roots.len(),
            project_count: self.projects.len(),
            total_source_files: total_files,
            loaded_symbol_tables: loaded_tables,
            orphan_files: self.orphan_files.len(),
        }
    }
}

impl Default for WorkspaceManager {
    fn default() -> Self {
        Self::new()
    }
}

/// Workspace statistics
#[derive(Debug, Clone)]
pub struct WorkspaceStats {
    pub root_count: usize,
    pub project_count: usize,
    pub total_source_files: usize,
    pub loaded_symbol_tables: usize,
    pub orphan_files: usize,
}

/// Normalize a path for comparison (lowercase on Windows)
fn normalize_path(path: &Path) -> PathBuf {
    let path = path.canonicalize().unwrap_or_else(|_| path.to_path_buf());

    #[cfg(windows)]
    {
        PathBuf::from(path.to_string_lossy().to_lowercase())
    }

    #[cfg(not(windows))]
    {
        path
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn test_workspace_manager_creation() {
        let manager = WorkspaceManager::new();
        assert_eq!(manager.projects().count(), 0);
    }
}
