//! VB6 Project
//!
//! Represents a VB6 project with its parsed VBP file and symbol tables.

use std::collections::HashMap;
use std::path::{Path, PathBuf};

use tower_lsp::lsp_types::{Location, Url};

use crate::analysis::{SymbolKind, SymbolTable, Visibility};

use super::vbp_parser::{ProjectMember, VbpFile, VbpParseError};

/// A VB6 project loaded from a .vbp file
#[derive(Debug)]
pub struct Vb6Project {
    /// The parsed VBP file
    pub vbp: VbpFile,

    /// Symbol tables for each file (keyed by absolute path)
    symbol_tables: HashMap<PathBuf, SymbolTable>,

    /// Index of public symbols: lowercase name -> (file_path, symbol_name)
    /// This is rebuilt when symbol tables change
    public_symbol_index: HashMap<String, Vec<(PathBuf, String)>>,
}

impl Vb6Project {
    /// Create a new project from a VBP file path
    pub fn from_vbp(vbp_path: &Path) -> Result<Self, VbpParseError> {
        let vbp = VbpFile::parse(vbp_path)?;
        Ok(Self::from_parsed_vbp(vbp))
    }

    /// Create a new project from an already-parsed VBP
    pub fn from_parsed_vbp(vbp: VbpFile) -> Self {
        Self {
            vbp,
            symbol_tables: HashMap::new(),
            public_symbol_index: HashMap::new(),
        }
    }

    /// Get the project name
    pub fn name(&self) -> &str {
        &self.vbp.name
    }

    /// Get the VBP file path
    pub fn vbp_path(&self) -> &Path {
        &self.vbp.path
    }

    /// Get the project root directory (where the VBP is located)
    pub fn root_dir(&self) -> &Path {
        self.vbp.path.parent().unwrap_or(Path::new("."))
    }

    /// Check if a file belongs to this project
    pub fn contains_file(&self, file_path: &Path) -> bool {
        self.vbp.contains_file(file_path)
    }

    /// Get a project member by file path
    pub fn get_member(&self, file_path: &Path) -> Option<&ProjectMember> {
        self.vbp.find_member(file_path)
    }

    /// Get a project member by name (e.g., "ModMain", "clsDatabase")
    pub fn get_member_by_name(&self, name: &str) -> Option<&ProjectMember> {
        self.vbp.find_member_by_name(name)
    }

    /// Get all source files in the project
    pub fn source_files(&self) -> impl Iterator<Item = &ProjectMember> {
        self.vbp.all_source_files()
    }

    /// Store a symbol table for a file
    pub fn set_symbol_table(&mut self, file_path: PathBuf, table: SymbolTable) {
        self.symbol_tables.insert(file_path, table);
        self.rebuild_public_index();
    }

    /// Get a symbol table for a file
    pub fn get_symbol_table(&self, file_path: &Path) -> Option<&SymbolTable> {
        self.symbol_tables.get(file_path)
    }

    /// Remove a symbol table (when file is closed or deleted)
    pub fn remove_symbol_table(&mut self, file_path: &Path) {
        self.symbol_tables.remove(file_path);
        self.rebuild_public_index();
    }

    /// Rebuild the public symbol index from all loaded symbol tables
    fn rebuild_public_index(&mut self) {
        self.public_symbol_index.clear();

        for (file_path, table) in &self.symbol_tables {
            // Get all public module-level symbols
            for symbol in table.module_symbols() {
                if symbol.visibility == Visibility::Public {
                    let key = symbol.name.to_lowercase();
                    self.public_symbol_index
                        .entry(key)
                        .or_default()
                        .push((file_path.clone(), symbol.name.clone()));
                }
            }
        }
    }

    /// Find a public symbol by name across all files in the project
    pub fn find_public_symbol(&self, name: &str) -> Option<Location> {
        let key = name.to_lowercase();

        if let Some(locations) = self.public_symbol_index.get(&key) {
            // Return the first match
            if let Some((file_path, symbol_name)) = locations.first() {
                if let Some(table) = self.symbol_tables.get(file_path) {
                    if let Some(symbol) = table.lookup_symbol(symbol_name, table.module_scope) {
                        let range = symbol.name_range.to_lsp();
                        return Some(Location {
                            uri: table.uri.clone(),
                            range,
                        });
                    }
                }
            }
        }

        None
    }

    /// Find all public symbols matching a prefix (for completion)
    pub fn find_public_symbols_with_prefix(&self, prefix: &str) -> Vec<(&str, &Path, SymbolKind)> {
        let prefix_lower = prefix.to_lowercase();
        let mut results = Vec::new();

        for (file_path, table) in &self.symbol_tables {
            for symbol in table.module_symbols() {
                if symbol.visibility == Visibility::Public
                    && symbol.name.to_lowercase().starts_with(&prefix_lower)
                {
                    results.push((symbol.name.as_str(), file_path.as_path(), symbol.kind));
                }
            }
        }

        results
    }

    /// Get all public symbols in the project (for workspace symbol search)
    pub fn all_public_symbols(&self) -> Vec<(&str, &Path, SymbolKind)> {
        let mut results = Vec::new();

        for (file_path, table) in &self.symbol_tables {
            for symbol in table.module_symbols() {
                if symbol.visibility == Visibility::Public {
                    results.push((symbol.name.as_str(), file_path.as_path(), symbol.kind));
                }
            }
        }

        results
    }

    /// Resolve a symbol reference to a module/class in this project
    /// E.g., "ModMain" -> Location of ModMain.bas
    ///       "clsDatabase" -> Location of clsDatabase.cls
    pub fn resolve_module_reference(&self, name: &str) -> Option<Location> {
        // First check if it's a module/class name
        if let Some(member) = self.vbp.find_member_by_name(name) {
            // Return location at the start of the file
            if let Ok(uri) = Url::from_file_path(&member.absolute_path) {
                return Some(Location {
                    uri,
                    range: tower_lsp::lsp_types::Range::default(),
                });
            }
        }

        None
    }

    /// Get statistics about the project
    pub fn stats(&self) -> ProjectStats {
        ProjectStats {
            module_count: self.vbp.modules.len(),
            class_count: self.vbp.classes.len(),
            form_count: self.vbp.forms.len(),
            user_control_count: self.vbp.user_controls.len(),
            loaded_symbol_tables: self.symbol_tables.len(),
            indexed_public_symbols: self.public_symbol_index.len(),
        }
    }
}

/// Statistics about a project
#[derive(Debug, Clone)]
pub struct ProjectStats {
    pub module_count: usize,
    pub class_count: usize,
    pub form_count: usize,
    pub user_control_count: usize,
    pub loaded_symbol_tables: usize,
    pub indexed_public_symbols: usize,
}

#[cfg(test)]
mod tests {
    use super::*;

    fn create_test_vbp() -> VbpFile {
        let content = r#"
Type=Exe
Name="TestProject"
Module=ModMain; ModMain.bas
Class=clsDatabase; clsDatabase.cls
Form=frmMain.frm
"#;
        VbpFile::parse_content(Path::new("C:\\Projects\\Test.vbp"), content).unwrap()
    }

    #[test]
    fn test_project_from_vbp() {
        let vbp = create_test_vbp();
        let project = Vb6Project::from_parsed_vbp(vbp);

        assert_eq!(project.name(), "TestProject");
        assert_eq!(project.source_files().count(), 3);
    }

    #[test]
    fn test_member_lookup() {
        let vbp = create_test_vbp();
        let project = Vb6Project::from_parsed_vbp(vbp);

        // By name
        let member = project.get_member_by_name("ModMain");
        assert!(member.is_some());
        assert_eq!(member.unwrap().name, "ModMain");

        // Case insensitive
        let member = project.get_member_by_name("modmain");
        assert!(member.is_some());
    }

    #[test]
    fn test_project_stats() {
        let vbp = create_test_vbp();
        let project = Vb6Project::from_parsed_vbp(vbp);
        let stats = project.stats();

        assert_eq!(stats.module_count, 1);
        assert_eq!(stats.class_count, 1);
        assert_eq!(stats.form_count, 1);
    }
}
