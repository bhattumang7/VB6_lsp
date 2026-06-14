//! VB6 Project
//!
//! Represents a VB6 project with its parsed VBP file.

use std::path::{Path, PathBuf};
use super::vbp_parser::{ProjectMember, VbpFile, VbpParseError};

/// Statistics for a single VB6 project
#[derive(Debug, Clone, Default)]
pub struct ProjectStats {
    pub module_count: usize,
    pub class_count: usize,
    pub form_count: usize,
    pub user_control_count: usize,
    pub property_page_count: usize,
    pub user_document_count: usize,
    pub designer_count: usize,
}

impl ProjectStats {
    pub fn total_files(&self) -> usize {
        self.module_count
            + self.class_count
            + self.form_count
            + self.user_control_count
            + self.property_page_count
            + self.user_document_count
            + self.designer_count
    }
}

/// A VB6 project loaded from a .vbp file
#[derive(Debug)]
pub struct Vb6Project {
    /// Path to the .vbp file
    vbp_path: PathBuf,
    /// The parsed VBP file
    pub vbp: VbpFile,
}

impl Vb6Project {
    /// Create a new project from a VBP file path
    pub fn from_vbp(vbp_path: &Path) -> Result<Self, VbpParseError> {
        let vbp = VbpFile::parse(vbp_path)?;
        Ok(Self { vbp_path: vbp_path.to_path_buf(), vbp })
    }

    /// Create a new project from an already-parsed VBP
    #[allow(dead_code)]
    pub fn from_parsed_vbp(vbp_path: PathBuf, vbp: VbpFile) -> Self {
        Self { vbp_path, vbp }
    }

    /// Path to the .vbp file
    pub fn vbp_path(&self) -> &Path {
        &self.vbp_path
    }

    /// Directory containing the .vbp file
    #[allow(dead_code)]
    pub fn root_dir(&self) -> &Path {
        self.vbp_path.parent().unwrap_or(Path::new("."))
    }

    /// Get the project name
    pub fn name(&self) -> &str {
        &self.vbp.name
    }

    /// Get all source files in the project
    pub fn source_files(&self) -> impl Iterator<Item = &ProjectMember> {
        self.vbp.all_source_files()
    }

    /// Find a member by logical name (case-insensitive)
    #[allow(dead_code)]
    pub fn get_member_by_name(&self, name: &str) -> Option<&ProjectMember> {
        self.vbp.find_member_by_name(name)
    }

    /// Project statistics (counts per member type)
    pub fn stats(&self) -> ProjectStats {
        ProjectStats {
            module_count: self.vbp.modules.len(),
            class_count: self.vbp.classes.len(),
            form_count: self.vbp.forms.len(),
            user_control_count: self.vbp.user_controls.len(),
            property_page_count: self.vbp.property_pages.len(),
            user_document_count: self.vbp.user_documents.len(),
            designer_count: self.vbp.designers.len(),
        }
    }
}

#[cfg(test)]
mod tests {
    use super::*;
    use std::path::Path;

    #[test]
    fn test_project_from_vbp() {
        let content = r#"
Type=Exe
Name="TestProject"
Module=ModMain; ModMain.bas
Class=clsDatabase; clsDatabase.cls
Form=frmMain.frm
"#;
        let vbp = VbpFile::parse_content(Path::new("C:\\Projects\\Test.vbp"), content).unwrap();
        let project = Vb6Project::from_parsed_vbp(
            PathBuf::from("C:\\Projects\\Test.vbp"),
            vbp,
        );

        assert_eq!(project.name(), "TestProject");
        assert_eq!(project.source_files().count(), 3);

        let stats = project.stats();
        assert_eq!(stats.module_count, 1);
        assert_eq!(stats.class_count, 1);
        assert_eq!(stats.form_count, 1);
        assert_eq!(stats.total_files(), 3);
    }
}
