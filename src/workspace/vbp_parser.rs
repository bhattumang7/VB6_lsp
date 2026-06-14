//! VBP Project File Parser
//!
//! Parses Visual Basic 6 project files (.vbp) to extract project structure.
//! VBP files are INI-style text files with key=value pairs.
//!
//! Features ported from previous implementations:
//! - Custom property sections ([MS Transaction Server], etc.)
//! - Strongly-typed properties (compilation, version info, threading)

use std::collections::HashMap;
use std::path::{Path, PathBuf};

use crate::utils::VB6FileReader;

/// A parsed VBP project file
#[derive(Debug, Clone)]
pub struct VbpFile {
    /// Project type (Exe, OleDll, Control, OleExe)
    pub project_type: ProjectType,

    /// Project name/title
    pub name: String,

    /// Standard modules (.bas files)
    pub modules: Vec<ProjectMember>,

    /// Class modules (.cls files)
    pub classes: Vec<ProjectMember>,

    /// Forms (.frm files)
    pub forms: Vec<ProjectMember>,

    /// User controls (.ctl files)
    pub user_controls: Vec<ProjectMember>,

    /// Property pages (.pag files)
    pub property_pages: Vec<ProjectMember>,

    /// User documents (.dob files)
    pub user_documents: Vec<ProjectMember>,

    /// Designers (.dsr files)
    pub designers: Vec<ProjectMember>,

    /// Related documents (not code, but tracked)
    pub related_documents: Vec<PathBuf>,

    /// Type library and subproject references (Reference= lines)
    pub references: Vec<TypeLibReference>,

    /// ActiveX/OCX object references (Object= lines)
    pub objects: Vec<ObjectReference>,

    /// Startup form or "Sub Main"
    pub startup: Option<String>,

    /// Output executable name
    pub exe_name: Option<String>,

    /// Version information
    pub version_info: VersionInfo,

    /// Compilation settings
    pub compilation: CompilationSettings,

    /// Threading model settings
    pub threading: ThreadingSettings,

    /// Compatibility settings
    pub compatibility: CompatibilitySettings,

    /// Custom property sections (e.g., [MS Transaction Server])
    pub custom_sections: HashMap<String, HashMap<String, String>>,

    /// All raw key-value pairs (for properties we don't specifically handle)
    pub properties: HashMap<String, String>,
}

/// A type library or subproject reference from a `Reference=` line
#[derive(Debug, Clone)]
pub enum TypeLibReference {
    /// A compiled type library: `*\G{GUID}#major.minor#lcid#path#description`
    Compiled {
        guid: String,
        major: u16,
        minor: u16,
        lcid: u32,
        path: PathBuf,
        description: String,
    },
    /// A VBP subproject reference: `*\Apath.vbp` or bare path
    Subproject { path: PathBuf },
}

/// An ActiveX/OCX object reference from an `Object=` line
///
/// Format: `{GUID}#version#0; filename.ocx`
#[derive(Debug, Clone)]
pub struct ObjectReference {
    pub guid: String,
    pub version: String,
    pub filename: String,
}

/// Project member (module, class, form, etc.)
#[derive(Debug, Clone)]
pub struct ProjectMember {
    /// Logical name from the VBP (e.g. `modMain`, `clsDatabase`)
    pub name: String,
    /// Path relative to the VBP file (as written in the file)
    pub relative_path: PathBuf,
    /// Absolute path (resolved when VBP is parsed)
    pub absolute_path: PathBuf,
}

/// Type of VB6 project
#[derive(Debug, Clone, Copy, PartialEq, Eq, Default)]
pub enum ProjectType {
    /// Standard EXE
    #[default]
    Exe,
    /// ActiveX DLL
    OleDll,
    /// ActiveX EXE
    OleExe,
    /// ActiveX Control
    Control,
}

impl ProjectType {
    fn from_str(s: &str) -> Self {
        match s.to_lowercase().as_str() {
            "exe" => ProjectType::Exe,
            "oledll" => ProjectType::OleDll,
            "oleexe" => ProjectType::OleExe,
            "control" => ProjectType::Control,
            _ => ProjectType::Exe,
        }
    }
}

/// Version information for the project
#[derive(Debug, Clone, Default)]
pub struct VersionInfo {
    /// Major version number
    pub major: u16,
    /// Minor version number
    pub minor: u16,
    /// Revision number
    pub revision: u16,
    /// Auto-increment revision on each compile
    pub auto_increment: u16,
    /// Company name
    pub company_name: Option<String>,
    /// File description
    pub file_description: Option<String>,
    /// Legal copyright
    pub legal_copyright: Option<String>,
    /// Legal trademarks
    pub legal_trademarks: Option<String>,
    /// Product name
    pub product_name: Option<String>,
    /// Comments
    pub comments: Option<String>,
}

/// Compilation settings
#[derive(Debug, Clone, Default)]
pub struct CompilationSettings {
    /// Compilation type: PCode (-1) or NativeCode (0)
    pub compilation_type: CompilationType,
    /// Optimization type for native code
    pub optimization_type: OptimizationType,
    /// Favor Pentium Pro instructions
    pub favor_pentium_pro: bool,
    /// Create CodeView debug info
    pub code_view_debug_info: bool,
    /// Assume no aliasing
    pub no_aliasing: bool,
    /// Array bounds checking
    pub bounds_check: bool,
    /// Integer overflow checking
    pub overflow_check: bool,
    /// Floating point error checking
    pub floating_point_check: bool,
    /// Pentium FDIV bug checking
    pub fdiv_check: bool,
    /// Allow unrounded floating point operations
    pub unrounded_fp: bool,
    /// Conditional compilation arguments
    pub conditional_compile: Option<String>,
}

/// Compilation type
#[derive(Debug, Clone, Copy, PartialEq, Eq, Default)]
pub enum CompilationType {
    /// P-Code (interpreted bytecode)
    #[default]
    PCode,
    /// Native code compilation
    NativeCode,
}

/// Optimization type for native code compilation
#[derive(Debug, Clone, Copy, PartialEq, Eq, Default)]
pub enum OptimizationType {
    /// No optimization
    None,
    /// Favor fast code (default)
    #[default]
    FavorFastCode,
    /// Favor small code
    FavorSmallCode,
}

/// Threading model settings
#[derive(Debug, Clone, Default)]
pub struct ThreadingSettings {
    /// Start mode (StandAlone or Automation)
    pub start_mode: StartMode,
    /// Unattended execution (no UI)
    pub unattended: bool,
    /// Retain DLL in memory
    pub retained: bool,
    /// Thread per object (-1 means use pool)
    pub thread_per_object: Option<u16>,
    /// Maximum number of threads
    pub max_threads: u16,
    /// Threading model
    pub threading_model: ThreadingModel,
}

/// Start mode
#[derive(Debug, Clone, Copy, PartialEq, Eq, Default)]
pub enum StartMode {
    /// Stand-alone application
    #[default]
    StandAlone,
    /// ActiveX automation component
    Automation,
}

/// Threading model
#[derive(Debug, Clone, Copy, PartialEq, Eq, Default)]
pub enum ThreadingModel {
    /// Single threaded
    SingleThreaded,
    /// Apartment threaded (default for VB6)
    #[default]
    ApartmentThreaded,
}

/// Compatibility settings
#[derive(Debug, Clone, Default)]
pub struct CompatibilitySettings {
    /// Compatibility mode
    pub mode: CompatibilityMode,
    /// Path to compatible executable
    pub compatible_exe: Option<PathBuf>,
    /// Upgrade ActiveX controls
    pub upgrade_controls: bool,
    /// Remove unused control info
    pub remove_unused_control_info: bool,
    /// Generate server support files
    pub server_support_files: bool,
}

/// Compatibility mode
#[derive(Debug, Clone, Copy, PartialEq, Eq, Default)]
pub enum CompatibilityMode {
    /// No compatibility - new GUIDs each compile
    NoCompatibility,
    /// Project compatibility - maintain type library ID
    #[default]
    Project,
    /// Binary compatibility - maintain class IDs
    Binary,
}

/// Error type for VBP parsing
#[derive(Debug, Clone)]
pub struct VbpParseError {
    pub message: String,
    pub line: Option<usize>,
}

impl std::fmt::Display for VbpParseError {
    fn fmt(&self, f: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        match self.line {
            Some(line) => write!(f, "VBP parse error at line {}: {}", line, self.message),
            None => write!(f, "VBP parse error: {}", self.message),
        }
    }
}

impl std::error::Error for VbpParseError {}

impl VbpFile {
    /// Parse a VBP file from its path
    ///
    /// Automatically detects encoding (UTF-8 or Windows-1252)
    pub fn parse(vbp_path: &Path) -> Result<Self, VbpParseError> {
        let file_content = VB6FileReader::read_file(vbp_path).map_err(|e| VbpParseError {
            message: format!("Failed to read file: {}", e),
            line: None,
        })?;

        if file_content.had_errors {
            tracing::warn!(
                "VBP file {} had encoding errors (detected as {})",
                vbp_path.display(),
                file_content.encoding.name()
            );
        }

        Self::parse_content(vbp_path, &file_content.text)
    }

    /// Parse VBP content (useful for testing)
    pub fn parse_content(vbp_path: &Path, content: &str) -> Result<Self, VbpParseError> {
        let vbp_dir = vbp_path.parent().unwrap_or(Path::new("."));

        let mut vbp = VbpFile {
            project_type: ProjectType::default(),
            name: vbp_path
                .file_stem()
                .and_then(|s| s.to_str())
                .unwrap_or("Untitled")
                .to_string(),
            modules: Vec::new(),
            classes: Vec::new(),
            forms: Vec::new(),
            user_controls: Vec::new(),
            property_pages: Vec::new(),
            user_documents: Vec::new(),
            designers: Vec::new(),
            related_documents: Vec::new(),
            references: Vec::new(),
            objects: Vec::new(),
            startup: None,
            exe_name: None,
            version_info: VersionInfo::default(),
            compilation: CompilationSettings::default(),
            threading: ThreadingSettings {
                max_threads: 1,
                ..Default::default()
            },
            compatibility: CompatibilitySettings::default(),
            custom_sections: HashMap::new(),
            properties: HashMap::new(),
        };

        let mut current_section: Option<String> = None;

        for (_line_num, line) in content.lines().enumerate() {
            let line = line.trim();

            // Skip empty lines
            if line.is_empty() {
                continue;
            }

            // Check for section headers like [MS Transaction Server]
            if let Some(section_name) = parse_section_header(line) {
                vbp.custom_sections
                    .entry(section_name.clone())
                    .or_insert_with(HashMap::new);
                current_section = Some(section_name);
                continue;
            }

            let Some((key, value)) = line.split_once('=') else {
                continue;
            };
            let key = key.trim();
            let value = value.trim();

            // If we're in a custom section, store there
            if let Some(ref section) = current_section {
                if let Some(section_map) = vbp.custom_sections.get_mut(section) {
                    section_map.insert(key.to_string(), value.to_string());
                }
                continue;
            }

            vbp.apply_property(key, value, vbp_dir);
        }

        Ok(vbp)
    }

    /// Apply a single standard `key=value` property to the project.
    fn apply_property(&mut self, key: &str, value: &str, vbp_dir: &Path) {
        if self.apply_member_property(key, value, vbp_dir) {
            return;
        }
        if self.apply_version_property(key, value) {
            return;
        }
        if self.apply_compilation_property(key, value) {
            return;
        }
        if self.apply_threading_property(key, value) {
            return;
        }
        if self.apply_compatibility_property(key, value) {
            return;
        }
        // Store other properties for potential future use
        self.properties.insert(key.to_string(), value.to_string());
    }

    /// Handle top-level identity, members, references and objects.
    /// Returns `true` if the key was recognized.
    fn apply_member_property(&mut self, key: &str, value: &str, vbp_dir: &Path) -> bool {
        match key {
            "Type" => self.project_type = ProjectType::from_str(value),
            "Name" | "Title" => self.name = unquote(value),
            "Module" => self.push_member(&mut |s| &mut s.modules, value, vbp_dir, ".bas"),
            "Class" => self.push_member(&mut |s| &mut s.classes, value, vbp_dir, ".cls"),
            "Form" => self.push_member(&mut |s| &mut s.forms, value, vbp_dir, ".frm"),
            "UserControl" => {
                self.push_member(&mut |s| &mut s.user_controls, value, vbp_dir, ".ctl")
            }
            "PropertyPage" => {
                self.push_member(&mut |s| &mut s.property_pages, value, vbp_dir, ".pag")
            }
            "UserDocument" => {
                self.push_member(&mut |s| &mut s.user_documents, value, vbp_dir, ".dob")
            }
            "Designer" => self.push_member(&mut |s| &mut s.designers, value, vbp_dir, ".dsr"),
            "RelatedDoc" => self.related_documents.push(vbp_dir.join(value)),
            "Reference" => {
                if let Some(r) = parse_reference(value, vbp_dir) {
                    self.references.push(r);
                }
            }
            "Object" => {
                if let Some(o) = parse_object(value) {
                    self.objects.push(o);
                }
            }
            "Startup" => {
                let startup_val = unquote(value);
                // Handle VB6's special "none" indicators
                if startup_val != "(None)" && !startup_val.is_empty() {
                    self.startup = Some(startup_val);
                }
            }
            "ExeName32" | "ExeName" => self.exe_name = Some(unquote(value)),
            _ => return false,
        }
        true
    }

    /// Push a parsed member into the vector selected by `select`.
    fn push_member(
        &mut self,
        select: &mut dyn FnMut(&mut Self) -> &mut Vec<ProjectMember>,
        value: &str,
        vbp_dir: &Path,
        default_ext: &str,
    ) {
        if let Some(member) = parse_member(value, vbp_dir, default_ext) {
            select(self).push(member);
        }
    }

    /// Handle version-info properties. Returns `true` if recognized.
    fn apply_version_property(&mut self, key: &str, value: &str) -> bool {
        let v = &mut self.version_info;
        match key {
            "MajorVer" => v.major = value.parse().unwrap_or(0),
            "MinorVer" => v.minor = value.parse().unwrap_or(0),
            "RevisionVer" => v.revision = value.parse().unwrap_or(0),
            "AutoIncrementVer" => v.auto_increment = value.parse().unwrap_or(0),
            "VersionCompanyName" => v.company_name = Some(unquote(value)),
            "VersionFileDescription" => v.file_description = Some(unquote(value)),
            "VersionLegalCopyright" => v.legal_copyright = Some(unquote(value)),
            "VersionLegalTrademarks" => v.legal_trademarks = Some(unquote(value)),
            "VersionProductName" => v.product_name = Some(unquote(value)),
            "VersionComments" => v.comments = Some(unquote(value)),
            _ => return false,
        }
        true
    }

    /// Handle compilation-settings properties. Returns `true` if recognized.
    fn apply_compilation_property(&mut self, key: &str, value: &str) -> bool {
        let c = &mut self.compilation;
        match key {
            "CompilationType" => {
                c.compilation_type = match value.parse::<i32>().unwrap_or(-1) {
                    0 => CompilationType::NativeCode,
                    _ => CompilationType::PCode,
                };
            }
            "OptimizationType" => {
                c.optimization_type = match value.parse::<i32>().unwrap_or(0) {
                    0 => OptimizationType::None,
                    1 => OptimizationType::FavorFastCode,
                    2 => OptimizationType::FavorSmallCode,
                    _ => OptimizationType::FavorFastCode,
                };
            }
            "FavorPentiumPro(tm)" => c.favor_pentium_pro = parse_bool(value),
            "CodeViewDebugInfo" => c.code_view_debug_info = parse_bool(value),
            "NoAliasing" => c.no_aliasing = parse_bool(value),
            "BoundsCheck" => c.bounds_check = parse_bool(value),
            "OverflowCheck" => c.overflow_check = parse_bool(value),
            "FlPointCheck" => c.floating_point_check = parse_bool(value),
            "FDIVCheck" => c.fdiv_check = parse_bool(value),
            "UnroundedFP" => c.unrounded_fp = parse_bool(value),
            "CondComp" => {
                let cond = unquote(value);
                if !cond.is_empty() {
                    c.conditional_compile = Some(cond);
                }
            }
            _ => return false,
        }
        true
    }

    /// Handle threading-settings properties. Returns `true` if recognized.
    fn apply_threading_property(&mut self, key: &str, value: &str) -> bool {
        let t = &mut self.threading;
        match key {
            "StartMode" => {
                t.start_mode = match value.parse::<i32>().unwrap_or(0) {
                    1 => StartMode::Automation,
                    _ => StartMode::StandAlone,
                };
            }
            "Unattended" => t.unattended = parse_bool(value),
            "Retained" => t.retained = parse_bool(value),
            "ThreadPerObject" => {
                let val = value.parse::<i32>().unwrap_or(-1);
                t.thread_per_object = if val < 0 { None } else { Some(val as u16) };
            }
            "MaxNumberOfThreads" => t.max_threads = value.parse().unwrap_or(1),
            "ThreadingModel" => {
                t.threading_model = match value.parse::<i32>().unwrap_or(1) {
                    0 => ThreadingModel::SingleThreaded,
                    _ => ThreadingModel::ApartmentThreaded,
                };
            }
            _ => return false,
        }
        true
    }

    /// Handle compatibility-settings properties. Returns `true` if recognized.
    fn apply_compatibility_property(&mut self, key: &str, value: &str) -> bool {
        let c = &mut self.compatibility;
        match key {
            "CompatibleMode" => {
                let mode_val = unquote(value);
                c.mode = match mode_val.parse::<i32>().unwrap_or(1) {
                    0 => CompatibilityMode::NoCompatibility,
                    2 => CompatibilityMode::Binary,
                    _ => CompatibilityMode::Project,
                };
            }
            "CompatibleEXE32" => {
                let path = unquote(value);
                if !path.is_empty() {
                    c.compatible_exe = Some(PathBuf::from(path));
                }
            }
            // NoControlUpgrade=1 means DON'T upgrade
            "NoControlUpgrade" => c.upgrade_controls = !parse_bool(value),
            "RemoveUnusedControlInfo" => c.remove_unused_control_info = parse_bool(value),
            "ServerSupportFiles" => c.server_support_files = parse_bool(value),
            _ => return false,
        }
        true
    }

    /// Get all source file members (modules, classes, forms, controls, etc.)
    pub fn all_source_files(&self) -> impl Iterator<Item = &ProjectMember> {
        self.modules
            .iter()
            .chain(self.classes.iter())
            .chain(self.forms.iter())
            .chain(self.user_controls.iter())
            .chain(self.property_pages.iter())
            .chain(self.user_documents.iter())
            .chain(self.designers.iter())
    }

    /// Get only the compiled type library references (TLB/DLL, not subprojects)
    pub fn get_compiled_references(&self) -> Vec<&TypeLibReference> {
        self.references
            .iter()
            .filter(|r| matches!(r, TypeLibReference::Compiled { .. }))
            .collect()
    }

    /// Get only the subproject references (.vbp files)
    pub fn get_subproject_references(&self) -> Vec<&TypeLibReference> {
        self.references
            .iter()
            .filter(|r| matches!(r, TypeLibReference::Subproject { .. }))
            .collect()
    }

    /// Find a member (any kind) by its logical name (case-insensitive)
    #[allow(dead_code)]
    pub fn find_member_by_name(&self, name: &str) -> Option<&ProjectMember> {
        let lower = name.to_lowercase();
        self.all_source_files()
            .find(|m| m.name.to_lowercase() == lower)
    }

    /// Get a named custom section (e.g., "MS Transaction Server")
    pub fn get_custom_section(&self, name: &str) -> Option<&HashMap<String, String>> {
        self.custom_sections.get(name)
    }
}

impl TypeLibReference {
    /// Return the GUID string if this is a compiled type library reference.
    pub fn uuid(&self) -> Option<&str> {
        match self {
            TypeLibReference::Compiled { guid, .. } => Some(guid.as_str()),
            TypeLibReference::Subproject { .. } => None,
        }
    }

    /// Human-readable description of the reference.
    pub fn description(&self) -> &str {
        match self {
            TypeLibReference::Compiled { description, .. } => description.as_str(),
            TypeLibReference::Subproject { path } => {
                path.to_str().unwrap_or("<subproject>")
            }
        }
    }
}

/// Parse a section header line like `[MS Transaction Server]`, returning the
/// section name without the brackets.
fn parse_section_header(line: &str) -> Option<String> {
    if line.starts_with('[') && line.ends_with(']') {
        Some(line[1..line.len() - 1].to_string())
    } else {
        None
    }
}

/// Parse a project member entry (Module, Class, Form, etc.)
/// Format: "name; path" or just "path"
fn parse_member(value: &str, vbp_dir: &Path, default_ext: &str) -> Option<ProjectMember> {
    let (name, raw_path) = if let Some((n, p)) = value.split_once(';') {
        (n.trim().to_string(), p.trim())
    } else {
        (String::new(), value.trim())
    };

    let relative_path = PathBuf::from(raw_path);

    // Ensure extension is present
    let relative_path = if relative_path.extension().is_none() {
        relative_path.with_extension(default_ext.trim_start_matches('.'))
    } else {
        relative_path
    };

    // Resolve to absolute path
    let absolute_path = if relative_path.is_absolute() {
        relative_path.clone()
    } else {
        vbp_dir.join(&relative_path)
    };

    Some(ProjectMember { name, relative_path, absolute_path })
}

/// Parse a `Reference=` value into a [`TypeLibReference`]
fn parse_reference(value: &str, vbp_dir: &Path) -> Option<TypeLibReference> {
    let value = value.strip_prefix(r"*\").unwrap_or(value);

    if let Some(rest) = value.strip_prefix('G') {
        // Compiled typelib: G{GUID}#major.minor#lcid#path#description
        let parts: Vec<&str> = rest.splitn(5, '#').collect();
        if parts.len() < 5 {
            return None;
        }
        let guid = parts[0].trim_matches(|c| c == '{' || c == '}').to_string();
        let ver: Vec<&str> = parts[1].split('.').collect();
        let major = ver.first().and_then(|v| v.parse().ok()).unwrap_or(0);
        let minor = ver.get(1).and_then(|v| v.parse().ok()).unwrap_or(0);
        let lcid = parts[2].parse().unwrap_or(0);
        let path = PathBuf::from(parts[3]);
        let description = parts[4].to_string();
        Some(TypeLibReference::Compiled { guid, major, minor, lcid, path, description })
    } else {
        // Subproject reference: A<path> or bare path
        let raw = value.strip_prefix('A').unwrap_or(value);
        let path = if Path::new(raw).is_absolute() {
            PathBuf::from(raw)
        } else {
            vbp_dir.join(raw)
        };
        Some(TypeLibReference::Subproject { path })
    }
}

/// Parse an `Object=` value into an [`ObjectReference`]
///
/// Format: `{GUID}#version#0; filename.ocx`
fn parse_object(value: &str) -> Option<ObjectReference> {
    let parts: Vec<&str> = value.splitn(4, '#').collect();
    if parts.len() < 3 {
        return None;
    }
    let guid = parts[0].trim().trim_matches(|c| c == '{' || c == '}').to_string();
    let version = parts[1].to_string();
    let filename = parts
        .get(2)
        .and_then(|p| p.split_once(';'))
        .map(|(_, f)| f.trim().to_string())
        .or_else(|| parts.get(3).map(|p| p.trim().to_string()))
        .unwrap_or_default();
    Some(ObjectReference { guid, version, filename })
}

/// Remove surrounding quotes from a string
fn unquote(s: &str) -> String {
    s.trim_matches('"').to_string()
}

/// Parse a VB6 boolean value (-1 = true, 0 = false)
fn parse_bool(s: &str) -> bool {
    let s = unquote(s);
    match s.parse::<i32>() {
        Ok(v) => v != 0,
        Err(_) => s.eq_ignore_ascii_case("true"),
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn test_parse_simple_vbp() {
        let content = r#"
Type=Exe
Name="TestProject"
Module=ModMain; ModMain.bas
Module=ModUtils; Utils\ModUtils.bas
Class=clsDatabase; clsDatabase.cls
Form=frmMain.frm
Startup="Sub Main"
ExeName32="TestProject.exe"
"#;

        let vbp = VbpFile::parse_content(Path::new("C:\\Projects\\Test.vbp"), content).unwrap();

        assert_eq!(vbp.project_type, ProjectType::Exe);
        assert_eq!(vbp.name, "TestProject");
        assert_eq!(vbp.modules.len(), 2);
        assert_eq!(vbp.classes.len(), 1);
        assert_eq!(vbp.forms.len(), 1);
        assert_eq!(vbp.startup, Some("Sub Main".to_string()));
        assert_eq!(vbp.exe_name, Some("TestProject.exe".to_string()));
    }

    #[test]
    fn test_project_type_parsing() {
        assert_eq!(ProjectType::from_str("Exe"), ProjectType::Exe);
        assert_eq!(ProjectType::from_str("OleDll"), ProjectType::OleDll);
        assert_eq!(ProjectType::from_str("OLEEXE"), ProjectType::OleExe);
        assert_eq!(ProjectType::from_str("Control"), ProjectType::Control);
    }

    #[test]
    fn test_version_info_parsing() {
        let content = r#"
Type=Exe
Name="TestProject"
MajorVer=1
MinorVer=2
RevisionVer=3
AutoIncrementVer=1
VersionCompanyName="Test Company"
VersionFileDescription="Test Description"
VersionLegalCopyright="Copyright 2024"
VersionProductName="Test Product"
"#;

        let vbp = VbpFile::parse_content(Path::new("C:\\Projects\\Test.vbp"), content).unwrap();

        assert_eq!(vbp.version_info.major, 1);
        assert_eq!(vbp.version_info.minor, 2);
        assert_eq!(vbp.version_info.revision, 3);
        assert_eq!(vbp.version_info.auto_increment, 1);
        assert_eq!(
            vbp.version_info.company_name,
            Some("Test Company".to_string())
        );
        assert_eq!(
            vbp.version_info.file_description,
            Some("Test Description".to_string())
        );
    }

    #[test]
    fn test_compilation_settings_parsing() {
        let content = r#"
Type=Exe
Name="TestProject"
CompilationType=0
OptimizationType=1
FavorPentiumPro(tm)=-1
CodeViewDebugInfo=-1
NoAliasing=-1
BoundsCheck=-1
OverflowCheck=-1
FlPointCheck=-1
FDIVCheck=-1
UnroundedFP=-1
CondComp="DEBUG=1"
"#;

        let vbp = VbpFile::parse_content(Path::new("C:\\Projects\\Test.vbp"), content).unwrap();

        assert_eq!(vbp.compilation.compilation_type, CompilationType::NativeCode);
        assert_eq!(
            vbp.compilation.optimization_type,
            OptimizationType::FavorFastCode
        );
        assert!(vbp.compilation.favor_pentium_pro);
        assert!(vbp.compilation.code_view_debug_info);
        assert!(vbp.compilation.no_aliasing);
        assert!(vbp.compilation.bounds_check);
        assert!(vbp.compilation.overflow_check);
        assert!(vbp.compilation.floating_point_check);
        assert!(vbp.compilation.fdiv_check);
        assert!(vbp.compilation.unrounded_fp);
        assert_eq!(
            vbp.compilation.conditional_compile,
            Some("DEBUG=1".to_string())
        );
    }

    #[test]
    fn test_threading_settings_parsing() {
        let content = r#"
Type=Exe
Name="TestProject"
StartMode=1
Unattended=-1
Retained=-1
ThreadPerObject=0
MaxNumberOfThreads=4
ThreadingModel=1
"#;

        let vbp = VbpFile::parse_content(Path::new("C:\\Projects\\Test.vbp"), content).unwrap();

        assert_eq!(vbp.threading.start_mode, StartMode::Automation);
        assert!(vbp.threading.unattended);
        assert!(vbp.threading.retained);
        assert_eq!(vbp.threading.thread_per_object, Some(0));
        assert_eq!(vbp.threading.max_threads, 4);
        assert_eq!(
            vbp.threading.threading_model,
            ThreadingModel::ApartmentThreaded
        );
    }

    #[test]
    fn test_custom_section_parsing() {
        let content = r#"
Type=Exe
Name="TestProject"

[MS Transaction Server]
AutoRefresh=1
"#;

        let vbp = VbpFile::parse_content(Path::new("C:\\Projects\\Test.vbp"), content).unwrap();

        let mts_section = vbp.get_custom_section("MS Transaction Server");
        assert!(mts_section.is_some());
        let mts = mts_section.unwrap();
        assert_eq!(mts.get("AutoRefresh"), Some(&"1".to_string()));
    }

    #[test]
    fn test_compatibility_settings_parsing() {
        let content = r#"
Type=Exe
Name="TestProject"
CompatibleMode="2"
CompatibleEXE32="C:\Projects\MyApp.exe"
NoControlUpgrade=1
RemoveUnusedControlInfo=-1
ServerSupportFiles=-1
"#;

        let vbp = VbpFile::parse_content(Path::new("C:\\Projects\\Test.vbp"), content).unwrap();

        assert_eq!(vbp.compatibility.mode, CompatibilityMode::Binary);
        assert_eq!(
            vbp.compatibility.compatible_exe,
            Some(PathBuf::from("C:\\Projects\\MyApp.exe"))
        );
        assert!(!vbp.compatibility.upgrade_controls); // NoControlUpgrade=1 means don't upgrade
        assert!(vbp.compatibility.remove_unused_control_info);
        assert!(vbp.compatibility.server_support_files);
    }

    #[test]
    fn test_get_subproject_references() {
        let content = r#"
Type=Exe
Name="TestProject"
Reference=*\G{00020430-0000-0000-C000-000000000046}#2.0#0#C:\Windows\System32\stdole2.tlb#OLE Automation
Reference=*\ACommonLib.vbp
Reference=*\AUtils.vbp
"#;

        let vbp = VbpFile::parse_content(Path::new("C:\\Projects\\Test.vbp"), content).unwrap();

        let subprojects = vbp.get_subproject_references();
        assert_eq!(subprojects.len(), 2);

        let compiled = vbp.get_compiled_references();
        assert_eq!(compiled.len(), 1);
    }
}
