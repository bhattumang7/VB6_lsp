//! VBP Project File Parser
//!
//! Parses Visual Basic 6 project files (.vbp) to extract project structure.
//! VBP files are INI-style text files with key=value pairs.
//!
//! Features ported from vb6parse-master:
//! - SubProject references (*\A<path> format)
//! - Custom property sections ([MS Transaction Server], etc.)
//! - Strongly-typed properties (compilation, version info, threading)
//! - UUID validation for type library references

use std::collections::HashMap;
use std::path::{Path, PathBuf};
use uuid::Uuid;

use crate::utils::VB6FileReader;

/// A parsed VBP project file
#[derive(Debug, Clone)]
pub struct VbpFile {
    /// Path to the .vbp file itself
    pub path: PathBuf,

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

    /// Type library references (including SubProject references)
    pub references: Vec<TypeLibReference>,

    /// OCX/ActiveX object references
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

/// Project member (module, class, form, etc.)
#[derive(Debug, Clone)]
pub struct ProjectMember {
    /// Logical name in the project (e.g., "ModMain")
    pub name: String,

    /// Relative path from VBP location (e.g., "Utils\ModMain.bas")
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

/// Reference to a type library or sub-project
#[derive(Debug, Clone)]
pub enum TypeLibReference {
    /// Compiled type library reference (.tlb, .olb, .dll)
    Compiled {
        /// Validated UUID of the type library
        uuid: Uuid,
        /// Version (major.minor format)
        version: String,
        /// Locale ID
        lcid: String,
        /// Path to the type library file
        path: Option<PathBuf>,
        /// Description/name of the library
        description: String,
    },
    /// Reference to another VB6 project
    SubProject {
        /// Path to the referenced .vbp file
        path: PathBuf,
    },
}

impl TypeLibReference {
    /// Get the UUID if this is a compiled reference
    pub fn uuid(&self) -> Option<&Uuid> {
        match self {
            TypeLibReference::Compiled { uuid, .. } => Some(uuid),
            TypeLibReference::SubProject { .. } => None,
        }
    }

    /// Get the description
    pub fn description(&self) -> &str {
        match self {
            TypeLibReference::Compiled { description, .. } => description,
            TypeLibReference::SubProject { path } => {
                path.file_stem()
                    .and_then(|s| s.to_str())
                    .unwrap_or("SubProject")
            }
        }
    }

    /// Check if this is a sub-project reference
    pub fn is_subproject(&self) -> bool {
        matches!(self, TypeLibReference::SubProject { .. })
    }
}

/// Reference to an OCX/ActiveX control
#[derive(Debug, Clone)]
pub struct ObjectReference {
    /// Validated UUID of the control
    pub uuid: Option<Uuid>,
    /// Raw GUID string (preserved if UUID parsing fails)
    pub guid_string: String,
    /// Version
    pub version: String,
    /// Filename (e.g., "MSCOMCTL.OCX")
    pub filename: Option<String>,
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
            path: vbp_path.to_path_buf(),
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

        for (line_num, line) in content.lines().enumerate() {
            let line = line.trim();

            // Skip empty lines
            if line.is_empty() {
                continue;
            }

            // Check for section headers like [MS Transaction Server]
            if line.starts_with('[') && line.ends_with(']') {
                let section_name = line[1..line.len() - 1].to_string();
                vbp.custom_sections
                    .entry(section_name.clone())
                    .or_insert_with(HashMap::new);
                current_section = Some(section_name);
                continue;
            }

            // Parse key=value
            if let Some((key, value)) = line.split_once('=') {
                let key = key.trim();
                let value = value.trim();

                // If we're in a custom section, store there
                if let Some(ref section) = current_section {
                    if let Some(section_map) = vbp.custom_sections.get_mut(section) {
                        section_map.insert(key.to_string(), value.to_string());
                    }
                    continue;
                }

                // Parse standard VBP properties
                match key {
                    "Type" => {
                        vbp.project_type = ProjectType::from_str(value);
                    }
                    "Name" | "Title" => {
                        vbp.name = unquote(value);
                    }
                    "Module" => {
                        if let Some(member) = parse_member(value, vbp_dir, ".bas") {
                            vbp.modules.push(member);
                        }
                    }
                    "Class" => {
                        if let Some(member) = parse_member(value, vbp_dir, ".cls") {
                            vbp.classes.push(member);
                        }
                    }
                    "Form" => {
                        if let Some(member) = parse_member(value, vbp_dir, ".frm") {
                            vbp.forms.push(member);
                        }
                    }
                    "UserControl" => {
                        if let Some(member) = parse_member(value, vbp_dir, ".ctl") {
                            vbp.user_controls.push(member);
                        }
                    }
                    "PropertyPage" => {
                        if let Some(member) = parse_member(value, vbp_dir, ".pag") {
                            vbp.property_pages.push(member);
                        }
                    }
                    "UserDocument" => {
                        if let Some(member) = parse_member(value, vbp_dir, ".dob") {
                            vbp.user_documents.push(member);
                        }
                    }
                    "Designer" => {
                        if let Some(member) = parse_member(value, vbp_dir, ".dsr") {
                            vbp.designers.push(member);
                        }
                    }
                    "RelatedDoc" => {
                        let path = vbp_dir.join(value);
                        vbp.related_documents.push(path);
                    }
                    "Reference" => {
                        if let Some(reference) = parse_reference(value, vbp_dir) {
                            vbp.references.push(reference);
                        }
                    }
                    "Object" => {
                        if let Some(object) = parse_object(value) {
                            vbp.objects.push(object);
                        }
                    }
                    "Startup" => {
                        let startup_val = unquote(value);
                        // Handle VB6's special "none" indicators
                        if startup_val != "(None)" && !startup_val.is_empty() {
                            vbp.startup = Some(startup_val);
                        }
                    }
                    "ExeName32" | "ExeName" => {
                        vbp.exe_name = Some(unquote(value));
                    }
                    // Version info
                    "MajorVer" => {
                        vbp.version_info.major = value.parse().unwrap_or(0);
                    }
                    "MinorVer" => {
                        vbp.version_info.minor = value.parse().unwrap_or(0);
                    }
                    "RevisionVer" => {
                        vbp.version_info.revision = value.parse().unwrap_or(0);
                    }
                    "AutoIncrementVer" => {
                        vbp.version_info.auto_increment = value.parse().unwrap_or(0);
                    }
                    "VersionCompanyName" => {
                        vbp.version_info.company_name = Some(unquote(value));
                    }
                    "VersionFileDescription" => {
                        vbp.version_info.file_description = Some(unquote(value));
                    }
                    "VersionLegalCopyright" => {
                        vbp.version_info.legal_copyright = Some(unquote(value));
                    }
                    "VersionLegalTrademarks" => {
                        vbp.version_info.legal_trademarks = Some(unquote(value));
                    }
                    "VersionProductName" => {
                        vbp.version_info.product_name = Some(unquote(value));
                    }
                    "VersionComments" => {
                        vbp.version_info.comments = Some(unquote(value));
                    }
                    // Compilation settings
                    "CompilationType" => {
                        vbp.compilation.compilation_type = match value.parse::<i32>().unwrap_or(-1)
                        {
                            0 => CompilationType::NativeCode,
                            _ => CompilationType::PCode,
                        };
                    }
                    "OptimizationType" => {
                        vbp.compilation.optimization_type =
                            match value.parse::<i32>().unwrap_or(0) {
                                0 => OptimizationType::None,
                                1 => OptimizationType::FavorFastCode,
                                2 => OptimizationType::FavorSmallCode,
                                _ => OptimizationType::FavorFastCode,
                            };
                    }
                    "FavorPentiumPro(tm)" => {
                        vbp.compilation.favor_pentium_pro = parse_bool(value);
                    }
                    "CodeViewDebugInfo" => {
                        vbp.compilation.code_view_debug_info = parse_bool(value);
                    }
                    "NoAliasing" => {
                        vbp.compilation.no_aliasing = parse_bool(value);
                    }
                    "BoundsCheck" => {
                        vbp.compilation.bounds_check = parse_bool(value);
                    }
                    "OverflowCheck" => {
                        vbp.compilation.overflow_check = parse_bool(value);
                    }
                    "FlPointCheck" => {
                        vbp.compilation.floating_point_check = parse_bool(value);
                    }
                    "FDIVCheck" => {
                        vbp.compilation.fdiv_check = parse_bool(value);
                    }
                    "UnroundedFP" => {
                        vbp.compilation.unrounded_fp = parse_bool(value);
                    }
                    "CondComp" => {
                        let cond = unquote(value);
                        if !cond.is_empty() {
                            vbp.compilation.conditional_compile = Some(cond);
                        }
                    }
                    // Threading settings
                    "StartMode" => {
                        vbp.threading.start_mode = match value.parse::<i32>().unwrap_or(0) {
                            1 => StartMode::Automation,
                            _ => StartMode::StandAlone,
                        };
                    }
                    "Unattended" => {
                        vbp.threading.unattended = parse_bool(value);
                    }
                    "Retained" => {
                        vbp.threading.retained = parse_bool(value);
                    }
                    "ThreadPerObject" => {
                        let val = value.parse::<i32>().unwrap_or(-1);
                        vbp.threading.thread_per_object = if val < 0 { None } else { Some(val as u16) };
                    }
                    "MaxNumberOfThreads" => {
                        vbp.threading.max_threads = value.parse().unwrap_or(1);
                    }
                    "ThreadingModel" => {
                        vbp.threading.threading_model = match value.parse::<i32>().unwrap_or(1) {
                            0 => ThreadingModel::SingleThreaded,
                            _ => ThreadingModel::ApartmentThreaded,
                        };
                    }
                    // Compatibility settings
                    "CompatibleMode" => {
                        let mode_val = unquote(value);
                        vbp.compatibility.mode = match mode_val.parse::<i32>().unwrap_or(1) {
                            0 => CompatibilityMode::NoCompatibility,
                            2 => CompatibilityMode::Binary,
                            _ => CompatibilityMode::Project,
                        };
                    }
                    "CompatibleEXE32" => {
                        let path = unquote(value);
                        if !path.is_empty() {
                            vbp.compatibility.compatible_exe = Some(PathBuf::from(path));
                        }
                    }
                    "NoControlUpgrade" => {
                        // NoControlUpgrade=1 means DON'T upgrade
                        vbp.compatibility.upgrade_controls = !parse_bool(value);
                    }
                    "RemoveUnusedControlInfo" => {
                        vbp.compatibility.remove_unused_control_info = parse_bool(value);
                    }
                    "ServerSupportFiles" => {
                        vbp.compatibility.server_support_files = parse_bool(value);
                    }
                    _ => {
                        // Store other properties for potential future use
                        vbp.properties.insert(key.to_string(), value.to_string());
                    }
                }
            }
        }

        Ok(vbp)
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

    /// Check if a file path belongs to this project
    pub fn contains_file(&self, file_path: &Path) -> bool {
        // Normalize paths for comparison
        let normalized = normalize_path(file_path);

        self.all_source_files()
            .any(|member| normalize_path(&member.absolute_path) == normalized)
    }

    /// Find a project member by its absolute path
    pub fn find_member(&self, file_path: &Path) -> Option<&ProjectMember> {
        let normalized = normalize_path(file_path);

        self.all_source_files()
            .find(|member| normalize_path(&member.absolute_path) == normalized)
    }

    /// Find a project member by its logical name
    pub fn find_member_by_name(&self, name: &str) -> Option<&ProjectMember> {
        let name_lower = name.to_lowercase();

        self.all_source_files()
            .find(|member| member.name.to_lowercase() == name_lower)
    }

    /// Get all sub-project references
    pub fn get_subproject_references(&self) -> Vec<&TypeLibReference> {
        self.references
            .iter()
            .filter(|r| r.is_subproject())
            .collect()
    }

    /// Get all compiled type library references
    pub fn get_compiled_references(&self) -> Vec<&TypeLibReference> {
        self.references
            .iter()
            .filter(|r| !r.is_subproject())
            .collect()
    }

    /// Get a custom section's properties
    pub fn get_custom_section(&self, name: &str) -> Option<&HashMap<String, String>> {
        self.custom_sections.get(name)
    }
}

/// Parse a project member entry (Module, Class, Form, etc.)
/// Format: "name; path" or just "path" (name derived from filename)
fn parse_member(value: &str, vbp_dir: &Path, default_ext: &str) -> Option<ProjectMember> {
    let (name, relative_path) = if let Some((n, p)) = value.split_once(';') {
        (n.trim().to_string(), PathBuf::from(p.trim()))
    } else {
        // No semicolon - value is just the path, derive name from filename
        let path = PathBuf::from(value.trim());
        let name = path
            .file_stem()
            .and_then(|s| s.to_str())
            .unwrap_or("Unknown")
            .to_string();
        (name, path)
    };

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

    Some(ProjectMember {
        name,
        relative_path,
        absolute_path,
    })
}

/// Parse a Reference entry
/// Format: *\G{GUID}#version#lcid#path#description (compiled)
/// Format: *\A<path> (sub-project)
fn parse_reference(value: &str, vbp_dir: &Path) -> Option<TypeLibReference> {
    // Check for sub-project reference: *\A<path>
    if value.starts_with("*\\A") {
        let path_str = value.trim_start_matches("*\\A").trim();
        let path = if Path::new(path_str).is_absolute() {
            PathBuf::from(path_str)
        } else {
            vbp_dir.join(path_str)
        };
        return Some(TypeLibReference::SubProject { path });
    }

    // Parse compiled reference: *\G{GUID}#version#lcid#path#description
    let value = value.trim_start_matches("*\\G");

    let parts: Vec<&str> = value.split('#').collect();
    if parts.len() < 5 {
        return None;
    }

    // Parse and validate UUID
    let guid_str = parts[0].trim_matches(|c| c == '{' || c == '}');
    let uuid = match Uuid::parse_str(guid_str) {
        Ok(uuid) => uuid,
        Err(_) => return None, // Skip invalid UUIDs
    };

    let version = parts[1].to_string();
    let lcid = parts[2].to_string();
    let path = if parts[3].is_empty() {
        None
    } else {
        Some(PathBuf::from(parts[3]))
    };
    let description = parts[4].to_string();

    Some(TypeLibReference::Compiled {
        uuid,
        version,
        lcid,
        path,
        description,
    })
}

/// Parse an Object entry
/// Format: {GUID}#version#0; filename
fn parse_object(value: &str) -> Option<ObjectReference> {
    let (guid_part, filename) = if let Some((g, f)) = value.split_once(';') {
        (g.trim(), Some(f.trim().to_string()))
    } else {
        (value.trim(), None)
    };

    let parts: Vec<&str> = guid_part.split('#').collect();
    if parts.is_empty() {
        return None;
    }

    let guid_str = parts[0].trim_matches(|c| c == '{' || c == '}');
    let uuid = Uuid::parse_str(guid_str).ok();
    let version = parts.get(1).unwrap_or(&"1.0").to_string();

    Some(ObjectReference {
        uuid,
        guid_string: guid_str.to_string(),
        version,
        filename,
    })
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

/// Normalize a path for comparison (lowercase on Windows, canonicalize if possible)
fn normalize_path(path: &Path) -> PathBuf {
    // Try to canonicalize, fall back to the original path
    let path = path.canonicalize().unwrap_or_else(|_| path.to_path_buf());

    // On Windows, paths are case-insensitive
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

        // Check module details
        assert_eq!(vbp.modules[0].name, "ModMain");
        assert_eq!(vbp.modules[1].name, "ModUtils");
        assert_eq!(
            vbp.modules[1].relative_path,
            PathBuf::from("Utils\\ModUtils.bas")
        );
    }

    #[test]
    fn test_parse_compiled_reference() {
        let ref_str =
            "*\\G{00020430-0000-0000-C000-000000000046}#2.0#0#C:\\Windows\\System32\\stdole2.tlb#OLE Automation";
        let reference = parse_reference(ref_str, Path::new("C:\\Projects")).unwrap();

        match reference {
            TypeLibReference::Compiled {
                uuid,
                version,
                description,
                ..
            } => {
                assert_eq!(
                    uuid,
                    Uuid::parse_str("00020430-0000-0000-C000-000000000046").unwrap()
                );
                assert_eq!(version, "2.0");
                assert_eq!(description, "OLE Automation");
            }
            _ => panic!("Expected compiled reference"),
        }
    }

    #[test]
    fn test_parse_subproject_reference() {
        let ref_str = "*\\ACommonLib.vbp";
        let reference = parse_reference(ref_str, Path::new("C:\\Projects")).unwrap();

        match reference {
            TypeLibReference::SubProject { path } => {
                assert!(path.ends_with("CommonLib.vbp"));
            }
            _ => panic!("Expected subproject reference"),
        }
    }

    #[test]
    fn test_parse_object() {
        let obj_str = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0; MSCOMCTL.OCX";
        let object = parse_object(obj_str).unwrap();

        assert!(object.uuid.is_some());
        assert_eq!(object.guid_string, "831FDD16-0C5C-11D2-A9FC-0000F8754DA1");
        assert_eq!(object.version, "2.0");
        assert_eq!(object.filename, Some("MSCOMCTL.OCX".to_string()));
    }

    #[test]
    fn test_member_without_semicolon() {
        // Some VBP files use just the filename without "name; path" format
        let member = parse_member("frmMain.frm", Path::new("C:\\Projects"), ".frm").unwrap();

        assert_eq!(member.name, "frmMain");
        assert_eq!(member.relative_path, PathBuf::from("frmMain.frm"));
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
