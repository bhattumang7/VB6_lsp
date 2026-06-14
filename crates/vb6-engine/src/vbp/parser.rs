/// VB6 `.vbp` project file parser.
///
/// VB6 writes every key=value pair in the project file in a fixed order; this
/// parser reads them back in the same order (and tolerates any order).
use std::fmt;

// ---------------------------------------------------------------------------
// Public types
// ---------------------------------------------------------------------------

/// A parsed VB6 project file.
#[derive(Debug, Clone, Default)]
pub struct ProjectFile {
    /// Project type string (first line of file), e.g. `"Standard EXE"`.
    pub project_type: String,

    /// Source-code modules: Form, Module, Class, UserControl, UserDocument,
    /// PropertyPage, MDIForm entries.
    pub modules: Vec<Module>,

    /// Type-library references (`Reference=*\G{...}` lines).
    pub references: Vec<Reference>,

    /// OCX ActiveX control registrations (`Object={progid}` lines).
    pub objects: Vec<OcxObject>,

    /// Resource file (`ResFile32="path"`), if present.
    pub res_file: Option<ResFile>,

    /// `Startup="..."` — startup object name.
    pub startup: Option<String>,

    /// `HelpFile="..."` — path to .hlp / .chm file.
    pub help_file: Option<String>,

    /// `Title="..."` — EXE/DLL title.
    pub title: Option<String>,

    /// `Name="..."` — internal project name.
    pub name: Option<String>,

    /// `ExeName32="..."` — output executable name.
    pub exe_name: Option<String>,

    /// `Description="..."` — project description.
    pub description: Option<String>,

    /// `HelpContextID="..."`.
    pub help_context_id: Option<String>,

    /// `Command32="..."` — command-line string for testing.
    pub command32: Option<String>,

    /// `CompatibleMode="..."`.
    pub compatible_mode: Option<String>,

    /// `CompatibleEXE32="..."`.
    pub compatible_exe32: Option<String>,

    /// `MajorVer`, `MinorVer`, `RevisionVer`.
    pub version: Option<ProjectVersion>,

    /// `AutoIncrementVer=N`.
    pub auto_increment_ver: Option<u32>,

    /// `CompilationType=N` (0=p-code, -1=native).
    pub compilation_type: Option<i32>,

    /// `OptimizationType=N`.
    pub optimization_type: Option<i32>,

    /// Compilation flags (NoAliasing, BoundsCheck, OverflowCheck, etc.)
    pub compile_flags: CompileFlags,

    /// `StartMode=N` (0=standalone, 1=ActiveX component).
    pub start_mode: Option<i32>,

    /// `Unattended=N`.
    pub unattended: Option<i32>,

    /// `Retained=N`.
    pub retained: Option<i32>,

    /// `ThreadPerObject=N`.
    pub thread_per_object: Option<i32>,

    /// `MaxNumberOfThreads=N`.
    pub max_threads: Option<i32>,

    /// `ThreadingModel=N`.
    pub threading_model: Option<i32>,

    /// `ServerSupportFiles=N`.
    pub server_support_files: Option<i32>,

    /// `DllBaseAddress=&HN`.
    pub dll_base_address: Option<u32>,

    /// `CondComp="..."` — conditional compilation arguments.
    pub cond_comp: Option<String>,

    /// `IconForm="..."`.
    pub icon_form: Option<String>,

    /// `RequireLicenseKey=1` — require license key for controls.
    pub require_license_key: bool,

    /// `NoControlUpgrade=1` — suppress control upgrade prompt.
    pub no_control_upgrade: bool,

    /// `VersionCompatible32="N"` — version compatibility mode.
    pub version_compatible32: Option<String>,

    /// `DebugStartupOption=N` (0=wait for component, 1=start with form, etc.)
    pub debug_startup_option: Option<i32>,

    /// `DebugStartupComponent=name` — component name for debug startup.
    pub debug_startup_component: Option<String>,

    /// `UseExistingBrowser=0` — reuse browser window for WebClass.
    pub use_existing_browser: bool,

    /// Version information strings (VersionComments, VersionCompanyName, etc.)
    pub version_info: VersionInfo,

    /// Any key=value pairs not recognized above (future-proof).
    pub extra: Vec<(String, String)>,
}

/// Version numbers embedded in the project.
#[derive(Debug, Clone, Default)]
pub struct ProjectVersion {
    pub major: u32,
    pub minor: u32,
    pub revision: u32,
}

/// Compilation boolean flags stored in `CompilationType` / various keys.
#[derive(Debug, Clone, Default)]
pub struct CompileFlags {
    pub favor_pentium_pro: bool,
    pub code_view_debug_info: bool,
    pub no_aliasing: bool,
    pub bounds_check: bool,
    pub overflow_check: bool,
    pub fl_point_check: bool,
    pub fdiv_check: bool,
    pub unrounded_fp: bool,
    pub remove_unused_control_info: bool,
}

/// Version-info strings written by the VB6 version-info dialog.
#[derive(Debug, Clone, Default)]
pub struct VersionInfo {
    pub comments: Option<String>,
    pub company_name: Option<String>,
    pub file_description: Option<String>,
    pub legal_copyright: Option<String>,
    pub legal_trademarks: Option<String>,
    pub product_name: Option<String>,
}

/// A source-code module entry (Form, Module, Class, etc.).
#[derive(Debug, Clone)]
pub struct Module {
    pub kind: ModuleKind,
    /// Module name as it appears in VB6 (without the path), e.g. `Form1`.
    /// For Form/MDIForm lines this is parsed from `; Name` after the path.
    pub name: Option<String>,
    /// Path to the source file, e.g. `Form1.frm`.
    pub path: String,
}

/// VBP module-type keywords.
#[derive(Debug, Clone, PartialEq)]
pub enum ModuleKind {
    Form,
    MdiForm,
    Module,
    Class,
    UserControl,
    UserDocument,
    PropertyPage,
    Resource,
    RelatedDoc,
    /// Any other `Key=...` line treated as a module (forward-compatible).
    Other(String),
}

/// A `Reference=*\G{guid}#major.minor#lcid#path#name` type-library reference.
///
/// VB6 writes each field in order.
#[derive(Debug, Clone)]
pub struct Reference {
    /// Full raw reference string (everything after `Reference=`).
    pub raw: String,
    /// GUID extracted from the `*\G{...}` prefix.
    pub guid: Option<String>,
    /// Version string, e.g. `"2.0"` (field 2 in `#`-delimited format).
    pub version: Option<String>,
    /// Locale ID (field 3), e.g. `0`.
    pub lcid: Option<u32>,
    /// Type-library path (field 4), e.g. `..\Windows\System32\stdole2.tlb`.
    pub path: Option<String>,
    /// Human-readable name (field 5 / last `#`-delimited field).
    pub name: Option<String>,
}

/// An `Object={progid}#ver#lcid ; filename` OCX control registration.
#[derive(Debug, Clone)]
pub struct OcxObject {
    /// Full progid including version, e.g. `{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0`.
    pub progid: String,
    /// Registered filename / description after `;`.
    pub filename: Option<String>,
}

/// A `ResFile32="path"` resource-file entry.
#[derive(Debug, Clone)]
pub struct ResFile {
    pub path: String,
}

// ---------------------------------------------------------------------------
// Error
// ---------------------------------------------------------------------------

#[derive(Debug)]
pub struct VbpError {
    pub line: usize,
    pub msg: String,
}

impl fmt::Display for VbpError {
    fn fmt(&self, f: &mut fmt::Formatter<'_>) -> fmt::Result {
        write!(f, "vbp parse error at line {}: {}", self.line, self.msg)
    }
}

impl std::error::Error for VbpError {}

// ---------------------------------------------------------------------------
// Parser
// ---------------------------------------------------------------------------

/// Parse a VB6 project file (`*.vbp`).
///
/// The file is a sequence of `Key=Value` lines.  Comments start with `'`.
/// Values may be quoted (`"..."`) or bare.  This function tolerates lines in
/// any order and unknown keys, matching real-world VBP files.
pub fn parse_vbp(src: &str) -> Result<ProjectFile, VbpError> {
    let mut proj = ProjectFile::default();
    let mut major_ver: Option<u32> = None;
    let mut minor_ver: Option<u32> = None;
    let mut revision_ver: Option<u32> = None;

    for (line_idx, raw) in src.lines().enumerate() {
        let _ln = line_idx + 1;
        let line = strip_comment(raw).trim();
        if line.is_empty() {
            continue;
        }

        // Split on first '='
        let (key, val) = match line.find('=') {
            Some(pos) => (&line[..pos], &line[pos + 1..]),
            None => continue, // malformed line — skip
        };
        let key = key.trim();
        let val = val.trim();

        match key {
            // First line: project type
            "Type" => proj.project_type = unquote(val),

            // Module entries
            "Form" => proj.modules.push(parse_module(ModuleKind::Form, val)),
            "MDIForm" => proj.modules.push(parse_module(ModuleKind::MdiForm, val)),
            "Module" => proj.modules.push(parse_module(ModuleKind::Module, val)),
            "Class" => proj.modules.push(parse_module(ModuleKind::Class, val)),
            "UserControl" => proj.modules.push(parse_module(ModuleKind::UserControl, val)),
            "UserDocument" => proj.modules.push(parse_module(ModuleKind::UserDocument, val)),
            "PropertyPage" => proj.modules.push(parse_module(ModuleKind::PropertyPage, val)),
            "Resource" => proj.modules.push(parse_module(ModuleKind::Resource, val)),
            "RelatedDoc" => proj.modules.push(parse_module(ModuleKind::RelatedDoc, val)),

            // References and OCX objects
            "Reference" => proj.references.push(parse_reference(val)),
            "Object" => proj.objects.push(parse_ocx_object(val)),
            "ResFile32" => proj.res_file = Some(ResFile { path: unquote(val) }),

            // Project settings
            "Name" => proj.name = Some(unquote(val)),
            "Startup" => proj.startup = Some(unquote(val)),
            "HelpFile" => proj.help_file = Some(unquote(val)),
            "Title" => proj.title = Some(unquote(val)),
            "ExeName32" => proj.exe_name = Some(unquote(val)),
            "Command32" => proj.command32 = Some(unquote(val)),
            "Description" => proj.description = Some(unquote(val)),
            "HelpContextID" => proj.help_context_id = Some(unquote(val)),
            "CompatibleMode" => proj.compatible_mode = Some(unquote(val)),
            "CompatibleEXE32" => proj.compatible_exe32 = Some(unquote(val)),
            "IconForm" => proj.icon_form = Some(unquote(val)),
            "CondComp" => proj.cond_comp = Some(unquote(val)),
            "RequireLicenseKey" => proj.require_license_key = val == "1",
            "NoControlUpgrade" => proj.no_control_upgrade = val == "1",
            "VersionCompatible32" => proj.version_compatible32 = Some(unquote(val)),
            "DebugStartupOption" => proj.debug_startup_option = parse_int(val),
            "DebugStartupComponent" => proj.debug_startup_component = Some(val.to_string()),
            "UseExistingBrowser" => proj.use_existing_browser = val == "0",

            // Version numbers
            "MajorVer" => major_ver = val.parse().ok(),
            "MinorVer" => minor_ver = val.parse().ok(),
            "RevisionVer" => revision_ver = val.parse().ok(),
            "AutoIncrementVer" => proj.auto_increment_ver = val.parse().ok(),

            // Compilation
            "CompilationType" => proj.compilation_type = parse_int(val),
            "OptimizationType" => proj.optimization_type = parse_int(val),
            "FavorPentiumPro(tm)" => proj.compile_flags.favor_pentium_pro = val == "-1",
            "CodeViewDebugInfo" => proj.compile_flags.code_view_debug_info = val == "-1",
            "NoAliasing" => proj.compile_flags.no_aliasing = val == "-1",
            "BoundsCheck" => proj.compile_flags.bounds_check = val == "-1",
            "OverflowCheck" => proj.compile_flags.overflow_check = val == "-1",
            "FlPointCheck" => proj.compile_flags.fl_point_check = val == "-1",
            "FDIVCheck" => proj.compile_flags.fdiv_check = val == "-1",
            "UnroundedFP" => proj.compile_flags.unrounded_fp = val == "-1",
            "RemoveUnusedControlInfo" => {
                proj.compile_flags.remove_unused_control_info = val == "-1"
            }

            // Runtime settings
            "StartMode" => proj.start_mode = parse_int(val),
            "Unattended" => proj.unattended = parse_int(val),
            "Retained" => proj.retained = parse_int(val),
            "ThreadPerObject" => proj.thread_per_object = parse_int(val),
            "MaxNumberOfThreads" => proj.max_threads = parse_int(val),
            "ThreadingModel" => proj.threading_model = parse_int(val),
            "ServerSupportFiles" => proj.server_support_files = parse_int(val),
            "DllBaseAddress" => proj.dll_base_address = parse_hex_or_int(val),

            // Version-info strings
            "VersionComments" => proj.version_info.comments = Some(unquote(val)),
            "VersionCompanyName" => proj.version_info.company_name = Some(unquote(val)),
            "VersionFileDescription" => proj.version_info.file_description = Some(unquote(val)),
            "VersionLegalCopyright" => proj.version_info.legal_copyright = Some(unquote(val)),
            "VersionLegalTrademarks" => proj.version_info.legal_trademarks = Some(unquote(val)),
            "VersionProductName" => proj.version_info.product_name = Some(unquote(val)),

            // Forward-compatible: store unknown keys verbatim
            _ => proj.extra.push((key.to_string(), val.to_string())),
        }
    }

    // Assemble version struct if any ver field was set
    if major_ver.is_some() || minor_ver.is_some() || revision_ver.is_some() {
        proj.version = Some(ProjectVersion {
            major: major_ver.unwrap_or(0),
            minor: minor_ver.unwrap_or(0),
            revision: revision_ver.unwrap_or(0),
        });
    }

    Ok(proj)
}

// ---------------------------------------------------------------------------
// Line parsers
// ---------------------------------------------------------------------------

/// Parse a module entry: `path` or `name; path` (for Form/MDIForm) or
/// `name = path` (for Module/Class).
///
/// VB6 writes Form entries as `Form=path; name` (with GUID for some types).
/// Module/Class entries: `Module=name; path` or `Class=name; path`.
fn parse_module(kind: ModuleKind, val: &str) -> Module {
    // Form/MDIForm: `Form1.frm` or `Form1.frm; Form1`
    // Module/Class: `modMain; modMain.bas`  (name ; path)
    if let Some(semi) = val.find(';') {
        let left = val[..semi].trim();
        let right = val[semi + 1..].trim();
        match kind {
            ModuleKind::Form | ModuleKind::MdiForm => {
                // left=path, right=name  (for GUID-bearing module lines the
                // right side may contain name after another ';')
                Module { kind, name: Some(unquote(right)), path: unquote(left) }
            }
            _ => {
                // Module/Class/UserControl: left=name, right=path
                Module { kind, name: Some(unquote(left)), path: unquote(right) }
            }
        }
    } else {
        // No ';' — entire value is the path
        let path = unquote(val);
        Module { kind, name: None, path }
    }
}

/// Parse a `Reference=*\G{guid}#major.minor#lcid#path#name` entry.
///
/// Field order: prefix`{guid}`#version#lcid#path#name.
fn parse_reference(val: &str) -> Reference {
    let raw = val.to_string();
    // The reference string before the first '#' contains the GUID in {}.
    let parts: Vec<&str> = val.splitn(5, '#').collect();
    let guid = parts.first().and_then(|s| {
        s.find('{').and_then(|i| s[i + 1..].find('}').map(|j| s[i + 1..i + 1 + j].to_string()))
    });
    let version = parts.get(1).map(|s| s.to_string());
    let lcid = parts.get(2).and_then(|s| s.parse().ok());
    let path = parts.get(3).map(|s| s.to_string());
    let name = parts.get(4).map(|s| s.to_string());
    Reference { raw, guid, version, lcid, path, name }
}

/// Parse an `Object={progid}#ver#lcid ; filename` entry.
fn parse_ocx_object(val: &str) -> OcxObject {
    if let Some(semi) = val.find(';') {
        let progid = val[..semi].trim().to_string();
        let filename = Some(val[semi + 1..].trim().to_string());
        OcxObject { progid, filename }
    } else {
        OcxObject { progid: val.to_string(), filename: None }
    }
}

// ---------------------------------------------------------------------------
// Utilities
// ---------------------------------------------------------------------------

fn strip_comment(s: &str) -> &str {
    // VB6 uses ' for comments.
    if let Some(pos) = s.find('\'') {
        &s[..pos]
    } else {
        s
    }
}

fn unquote(s: &str) -> String {
    let s = s.trim();
    if s.starts_with('"') && s.ends_with('"') && s.len() >= 2 {
        // Unescape "" → "
        s[1..s.len() - 1].replace("\"\"", "\"")
    } else {
        s.to_string()
    }
}

fn parse_int(s: &str) -> Option<i32> {
    s.trim().parse().ok()
}

fn parse_hex_or_int(s: &str) -> Option<u32> {
    let s = s.trim();
    if let Some(hex) = s.strip_prefix("&H").or_else(|| s.strip_prefix("&h")) {
        u32::from_str_radix(hex.trim_end_matches('&'), 16).ok()
    } else {
        s.parse().ok()
    }
}

// ---------------------------------------------------------------------------
// Tests
// ---------------------------------------------------------------------------

#[cfg(test)]
mod tests {
    use super::*;

    const MINIMAL_VBP: &str = r#"Type=Exe
Form=Form1.frm
Module=modMain; modMain.bas
Startup="Form1"
Name="Project1"
HelpContextID="0"
CompatibleMode="0"
MajorVer=1
MinorVer=0
RevisionVer=0
AutoIncrementVer=0
ServerSupportFiles=0
VersionCompatible32="393222000"
CompilationType=0
OptimizationType=0
FavorPentiumPro(tm)=0
CodeViewDebugInfo=0
NoAliasing=0
BoundsCheck=0
OverflowCheck=0
FlPointCheck=0
FDIVCheck=0
UnroundedFP=0
StartMode=0
Unattended=0
Retained=0
ThreadPerObject=0
MaxNumberOfThreads=1
"#;

    #[test]
    fn parse_minimal_project() {
        let p = parse_vbp(MINIMAL_VBP).unwrap();
        assert_eq!(p.project_type, "Exe");
        assert_eq!(p.modules.len(), 2);
        assert_eq!(p.modules[0].kind, ModuleKind::Form);
        assert_eq!(p.modules[0].path, "Form1.frm");
        assert_eq!(p.modules[1].kind, ModuleKind::Module);
        assert_eq!(p.modules[1].name, Some("modMain".into()));
        assert_eq!(p.modules[1].path, "modMain.bas");
        assert_eq!(p.startup, Some("Form1".into()));
        assert_eq!(p.name, Some("Project1".into()));
        assert_eq!(p.max_threads, Some(1));
    }

    #[test]
    fn parse_reference_line() {
        let r = parse_reference(
            r#"*\G{00020430-0000-0000-C000-000000000046}#2.0#0#..\Windows\System32\stdole2.tlb#OLE Automation"#,
        );
        assert_eq!(r.guid.as_deref(), Some("00020430-0000-0000-C000-000000000046"));
        assert_eq!(r.name.as_deref(), Some("OLE Automation"));
    }

    #[test]
    fn parse_ocx_object_line() {
        let o = parse_ocx_object(
            "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0; MSCOMCTL.OCX",
        );
        assert!(o.progid.contains("831FDD16"));
        assert_eq!(o.filename.as_deref(), Some("MSCOMCTL.OCX"));
    }

    #[test]
    fn tolerates_unknown_keys() {
        let src = "Type=Exe\nFutureKey=some_value\n";
        let p = parse_vbp(src).unwrap();
        assert_eq!(p.extra.len(), 1);
        assert_eq!(p.extra[0].0, "FutureKey");
    }
}
