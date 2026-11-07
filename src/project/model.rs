//! A loaded VB6 project (or project group): its members, its references, and one engine
//! session per project.

use std::collections::HashMap;
use std::path::{Path, PathBuf};
use std::sync::{Arc, RwLock};

use serde::Serialize;
use crate::typelib::{load_library_noted, TypeLibraryRequest};
use vb6_engine::session::{Library, ProjectKind, Session, SessionConfig};

use super::paths::{engine_key_for_path, lookup_key, normalize};
use crate::workspace::{ProjectMember, ProjectType, VbgFile, VbpFile, VbpReference};

/// How serious a problem found while loading is.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Serialize)]
#[serde(rename_all = "lowercase")]
pub enum Severity {
    Error,
    Warning,
    Information,
}

/// Something wrong with a project, located at a line of the `.vbp`/`.vbg` that declares it.
#[derive(Debug, Clone, PartialEq, Eq, Serialize)]
pub struct Problem {
    pub file: PathBuf,
    /// 1-based; `0` when the problem belongs to the file as a whole.
    pub line: usize,
    pub severity: Severity,
    pub message: String,
}

/// A source file listed in a project.
#[derive(Debug, Clone, PartialEq, Eq, Serialize)]
pub struct MemberFile {
    /// The member's name in the project (`frmMain`); empty when the project does not give one.
    pub name: String,
    /// Lower-case extension (`bas`, `frm`...).
    pub extension: String,
    pub path: PathBuf,
    /// The key the engine knows the file by.
    #[serde(skip)]
    pub key: String,
    /// [`lookup_key`] of the path.
    #[serde(skip)]
    pub lookup: String,
    /// Line of the entry in the project file.
    pub line: usize,
    pub exists: bool,
}

/// What a reference is.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Serialize)]
#[serde(rename_all = "camelCase")]
pub enum ReferenceKind {
    TypeLibrary,
    Project,
    Control,
}

/// Whether the analysis can see a reference's types.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Serialize)]
#[serde(rename_all = "camelCase")]
pub enum ReferenceStatus {
    /// A library whose types are available: a runtime library the engine carries, or one read
    /// from its type library.
    Loaded,
    /// A library that could not be read: its types are not available.
    NotLoaded,
    /// A project whose public classes are available.
    ProjectLoaded,
    /// A project that is not on disk.
    ProjectMissing,
    /// A project that (indirectly) references this one.
    ProjectCircular,
    /// A project that exposes nothing to others (a Standard EXE).
    ProjectNotExposed,
}

#[derive(Debug, Clone, PartialEq, Eq, Serialize)]
pub struct ReferenceInfo {
    pub label: String,
    pub kind: ReferenceKind,
    pub status: ReferenceStatus,
    pub line: usize,
    pub detail: String,
}

/// One project of the selection with its analysis.
pub struct Unit {
    pub vbp_path: PathBuf,
    pub name: String,
    pub kind: ProjectKind,
    /// Listed in the opened group (or the opened project itself) rather than reached only
    /// through a reference.
    pub listed: bool,
    pub members: Vec<MemberFile>,
    pub references: Vec<ReferenceInfo>,
    pub session: Arc<RwLock<Session>>,
    /// Units whose public classes this project's analysis sees.
    pub dependencies: Vec<usize>,
    /// The type libraries and control libraries the project references that were read.
    pub type_libraries: Vec<Arc<Library>>,
    /// What the last analysis exported to dependents.
    pub exported: Library,
    pub conditional_compilation: Option<String>,
}

/// The opened project or group.
pub struct Selection {
    /// The `.vbp` or `.vbg` the user opened.
    pub path: PathBuf,
    pub is_group: bool,
    pub name: String,
    /// Listed projects first (startup project first), then the projects reached by reference.
    pub units: Vec<Unit>,
    pub problems: Vec<Problem>,
}

/// Why a selection could not be opened.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct LoadError {
    pub file: PathBuf,
    pub line: Option<usize>,
    pub message: String,
}

impl std::fmt::Display for LoadError {
    fn fmt(&self, f: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        write!(f, "{}: {}", self.file.display(), self.message)
    }
}

impl std::error::Error for LoadError {}

/// Type libraries the engine carries, by GUID.
const BUNDLED_LIBRARY_GUIDS: [&str; 3] = [
    "00020430-0000-0000-c000-000000000046", // stdole
    "000204ef-0000-0000-c000-000000000046", // VBA
    "fcfb3d2e-a0fa-1068-a738-08002b3371b5", // VB
];

/// Unsaved editor text: the Windows-1252 bytes of a file by [`lookup_key`], when it is open.
pub type Overlay<'a> = &'a dyn Fn(&str) -> Option<Vec<u8>>;

pub fn project_kind(t: ProjectType) -> ProjectKind {
    match t {
        ProjectType::Exe => ProjectKind::Exe,
        ProjectType::OleDll => ProjectKind::OleDll,
        ProjectType::OleExe => ProjectKind::OleExe,
        ProjectType::Control => ProjectKind::Control,
    }
}

fn is_extension(path: &Path, ext: &str) -> bool {
    path.extension().is_some_and(|e| e.eq_ignore_ascii_case(ext))
}

/// The version an `Object=` line gives (`2.0`, `1.3`) as major and minor numbers.
fn parse_version(version: &str) -> (u16, u16) {
    let mut parts = version.trim().split('.');
    let major = parts.next().and_then(|p| p.trim().parse().ok()).unwrap_or(0);
    let minor = parts.next().and_then(|p| p.trim().parse().ok()).unwrap_or(0);
    (major, minor)
}

fn reference_line(r: &VbpReference) -> usize {
    match r {
        VbpReference::Subproject { line, .. } | VbpReference::Compiled { line, .. } => *line,
    }
}

/// A project read from its file, before analysis.
struct Draft {
    path: PathBuf,
    vbp: VbpFile,
    listed: bool,
    members: Vec<MemberFile>,
    references: Vec<ReferenceInfo>,
    /// (index into `vbp.references`, target draft) for each project reference that resolved.
    edges: Vec<(usize, usize)>,
    /// The type libraries and control libraries the project references that were read.
    type_libraries: Vec<Arc<Library>>,
}

fn member_file(m: &ProjectMember) -> MemberFile {
    let path = normalize(&m.absolute_path);
    MemberFile {
        name: m.name.clone(),
        extension: path.extension().map(|e| e.to_string_lossy().to_lowercase()).unwrap_or_default(),
        key: engine_key_for_path(&path),
        lookup: lookup_key(&path),
        line: m.line,
        exists: path.exists(),
        path,
    }
}

/// The Windows-1252 bytes of a member: unsaved editor text if it is open, else the file.
pub fn read_member(m: &mut MemberFile, overlay: Overlay) -> Option<Vec<u8>> {
    if let Some(bytes) = overlay(&m.lookup) {
        m.exists = true;
        return Some(bytes);
    }
    match std::fs::read(&m.path) {
        Ok(bytes) => {
            m.exists = true;
            Some(bytes)
        }
        Err(_) => {
            m.exists = false;
            None
        }
    }
}

/// The libraries a project's analysis is told about: the runtime libraries, the type libraries
/// and control libraries it references, and what each of its referenced projects exports.
pub fn libraries_for(
    type_libraries: &[Arc<Library>],
    dependencies: &[usize],
    exported: impl Fn(usize) -> Option<Library>,
) -> Vec<Library> {
    let mut libraries = Library::bundled();
    libraries.extend(type_libraries.iter().map(|l| (**l).clone()));
    libraries.extend(dependencies.iter().filter_map(|&d| exported(d)));
    libraries
}

fn visit(
    at: usize,
    drafts: &[Draft],
    state: &mut [u8],
    order: &mut Vec<usize>,
    kept: &mut [Vec<usize>],
    cycles: &mut Vec<(usize, usize)>,
) {
    state[at] = 1;
    for &(ri, to) in &drafts[at].edges {
        if state[to] == 1 {
            cycles.push((at, ri));
            continue;
        }
        if state[to] == 0 {
            visit(to, drafts, state, order, kept, cycles);
        }
        if !kept[at].contains(&to) {
            kept[at].push(to);
        }
    }
    state[at] = 2;
    order.push(at);
}

/// The projects of an opened file, read but not yet analysed.
struct Drafts {
    path: PathBuf,
    is_group: bool,
    group_name: String,
    drafts: Vec<Draft>,
    problems: Vec<Problem>,
}

/// The order to analyse projects in (a project after those it references) and, for each
/// project, the projects whose public classes it sees.
struct DependencyOrder {
    order: Vec<usize>,
    kept: Vec<Vec<usize>>,
}

impl Selection {
    /// Read the opened file and every project it lists or references.
    fn read_drafts(path: &Path) -> Result<Drafts, LoadError> {
        let path = normalize(path);
        let mut problems = Vec::new();

        let (is_group, group_name, listed_paths) = Self::listed_projects(&path, &mut problems)?;

        // Listed projects, then (breadth first) every project they reference that can be read.
        let mut drafts: Vec<Draft> = Vec::new();
        let mut index: HashMap<String, usize> = HashMap::new();
        let mut queue: Vec<(PathBuf, PathBuf, usize, bool)> =
            listed_paths.into_iter().map(|(p, l)| (p, path.clone(), l, true)).collect();
        let mut cursor = 0;
        while cursor < queue.len() {
            let (project_path, declared_in, declared_line, listed) = queue[cursor].clone();
            cursor += 1;
            let key = lookup_key(&project_path);
            if index.contains_key(&key) {
                continue;
            }
            let vbp = match VbpFile::parse(&project_path) {
                Ok(v) => v,
                Err(e) if !is_group && listed => {
                    return Err(LoadError { file: project_path, line: e.line, message: e.message });
                }
                // A referenced project that cannot be read is reported with the reference.
                Err(_) if !listed => continue,
                Err(e) => {
                    problems.push(Problem {
                        file: declared_in,
                        line: declared_line,
                        severity: Severity::Error,
                        message: format!("Cannot read project {}: {}", project_path.display(), e.message),
                    });
                    continue;
                }
            };
            for r in vbp.get_subproject_references() {
                if let VbpReference::Subproject { path: target, line } = r {
                    queue.push((normalize(target), project_path.clone(), *line, false));
                }
            }
            index.insert(key, drafts.len());
            let members = vbp.all_source_files().map(member_file).collect();
            drafts.push(Draft {
                path: project_path,
                vbp,
                listed,
                members,
                references: Vec::new(),
                edges: Vec::new(),
                type_libraries: Vec::new(),
            });
        }

        for di in 0..drafts.len() {
            Self::resolve_references(di, &mut drafts, &index, &mut problems);
        }
        Ok(Drafts { path, is_group, group_name, drafts, problems })
    }

    /// Work out the analysis order. An edge that closes a cycle is dropped and reported; a
    /// reference to a Standard EXE, which exposes nothing, is reported and dropped.
    fn dependency_order(drafts: &mut [Draft], problems: &mut Vec<Problem>) -> DependencyOrder {
        let count = drafts.len();
        let mut state = vec![0u8; count];
        let mut order = Vec::new();
        let mut kept: Vec<Vec<usize>> = vec![Vec::new(); count];
        let mut cycles = Vec::new();
        for i in 0..count {
            if state[i] == 0 {
                visit(i, drafts, &mut state, &mut order, &mut kept, &mut cycles);
            }
        }
        for (di, ri) in cycles {
            let line = reference_line(&drafts[di].vbp.references[ri]);
            problems.push(Problem {
                file: drafts[di].path.clone(),
                line,
                severity: Severity::Error,
                message: "Cyclic reference of projects not allowed".to_string(),
            });
            if let Some(info) = drafts[di].references.iter_mut().find(|i| i.line == line) {
                info.status = ReferenceStatus::ProjectCircular;
            }
        }
        for di in 0..count {
            let mut dependencies = Vec::new();
            for &dep in &kept[di] {
                if project_kind(drafts[dep].vbp.project_type) != ProjectKind::Exe {
                    dependencies.push(dep);
                    continue;
                }
                let edge = drafts[di].edges.iter().find(|(_, to)| *to == dep).map(|(ri, _)| *ri);
                let line = edge.map(|ri| reference_line(&drafts[di].vbp.references[ri])).unwrap_or(0);
                problems.push(Problem {
                    file: drafts[di].path.clone(),
                    line,
                    severity: Severity::Error,
                    message: format!(
                        "Project '{}' can not be referenced because it's project type is EXE.",
                        drafts[dep].path.display()
                    ),
                });
                if let Some(info) = drafts[di].references.iter_mut().find(|i| i.line == line) {
                    info.status = ReferenceStatus::ProjectNotExposed;
                }
            }
            kept[di] = dependencies;
        }
        DependencyOrder { order, kept }
    }

    /// Describe `path` (a `.vbp` or `.vbg`) without analysing it: its projects, members,
    /// references and problems, as opening it would find them.
    pub fn preview(path: &Path) -> Result<super::info::ProjectInfo, LoadError> {
        use super::info::{ProjectInfo, UnitInfo};
        let Drafts { path, is_group, group_name, mut drafts, mut problems } = Self::read_drafts(path)?;
        Self::dependency_order(&mut drafts, &mut problems);
        for d in &drafts {
            for m in d.members.iter().filter(|m| !m.exists) {
                problems.push(Problem {
                    file: d.path.clone(),
                    line: m.line,
                    severity: Severity::Error,
                    message: format!("File not found: {}", m.path.display()),
                });
            }
        }
        problems.sort_by(|a, b| (lookup_key(&a.file), a.line).cmp(&(lookup_key(&b.file), b.line)));
        let name = if is_group { group_name } else { drafts.first().map(|d| d.vbp.name.clone()).unwrap_or_default() };
        let projects = drafts
            .iter()
            .enumerate()
            .map(|(i, d)| {
                UnitInfo::new(&d.path, &d.vbp.name, project_kind(d.vbp.project_type), d.listed, i == 0, &d.members, &d.references)
            })
            .collect();
        Ok(ProjectInfo::new(&path, &name, is_group, projects, problems))
    }

    /// Open `path` (a `.vbp` or `.vbg`), reading each member from `overlay` or from disk.
    pub fn load(path: &Path, overlay: Overlay) -> Result<Selection, LoadError> {
        let Drafts { path, is_group, group_name, mut drafts, mut problems } = Self::read_drafts(path)?;
        let count = drafts.len();
        let DependencyOrder { order, mut kept } = Self::dependency_order(&mut drafts, &mut problems);

        // Analyse a project after the projects it references.
        let mut sessions: Vec<Option<Arc<RwLock<Session>>>> = (0..count).map(|_| None).collect();
        let mut exported: Vec<Option<Library>> = (0..count).map(|_| None).collect();
        for &di in &order {
            let kind = project_kind(drafts[di].vbp.project_type);

            let d = &mut drafts[di];
            let mut files = Vec::new();
            for m in &mut d.members {
                match read_member(m, overlay) {
                    Some(bytes) => files.push((m.key.clone(), bytes)),
                    None => problems.push(Problem {
                        file: d.path.clone(),
                        line: m.line,
                        severity: Severity::Error,
                        message: format!("File not found: {}", m.path.display()),
                    }),
                }
            }
            let config = SessionConfig {
                project_name: Some(d.vbp.name.clone()),
                libraries: libraries_for(&d.type_libraries, &kept[di], |dep| exported[dep].clone()),
                conditional_compilation: d.vbp.compilation.conditional_compile.clone(),
            };
            let session = Session::with_config(files, config);
            exported[di] = Some(session.project_library(kind));
            sessions[di] = Some(Arc::new(RwLock::new(session)));
        }

        problems.sort_by(|a, b| (lookup_key(&a.file), a.line).cmp(&(lookup_key(&b.file), b.line)));

        let name = if is_group {
            group_name
        } else {
            drafts.first().map(|d| d.vbp.name.clone()).unwrap_or_default()
        };
        let units = drafts
            .into_iter()
            .enumerate()
            .map(|(i, d)| Unit {
                kind: project_kind(d.vbp.project_type),
                name: d.vbp.name.clone(),
                vbp_path: d.path,
                listed: d.listed,
                members: d.members,
                references: d.references,
                session: sessions[i].take().expect("every project is analysed"),
                dependencies: std::mem::take(&mut kept[i]),
                type_libraries: d.type_libraries,
                exported: exported[i].take().unwrap_or(Library { name: String::new(), types: Vec::new() }),
                conditional_compilation: d.vbp.compilation.conditional_compile,
            })
            .collect();
        Ok(Selection { path, is_group, name, units, problems })
    }

    /// The projects the opened file lists, startup project first, with the line each is
    /// listed on.
    fn listed_projects(
        path: &Path,
        problems: &mut Vec<Problem>,
    ) -> Result<(bool, String, Vec<(PathBuf, usize)>), LoadError> {
        if is_extension(path, "vbg") {
            let group = VbgFile::parse(path).map_err(|e| LoadError {
                file: path.to_path_buf(),
                line: e.line,
                message: e.message,
            })?;
            // The startup project is a member of the group in its own right: it is listed on the
            // `StartupProject=` line and not again. A project listed twice is refused by VB6.
            let mut entries: Vec<(usize, PathBuf)> =
                group.projects.iter().map(|p| (p.line, normalize(&p.absolute_path))).collect();
            if let (Some(startup), Some(line)) = (&group.startup, group.startup_line) {
                entries.push((line, normalize(startup)));
            }
            entries.sort_by_key(|(line, _)| *line);
            let startup_key = group.startup.as_ref().map(|s| lookup_key(&normalize(s)));
            let mut listed: Vec<(PathBuf, usize)> = Vec::new();
            for (line, project) in entries {
                let key = lookup_key(&project);
                if listed.iter().any(|(p, _)| lookup_key(p) == key) {
                    problems.push(Problem {
                        file: path.to_path_buf(),
                        line,
                        severity: Severity::Error,
                        message: format!("The file '{}' is already open.", project.display()),
                    });
                    continue;
                }
                if startup_key.as_deref() == Some(key.as_str()) {
                    listed.insert(0, (project, line));
                } else {
                    listed.push((project, line));
                }
            }
            let name = path.file_stem().map(|s| s.to_string_lossy().into_owned()).unwrap_or_default();
            Ok((true, name, listed))
        } else if is_extension(path, "vbp") {
            Ok((false, String::new(), vec![(path.to_path_buf(), 0)]))
        } else {
            Err(LoadError {
                file: path.to_path_buf(),
                line: None,
                message: "Not a VB6 project (.vbp) or project group (.vbg)".to_string(),
            })
        }
    }

    /// Describe project `di`'s references and record which resolve to another project.
    fn resolve_references(
        di: usize,
        drafts: &mut [Draft],
        index: &HashMap<String, usize>,
        problems: &mut Vec<Problem>,
    ) {
        let vbp_path = drafts[di].path.clone();
        let vbp_dir = vbp_path.parent().map(Path::to_path_buf).unwrap_or_default();
        let references: Vec<VbpReference> = drafts[di].vbp.references.clone();
        let mut infos = Vec::new();
        let mut edges = Vec::new();
        let mut type_libraries: Vec<Arc<Library>> = Vec::new();
        for (ri, r) in references.iter().enumerate() {
            match r {
                VbpReference::Subproject { path: target, line } => {
                    let target = normalize(target);
                    let label = target.file_stem().map(|s| s.to_string_lossy().into_owned()).unwrap_or_default();
                    let status = match index.get(&lookup_key(&target)) {
                        Some(&to) => {
                            edges.push((ri, to));
                            if !drafts[to].listed {
                                problems.push(Problem {
                                    file: vbp_path.clone(),
                                    line: *line,
                                    severity: Severity::Warning,
                                    message: format!(
                                        "Can't find project or library: {} is not part of the opened project group",
                                        target.display()
                                    ),
                                });
                            }
                            ReferenceStatus::ProjectLoaded
                        }
                        None => {
                            problems.push(Problem {
                                file: vbp_path.clone(),
                                line: *line,
                                severity: Severity::Error,
                                message: if target.exists() {
                                    format!("Cannot read referenced project {}", target.display())
                                } else {
                                    format!("Can't find project or library: {}", target.display())
                                },
                            });
                            ReferenceStatus::ProjectMissing
                        }
                    };
                    infos.push(ReferenceInfo {
                        label,
                        kind: ReferenceKind::Project,
                        status,
                        line: *line,
                        detail: target.display().to_string(),
                    });
                }
                VbpReference::Compiled { guid, major, minor, lcid, path: file, description, line } => {
                    let bundled = BUNDLED_LIBRARY_GUIDS.iter().any(|g| g.eq_ignore_ascii_case(guid));
                    let status = if bundled {
                        ReferenceStatus::Loaded
                    } else {
                        let file = if file.is_absolute() { file.clone() } else { vbp_dir.join(file) };
                        let request = TypeLibraryRequest {
                            guid: guid.clone(),
                            major: *major,
                            minor: *minor,
                            lcid: *lcid,
                            paths: vec![file],
                            description: description.clone(),
                        };
                        Self::attach_library(&mut type_libraries, &request, &vbp_path, *line, problems)
                    };
                    infos.push(ReferenceInfo {
                        label: description.clone(),
                        kind: ReferenceKind::TypeLibrary,
                        status,
                        line: *line,
                        detail: format!("{{{}}}", guid),
                    });
                }
            }
        }
        for o in &drafts[di].vbp.objects {
            let (major, minor) = parse_version(&o.version);
            let mut paths = vec![vbp_dir.join(&o.filename)];
            for dir in ["SysWOW64", "System32"] {
                if let Some(root) = std::env::var_os("SystemRoot") {
                    paths.push(Path::new(&root).join(dir).join(&o.filename));
                }
            }
            let request = TypeLibraryRequest {
                guid: o.guid.clone(),
                major,
                minor,
                lcid: 0,
                paths,
                description: o.filename.clone(),
            };
            let status = Self::attach_library(&mut type_libraries, &request, &vbp_path, o.line, problems);
            infos.push(ReferenceInfo {
                label: o.filename.clone(),
                kind: ReferenceKind::Control,
                status,
                line: o.line,
                detail: format!("{{{}}} {}", o.guid, o.version),
            });
        }
        drafts[di].type_libraries = type_libraries;
        drafts[di].references = infos;
        drafts[di].edges = edges;
    }

    /// Read the library `request` names and add it to `libraries`; a library that cannot be read
    /// is reported on the reference's line.
    fn attach_library(
        libraries: &mut Vec<Arc<Library>>,
        request: &TypeLibraryRequest,
        vbp_path: &Path,
        line: usize,
        problems: &mut Vec<Problem>,
    ) -> ReferenceStatus {
        match load_library_noted(request) {
            Ok(loaded) => {
                for file in &loaded.missing_imports {
                    problems.push(Problem {
                        file: vbp_path.to_path_buf(),
                        line,
                        severity: Severity::Warning,
                        message: format!(
                            "Type library {} takes types from {file}, which was not found. Those types are not available to the analysis",
                            request.display_name()
                        ),
                    });
                }
                if !libraries.iter().any(|l| l.name.eq_ignore_ascii_case(&loaded.library.name)) {
                    libraries.push(loaded.library);
                }
                ReferenceStatus::Loaded
            }
            Err(e) => {
                problems.push(Problem {
                    file: vbp_path.to_path_buf(),
                    line,
                    severity: Severity::Warning,
                    message: format!("{e}. Its types are not available to the analysis"),
                });
                ReferenceStatus::NotLoaded
            }
        }
    }

    /// The unit and member owning `lookup` (a [`lookup_key`]); earlier units win a shared file.
    pub fn locate(&self, lookup: &str) -> Option<(usize, usize)> {
        self.units.iter().enumerate().find_map(|(u, unit)| {
            unit.members.iter().position(|m| m.lookup == lookup).map(|m| (u, m))
        })
    }

    /// The unit whose project file is `lookup`.
    pub fn unit_of_project_file(&self, lookup: &str) -> Option<usize> {
        self.units.iter().position(|u| lookup_key(&u.vbp_path) == lookup)
    }
}

impl Selection {
    /// After project `changed` was edited, refresh what the projects that reference it see.
    ///
    /// A project exposes its public classes to its dependents; when an edit changes what is
    /// exposed, each dependent is analysed again with the new description. Returns the
    /// projects that were analysed again (transitively).
    pub fn propagate_exports(&mut self, changed: usize) -> Vec<usize> {
        let mut reanalysed = Vec::new();
        let mut work = vec![changed];
        while let Some(u) = work.pop() {
            if !self.units.iter().any(|other| other.dependencies.contains(&u)) {
                continue;
            }
            let library = self.units[u].session.read().unwrap().project_library(self.units[u].kind);
            if library == self.units[u].exported {
                continue;
            }
            self.units[u].exported = library;
            for j in 0..self.units.len() {
                if !self.units[j].dependencies.contains(&u) {
                    continue;
                }
                let libraries = libraries_for(&self.units[j].type_libraries, &self.units[j].dependencies, |d| {
                    Some(self.units[d].exported.clone())
                });
                let config = SessionConfig {
                    project_name: Some(self.units[j].name.clone()),
                    libraries,
                    conditional_compilation: self.units[j].conditional_compilation.clone(),
                };
                self.units[j].session.write().unwrap().set_config(config);
                if !reanalysed.contains(&j) {
                    reanalysed.push(j);
                }
                work.push(j);
            }
        }
        reanalysed
    }

    /// Record whether member `key` of project `unit` exists on disk.
    pub fn set_member_exists(&mut self, unit: usize, key: &str, exists: bool) {
        if let Some(m) = self.units[unit].members.iter_mut().find(|m| m.key == key) {
            m.exists = exists;
        }
    }
}
