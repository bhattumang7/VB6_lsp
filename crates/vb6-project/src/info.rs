//! A serializable description of the opened project, for an editor's project views.

use serde::Serialize;
use vb6_engine::session::ProjectKind;

use super::model::{MemberFile, Problem, ReferenceInfo, Selection};

#[derive(Debug, Clone, Serialize)]
#[serde(rename_all = "camelCase")]
pub struct MemberInfo {
    pub name: String,
    pub extension: String,
    pub path: String,
    pub uri: String,
    /// 1-based line of the entry in the project file.
    pub line: usize,
    pub missing: bool,
}

#[derive(Debug, Clone, Serialize)]
#[serde(rename_all = "camelCase")]
pub struct UnitInfo {
    pub path: String,
    pub uri: String,
    pub name: String,
    /// `exe`, `oledll`, `oleexe` or `control`.
    pub kind: &'static str,
    /// Listed in the opened group; false for a project reached only through a reference.
    pub listed: bool,
    pub startup: bool,
    pub members: Vec<MemberInfo>,
    pub references: Vec<ReferenceInfo>,
}

#[derive(Debug, Clone, Serialize)]
#[serde(rename_all = "camelCase")]
pub struct ProjectInfo {
    pub path: String,
    pub uri: String,
    pub name: String,
    pub is_group: bool,
    pub projects: Vec<UnitInfo>,
    pub problems: Vec<Problem>,
}

fn kind_name(kind: ProjectKind) -> &'static str {
    match kind {
        ProjectKind::Exe => "exe",
        ProjectKind::OleDll => "oledll",
        ProjectKind::OleExe => "oleexe",
        ProjectKind::Control => "control",
    }
}

fn uri_text(path: &std::path::Path) -> String {
    super::paths::engine_key_for_path(path)
}

impl UnitInfo {
    pub fn new(
        path: &std::path::Path,
        name: &str,
        kind: ProjectKind,
        listed: bool,
        startup: bool,
        members: &[MemberFile],
        references: &[ReferenceInfo],
    ) -> Self {
        UnitInfo {
            path: path.display().to_string(),
            uri: uri_text(path),
            name: name.to_string(),
            kind: kind_name(kind),
            listed,
            startup,
            members: members
                .iter()
                .map(|m| MemberInfo {
                    name: m.name.clone(),
                    extension: m.extension.clone(),
                    path: m.path.display().to_string(),
                    uri: m.key.clone(),
                    line: m.line,
                    missing: !m.exists,
                })
                .collect(),
            references: references.to_vec(),
        }
    }
}

impl ProjectInfo {
    pub fn new(path: &std::path::Path, name: &str, is_group: bool, projects: Vec<UnitInfo>, problems: Vec<Problem>) -> Self {
        ProjectInfo {
            path: path.display().to_string(),
            uri: uri_text(path),
            name: name.to_string(),
            is_group,
            projects,
            problems,
        }
    }

    pub fn of(selection: &Selection) -> Self {
        let projects = selection
            .units
            .iter()
            .enumerate()
            .map(|(i, u)| UnitInfo::new(&u.vbp_path, &u.name, u.kind, u.listed, i == 0, &u.members, &u.references))
            .collect();
        ProjectInfo::new(&selection.path, &selection.name, selection.is_group, projects, selection.problems.clone())
    }
}
