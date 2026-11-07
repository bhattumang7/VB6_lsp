//! The opened project or group inside the language server: routing a document to the session
//! that analyses it, keeping every project member analysed whether or not it is open, and
//! publishing diagnostics for the files an edit or a reload affects.

use std::collections::{HashMap, HashSet};
use std::path::{Path, PathBuf};
use std::sync::{Arc, RwLock};

use serde::{Deserialize, Serialize};
use serde_json::Value;
use tower_lsp::lsp_types::notification::Notification;
use tower_lsp::lsp_types::*;
use vb6_engine::session::Session;

use super::Vb6LanguageServer;
use crate::engine_glue;
use crate::project::info::ProjectInfo;
use crate::project::paths::{self, lookup_key};
use crate::project::{LoadError, Problem, Selection, Severity};

/// The session that analyses a document and the key the session knows it by.
pub(super) struct Route {
    pub session: Arc<RwLock<Session>>,
    pub key: String,
    /// The project (index into the selection's units) the document belongs to.
    pub unit: Option<usize>,
    /// For a project member, its file.
    pub path: Option<PathBuf>,
}

/// Where the opened project stands, sent to the editor as `vb6/projectStatus`.
#[derive(Debug, Clone, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct ProjectStatus {
    /// `none`, `loading`, `loaded` or `error`.
    pub state: String,
    pub path: Option<String>,
    pub name: Option<String>,
    pub is_group: bool,
    pub project_count: usize,
    pub member_count: usize,
    pub missing_count: usize,
    pub error_count: usize,
    pub warning_count: usize,
    pub information_count: usize,
    pub message: Option<String>,
}

impl ProjectStatus {
    fn none() -> Self {
        ProjectStatus {
            state: "none".to_string(),
            path: None,
            name: None,
            is_group: false,
            project_count: 0,
            member_count: 0,
            missing_count: 0,
            error_count: 0,
            warning_count: 0,
            information_count: 0,
            message: None,
        }
    }

    fn loading(path: &Path) -> Self {
        ProjectStatus { state: "loading".to_string(), path: Some(path.display().to_string()), ..Self::none() }
    }

    fn failed(path: &Path, message: String) -> Self {
        ProjectStatus {
            state: "error".to_string(),
            path: Some(path.display().to_string()),
            message: Some(message),
            ..Self::none()
        }
    }

    fn of(selection: &Selection) -> Self {
        let listed = || selection.units.iter().filter(|u| u.listed);
        let count = |s: Severity| selection.problems.iter().filter(|p| p.severity == s).count();
        ProjectStatus {
            state: "loaded".to_string(),
            path: Some(selection.path.display().to_string()),
            name: Some(selection.name.clone()),
            is_group: selection.is_group,
            project_count: listed().count(),
            member_count: listed().map(|u| u.members.len()).sum(),
            missing_count: listed().map(|u| u.members.iter().filter(|m| !m.exists).count()).sum(),
            error_count: count(Severity::Error),
            warning_count: count(Severity::Warning),
            information_count: count(Severity::Information),
            message: None,
        }
    }
}

/// The `vb6/projectStatus` notification.
pub enum ProjectStatusNotification {}

impl Notification for ProjectStatusNotification {
    type Params = ProjectStatus;
    const METHOD: &'static str = "vb6/projectStatus";
}

/// How a watched file changed on disk.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum DiskChange {
    Written,
    Deleted,
}

/// What an update did beyond the file it touched.
#[derive(Default)]
struct Reach {
    /// Projects whose every member needs its diagnostics published again.
    units: HashSet<usize>,
    /// The documents analysed outside any project need theirs published again.
    loose: bool,
}

impl Vb6LanguageServer {
    /// Where the referenced project that `target` belongs to declares it. The referenced project
    /// is one `route`'s project depends on, found by the library's name, and is asked for its
    /// current declaration so the answer follows the project's edits.
    pub(super) fn external_definition(
        &self,
        route: &Route,
        target: &vb6_engine::session::ExternalTarget,
    ) -> Option<Location> {
        let guard = self.selection.read().unwrap();
        let selection = guard.as_ref()?;
        let unit = selection.units.get(route.unit?)?;
        for &dependency in &unit.dependencies {
            let referenced = &selection.units[dependency];
            if !referenced.name.eq_ignore_ascii_case(&target.library) {
                continue;
            }
            let session = referenced.session.read().unwrap();
            let member = target.member.as_ref().map(|(name, kind)| (name.as_str(), *kind));
            if let Some(found) = session.declaration_of(&target.type_name, member) {
                return crate::engine_glue::location(&session, found);
            }
        }
        None
    }

    /// The session that analyses `uri`: the opened project that lists the file, else the
    /// session of files outside any project.
    pub(super) fn route(&self, uri: &Url) -> Route {
        let lookup = paths::uri_lookup_key(uri);
        if let Some(selection) = self.selection.read().unwrap().as_ref() {
            if let Some((u, m)) = selection.locate(&lookup) {
                let unit = &selection.units[u];
                return Route {
                    session: unit.session.clone(),
                    key: unit.members[m].key.clone(),
                    unit: Some(u),
                    path: Some(unit.members[m].path.clone()),
                };
            }
        }
        Route { session: self.loose.clone(), key: engine_glue::doc_key(uri), unit: None, path: None }
    }

    /// Windows-1252 text of every open document, by [`lookup_key`].
    fn overlay_map(&self) -> HashMap<String, Vec<u8>> {
        self.documents
            .iter()
            .map(|d| (paths::uri_lookup_key(d.key()), engine_glue::to_cp1252(&d.content.to_string())))
            .collect()
    }

    /// The editor's URI and version for each open document, by [`lookup_key`].
    fn open_documents(&self) -> HashMap<String, (Url, i32)> {
        self.documents
            .iter()
            .map(|d| (paths::uri_lookup_key(d.key()), (d.key().clone(), d.version)))
            .collect()
    }

    // ── Opening and closing a project ───────────────────────────────────────────

    /// Open `path` (a `.vbp` or `.vbg`) as the project, replacing the one that is open.
    pub(super) async fn open_selection(&self, path: PathBuf) -> std::result::Result<ProjectInfo, LoadError> {
        let _serial = self.load_lock.lock().await;
        self.send_status(ProjectStatus::loading(&path)).await;

        let overlay = self.overlay_map();
        let load_path = path.clone();
        let loaded = tokio::task::spawn_blocking(move || {
            Selection::load(&load_path, &|key: &str| overlay.get(key).cloned())
        })
        .await
        .unwrap_or_else(|e| {
            Err(LoadError { file: path.clone(), line: None, message: format!("loading failed: {e}") })
        });

        match loaded {
            Ok(selection) => {
                let info = ProjectInfo::of(&selection);
                let status = ProjectStatus::of(&selection);
                *self.selection.write().unwrap() = Some(selection);
                self.rebuild_loose();
                self.publish_everything().await;
                self.send_status(status).await;
                Ok(info)
            }
            Err(error) => {
                let status = match self.selection.read().unwrap().as_ref() {
                    // The previous project stays open; say why the new one did not open.
                    Some(current) => ProjectStatus {
                        message: Some(error.to_string()),
                        ..ProjectStatus::of(current)
                    },
                    None => ProjectStatus::failed(&path, error.to_string()),
                };
                self.send_status(status).await;
                self.client.show_message(MessageType::ERROR, format!("Cannot open project: {error}")).await;
                let problem = Problem {
                    file: error.file.clone(),
                    line: error.line.unwrap_or(0),
                    severity: Severity::Error,
                    message: error.message.clone(),
                };
                self.publish_problems(&[problem]).await;
                Err(error)
            }
        }
    }

    /// Stop analysing the opened project; its files are analysed on their own when open.
    pub(super) async fn close_selection(&self) {
        let _serial = self.load_lock.lock().await;
        *self.selection.write().unwrap() = None;
        self.rebuild_loose();
        self.publish_everything().await;
        self.send_status(ProjectStatus::none()).await;
    }

    pub(super) fn project_info(&self) -> Option<ProjectInfo> {
        self.selection.read().unwrap().as_ref().map(ProjectInfo::of)
    }

    /// The session of files outside the opened project holds exactly the open documents the
    /// project does not list.
    fn rebuild_loose(&self) {
        let selection = self.selection.read().unwrap();
        let files: Vec<(String, Vec<u8>)> = self
            .documents
            .iter()
            .filter(|d| {
                let lookup = paths::uri_lookup_key(d.key());
                selection.as_ref().is_none_or(|s| s.locate(&lookup).is_none())
            })
            .map(|d| (engine_glue::doc_key(d.key()), engine_glue::to_cp1252(&d.content.to_string())))
            .collect();
        *self.loose.write().unwrap() = Session::from_sources(files);
    }

    async fn send_status(&self, status: ProjectStatus) {
        self.client.send_notification::<ProjectStatusNotification>(status).await;
    }

    // ── Document changes ────────────────────────────────────────────────────────

    /// Put the document's current text into the analysis and publish diagnostics for every
    /// file the change reaches.
    pub(super) async fn sync_document(&self, uri: &Url) {
        let (content, version) = match self.documents.get(uri) {
            Some(doc) => (doc.content.to_string(), doc.version),
            None => return,
        };
        let route = self.route(uri);
        let bytes = engine_glue::to_cp1252(&content);
        let reach = self.update_one(&route, bytes);
        let diagnostics = {
            let session = route.session.read().unwrap();
            engine_glue::diagnostics_for(&session, &route.key)
        };
        self.publish(uri.clone(), diagnostics, Some(version), true).await;
        self.publish_reach(reach, Some(uri)).await;
    }

    /// Replace one file's content in its session. Whether the rest of the project is affected
    /// is read off how many modules the session bound again.
    fn update_one(&self, route: &Route, bytes: Vec<u8>) -> Reach {
        let full = {
            let mut session = route.session.write().unwrap();
            let before = session.bind_count();
            session.update_file(&route.key, bytes);
            session.bind_count() - before > 1
        };
        self.reach_of(route.unit, full)
    }

    fn reach_of(&self, unit: Option<usize>, full: bool) -> Reach {
        let mut reach = Reach::default();
        match unit {
            None => reach.loose = full,
            Some(u) => {
                if full {
                    reach.units.insert(u);
                }
                if let Some(selection) = self.selection.write().unwrap().as_mut() {
                    reach.units.extend(selection.propagate_exports(u));
                }
            }
        }
        reach
    }

    /// A document is closed: a project member goes back to its content on disk; any other
    /// file leaves the analysis.
    pub(super) async fn close_document(&self, uri: &Url) {
        let route = self.route(uri);
        self.documents.remove(uri);
        let reach = match (&route.unit, &route.path) {
            (Some(_), Some(path)) => match std::fs::read(path) {
                Ok(bytes) => self.update_one(&route, bytes),
                Err(_) => {
                    let full = {
                        let mut session = route.session.write().unwrap();
                        let before = session.bind_count();
                        session.remove_file(&route.key);
                        session.bind_count() - before > 1
                    };
                    self.reach_of(route.unit, full)
                }
            },
            _ => {
                let full = {
                    let mut session = route.session.write().unwrap();
                    let before = session.bind_count();
                    session.remove_file(&route.key);
                    session.bind_count() - before > 1
                };
                self.published.remove(uri);
                self.client.publish_diagnostics(uri.clone(), Vec::new(), None).await;
                self.reach_of(None, full)
            }
        };
        if route.unit.is_some() {
            self.republish_member(&route).await;
        }
        self.publish_reach(reach, None).await;
    }

    /// Publish the diagnostics of one project member from its session.
    async fn republish_member(&self, route: &Route) {
        let Some(path) = &route.path else { return };
        let uri = paths::uri_for_path(path);
        let Some(uri) = uri else { return };
        let diagnostics = {
            let session = route.session.read().unwrap();
            engine_glue::diagnostics_for(&session, &route.key)
        };
        self.publish(uri, diagnostics, None, false).await;
    }

    // ── Changes made outside the editor ─────────────────────────────────────────

    /// Files changed on disk: reload the opened project when its own file changed, and the
    /// members of the opened project that the editor does not have open.
    pub(super) async fn apply_disk_changes(&self, changes: Vec<(Url, DiskChange)>) {
        let mut reload: Option<PathBuf> = None;
        let mut closed: Option<()> = None;
        let mut by_unit: HashMap<usize, (Vec<(String, Vec<u8>)>, Vec<String>)> = HashMap::new();
        {
            let selection = self.selection.read().unwrap();
            let Some(selection) = selection.as_ref() else { return };
            let open = self.open_documents();
            for (uri, change) in &changes {
                let lookup = paths::uri_lookup_key(uri);
                let is_project_file = lookup_key(&selection.path) == lookup
                    || selection.unit_of_project_file(&lookup).is_some();
                if is_project_file {
                    match change {
                        DiskChange::Written => reload = Some(selection.path.clone()),
                        DiskChange::Deleted if lookup_key(&selection.path) == lookup => closed = Some(()),
                        DiskChange::Deleted => reload = Some(selection.path.clone()),
                    }
                    continue;
                }
                if open.contains_key(&lookup) {
                    continue;
                }
                let Some((u, m)) = selection.locate(&lookup) else { continue };
                let member = &selection.units[u].members[m];
                let entry = by_unit.entry(u).or_default();
                match change {
                    DiskChange::Written => {
                        if let Ok(bytes) = std::fs::read(&member.path) {
                            entry.0.push((member.key.clone(), bytes));
                        }
                    }
                    DiskChange::Deleted => entry.1.push(member.key.clone()),
                }
            }
        }

        if closed.is_some() {
            self.close_selection().await;
            self.client
                .show_message(MessageType::WARNING, "The open VB6 project file was deleted; the project is closed")
                .await;
            return;
        }
        if let Some(path) = reload {
            let _ = self.open_selection(path).await;
            return;
        }

        let mut reach = Reach::default();
        for (u, (writes, removals)) in by_unit {
            let session = {
                let selection = self.selection.read().unwrap();
                let Some(selection) = selection.as_ref() else { return };
                selection.units[u].session.clone()
            };
            let full = {
                let mut s = session.write().unwrap();
                let before = s.bind_count();
                s.update_files(writes.clone());
                s.remove_files(&removals);
                s.bind_count() != before
            };
            if let Some(selection) = self.selection.write().unwrap().as_mut() {
                for (key, _) in &writes {
                    selection.set_member_exists(u, key, true);
                }
                for key in &removals {
                    selection.set_member_exists(u, key, false);
                }
                reach.units.extend(selection.propagate_exports(u));
            }
            if full {
                reach.units.insert(u);
            }
        }
        self.publish_reach(reach, None).await;
    }

    // ── Publishing ──────────────────────────────────────────────────────────────

    /// Publish `diagnostics` for `uri` unless they are what the editor already has.
    async fn publish(&self, uri: Url, diagnostics: Vec<Diagnostic>, version: Option<i32>, force: bool) {
        if !force && self.published.get(&uri).is_some_and(|d| *d == diagnostics) {
            return;
        }
        self.published.insert(uri.clone(), diagnostics.clone());
        self.client.publish_diagnostics(uri, diagnostics, version).await;
    }

    async fn publish_reach(&self, reach: Reach, except: Option<&Url>) {
        let open = self.open_documents();
        let mut batch: Vec<(Url, Vec<Diagnostic>, Option<i32>)> = Vec::new();
        {
            let selection = self.selection.read().unwrap();
            if let Some(selection) = selection.as_ref() {
                for &u in &reach.units {
                    batch.extend(unit_diagnostics(selection, u, &open));
                }
            }
        }
        if reach.loose {
            batch.extend(self.loose_diagnostics(&open));
        }
        let skip = except.map(paths::uri_lookup_key);
        for (uri, diagnostics, version) in batch {
            if skip.as_deref() == Some(paths::uri_lookup_key(&uri).as_str()) {
                continue;
            }
            self.publish(uri, diagnostics, version, false).await;
        }
    }

    fn loose_diagnostics(&self, open: &HashMap<String, (Url, i32)>) -> Vec<(Url, Vec<Diagnostic>, Option<i32>)> {
        let session = self.loose.read().unwrap();
        let selection = self.selection.read().unwrap();
        open.iter()
            .filter(|(lookup, _)| selection.as_ref().is_none_or(|s| s.locate(lookup).is_none()))
            .map(|(_, (uri, version))| {
                let diagnostics = engine_glue::diagnostics_for(&session, &engine_glue::doc_key(uri));
                (uri.clone(), diagnostics, Some(*version))
            })
            .collect()
    }

    /// Bring the editor's diagnostics in line with the analysis: publish every project member's
    /// and every problem found in the project files, and clear what no longer applies.
    async fn publish_everything(&self) {
        let open = self.open_documents();
        let mut batch: Vec<(Url, Vec<Diagnostic>, Option<i32>)> = Vec::new();
        let problems = {
            let selection = self.selection.read().unwrap();
            match selection.as_ref() {
                Some(selection) => {
                    for u in 0..selection.units.len() {
                        batch.extend(unit_diagnostics(selection, u, &open));
                    }
                    selection.problems.clone()
                }
                None => Vec::new(),
            }
        };
        batch.extend(self.loose_diagnostics(&open));
        batch.extend(problem_diagnostics(&problems));

        let wanted: HashSet<Url> = batch.iter().map(|(u, _, _)| u.clone()).collect();
        let stale: Vec<Url> = self.published.iter().map(|e| e.key().clone()).filter(|u| !wanted.contains(u)).collect();
        for uri in stale {
            self.published.remove(&uri);
            self.client.publish_diagnostics(uri, Vec::new(), None).await;
        }
        for (uri, diagnostics, version) in batch {
            self.publish(uri, diagnostics, version, false).await;
        }
    }

    /// Publish problems found in project files (a failed load).
    async fn publish_problems(&self, problems: &[Problem]) {
        for (uri, diagnostics, version) in problem_diagnostics(problems) {
            self.publish(uri, diagnostics, version, true).await;
        }
    }
}

/// Diagnostics of every file of project `u`, addressed by the editor's URI for open files.
fn unit_diagnostics(
    selection: &Selection,
    u: usize,
    open: &HashMap<String, (Url, i32)>,
) -> Vec<(Url, Vec<Diagnostic>, Option<i32>)> {
    let unit = &selection.units[u];
    if !unit.listed {
        return Vec::new();
    }
    let session = unit.session.read().unwrap();
    unit.members
        .iter()
        .filter(|m| session.module_of(&m.key).is_some())
        .filter_map(|m| {
            let diagnostics = engine_glue::diagnostics_for(&session, &m.key);
            match open.get(&m.lookup) {
                Some((uri, version)) => Some((uri.clone(), diagnostics, Some(*version))),
                None => paths::uri_for_path(&m.path).map(|uri| (uri, diagnostics, None)),
            }
        })
        .collect()
}

/// One diagnostic per problem, on the problem's line of its project file; every project file
/// named by a problem appears, so a file whose problems are gone is cleared.
fn problem_diagnostics(problems: &[Problem]) -> Vec<(Url, Vec<Diagnostic>, Option<i32>)> {
    let mut files: Vec<PathBuf> = Vec::new();
    let mut by_file: HashMap<String, Vec<Diagnostic>> = HashMap::new();
    for p in problems {
        let key = lookup_key(&p.file);
        if !by_file.contains_key(&key) {
            files.push(p.file.clone());
        }
        by_file.entry(key).or_default().push(Diagnostic {
            range: line_range(&p.file, p.line),
            severity: Some(match p.severity {
                Severity::Error => DiagnosticSeverity::ERROR,
                Severity::Warning => DiagnosticSeverity::WARNING,
                Severity::Information => DiagnosticSeverity::INFORMATION,
            }),
            message: p.message.clone(),
            source: Some("vb6-lsp".to_string()),
            ..Default::default()
        });
    }
    files
        .into_iter()
        .filter_map(|f| {
            let uri = paths::uri_for_path(&f)?;
            let diagnostics = by_file.remove(&lookup_key(&f))?;
            Some((uri, diagnostics, None))
        })
        .collect()
}

/// The whole of 1-based `line` of `file` (the first line when `line` is 0).
fn line_range(file: &Path, line: usize) -> Range {
    let row = line.saturating_sub(1) as u32;
    let width = std::fs::read(file)
        .ok()
        .and_then(|bytes| {
            let text = String::from_utf8_lossy(&bytes).into_owned();
            text.lines().nth(row as usize).map(|l| l.trim_end().chars().count() as u32)
        })
        .unwrap_or(0);
    Range { start: Position { line: row, character: 0 }, end: Position { line: row, character: width } }
}

/// The path a command argument names: a `file:` URI or a plain path.
pub(super) fn path_argument(value: &Value) -> Option<PathBuf> {
    let text = value.as_str()?;
    match Url::parse(text) {
        Ok(url) if url.scheme() == "file" => paths::path_for_uri(&url),
        _ => Some(PathBuf::from(text)),
    }
}
