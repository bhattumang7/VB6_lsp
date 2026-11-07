//! Path identity. The editor names a file by URI (`file:///c%3A/Proj/a.bas`), a project file
//! by a path relative to itself (`..\Proj\A.bas`); both must lead to the same engine key.

use std::path::{Component, Path, PathBuf};

use tower_lsp::lsp_types::Url;

/// Resolve `.` and `..` components without touching the file system.
pub fn normalize(path: &Path) -> PathBuf {
    let mut out = PathBuf::new();
    for component in path.components() {
        match component {
            Component::CurDir => {}
            Component::ParentDir => {
                let can_pop = matches!(out.components().next_back(), Some(Component::Normal(_)));
                if can_pop {
                    out.pop();
                } else if !matches!(
                    out.components().next_back(),
                    Some(Component::RootDir | Component::Prefix(_))
                ) {
                    out.push("..");
                }
            }
            other => out.push(other.as_os_str()),
        }
    }
    out
}

/// A key under which two spellings of one file compare equal: normalized, and
/// case-folded where the file system is case-insensitive.
pub fn lookup_key(path: &Path) -> String {
    let text = normalize(path).to_string_lossy().replace('/', std::path::MAIN_SEPARATOR_STR);
    if cfg!(any(windows, target_os = "macos")) {
        text.to_lowercase()
    } else {
        text
    }
}

/// The URI a file is known by to the editor and the engine.
pub fn uri_for_path(path: &Path) -> Option<Url> {
    Url::from_file_path(normalize(path)).ok()
}

/// The engine's key for a file: its URI. Falls back to the path text when the path cannot
/// be a URI (a relative path).
pub fn engine_key_for_path(path: &Path) -> String {
    match uri_for_path(path) {
        Some(uri) => uri.to_string(),
        None => normalize(path).to_string_lossy().into_owned(),
    }
}

/// The file a `file:` URI names; `None` for other schemes.
pub fn path_for_uri(uri: &Url) -> Option<PathBuf> {
    if uri.scheme() != "file" {
        return None;
    }
    uri.to_file_path().ok()
}

/// The lookup key of a document URI (its path's key); non-file URIs key as their text.
pub fn uri_lookup_key(uri: &Url) -> String {
    match path_for_uri(uri) {
        Some(path) => lookup_key(&path),
        None => uri.to_string(),
    }
}
