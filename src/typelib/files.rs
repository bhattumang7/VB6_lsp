//! Reading a type library from the files a project names, without the operating system.

use super::imports::{guid_key, SearchImports};
use super::msft::{library_guid, read_library};
use super::pe::type_library_bytes;
use super::{TypeLibraryError, TypeLibraryRequest};
use std::path::{Path, PathBuf};
use vb6_engine::session::Library;

/// `text` with the separators of the running platform: a project written on Windows names its
/// files with backslashes.
pub fn with_native_separators(text: &str) -> String {
    if std::path::MAIN_SEPARATOR == '\\' {
        text.to_string()
    } else {
        text.replace('\\', "/")
    }
}

/// The existing file `path` names, matching each part of the path without regard to case (a
/// project written on Windows does not agree with the case of the file system it is read on).
pub fn find_existing(path: &Path) -> Option<PathBuf> {
    if path.exists() {
        return Some(path.to_path_buf());
    }
    let mut found = PathBuf::new();
    for component in path.components() {
        let part = component.as_os_str();
        let direct = found.join(part);
        if direct.exists() || found.as_os_str().is_empty() && !matches!(component, std::path::Component::Normal(_)) {
            found = direct;
            continue;
        }
        let base = if found.as_os_str().is_empty() { PathBuf::from(".") } else { found.clone() };
        let wanted = part.to_string_lossy().to_lowercase();
        let matched = std::fs::read_dir(&base)
            .ok()?
            .flatten()
            .find(|entry| entry.file_name().to_string_lossy().to_lowercase() == wanted)?;
        found = found.join(matched.file_name());
    }
    found.exists().then_some(found)
}

/// The file and resource index a path names: `file.dll\2` selects the second type library of
/// `file.dll`.
fn split_index(path: &Path) -> (PathBuf, Option<u32>) {
    if path.exists() {
        return (path.to_path_buf(), None);
    }
    let text = path.to_string_lossy();
    if let Some((file, index)) = text.rsplit_once(['\\', '/']) {
        if let Ok(n) = index.parse::<u32>() {
            return (PathBuf::from(file), Some(n));
        }
    }
    (path.to_path_buf(), None)
}

/// Read the library of the first file of `request.paths` that holds it, with the files of
/// imported libraries that were not found. Libraries it imports are looked for next to the files
/// named.
pub fn read_from_files(request: &TypeLibraryRequest) -> Result<(Library, Vec<String>), TypeLibraryError> {
    read_from_files_searching(request, &[])
}

/// As [`read_from_files`], also looking for imported libraries in `extra_directories`.
pub fn read_from_files_searching(
    request: &TypeLibraryRequest,
    extra_directories: &[PathBuf],
) -> Result<(Library, Vec<String>), TypeLibraryError> {
    let wanted = guid_key(&request.guid);
    let mut tried: Vec<String> = Vec::new();
    let mut reason = String::new();
    let mut directories: Vec<PathBuf> = request
        .paths
        .iter()
        .filter_map(|p| split_index(p).0.parent().map(Path::to_path_buf))
        .collect();
    directories.extend(extra_directories.iter().cloned());
    let imports = SearchImports { directories };
    for written in &request.paths {
        let path = PathBuf::from(with_native_separators(&written.to_string_lossy()));
        tried.push(written.display().to_string());
        let (candidate, index) = split_index(&path);
        let Some(file) = find_existing(&candidate).filter(|f| f.is_file()) else { continue };
        let bytes = match std::fs::read(&file) {
            Ok(b) => b,
            Err(e) => {
                reason = format!("{}: {e}", written.display());
                continue;
            }
        };
        let image = match type_library_bytes(&bytes, index) {
            Ok(i) => i,
            Err(e) => {
                reason = format!("{}: {e}", written.display());
                continue;
            }
        };
        match library_guid(image) {
            Some(found) if found == wanted => {}
            Some(_) => {
                reason = format!("{} holds a different library", written.display());
                continue;
            }
            None => {
                reason = format!("{}: not a readable type library", written.display());
                continue;
            }
        }
        return read_library(image, &imports)
            .map_err(|e| TypeLibraryError::Unreadable { what: request.display_name(), reason: e });
    }
    Err(TypeLibraryError::NotFound { what: request.display_name(), tried, reason })
}
