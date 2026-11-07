//! Reading the type libraries a project references (`Reference=*\G{...}` and `Object={...}`
//! lines of a `.vbp`) so the analysis sees their types.
//!
//! A library is located by its GUID and version in the registry, else by the file the project
//! names, and described as an engine [`Library`]. On Windows the registry locates the library and
//! the automation type-information interfaces read it; elsewhere the file the project names is
//! read directly (a `.tlb`, or the `TYPELIB` resource of a DLL or OCX).

pub mod files;
pub mod imports;
pub mod locate;
pub mod msft;
pub mod names;
pub mod pe;
mod stdole_table;

#[cfg(windows)]
mod com;
#[cfg(windows)]
mod registry;

use std::collections::HashMap;
use std::path::PathBuf;
use std::sync::{Arc, Mutex, OnceLock};
use vb6_engine::session::Library;

pub use stdole_table::stdole_type;

/// A library a project references.
#[derive(Debug, Clone, PartialEq, Eq, Hash)]
pub struct TypeLibraryRequest {
    /// The library's GUID, with or without braces.
    pub guid: String,
    pub major: u16,
    pub minor: u16,
    pub lcid: u32,
    /// Files to read the library from when the registry does not name it, in order.
    pub paths: Vec<PathBuf>,
    /// What the project calls the reference, for messages.
    pub description: String,
}

impl TypeLibraryRequest {
    /// How messages name the library.
    pub fn display_name(&self) -> String {
        let guid = self.guid.trim().trim_matches(|c| c == '{' || c == '}');
        if self.description.is_empty() {
            format!("{{{guid}}} {}.{}", self.major, self.minor)
        } else {
            format!("{} {{{guid}}} {}.{}", self.description, self.major, self.minor)
        }
    }

    fn cache_key(&self) -> (String, u16, u16, u32, Vec<PathBuf>) {
        (
            self.guid.trim().trim_matches(|c| c == '{' || c == '}').to_ascii_lowercase(),
            self.major,
            self.minor,
            self.lcid,
            self.paths.clone(),
        )
    }
}

/// Why a type library could not be read.
#[derive(Debug, Clone, PartialEq, Eq)]
pub enum TypeLibraryError {
    /// Neither the registry nor the files the project names hold the library.
    NotFound { what: String, tried: Vec<String>, reason: String },
    /// The library was found but its description could not be read.
    Unreadable { what: String, reason: String },
}

impl std::fmt::Display for TypeLibraryError {
    fn fmt(&self, f: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        match self {
            TypeLibraryError::NotFound { what, tried, reason } => {
                if !reason.is_empty() {
                    write!(f, "Type library {what} could not be loaded: {reason}")
                } else if tried.is_empty() {
                    write!(f, "Type library {what} is not registered and the project names no file for it")
                } else {
                    write!(f, "Type library {what} is not registered and was not found at {}", tried.join(", "))
                }
            }
            TypeLibraryError::Unreadable { what, reason } => {
                write!(f, "Type library {what} could not be read: {reason}")
            }
        }
    }
}

impl std::error::Error for TypeLibraryError {}

type Key = (String, u16, u16, u32, Vec<PathBuf>);

/// A library read, with what the reader could not describe.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct LoadedLibrary {
    pub library: Arc<Library>,
    /// Files of libraries this one takes types from that were not found: the types taken from
    /// them are left undescribed.
    pub missing_imports: Vec<String>,
}

fn cache() -> &'static Mutex<HashMap<Key, LoadedLibrary>> {
    static CACHE: OnceLock<Mutex<HashMap<Key, LoadedLibrary>>> = OnceLock::new();
    CACHE.get_or_init(|| Mutex::new(HashMap::new()))
}

/// Read the library `request` names. A library read once is kept for the life of the process,
/// and two requests for the same library never read it twice at the same time.
pub fn load_library_noted(request: &TypeLibraryRequest) -> Result<LoadedLibrary, TypeLibraryError> {
    let key = request.cache_key();
    let mut kept = cache().lock().unwrap_or_else(|poisoned| poisoned.into_inner());
    if let Some(found) = kept.get(&key) {
        return Ok(found.clone());
    }
    let (library, missing_imports) = read(request)?;
    let loaded = LoadedLibrary { library: Arc::new(library), missing_imports };
    kept.insert(key, loaded.clone());
    Ok(loaded)
}

/// As [`load_library_noted`], without the notes.
pub fn load_library(request: &TypeLibraryRequest) -> Result<Arc<Library>, TypeLibraryError> {
    load_library_noted(request).map(|loaded| loaded.library)
}

#[cfg(windows)]
fn read(request: &TypeLibraryRequest) -> Result<(Library, Vec<String>), TypeLibraryError> {
    com::read(request).map(|library| (library, Vec::new()))
}

#[cfg(not(windows))]
fn read(request: &TypeLibraryRequest) -> Result<(Library, Vec<String>), TypeLibraryError> {
    files::read_from_files(request)
}
