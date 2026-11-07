//! Types a library takes from another library (`stdole.IDispatch` and the like), found by the
//! GUID the importing library records for them.

use super::pe::type_library_bytes;
use msft_typelib::TypeLib;
use std::path::PathBuf;

/// A type of another library.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct ImportedType {
    /// The name of the library that declares the type.
    pub library: String,
    pub name: String,
    /// An interface, dispatch interface or class: a type whose values are object pointers.
    pub object: bool,
}

/// What an importing library records about an imported type.
#[derive(Debug, Clone, Copy)]
pub struct ImportRequest<'a> {
    /// The imported library's file name, as the importing library writes it.
    pub file: &'a str,
    /// The imported library's GUID, lower case without braces.
    pub library_guid: &'a str,
    /// The imported type's GUID, lower case without braces, when the import is by GUID.
    pub type_guid: Option<&'a str>,
    /// The imported type's index in its library, when the import is by position.
    pub type_index: Option<usize>,
}

pub trait ImportResolver {
    fn resolve(&self, request: &ImportRequest<'_>) -> Option<ImportedType>;
}

/// A GUID as the comparison key: lower case, without braces.
pub fn guid_key(text: &str) -> String {
    text.trim().trim_matches(|c| c == '{' || c == '}').to_ascii_lowercase()
}

const STDOLE_GUID: &str = "00020430-0000-0000-c000-000000000046";

/// Finds the imported library as a file in `directories`, else knows the OLE automation types.
pub struct SearchImports {
    pub directories: Vec<PathBuf>,
}

impl SearchImports {
    /// The bytes of the file `name` names, found in the directories without regard to case, and
    /// the resource index the name ends with (`MSVBVM60.DLL`), if any.
    fn find_file(&self, name: &str) -> Option<(Vec<u8>, Option<u32>)> {
        let last = name.rsplit(['\\', '/']).next()?;
        let (file_name, index) = match name.rsplit_once(['\\', '/']) {
            Some((before, tail)) if tail.chars().all(|c| c.is_ascii_digit()) && !tail.is_empty() => {
                (before.rsplit(['\\', '/']).next().unwrap_or(before), tail.parse().ok())
            }
            _ => (last, None),
        };
        for dir in &self.directories {
            let Ok(entries) = std::fs::read_dir(dir) else { continue };
            for entry in entries.flatten() {
                if entry.file_name().to_string_lossy().eq_ignore_ascii_case(file_name) {
                    if let Ok(bytes) = std::fs::read(entry.path()) {
                        return Some((bytes, index));
                    }
                }
            }
        }
        None
    }
}

impl ImportResolver for SearchImports {
    fn resolve(&self, request: &ImportRequest<'_>) -> Option<ImportedType> {
        if let Some((bytes, index)) = self.find_file(request.file) {
            if let Ok(image) = type_library_bytes(&bytes, index) {
                if let Ok(lib) = TypeLib::parse(image) {
                    if let Some(found) = type_in(&lib, request) {
                        return Some(found);
                    }
                }
            }
        }
        if request.library_guid == STDOLE_GUID {
            return super::stdole_table::stdole_type(request.type_guid, request.type_index);
        }
        None
    }
}

fn type_in(lib: &TypeLib<'_>, request: &ImportRequest<'_>) -> Option<ImportedType> {
    let index = match (request.type_guid, request.type_index) {
        (Some(wanted), _) => (0..lib.typeinfo_count()).find(|&i| {
            lib.typeinfo(i)
                .ok()
                .and_then(|ti| lib.guid(ti.guid_offset()))
                .is_some_and(|g| guid_key(&g.to_string()) == wanted)
        })?,
        (None, Some(i)) => i,
        (None, None) => return None,
    };
    super::msft::imported_type(lib, index)
}
