//! The file reader describes every installed type library exactly as the operating system's
//! type-information interfaces do.
#![cfg(windows)]

use std::path::{Path, PathBuf};
use vb6_engine::session::Library;
use vb6_lsp::typelib::files::read_from_files_searching;
use vb6_lsp::typelib::msft::exported_types;
use vb6_lsp::typelib::pe::type_library_bytes;
use vb6_lsp::typelib::{load_library, TypeLibraryRequest};

const SYSTEM_DIRECTORIES: [&str; 3] =
    [r"C:\Windows\System32", r"C:\Windows\SysWOW64", r"C:\Program Files (x86)\Common Files\System\ado"];

/// The installed libraries to compare: type libraries, OCX controls and ADO.
fn corpus() -> Vec<PathBuf> {
    let mut files = Vec::new();
    for dir in SYSTEM_DIRECTORIES {
        let Ok(entries) = std::fs::read_dir(dir) else { continue };
        for entry in entries.flatten() {
            let path = entry.path();
            let name = path.file_name().map(|n| n.to_string_lossy().to_ascii_lowercase()).unwrap_or_default();
            let ext = path.extension().map(|e| e.to_string_lossy().to_ascii_lowercase()).unwrap_or_default();
            if ["tlb", "olb", "ocx"].contains(&ext.as_str()) || name == "msado15.dll" {
                files.push(path);
            }
        }
    }
    files
}

fn request_for(path: &Path, bytes: &[u8]) -> Option<TypeLibraryRequest> {
    let image = type_library_bytes(bytes, None).ok()?;
    let lib = msft_typelib::TypeLib::parse(image).ok()?;
    Some(TypeLibraryRequest {
        guid: lib.lib_guid()?.to_string(),
        major: 0,
        minor: 0,
        lcid: 0,
        paths: vec![path.to_path_buf()],
        description: String::new(),
    })
}

/// The first place two descriptions of one library disagree, or `None`.
fn first_difference(file: &Library, system: &Library) -> Option<String> {
    if file == system {
        return None;
    }
    if file.name != system.name || file.types.len() != system.types.len() {
        return Some(format!("library {} ({} types) against {} ({} types)", file.name, file.types.len(), system.name, system.types.len()));
    }
    for (a, b) in file.types.iter().zip(&system.types) {
        if a == b {
            continue;
        }
        for (m, n) in a.members.iter().zip(&b.members) {
            if m != n {
                return Some(format!("type {} member {}:\n  file   {m:?}\n  system {n:?}", a.name, m.name));
            }
        }
        return Some(format!("type {}:\n  file   {a:?}\n  system {b:?}", a.name));
    }
    Some("different".to_string())
}

#[test]
fn every_installed_library_reads_the_same_from_its_file_as_through_the_system() {
    let search: Vec<PathBuf> = SYSTEM_DIRECTORIES.iter().map(PathBuf::from).collect();
    let (mut compared, mut differences) = (0, Vec::new());
    for path in corpus() {
        let Ok(bytes) = std::fs::read(&path) else { continue };
        let Some(request) = request_for(&path, &bytes) else { continue };
        let (Ok((from_file, missing)), Ok(from_system)) = (read_from_files_searching(&request, &search), load_library(&request)) else {
            continue;
        };
        if !missing.is_empty() {
            continue;
        }
        compared += 1;
        if let Some(difference) = first_difference(&from_file, &from_system) {
            differences.push(format!("{}: {difference}", path.display()));
        }
    }
    assert!(compared >= 10, "only {compared} installed libraries could be compared");
    assert_eq!(differences, Vec::<String>::new());
}

#[test]
fn libraries_that_take_types_only_from_the_automation_library_read_alone_the_same_way() {
    // Copied away from every other library, a file can resolve only what it needs from the
    // automation library, which the reader knows.
    let dir = std::env::temp_dir().join(format!("vb6lsp-alone-{}", std::process::id()));
    std::fs::create_dir_all(&dir).unwrap();
    let (mut compared, mut differences) = (0, Vec::new());
    for path in corpus() {
        let Ok(bytes) = std::fs::read(&path) else { continue };
        let Some(mut request) = request_for(&path, &bytes) else { continue };
        let Ok(from_system) = load_library(&request) else { continue };
        let copy = dir.join(path.file_name().unwrap());
        std::fs::write(&copy, &bytes).unwrap();
        request.paths = vec![copy.clone()];
        let result = read_from_files_searching(&request, &[]);
        std::fs::remove_file(&copy).unwrap();
        let Ok((from_file, missing)) = result else { continue };
        if !missing.is_empty() {
            continue;
        }
        compared += 1;
        if let Some(difference) = first_difference(&from_file, &from_system) {
            differences.push(format!("{}: {difference}", path.display()));
        }
    }
    std::fs::remove_dir_all(&dir).unwrap();
    assert!(compared >= 5, "only {compared} installed libraries could be compared");
    assert_eq!(differences, Vec::<String>::new());
}

#[test]
fn the_built_in_description_of_the_automation_library_matches_the_installed_one() {
    for dir in SYSTEM_DIRECTORIES {
        let path = Path::new(dir).join("stdole2.tlb");
        let Ok(bytes) = std::fs::read(&path) else { continue };
        let image = type_library_bytes(&bytes, None).unwrap();
        let installed = exported_types(image).unwrap();
        let built_in: Vec<(usize, Option<String>, String, String, bool)> = installed
            .iter()
            .map(|(index, guid, t)| (*index, guid.clone(), t.library.clone(), t.name.clone(), t.object))
            .collect();
        for (index, guid, library, name, object) in built_in {
            let by_index = vb6_lsp::typelib::stdole_type(None, Some(index)).unwrap();
            assert_eq!((by_index.library.as_str(), by_index.name.as_str(), by_index.object), (library.as_str(), name.as_str(), object), "index {index}");
            if let Some(guid) = guid {
                let by_guid = vb6_lsp::typelib::stdole_type(Some(&guid), None).unwrap();
                assert_eq!(by_guid, by_index, "guid {guid}");
            }
        }
        return;
    }
    panic!("the automation library is not installed");
}
