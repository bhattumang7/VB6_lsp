//! One file, many spellings: the editor's URIs and a project's relative paths must agree.

use std::path::{Path, PathBuf};

use tower_lsp::lsp_types::Url;
use vb6_lsp::engine_glue::doc_key;
use vb6_lsp::project::paths::{
    engine_key_for_path, lookup_key, normalize, path_for_uri, uri_for_path, uri_lookup_key,
};

#[test]
fn normalize_resolves_current_and_parent_components() {
    assert_eq!(normalize(Path::new("a/./b/../c")), PathBuf::from("a").join("c"));
}

#[test]
fn normalize_keeps_a_leading_parent_of_a_relative_path() {
    assert_eq!(normalize(Path::new("../x/./y")), PathBuf::from("..").join("x").join("y"));
}

#[test]
fn normalize_does_not_climb_above_the_root() {
    let root = std::env::temp_dir().ancestors().last().unwrap().to_path_buf();
    let climbing = root.join("..").join("x");

    assert_eq!(normalize(&climbing), root.join("x"));
}

#[test]
fn normalize_leaves_a_normal_path_alone() {
    let p = std::env::temp_dir().join("Proj").join("a.bas");

    assert_eq!(normalize(&p), p);
}

#[cfg(windows)]
#[test]
fn spellings_differing_in_case_dots_and_separators_share_a_key() {
    let a = lookup_key(Path::new(r"C:\Proj\Mod\A.bas"));
    let b = lookup_key(Path::new(r"c:\proj\Other\..\Mod\a.BAS"));
    let c = lookup_key(Path::new("C:/Proj/Mod/A.bas"));

    assert_eq!(a, r"c:\proj\mod\a.bas");
    assert_eq!(a, b);
    assert_eq!(a, c);
}

#[cfg(windows)]
#[test]
fn the_editors_encoded_uri_matches_the_project_path() {
    let from_editor = Url::parse("file:///c%3A/Proj/Mod/A.bas").unwrap();
    let from_project = Url::parse("file:///C:/Proj/Other/../Mod/a.bas").unwrap();

    assert_eq!(path_for_uri(&from_editor), Some(PathBuf::from(r"c:\Proj\Mod\A.bas")));
    assert_eq!(uri_lookup_key(&from_editor), uri_lookup_key(&from_project));
    assert_eq!(uri_lookup_key(&from_editor), lookup_key(Path::new(r"C:\Proj\Mod\A.bas")));
}

#[cfg(windows)]
#[test]
fn the_engine_key_normalizes_the_uri_but_keeps_the_spelling_of_the_path() {
    let uri = Url::parse("file:///C:/Proj/Other/../Mod/A.bas").unwrap();

    assert_eq!(doc_key(&uri), "file:///C:/Proj/Mod/A.bas");
    assert_eq!(engine_key_for_path(Path::new(r"C:\Proj\Other\..\Mod\A.bas")), "file:///C:/Proj/Mod/A.bas");
}

#[test]
fn a_uri_that_names_no_file_is_keyed_by_its_text() {
    let uri = Url::parse("untitled:Untitled-1").unwrap();

    assert_eq!(path_for_uri(&uri), None);
    assert_eq!(uri_lookup_key(&uri), "untitled:Untitled-1");
    assert_eq!(doc_key(&uri), "untitled:Untitled-1");
}

#[test]
fn a_path_round_trips_through_its_uri() {
    let p = std::env::temp_dir().join("Proj").join("My Module.bas");

    let uri = uri_for_path(&p).unwrap();

    assert_eq!(path_for_uri(&uri), Some(p.clone()));
    assert_eq!(engine_key_for_path(&p), uri.to_string());
}

#[test]
fn a_relative_path_has_no_uri_and_keys_as_its_text() {
    let p = Path::new("rel").join("a.bas");

    assert_eq!(uri_for_path(&p), None);
    assert_eq!(engine_key_for_path(&p), p.to_string_lossy());
}
