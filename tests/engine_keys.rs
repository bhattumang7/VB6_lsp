//! The engine's key for a document is the path's key, whichever way the editor spells the URI.

use std::path::Path;

use tower_lsp::lsp_types::Url;
use vb6_lsp::engine_glue::doc_key;
use vb6_project::paths::{engine_key_for_path, path_for_uri, uri_lookup_key};

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

