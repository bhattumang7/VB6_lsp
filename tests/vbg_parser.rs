//! `.vbg` project-group files.

use std::path::{Path, PathBuf};

use vb6_lsp::workspace::{GroupProject, VbgFile, VbgParseError};

fn group_path() -> PathBuf {
    std::env::temp_dir().join("groups").join("All.vbg")
}

fn dir() -> PathBuf {
    group_path().parent().unwrap().to_path_buf()
}

#[test]
fn reads_the_header_projects_and_startup_project() {
    let text = "VBGROUP 5.0\nStartupProject=App\\App.vbp\nProject=Lib\\Lib.vbp\nProject=App\\App.vbp\nPackage=ignored\n";

    let g = VbgFile::parse_content(&group_path(), text).unwrap();

    assert_eq!(
        g,
        VbgFile {
            version: "5.0".to_string(),
            projects: vec![
                GroupProject {
                    relative_path: PathBuf::from("Lib\\Lib.vbp"),
                    absolute_path: dir().join("Lib\\Lib.vbp"),
                    line: 3,
                },
                GroupProject {
                    relative_path: PathBuf::from("App\\App.vbp"),
                    absolute_path: dir().join("App\\App.vbp"),
                    line: 4,
                },
            ],
            startup: Some(dir().join("App\\App.vbp")),
            startup_line: Some(2),
        }
    );
}

#[test]
fn a_group_without_a_startup_entry_has_none() {
    let g = VbgFile::parse_content(&group_path(), "VBGROUP 5.0\nProject=A.vbp\n").unwrap();

    assert_eq!(g.startup, None);
    assert_eq!(g.startup_line, None);
    assert_eq!(g.projects.len(), 1);
}

#[test]
fn quotes_blank_lines_and_key_case_are_tolerated() {
    let text = "\n\nvbgroup 6.0\n\nproject=\"Quoted\\Q.vbp\"\nSTARTUPPROJECT=Quoted\\Q.vbp\nProject=\nnot a setting\n";

    let g = VbgFile::parse_content(&group_path(), text).unwrap();

    assert_eq!(g.version, "6.0");
    assert_eq!(
        g.projects,
        vec![GroupProject {
            relative_path: PathBuf::from("Quoted\\Q.vbp"),
            absolute_path: dir().join("Quoted\\Q.vbp"),
            line: 5,
        }]
    );
    assert_eq!(g.startup_line, Some(6));
}

#[test]
fn an_absolute_project_path_is_kept() {
    let absolute = std::env::temp_dir().join("elsewhere").join("Z.vbp");
    let text = format!("VBGROUP 5.0\nProject={}\n", absolute.display());

    let g = VbgFile::parse_content(&group_path(), &text).unwrap();

    assert_eq!(g.projects[0].absolute_path, absolute);
}

#[test]
fn a_missing_header_is_an_error() {
    let err = VbgFile::parse_content(Path::new("All.vbg"), "\n  \n").unwrap_err();

    assert_eq!(err, VbgParseError { message: "missing `VBGROUP` header".to_string(), line: None });
}

#[test]
fn a_first_line_that_is_not_the_header_is_an_error() {
    let err = VbgFile::parse_content(Path::new("All.vbg"), "\nProject=A.vbp\n").unwrap_err();

    assert_eq!(
        err,
        VbgParseError { message: "the first line must be the `VBGROUP` header".to_string(), line: Some(2) }
    );
}

#[test]
fn a_file_that_cannot_be_read_is_an_error() {
    let err = VbgFile::parse(&dir().join("does-not-exist.vbg")).unwrap_err();

    assert_eq!(err.line, None);
    assert!(err.message.starts_with("Failed to read file"), "{}", err.message);
}
