//! Every member and reference of a `.vbp` knows its line, so a problem can be shown there.

use std::path::{Path, PathBuf};

use vb6_formats::vbp::{TypeLibReference as VbpReference, VbpFile};

const VBP: &str = "Type=Exe\n\
Reference=*\\G{00020430-0000-0000-C000-000000000046}#2.0#0#..\\stdole2.tlb#OLE Automation\n\
Reference=*\\A..\\Lib\\Lib.vbp\n\
Object={831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0; MSCOMCTL.OCX\n\
\n\
Module=modMain; modMain.bas\n\
Class=clsThing; clsThing.cls\n\
Form=frmMain.frm\n\
Name=\"App\"\n\
Title=\"Shown title\"\n";

fn project() -> VbpFile {
    VbpFile::parse_content(&std::env::temp_dir().join("p").join("App.vbp"), VBP).unwrap()
}

#[test]
fn members_carry_the_line_of_their_entry() {
    let vbp = project();

    let lines: Vec<(&str, usize)> = vbp.all_source_files().map(|m| (m.name.as_str(), m.line)).collect();

    assert_eq!(lines, vec![("modMain", 6), ("clsThing", 7), ("", 8)]);
}

#[test]
fn references_and_objects_carry_the_line_of_their_entry() {
    let vbp = project();

    match &vbp.references[0] {
        VbpReference::Compiled { guid, major, minor, lcid, description, line, .. } => {
            assert_eq!(guid, "00020430-0000-0000-C000-000000000046");
            assert_eq!((*major, *minor, *lcid), (2, 0, 0));
            assert_eq!(description, "OLE Automation");
            assert_eq!(*line, 2);
        }
        other => panic!("expected a compiled reference, got {other:?}"),
    }
    match &vbp.references[1] {
        VbpReference::Subproject { path, line } => {
            assert_eq!(path, &std::env::temp_dir().join("p").join("..\\Lib\\Lib.vbp"));
            assert_eq!(*line, 3);
        }
        other => panic!("expected a project reference, got {other:?}"),
    }
    assert_eq!(vbp.objects.len(), 1);
    assert_eq!(vbp.objects[0].filename, "MSCOMCTL.OCX");
    assert_eq!(vbp.objects[0].line, 4);
}

#[test]
fn the_project_name_is_the_name_entry_not_the_title() {
    assert_eq!(project().name, "App");
}

#[test]
fn without_a_name_entry_the_project_is_named_after_its_file() {
    let vbp = VbpFile::parse_content(Path::new("C:\\p\\Fallback.vbp"), "Type=Exe\nTitle=\"Shown\"\n").unwrap();

    assert_eq!(vbp.name, "Fallback");
}

#[test]
fn member_paths_resolve_against_the_project_folder() {
    let vbp = project();

    let m = vbp.all_source_files().next().unwrap();

    assert_eq!(m.relative_path, PathBuf::from("modMain.bas"));
    assert_eq!(m.absolute_path, std::env::temp_dir().join("p").join("modMain.bas"));
}
