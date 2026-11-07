//! Opening a `.vbp` or `.vbg`: which projects and members are found, what is wrong with them,
//! and how each project's analysis is set up.

mod common;

use common::TempDir;
use vb6_engine::session::ProjectKind;
use vb6_lsp::project::{
    LoadError, MemberFile, Problem, ReferenceInfo, ReferenceKind, ReferenceStatus, Selection, Severity,
};

fn no_overlay(_: &str) -> Option<Vec<u8>> {
    None
}

const MAIN: &str = "Option Explicit\nSub Run()\n    Greet\nEnd Sub\n";
const UTIL: &str = "Option Explicit\nPublic Sub Greet()\nEnd Sub\n";

const EXPOSED_CLASS: &str = "VERSION 1.0 CLASS\nBEGIN\n  MultiUse = -1  'True\nEND\n\
Attribute VB_Name = \"clsThing\"\nAttribute VB_GlobalNameSpace = False\nAttribute VB_Creatable = True\n\
Attribute VB_PredeclaredId = False\nAttribute VB_Exposed = True\nPublic Sub Hello()\nEnd Sub\n";

fn member(name: &str, ext: &str, path: std::path::PathBuf, line: usize, exists: bool) -> MemberFile {
    let lookup = vb6_lsp::project::paths::lookup_key(&path);
    let key = vb6_lsp::project::paths::engine_key_for_path(&path);
    MemberFile { name: name.into(), extension: ext.into(), path, key, lookup, line, exists }
}

#[test]
fn a_project_lists_its_members_and_analyses_them_together() {
    let dir = TempDir::new("single");
    let vbp = dir.write(
        "App.vbp",
        "Type=Exe\nModule=modMain; modMain.bas\nModule=modUtil; src\\modUtil.bas\nName=\"App\"\nTitle=\"Shown title\"\nCondComp=\"DEBUG = 1\"\n",
    );
    dir.write("modMain.bas", MAIN);
    dir.write("src/modUtil.bas", UTIL);

    let s = Selection::load(&vbp, &no_overlay).unwrap();

    assert_eq!(s.path, vbp);
    assert!(!s.is_group);
    assert_eq!(s.name, "App");
    assert_eq!(s.problems, Vec::<Problem>::new());
    assert_eq!(s.units.len(), 1);
    let u = &s.units[0];
    assert_eq!(u.name, "App");
    assert_eq!(u.vbp_path, vbp);
    assert_eq!(u.kind, ProjectKind::Exe);
    assert!(u.listed);
    assert_eq!(u.conditional_compilation.as_deref(), Some("DEBUG = 1"));
    assert_eq!(
        u.members,
        vec![
            member("modMain", "bas", dir.path("modMain.bas"), 2, true),
            member("modUtil", "bas", dir.path("src/modUtil.bas"), 3, true),
        ]
    );
    // Both files are analysed, so the call in one resolves to the other: no undeclared name.
    let session = u.session.read().unwrap();
    assert_eq!(session.module_count(), 2);
    let main = session.module_of(&u.members[0].key).unwrap();
    assert_eq!(session.diagnostics(main).len(), 0);
}

#[test]
fn a_project_is_linked_once_however_many_files_it_has() {
    let dir = TempDir::new("linear");
    let mut vbp = String::from("Type=Exe\nName=\"Big\"\n");
    for i in 0..40 {
        vbp.push_str(&format!("Module=mod{i}; mod{i}.bas\n"));
        dir.write(&format!("mod{i}.bas"), &format!("Option Explicit\nPublic Sub P{i}()\nEnd Sub\n"));
    }
    let vbp = dir.write("Big.vbp", &vbp);

    let s = Selection::load(&vbp, &no_overlay).unwrap();

    let session = s.units[0].session.read().unwrap();
    assert_eq!(session.module_count(), 40);
    assert_eq!(session.bind_count(), 40);
}

#[test]
fn a_member_missing_from_disk_is_reported_on_its_line() {
    let dir = TempDir::new("missing");
    let vbp = dir.write("App.vbp", "Type=Exe\nName=\"App\"\nModule=modMain; modMain.bas\nModule=modGone; modGone.bas\n");
    dir.write("modMain.bas", "Option Explicit\n");

    let s = Selection::load(&vbp, &no_overlay).unwrap();

    assert_eq!(
        s.problems,
        vec![Problem {
            file: vbp.clone(),
            line: 4,
            severity: Severity::Error,
            message: format!("File not found: {}", dir.path("modGone.bas").display()),
        }]
    );
    let u = &s.units[0];
    assert_eq!(
        u.members,
        vec![
            member("modMain", "bas", dir.path("modMain.bas"), 3, true),
            member("modGone", "bas", dir.path("modGone.bas"), 4, false),
        ]
    );
    assert_eq!(u.session.read().unwrap().module_count(), 1);
}

#[test]
fn unsaved_editor_text_replaces_the_file_on_disk() {
    let dir = TempDir::new("overlay");
    let vbp = dir.write("App.vbp", "Type=Exe\nName=\"App\"\nModule=modMain; modMain.bas\n");
    let main = dir.write("modMain.bas", "Option Explicit\n");
    let edited = "Option Explicit\nSub Run()\n    ghost\nEnd Sub\n";
    let lookup = vb6_lsp::project::paths::lookup_key(&main);

    let s = Selection::load(&vbp, &|key: &str| (key == lookup).then(|| edited.as_bytes().to_vec())).unwrap();

    let u = &s.units[0];
    let session = u.session.read().unwrap();
    let m = session.module_of(&u.members[0].key).unwrap();
    assert_eq!(session.diagnostics(m).len(), 1);
}

#[test]
fn a_group_lists_its_projects_startup_first_and_wires_project_references() {
    let dir = TempDir::new("group");
    let vbg = dir.write(
        "All.vbg",
        "VBGROUP 5.0\nProject=Lib\\Lib.vbp\nStartupProject=App\\App.vbp\n",
    );
    dir.write(
        "Lib/Lib.vbp",
        "Type=OleDll\nName=\"Lib\"\nClass=clsThing; clsThing.cls\n",
    );
    dir.write("Lib/clsThing.cls", EXPOSED_CLASS);
    dir.write(
        "App/App.vbp",
        "Type=Exe\nReference=*\\A..\\Lib\\Lib.vbp\nModule=modMain; modMain.bas\nName=\"App\"\n",
    );
    let main = "Option Explicit\nSub Run()\n    Dim t As Lib.clsThing\n    t.Hello\nEnd Sub\n";
    dir.write("App/modMain.bas", main);

    let s = Selection::load(&vbg, &no_overlay).unwrap();

    assert!(s.is_group);
    assert_eq!(s.name, "All");
    assert_eq!(s.problems, Vec::<Problem>::new());
    let names: Vec<&str> = s.units.iter().map(|u| u.name.as_str()).collect();
    assert_eq!(names, vec!["App", "Lib"]);
    assert!(s.units.iter().all(|u| u.listed));

    let (app, lib) = (&s.units[0], &s.units[1]);
    assert_eq!(app.dependencies, vec![1]);
    assert_eq!(lib.dependencies, Vec::<usize>::new());
    assert_eq!(
        app.references,
        vec![ReferenceInfo {
            label: "Lib".into(),
            kind: ReferenceKind::Project,
            status: ReferenceStatus::ProjectLoaded,
            line: 2,
            detail: vb6_lsp::project::paths::normalize(&dir.path("Lib/Lib.vbp")).display().to_string(),
        }]
    );
    assert_eq!(lib.exported.name, "Lib");
    assert_eq!(
        lib.exported.types.iter().map(|t| t.name.as_str()).collect::<Vec<_>>(),
        vec!["clsThing"]
    );
    assert_eq!(
        lib.exported.types[0].members.iter().map(|m| m.name.as_str()).collect::<Vec<_>>(),
        vec!["Hello"]
    );

    // The application's call to the library class's method resolves against the library.
    let session = app.session.read().unwrap();
    let m = session.module_of(&app.members[0].key).unwrap();
    assert_eq!(session.diagnostics(m).len(), 0);
    let at = main.find("Hello").unwrap() as u32;
    assert_eq!(
        session.library_member_at(m, at),
        Some(("lib".to_string(), "clsThing".to_string(), "hello".to_string()))
    );
}

#[test]
fn a_project_reached_only_by_reference_is_analysed_but_not_listed() {
    let dir = TempDir::new("hidden");
    dir.write("Lib/Lib.vbp", "Type=OleDll\nName=\"Lib\"\nClass=clsThing; clsThing.cls\n");
    dir.write("Lib/clsThing.cls", EXPOSED_CLASS);
    let vbp = dir.write(
        "App/App.vbp",
        "Type=Exe\nReference=*\\A..\\Lib\\Lib.vbp\nModule=modMain; modMain.bas\nName=\"App\"\n",
    );
    dir.write("App/modMain.bas", "Option Explicit\nSub Run()\n    Dim t As Lib.clsThing\n    t.Hello\nEnd Sub\n");

    let s = Selection::load(&vbp, &no_overlay).unwrap();

    let listed: Vec<(&str, bool)> = s.units.iter().map(|u| (u.name.as_str(), u.listed)).collect();
    assert_eq!(listed, vec![("App", true), ("Lib", false)]);
    assert_eq!(s.units[0].dependencies, vec![1]);
    assert_eq!(
        s.problems,
        vec![Problem {
            file: vbp.clone(),
            line: 2,
            severity: Severity::Warning,
            message: format!(
                "Can't find project or library: {} is not part of the opened project group",
                vb6_lsp::project::paths::normalize(&dir.path("App/../Lib/Lib.vbp")).display()
            ),
        }]
    );
}

#[test]
fn a_reference_to_a_project_that_is_not_there_is_reported_on_its_line() {
    let dir = TempDir::new("noproject");
    let vbp = dir.write(
        "App.vbp",
        "Type=Exe\nName=\"App\"\nReference=*\\A..\\Nowhere\\Nowhere.vbp\nModule=modMain; modMain.bas\n",
    );
    dir.write("modMain.bas", "Option Explicit\n");

    let s = Selection::load(&vbp, &no_overlay).unwrap();

    let missing = vb6_lsp::project::paths::normalize(&dir.path("..\\Nowhere\\Nowhere.vbp"));
    assert_eq!(
        s.problems,
        vec![Problem {
            file: vbp.clone(),
            line: 3,
            severity: Severity::Error,
            message: format!("Can't find project or library: {}", missing.display()),
        }]
    );
    assert_eq!(s.units[0].references[0].status, ReferenceStatus::ProjectMissing);
    assert_eq!(s.units[0].dependencies, Vec::<usize>::new());
}

#[test]
fn projects_that_reference_each_other_are_reported_where_the_cycle_closes() {
    let dir = TempDir::new("cycle");
    let vbg = dir.write("All.vbg", "VBGROUP 5.0\nProject=A.vbp\nProject=B.vbp\n");
    dir.write("A.vbp", "Type=OleDll\nName=\"A\"\nReference=*\\AB.vbp\n");
    let b = dir.write("B.vbp", "Type=OleDll\nName=\"B\"\nName2=x\nReference=*\\AA.vbp\n");

    let s = Selection::load(&vbg, &no_overlay).unwrap();

    assert_eq!(
        s.problems,
        vec![Problem {
            file: b,
            line: 4,
            severity: Severity::Error,
            message: "Cyclic reference of projects not allowed".to_string(),
        }]
    );
    let a = &s.units[0];
    let b = &s.units[1];
    assert_eq!(a.dependencies, vec![1]);
    assert_eq!(b.dependencies, Vec::<usize>::new());
    assert_eq!(b.references[0].status, ReferenceStatus::ProjectCircular);
}

#[test]
fn a_standard_exe_exposes_nothing_to_a_project_that_references_it() {
    let dir = TempDir::new("exeref");
    let vbg = dir.write("All.vbg", "VBGROUP 5.0\nProject=Main.vbp\nProject=Tool.vbp\n");
    let main = dir.write("Main.vbp", "Type=Exe\nName=\"Main\"\nReference=*\\ATool.vbp\n");
    let tool = dir.write("Tool.vbp", "Type=Exe\nName=\"Tool\"\n");

    let s = Selection::load(&vbg, &no_overlay).unwrap();

    assert_eq!(
        s.problems,
        vec![Problem {
            file: main,
            line: 3,
            severity: Severity::Error,
            message: format!("Project '{}' can not be referenced because it's project type is EXE.", tool.display()),
        }]
    );
    assert_eq!(s.units[0].references[0].status, ReferenceStatus::ProjectNotExposed);
    assert_eq!(s.units[0].dependencies, Vec::<usize>::new());
}

#[test]
fn type_libraries_the_engine_does_not_carry_are_reported_not_loaded() {
    let dir = TempDir::new("typelibs");
    let vbp = dir.write(
        "App.vbp",
        "Type=Exe\n\
Reference=*\\G{00020430-0000-0000-C000-000000000046}#2.0#0#..\\Windows\\stdole2.tlb#OLE Automation\n\
Reference=*\\G{DEADBEEF-0000-4000-8000-000000000001}#2.5#0#C:\\msado25.tlb#Microsoft ActiveX Data Objects 2.5 Library\n\
Object={DEADBEEF-0000-4000-8000-000000000002}#2.0#0; Missing.OCX\n\
Name=\"App\"\n",
    );

    let s = Selection::load(&vbp, &no_overlay).unwrap();

    assert_eq!(
        s.problems,
        vec![
            Problem {
                file: vbp.clone(),
                line: 3,
                severity: Severity::Warning,
                message: common::missing_library_message(
                    "Microsoft ActiveX Data Objects 2.5 Library {DEADBEEF-0000-4000-8000-000000000001} 2.5",
                    r"C:\msado25.tlb",
                ),
            },
            Problem {
                file: vbp.clone(),
                line: 4,
                severity: Severity::Warning,
                message: common::missing_library_message(
                    "Missing.OCX {DEADBEEF-0000-4000-8000-000000000002} 2.0",
                    &format!(
                        "{}, {}, {}",
                        dir.path("Missing.OCX").display(),
                        std::path::Path::new(&std::env::var("SystemRoot").unwrap_or_default()).join("SysWOW64").join("Missing.OCX").display(),
                        std::path::Path::new(&std::env::var("SystemRoot").unwrap_or_default()).join("System32").join("Missing.OCX").display(),
                    ),
                ),
            },
        ]
    );
    let statuses: Vec<(ReferenceKind, ReferenceStatus)> =
        s.units[0].references.iter().map(|r| (r.kind, r.status)).collect();
    assert_eq!(
        statuses,
        vec![
            (ReferenceKind::TypeLibrary, ReferenceStatus::Loaded),
            (ReferenceKind::TypeLibrary, ReferenceStatus::NotLoaded),
            (ReferenceKind::Control, ReferenceStatus::NotLoaded),
        ]
    );
}

#[test]
fn the_startup_project_is_a_member_of_the_group_listed_on_its_own_line() {
    let dir = TempDir::new("startupmember");
    let vbg = dir.write("All.vbg", "VBGROUP 5.0\nProject=Lib.vbp\nStartupProject=App.vbp\n");
    dir.write("Lib.vbp", "Type=OleDll\nName=\"Lib\"\n");
    dir.write("App.vbp", "Type=Exe\nName=\"App\"\n");

    let s = Selection::load(&vbg, &no_overlay).unwrap();

    assert_eq!(s.problems, Vec::<Problem>::new());
    let names: Vec<(&str, bool)> = s.units.iter().map(|u| (u.name.as_str(), u.listed)).collect();
    assert_eq!(names, vec![("App", true), ("Lib", true)]);
}

#[test]
fn a_project_listed_twice_in_a_group_is_refused_on_the_second_line_and_loaded_once() {
    let dir = TempDir::new("twice");
    let vbg = dir.write("All.vbg", "VBGROUP 5.0\nStartupProject=App.vbp\nProject=Lib.vbp\nProject=App.vbp\n");
    let app = dir.write("App.vbp", "Type=Exe\nName=\"App\"\n");
    dir.write("Lib.vbp", "Type=OleDll\nName=\"Lib\"\n");

    let s = Selection::load(&vbg, &no_overlay).unwrap();

    assert_eq!(
        s.problems,
        vec![Problem {
            file: vbg.clone(),
            line: 4,
            severity: Severity::Error,
            message: format!("The file '{}' is already open.", app.display()),
        }]
    );
    let names: Vec<&str> = s.units.iter().map(|u| u.name.as_str()).collect();
    assert_eq!(names, vec!["App", "Lib"]);
}

#[test]
fn a_startup_project_that_is_not_there_is_reported_on_its_line() {
    let dir = TempDir::new("badstartup");
    let vbg = dir.write("All.vbg", "VBGROUP 5.0\nStartupProject=Other.vbp\nProject=A.vbp\n");
    dir.write("A.vbp", "Type=Exe\nName=\"A\"\n");

    let s = Selection::load(&vbg, &no_overlay).unwrap();

    assert_eq!(s.problems.len(), 1);
    assert_eq!(s.problems[0].file, vbg);
    assert_eq!(s.problems[0].line, 2);
    assert_eq!(s.problems[0].severity, Severity::Error);
    assert!(
        s.problems[0].message.starts_with(&format!("Cannot read project {}: ", dir.path("Other.vbp").display())),
        "{}",
        s.problems[0].message
    );
    let names: Vec<&str> = s.units.iter().map(|u| u.name.as_str()).collect();
    assert_eq!(names, vec!["A"]);
}

#[test]
fn a_project_that_references_itself_is_cyclic() {
    let dir = TempDir::new("selfref");
    let vbp = dir.write("Own.vbp", "Type=OleDll\nName=\"Own\"\nReference=*\\AOwn.vbp\n");

    let s = Selection::load(&vbp, &no_overlay).unwrap();

    assert_eq!(
        s.problems,
        vec![Problem {
            file: vbp,
            line: 3,
            severity: Severity::Error,
            message: "Cyclic reference of projects not allowed".to_string(),
        }]
    );
    assert_eq!(s.units[0].references[0].status, ReferenceStatus::ProjectCircular);
}

#[test]
fn a_project_of_a_group_that_cannot_be_read_is_reported_on_the_group_line() {
    let dir = TempDir::new("badmember");
    let vbg = dir.write("All.vbg", "VBGROUP 5.0\nProject=Gone.vbp\nProject=A.vbp\n");
    dir.write("A.vbp", "Type=Exe\nName=\"A\"\n");

    let s = Selection::load(&vbg, &no_overlay).unwrap();

    assert_eq!(s.problems.len(), 1);
    let p = &s.problems[0];
    assert_eq!((p.file.clone(), p.line, p.severity), (vbg.clone(), 2, Severity::Error));
    assert!(p.message.starts_with("Cannot read project "), "{}", p.message);
    assert_eq!(s.units.iter().map(|u| u.name.as_str()).collect::<Vec<_>>(), vec!["A"]);
}

#[test]
fn opening_something_that_is_not_a_project_fails() {
    let dir = TempDir::new("notproject");
    let bas = dir.write("a.bas", "Option Explicit\n");

    let err = match Selection::load(&bas, &no_overlay) {
        Err(e) => e,
        Ok(_) => panic!("a .bas is not a project"),
    };

    assert_eq!(
        err,
        LoadError { file: bas, line: None, message: "Not a VB6 project (.vbp) or project group (.vbg)".to_string() }
    );
}

#[test]
fn a_project_file_that_does_not_exist_fails_to_open() {
    let dir = TempDir::new("nofile");
    let vbp = dir.path("Absent.vbp");

    let err = match Selection::load(&vbp, &no_overlay) {
        Err(e) => e,
        Ok(_) => panic!("there is no such project"),
    };

    assert_eq!(err.file, vbp);
    assert_eq!(err.line, None);
    assert!(err.message.starts_with("VBP parse error") || err.message.starts_with("Failed to read file"), "{}", err.message);
}

#[test]
fn a_group_file_without_its_header_fails_to_open() {
    let dir = TempDir::new("nohead");
    let vbg = dir.write("All.vbg", "Project=A.vbp\n");

    let err = match Selection::load(&vbg, &no_overlay) {
        Err(e) => e,
        Ok(_) => panic!("a group needs its header"),
    };

    assert_eq!(
        err,
        LoadError {
            file: vbg,
            line: Some(1),
            message: "the first line must be the `VBGROUP` header".to_string()
        }
    );
}

#[test]
fn editing_a_class_interface_refreshes_the_projects_that_reference_it() {
    let dir = TempDir::new("propagate");
    let vbg = dir.write("All.vbg", "VBGROUP 5.0\nProject=App\\App.vbp\nProject=Lib\\Lib.vbp\n");
    dir.write("Lib/Lib.vbp", "Type=OleDll\nName=\"Lib\"\nClass=clsThing; clsThing.cls\n");
    dir.write("Lib/clsThing.cls", EXPOSED_CLASS);
    dir.write(
        "App/App.vbp",
        "Type=Exe\nReference=*\\A..\\Lib\\Lib.vbp\nModule=modMain; modMain.bas\nName=\"App\"\n",
    );
    let main = "Option Explicit\nSub Run()\n    Dim t As Lib.clsThing\n    t.Hello\nEnd Sub\n";
    dir.write("App/modMain.bas", main);
    let mut s = Selection::load(&vbg, &no_overlay).unwrap();

    // Hello is no longer public in the library.
    let lib_key = s.units[1].members[0].key.clone();
    s.units[1]
        .session
        .write()
        .unwrap()
        .update_file(&lib_key, EXPOSED_CLASS.replace("Public Sub Hello", "Private Sub Hello").into_bytes());
    let reanalysed = s.propagate_exports(1);

    assert_eq!(reanalysed, vec![0]);
    assert_eq!(s.units[1].exported.types[0].members.len(), 0);
    let app = &s.units[0];
    let session = app.session.read().unwrap();
    let m = session.module_of(&app.members[0].key).unwrap();
    let at = main.find("Hello").unwrap() as u32;
    assert_eq!(session.library_member_at(m, at), None);

    // An edit that leaves the exported interface alone reanalyses nothing.
    drop(session);
    s.units[1]
        .session
        .write()
        .unwrap()
        .update_file(&lib_key, EXPOSED_CLASS.replace("Public Sub Hello", "Private Sub Hello\n    Dim x\n    x = 1\n    Rem").into_bytes());
    assert_eq!(s.propagate_exports(1), Vec::<usize>::new());
}
