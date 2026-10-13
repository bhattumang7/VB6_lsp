//! Editing a file keeps member resolution correct: an edit that leaves a
//! module's interface alone binds only that module; an interface edit binds
//! the project again, and dependents see the change.

mod common;

use common::{at, coded};
use vb6_core::sema::MemberStatus;
use vb6_core::session::{DeclId, Library, Session, SessionConfig};

const CLS: &str = "Public Sub Save()\nEnd Sub\n";
const MAIN: &str = "Option Explicit\nSub Run()\n    Dim o As clsTrip\n    o.Save\n    o.Load\n    ghost\nEnd Sub\n";

fn new_session() -> Session {
    Session::from_sources(vec![
        ("clsTrip.cls".into(), CLS.as_bytes().to_vec()),
        ("modMain.bas".into(), MAIN.as_bytes().to_vec()),
    ])
}

#[test]
fn edit_inside_a_body_binds_only_that_module() {
    let mut s = new_session();
    let before = s.bind_count();
    let diags_before = s.diagnostics(1).len();

    let edited = "Public Sub Save()\n    Dim x As Long\n    x = 1\nEnd Sub\n";
    s.update_file("clsTrip.cls", edited.as_bytes().to_vec());

    assert_eq!(s.bind_count() - before, 1);
    // The dependent module is as it was: same resolution, same diagnostics.
    assert_eq!(s.declaration_at(1, at(MAIN, "o.Save", "Save")), Some(DeclId::Proc { module: 0, idx: 0 }));
    assert_eq!(s.diagnostics(1).len(), diags_before);
    // The class's own locals resolve in the new text.
    assert_eq!(
        s.declaration_at(0, at(edited, "x = 1", "x")),
        Some(DeclId::Local { module: 0, proc: 0, idx: 0 })
    );
}

#[test]
fn repeated_body_edits_do_not_duplicate_diagnostics() {
    let mut s = new_session();
    let main_diags = s.diagnostics(1);
    for round in 0..4 {
        let src = format!("Option Explicit\nSub Run()\n    Dim o As clsTrip\n    o.Save\n    o.Load\n    ghost\n    ' edit {round}\nEnd Sub\n");
        s.update_file("modMain.bas", src.as_bytes().to_vec());
        let d = s.diagnostics(1);
        assert_eq!(d.len(), main_diags.len(), "round {round}");
        assert_eq!(coded(&s, &src, 1, 461).len(), 1, "round {round}");
    }
}

#[test]
fn adding_a_public_member_resolves_a_reported_use() {
    let mut s = new_session();
    assert_eq!(s.member_status_at(1, at(MAIN, "o.Load", "Load")), Some(MemberStatus::NoSuchMember));
    assert_eq!(coded(&s, MAIN, 1, 461).len(), 1);
    let before = s.bind_count();

    let edited = "Public Sub Save()\nEnd Sub\nPublic Sub Load()\nEnd Sub\n";
    s.update_file("clsTrip.cls", edited.as_bytes().to_vec());

    assert_eq!(s.bind_count() - before, s.module_count());
    assert_eq!(s.declaration_at(1, at(MAIN, "o.Load", "Load")), Some(DeclId::Proc { module: 0, idx: 1 }));
    assert_eq!(coded(&s, MAIN, 1, 461), vec![]);
}

#[test]
fn removing_a_public_member_reports_its_uses() {
    let mut s = new_session();
    s.update_file("clsTrip.cls", b"Public Sub Save()\nEnd Sub\nPublic Sub Load()\nEnd Sub\n".to_vec());
    assert_eq!(coded(&s, MAIN, 1, 461), vec![]);
    s.update_file("clsTrip.cls", b"Public Sub Save()\nEnd Sub\n".to_vec());
    assert_eq!(s.declaration_at(1, at(MAIN, "o.Load", "Load")), None);
    assert_eq!(coded(&s, MAIN, 1, 461).len(), 1);
}

#[test]
fn changing_visibility_is_an_interface_edit() {
    let mut s = new_session();
    let before = s.bind_count();
    s.update_file("clsTrip.cls", b"Private Sub Save()\nEnd Sub\n".to_vec());
    assert_eq!(s.bind_count() - before, s.module_count());
    assert_eq!(s.declaration_at(1, at(MAIN, "o.Save", "Save")), None);
    assert_eq!(coded(&s, MAIN, 1, 461).len(), 2);
}

#[test]
fn changing_a_declared_type_is_an_interface_edit() {
    let cls = "Public Owner As clsPerson\n";
    let person = "Public Sub Greet()\nEnd Sub\n";
    let other = "Public Sub Greet()\nEnd Sub\nPublic Sub Wave()\nEnd Sub\n";
    let main = "Sub Run()\n    Dim o As clsTrip\n    o.Owner.Wave\nEnd Sub\n";
    let mut s = Session::from_sources(vec![
        ("clsTrip.cls".into(), cls.as_bytes().to_vec()),
        ("clsPerson.cls".into(), person.as_bytes().to_vec()),
        ("clsOther.cls".into(), other.as_bytes().to_vec()),
        ("modMain.bas".into(), main.as_bytes().to_vec()),
    ]);
    assert_eq!(s.member_status_at(3, at(main, "o.Owner.Wave", "Wave")), Some(MemberStatus::NoSuchMember));
    s.update_file("clsTrip.cls", b"Public Owner As clsOther\n".to_vec());
    assert_eq!(s.declaration_at(3, at(main, "o.Owner.Wave", "Wave")), Some(DeclId::Proc { module: 2, idx: 1 }));
}

#[test]
fn adding_a_file_binds_the_project_again() {
    let mut s = new_session();
    let before = s.bind_count();
    s.update_file("modExtra.bas", b"Public Sub Extra()\nEnd Sub\n".to_vec());
    assert_eq!(s.bind_count() - before, 3);
}

#[test]
fn set_config_makes_library_types_resolve() {
    let main = "Sub Run()\n    Dim rs As ADODB.Recordset\n    rs.MoveNext\nEnd Sub\n";
    let mut s = Session::from_sources(vec![("modMain.bas".into(), main.as_bytes().to_vec())]);
    let site = at(main, "rs.MoveNext", "MoveNext");
    assert_eq!(s.member_status_at(0, site), Some(MemberStatus::LibraryNotLoaded("adodb".to_string())));
    let libs = Library::from_json(
        r#"[{"name":"ADODB","types":[{"name":"Recordset","kind":"class","members":[{"name":"MoveNext","kind":"method"}]}]}]"#,
    )
    .unwrap();
    s.set_config(SessionConfig { project_name: None, libraries: libs });
    assert_eq!(s.member_status_at(0, site), Some(MemberStatus::Resolved));
    assert_eq!(s.external_member_references("ADODB", "Recordset", "MoveNext").len(), 1);
}
