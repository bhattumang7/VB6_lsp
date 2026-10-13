//! Coverage invariants: every name is bound, and every member access either
//! resolves or is reported with a reason - none is silently dropped.

mod common;

use std::fs;
use std::path::Path;

use common::{coded, session, session_with};
use vb6_core::sema::MemberStatus;
use vb6_core::session::{Library, SessionConfig};

const CLS_TRIP: &str = "Attribute VB_Name = \"clsTrip\"\n\
Public Name As String\n\
Public Sub Save()\nEnd Sub\n\
Public Property Get Item(ByVal i As Long) As clsTrip\nAttribute Item.VB_UserMemId = 0\nEnd Property\n\
Public Property Let Title(ByVal v As String)\nEnd Property\n\
Public Property Get Title() As String\nEnd Property\n\
Public Property Set Parent(ByVal o As clsTrip)\nEnd Property\n\
Public Property Get Parent() As clsTrip\nEnd Property\n\
Public Function All() As clsTrip()\nEnd Function\n\
Public Sub Inside()\n    Me.Save\n    Me.Name = \"n\"\n    With Me\n        .Save\n    End With\nEnd Sub\n";

const MOD_TYPES: &str = "Attribute VB_Name = \"modTypes\"\n\
Public Enum Color\n    Red\n    Green\nEnd Enum\n\
Public Type TPoint\n    X As Long\n    Tags(3) As String\nEnd Type\n\
Public Type TLine\n    A As TPoint\n    B As TPoint\nEnd Type\n\
Public gLine As TLine\n\
Public Function MakeTrip() As clsTrip\nEnd Function\n";

const MOD_MAIN: &str = "Attribute VB_Name = \"modMain\"\nOption Explicit\n\
Private mTrips(3) As clsTrip\n\
Sub Everything(ByVal pTrip As clsTrip, vVar As Variant, oObj As Object)\n\
    Dim t As clsTrip\n    Dim n As Long\n    Dim s As String\n    Dim l As TLine\n    Dim aPts(2) As TPoint\n    Dim u As clsMissing\n\
    Set t = New clsTrip\n\
    t.Save\n    t.Name = \"x\"\n    s = t.Name\n    t.Title = \"t\"\n    s = t.Title\n\
    Set t.Parent = pTrip\n    Set t = t.Parent\n    t.Parent.Parent.Save\n\
    t(1).Save\n    t.Item(2).Name = \"q\"\n    t.All(1).Save\n    t.All()(0).Save\n\
    mTrips(1).Save\n    mTrips(n).Parent.Save\n\
    MakeTrip().Save\n    MakeTrip.Name = \"m\"\n    modTypes.MakeTrip.Save\n\
    l.A.X = 1\n    l.B.Tags(1) = \"a\"\n    aPts(1).X = 2\n    gLine.A.X = 3\n    n = Green + modTypes.Red + Color.Green\n\
    ReDim aPts(5)\n    ReDim l.A.Tags(5)\n\
    With t\n        .Save\n        With .Parent\n            .Name = \"p\"\n        End With\n        .Item(1).Save\n    End With\n\
    With New clsTrip\n        .Save\n    End With\n\
    (New clsTrip).Save\n\
    vVar.Anything\n    oObj.Anything\n    vVar!Key = 1\n    u.Foo\n    n.Foo\n    t.NoSuch\n    ghost.Thing\n\
End Sub\n";

fn assert_invariants(s: &vb6_core::session::Session, label: &str) {
    for m in 0..s.module_count() {
        assert_eq!(s.unaccounted_member_accesses(m), vec![], "{label}: member access dropped");
        assert_eq!(s.unbound_namerefs(m), vec![], "{label}: name reference not bound");
    }
}

#[test]
fn every_construct_is_accounted_for() {
    let s = session(&[("clsTrip.cls", CLS_TRIP), ("modTypes.bas", MOD_TYPES), ("modMain.bas", MOD_MAIN)]);
    assert_invariants(&s, "constructed");
    // Spot-check a few outcomes among the members.
    let statuses: Vec<MemberStatus> = s.member_accesses(2).into_iter().map(|(_, st)| st).collect();
    assert!(statuses.contains(&MemberStatus::Resolved));
    assert!(statuses.contains(&MemberStatus::LateBound));
    assert!(statuses.contains(&MemberStatus::Bang));
    assert!(statuses.contains(&MemberStatus::NotAnObject));
    assert!(statuses.contains(&MemberStatus::NoSuchMember));
    assert!(statuses.contains(&MemberStatus::BaseUnresolved));
    assert!(statuses.contains(&MemberStatus::UnknownType("clsmissing".to_string())));
    // Exactly one member of a known class is missing, and it is the only 461.
    assert_eq!(coded(&s, MOD_MAIN, 2, 461).len(), 1);
}

#[test]
fn summary_counts_every_access_once() {
    let s = session(&[("clsTrip.cls", CLS_TRIP), ("modTypes.bas", MOD_TYPES), ("modMain.bas", MOD_MAIN)]);
    let total: usize = (0..s.module_count()).map(|m| s.member_accesses(m).len()).sum();
    let sum = s.member_access_summary();
    let counted = sum.resolved
        + sum.late_bound
        + sum.bang
        + sum.not_an_object
        + sum.unknown_type
        + sum.library_not_loaded
        + sum.no_such_member
        + sum.missing_accessor
        + sum.base_unresolved;
    assert_eq!(counted, total);
    assert!(sum.resolved > 20);
}

#[test]
fn interface_and_library_modules_are_accounted_for() {
    let libs = Library::from_json(r#"[{"name":"VB","types":[
        {"name":"TextBox","kind":"class","members":[{"name":"Text","kind":"field","return_type":"String"}]},
        {"name":"Form","kind":"class","members":[{"name":"Caption","kind":"field","return_type":"String"}]}]}]"#)
        .unwrap();
    let frm = "VERSION 5.00\nBegin VB.Form frmA\n   Begin VB.TextBox txtA\n   End\nEnd\nAttribute VB_Name = \"frmA\"\n\
               Option Explicit\nPrivate Sub Form_Load()\n    txtA.Text = \"a\"\n    Me.Caption = \"c\"\n    Caption = \"d\"\nEnd Sub\n";
    let s = session_with(
        &[("frmA.frm", frm)],
        SessionConfig { project_name: Some("P".into()), libraries: libs },
    );
    assert_invariants(&s, "designer");
}

fn fixture_dirs() -> Vec<std::path::PathBuf> {
    let root = Path::new(env!("CARGO_MANIFEST_DIR")).join("../vb6-codegen/tests/fixtures");
    let mut dirs: Vec<_> = fs::read_dir(root)
        .expect("fixtures directory")
        .filter_map(Result::ok)
        .map(|e| e.path())
        .filter(|p| p.is_dir())
        .collect();
    dirs.sort();
    dirs
}

/// Each fixture is a complete, valid program, so no member access in it may be
/// reported as missing, and none may be dropped.
#[test]
fn valid_programs_have_no_missing_members() {
    let mut programs = 0usize;
    let mut resolved = 0usize;
    for dir in fixture_dirs() {
        let mut files: Vec<(String, Vec<u8>)> = Vec::new();
        for name in ["input.bas", "input.cls", "input2.cls"] {
            let p = dir.join(name);
            if let Ok(bytes) = fs::read(&p) {
                files.push((format!("{}/{name}", dir.file_name().unwrap().to_string_lossy()), bytes));
            }
        }
        if files.is_empty() {
            continue;
        }
        programs += 1;
        let s = vb6_core::session::Session::from_sources(files);
        let label = dir.file_name().unwrap().to_string_lossy().into_owned();
        for m in 0..s.module_count() {
            assert_eq!(s.unaccounted_member_accesses(m), vec![], "{label}: member access dropped");
            assert_eq!(s.unbound_namerefs(m), vec![], "{label}: name reference not bound");
            for d in s.diagnostics(m) {
                assert_ne!(d.code, 461, "{label}: member of a known class reported missing");
            }
            for (span, status) in s.member_accesses(m) {
                assert_ne!(status, MemberStatus::NoSuchMember, "{label}: at {}", span.start);
                if status == MemberStatus::Resolved {
                    resolved += 1;
                }
            }
        }
    }
    assert!(programs > 300, "fixtures were found: {programs}");
    assert!(resolved > 50, "class members resolved across the fixtures: {resolved}");
}
