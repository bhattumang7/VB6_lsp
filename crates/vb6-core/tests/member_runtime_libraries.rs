//! The runtime libraries the engine bundles: their functions, constants,
//! `Err` and `Collection` resolve without the host describing anything.

mod common;

use common::{at, session, session_with};
use vb6_core::sema::MemberStatus;
use vb6_core::session::{Library, SessionConfig, SemTokenKind};

fn uses(s: &vb6_core::session::Session, member: &str) -> Vec<u32> {
    s.external_member_references("VBA", "Strings", member).iter().map(|l| l.span.start).collect()
}

#[test]
fn string_functions_resolve_to_the_library_and_the_dollar_form_to_its_string_variant() {
    let src = "Option Explicit\nSub Run()\n    Dim s As String\n    s = Left(s, 1)\n    s = Left$(s, 1)\nEnd Sub\n";
    let s = session(&[("modMain.bas", src)]);
    assert_eq!(s.diagnostics(0), vec![]);
    assert_eq!(uses(&s, "_B_var_Left"), vec![at(src, "= Left(s", "Left")]);
    assert_eq!(uses(&s, "_B_str_Left"), vec![at(src, "Left$(s", "Left")]);
}

#[test]
fn constants_are_known_names() {
    let src = "Option Explicit\nSub Run()\n    Dim s As String\n    s = \"a\" & vbCrLf & vbTab\n    n = vbOKOnly\nEnd Sub\n";
    let s = session(&[("modMain.bas", src)]);
    // Only the undeclared `n` is reported.
    let d = s.diagnostics(0);
    assert_eq!(d.len(), 1);
    assert_eq!(d[0].span.start, at(src, "n = vbOKOnly", "n"));
    let crlf = s.external_member_references("VBA", "Constants", "vbCrLf");
    assert_eq!(crlf.len(), 1);
    assert_eq!(crlf[0].span.start, at(src, "vbCrLf", "vbCrLf"));
}

#[test]
fn err_object_members_resolve() {
    let src = "Sub Run()\n    Dim n As Long\n    n = Err.Number\n    Err.Raise 5\n    s = Err.Description\nEnd Sub\n";
    let s = session(&[("modMain.bas", src)]);
    for (ctx, member) in [("Err.Number", "Number"), ("Err.Raise", "Raise"), ("Err.Description", "Description")] {
        assert_eq!(s.member_status_at(0, at(src, ctx, member)), Some(MemberStatus::Resolved), "{ctx}");
    }
    assert_eq!(s.external_member_references("VBA", "ErrObject", "Number").len(), 1);
}

#[test]
fn collection_members_and_default_member() {
    let src = "Sub Run()\n    Dim c As New Collection\n    c.Add 1\n    n = c.Count\n    v = c(1)\n    c.Item(1)\nEnd Sub\n";
    let s = session(&[("modMain.bas", src)]);
    for (ctx, member) in [("c.Add", "Add"), ("c.Count", "Count"), ("c.Item(1)", "Item")] {
        assert_eq!(s.member_status_at(0, at(src, ctx, member)), Some(MemberStatus::Resolved), "{ctx}");
    }
    assert_eq!(s.external_member_references("VBA", "Collection", "Item").len(), 1);
}

#[test]
fn project_and_local_names_win_over_the_library() {
    let src = "Public Function Left(s As String) As String\nEnd Function\nSub Run()\n    Dim Mid As Long\n    Dim t As String\n    t = Left(t)\n    Mid = 3\nEnd Sub\n";
    let s = session(&[("modMain.bas", src)]);
    assert_eq!(s.external_member_references("VBA", "Strings", "_B_var_Left"), vec![]);
    assert_eq!(s.external_member_references("VBA", "Strings", "_B_var_Mid"), vec![]);
}

#[test]
fn runtime_object_library_types_resolve() {
    let src = "Sub Run(pb As PropertyBag)\n    pb.WriteProperty \"a\", 1\n    x = pb.ReadProperty(\"a\")\nEnd Sub\n";
    let s = session(&[("clsA.cls", src)]);
    assert_eq!(s.member_status_at(0, at(src, "pb.WriteProperty", "WriteProperty")), Some(MemberStatus::Resolved));
    assert_eq!(s.member_status_at(0, at(src, "pb.ReadProperty", "ReadProperty")), Some(MemberStatus::Resolved));
}

#[test]
fn library_member_tokens_follow_the_member_kind() {
    let src = "Sub Run()\n    Dim s As String\n    s = Left(s, 1) & vbCrLf\nEnd Sub\n";
    let s = session(&[("modMain.bas", src)]);
    let toks = s.semantic_tokens(0);
    let kind_at = |off: u32| toks.iter().find(|t| t.span.start == off).map(|t| t.kind);
    assert_eq!(kind_at(at(src, "Left(s", "Left")), Some(SemTokenKind::Function));
    assert_eq!(kind_at(at(src, "vbCrLf", "vbCrLf")), Some(SemTokenKind::EnumMember));
}

#[test]
fn no_libraries_means_the_names_are_unknown() {
    let src = "Option Explicit\nSub Run()\n    Dim s As String\n    s = \"a\" & vbCrLf\nEnd Sub\n";
    let none = SessionConfig { project_name: None, libraries: Vec::new() };
    let s = session_with(&[("modMain.bas", src)], none);
    assert_eq!(s.diagnostics(0).len(), 1);
    assert_eq!(Library::bundled().len(), 2);
}

#[test]
fn library_name_qualifies_its_functions_and_enums() {
    let src = "Option Explicit\nSub Run()\n    Dim b As Boolean\n    Dim n As Long\n    b = VBA.IsError(n)\n\
               \x20   n = VBRUN.CheckBoxConstants.vbChecked\n    n = CheckBoxConstants.vbUnchecked\n    n = VBA.vbCrLf\nEnd Sub\n";
    let s = session(&[("modMain.bas", src)]);
    assert_eq!(s.diagnostics(0), vec![]);
    for (ctx, member) in [
        ("VBA.IsError", "IsError"),
        ("VBRUN.CheckBoxConstants.vbChecked", "vbChecked"),
        ("CheckBoxConstants.vbUnchecked", "vbUnchecked"),
        ("VBA.vbCrLf", "vbCrLf"),
    ] {
        assert_eq!(s.member_status_at(0, at(src, ctx, member)), Some(MemberStatus::Resolved), "{ctx}");
    }
    assert_eq!(s.external_member_references("VBRUN", "CheckBoxConstants", "vbChecked").len(), 1);
}

#[test]
fn unknown_member_of_a_library_is_reported() {
    let src = "Sub Run()\n    x = VBA.NoSuchFunction(1)\n    n = Err.Desc\nEnd Sub\n";
    let s = session(&[("modMain.bas", src)]);
    assert_eq!(s.member_status_at(0, at(src, "VBA.NoSuchFunction", "NoSuchFunction")), Some(MemberStatus::NoSuchMember));
    assert_eq!(s.member_status_at(0, at(src, "Err.Desc", "Desc")), Some(MemberStatus::NoSuchMember));
}
