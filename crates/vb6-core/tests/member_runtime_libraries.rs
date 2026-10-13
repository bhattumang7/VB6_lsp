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
    let names: Vec<String> = Library::bundled().into_iter().map(|l| l.name).collect();
    assert_eq!(names, vec!["VBA", "VB", "VBRUN", "stdole"]);
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

#[test]
fn global_procedures_and_objects_of_the_vb_library_resolve() {
    let src = "Option Explicit\nSub Run()\n    Dim p As Object\n    Dim s As String\n    s = LoadResString(101)\n\
               \x20   Set p = LoadPicture(\"a.bmp\")\n    s = App.Path\n    Screen.MousePointer = 11\n    s = Clipboard.GetText\nEnd Sub\n";
    let s = session(&[("modMain.bas", src)]);
    assert_eq!(s.diagnostics(0), vec![]);
    for (ctx, member) in [
        ("App.Path", "Path"),
        ("Screen.MousePointer", "MousePointer"),
        ("Clipboard.GetText", "GetText"),
    ] {
        assert_eq!(s.member_status_at(0, at(src, ctx, member)), Some(MemberStatus::Resolved), "{ctx}");
    }
    assert_eq!(s.external_member_references("VB", "VBGlobal", "LoadResString").len(), 1);
    assert_eq!(s.external_member_references("VB", "VBGlobal", "LoadPicture").len(), 1);
}

#[test]
fn form_members_are_reached_by_bare_name_in_a_form_and_a_user_control() {
    let frm = "VERSION 5.00\r\nBegin VB.Form frmA\r\nEnd\r\nAttribute VB_Name = \"frmA\"\r\nOption Explicit\r\n\
               Private Sub Go()\r\n    Dim x As Single\r\n    Hide\r\n    x = ScaleX(1, 1, 3)\r\n    PopupMenu mnuA\r\n    Caption = \"c\"\r\nEnd Sub\r\n";
    let ctl = "VERSION 5.00\r\nBegin VB.UserControl ctlA\r\nEnd\r\nAttribute VB_Name = \"ctlA\"\r\nOption Explicit\r\n\
               Private Sub Go()\r\n    Dim w As Single\r\n    PropertyChanged \"Value\"\r\n    w = TextWidth(\"abc\")\r\nEnd Sub\r\n";
    let s = session(&[("frmA.frm", frm), ("ctlA.ctl", ctl)]);
    let lines = |file: usize| -> Vec<u32> { s.diagnostics(file).iter().map(|d| d.span.start).collect() };
    // Only `mnuA` (a menu that is not declared) is undefined in the form.
    assert_eq!(lines(0), vec![at(frm, "PopupMenu mnuA", "mnuA")]);
    assert_eq!(lines(1), Vec::<u32>::new());
    for m in ["Hide", "ScaleX", "PopupMenu", "Caption"] {
        assert_eq!(s.external_member_references("VB", "Form", m).len(), 1, "{m}");
    }
    for m in ["PropertyChanged", "TextWidth"] {
        assert_eq!(s.external_member_references("VB", "UserControl", m).len(), 1, "{m}");
    }
}

#[test]
fn control_types_of_the_vb_library_type_the_controls() {
    let frm = "VERSION 5.00\r\nBegin VB.Form frmA\r\n   Begin VB.TextBox txtA\r\n   End\r\n   Begin VB.ListBox lstA\r\n   End\r\nEnd\r\n\
               Attribute VB_Name = \"frmA\"\r\nOption Explicit\r\nPrivate Sub Go()\r\n    txtA.Text = \"a\"\r\n    lstA.AddItem \"b\"\r\n    n = lstA.ListCount\r\nEnd Sub\r\n";
    let s = session(&[("frmA.frm", frm)]);
    for (ctx, member) in [("txtA.Text", "Text"), ("lstA.AddItem", "AddItem"), ("lstA.ListCount", "ListCount")] {
        assert_eq!(s.member_status_at(0, at(frm, ctx, member)), Some(MemberStatus::Resolved), "{ctx}");
    }
}

#[test]
fn font_and_picture_properties_resolve_through_stdole() {
    let src = "Sub Run(t As TextBox)\n    t.Font.Bold = True\n    t.Font.Name = \"Arial\"\nEnd Sub\n";
    let s = session(&[("clsA.cls", src)]);
    assert_eq!(s.member_status_at(0, at(src, "t.Font.Bold", "Font")), Some(MemberStatus::Resolved));
    assert_eq!(s.member_status_at(0, at(src, "t.Font.Bold", "Bold")), Some(MemberStatus::Resolved));
}

#[test]
fn a_user_control_seen_from_a_form_has_the_container_members() {
    let ctl = "VERSION 5.00\r\nBegin VB.UserControl ctlA\r\nEnd\r\nAttribute VB_Name = \"ctlA\"\r\nPublic Sub Reload()\r\nEnd Sub\r\n";
    let frm = "VERSION 5.00\r\nBegin VB.Form frmA\r\n   Begin Project1.ctlA ctlMain\r\n   End\r\nEnd\r\nAttribute VB_Name = \"frmA\"\r\n\
               Option Explicit\r\nPrivate Sub Go()\r\n    ctlMain.Reload\r\n    ctlMain.Visible = True\r\n    ctlMain.Left = 5\r\n    ctlMain.Move 1, 2\r\n    ctlMain.Nope\r\nEnd Sub\r\n";
    let cfg = SessionConfig { project_name: Some("Project1".into()), ..SessionConfig::default() };
    let s = session_with(&[("ctlA.ctl", ctl), ("frmA.frm", frm)], cfg);
    for (ctx, member) in [
        ("ctlMain.Reload", "Reload"),
        ("ctlMain.Visible", "Visible"),
        ("ctlMain.Left", "Left"),
        ("ctlMain.Move", "Move"),
    ] {
        assert_eq!(s.member_status_at(1, at(frm, ctx, member)), Some(MemberStatus::Resolved), "{ctx}");
    }
    assert_eq!(s.member_status_at(1, at(frm, "ctlMain.Nope", "Nope")), Some(MemberStatus::NoSuchMember));
}

#[test]
fn control_is_a_generic_extensible_type() {
    let src = "Sub Run(oCtrl As Control, oForm As Form)\n    oCtrl.Visible = False\n    oCtrl.Whatever\nEnd Sub\n";
    let s = session(&[("modMain.bas", src)]);
    assert_eq!(s.member_status_at(0, at(src, "oCtrl.Visible", "Visible")), Some(MemberStatus::LateBound));
    assert_eq!(s.member_status_at(0, at(src, "oCtrl.Whatever", "Whatever")), Some(MemberStatus::LateBound));
    assert_eq!(s.diagnostics(0), vec![]);
}

#[test]
fn generic_form_type_looks_unknown_members_up_at_run_time() {
    let src = "Option Explicit\nSub Run(f As Form, c As Collection)\n    f.SetStatusBarText \"x\"\n    f.Caption = \"c\"\n    c.NoSuch\nEnd Sub\n";
    let s = session(&[("modMain.bas", src)]);
    // `Form` is extensible: a public procedure of the real form is found late.
    assert_eq!(s.member_status_at(0, at(src, "f.SetStatusBarText", "SetStatusBarText")), Some(MemberStatus::LateBound));
    assert_eq!(s.member_status_at(0, at(src, "f.Caption", "Caption")), Some(MemberStatus::Resolved));
    // `Collection` is a closed dual interface: an unknown member is an error.
    assert_eq!(s.member_status_at(0, at(src, "c.NoSuch", "NoSuch")), Some(MemberStatus::NoSuchMember));
    assert_eq!(s.diagnostics(0).len(), 1);
}
