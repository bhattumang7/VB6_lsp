//! Controls of a form or UserControl are names of that module, typed by the
//! control's type; control arrays index like arrays; designers extend a
//! library type; event handlers stay ordinary procedures.

mod common;

use common::{at, nth, session, session_with, sites};
use vb6_core::sema::MemberStatus;
use vb6_core::session::{DeclId, Library, SessionConfig};

const FRM: &str = "VERSION 5.00\r\n\
Begin VB.Form frmMain\r\n\
   Caption         =   \"Main\"\r\n\
   Begin VB.TextBox txtDate\r\n\
      Index           =   0\r\n\
   End\r\n\
   Begin VB.TextBox txtDate\r\n\
      Index           =   1\r\n\
   End\r\n\
   Begin VB.CommandButton cmdOK\r\n\
      Caption         =   \"OK\"\r\n\
   End\r\n\
   Begin MSComctlLib.ListView lvwTrips\r\n\
   End\r\n\
End\r\n\
Attribute VB_Name = \"frmMain\"\r\n\
Attribute VB_PredeclaredId = True\r\n\
Option Explicit\r\n\
Public Sub RefreshView()\r\n\
cmdOK_Click\r\n\
End Sub\r\n\
Private Sub cmdOK_Click()\r\n\
    txtDate(0).Text = \"x\"\r\n\
    cmdOK.Caption = \"OK\"\r\n\
    Me.cmdOK.Caption = \"y\"\r\n\
    lvwTrips.ListItems.Clear\r\n\
    Me.Show\r\n\
End Sub\r\n";

const VB_LIB: &str = r#"{
  "name": "VB",
  "types": [
    {"name": "TextBox", "kind": "class", "members": [
      {"name": "Text", "kind": "field", "return_type": "String"}]},
    {"name": "CommandButton", "kind": "class", "members": [
      {"name": "Caption", "kind": "field", "return_type": "String"}]},
    {"name": "Form", "kind": "class", "members": [
      {"name": "Show", "kind": "method", "param_count": 0}]}
  ]
}"#;

fn vb_config() -> SessionConfig {
    SessionConfig { project_name: None, libraries: Library::from_json(VB_LIB).unwrap() }
}

fn no_libraries() -> SessionConfig {
    SessionConfig { project_name: None, libraries: Vec::new() }
}

fn control(idx: usize) -> DeclId {
    DeclId::Control { module: 0, idx }
}

#[test]
fn bare_control_name_resolves_to_the_control() {
    let s = session(&[("frmMain.frm", FRM)]);
    // Controls in declaration order, a control array being one: txtDate, cmdOK, lvwTrips.
    assert_eq!(s.declaration_at(0, at(FRM, "cmdOK.Caption = ", "cmdOK")), Some(control(1)));
    assert_eq!(s.declaration_at(0, at(FRM, "lvwTrips.ListItems", "lvwTrips")), Some(control(2)));
    // No "Variable not defined" for a control under Option Explicit.
    assert_eq!(s.diagnostics(0), vec![]);
}

#[test]
fn control_definition_is_its_begin_line() {
    let s = session(&[("frmMain.frm", FRM)]);
    let def = s.definition(0, at(FRM, "cmdOK.Caption = ", "cmdOK")).unwrap();
    assert_eq!(def.module, 0);
    assert_eq!(def.span.start, at(FRM, "VB.CommandButton cmdOK", "cmdOK"));
    assert_eq!(def.span.len, 5);
}

#[test]
fn control_array_is_one_control_with_every_element_declared() {
    let s = session(&[("frmMain.frm", FRM)]);
    let first = nth(FRM, "txtDate", 0);
    let second = nth(FRM, "txtDate", 1);
    // Both declarations are the same control.
    assert_eq!(s.declaration_at(0, first), Some(control(0)));
    assert_eq!(s.declaration_at(0, second), Some(control(0)));
    // A use is a reference to it; the first element is its definition.
    let use_at = at(FRM, "txtDate(0).Text", "txtDate");
    assert_eq!(s.declaration_at(0, use_at), Some(control(0)));
    assert_eq!(s.definition(0, use_at).unwrap().span.start, first);
}

#[test]
fn control_member_of_an_external_type_is_not_loaded_when_the_library_is_absent() {
    let s = session_with(&[("frmMain.frm", FRM)], no_libraries());
    let site = at(FRM, "txtDate(0).Text", "Text");
    assert_eq!(s.declaration_at(0, site), None);
    assert_eq!(s.member_status_at(0, site), Some(MemberStatus::LibraryNotLoaded("vb".to_string())));
    assert_eq!(s.diagnostics(0), vec![]);
}

#[test]
fn control_member_resolves_through_the_supplied_library_type() {
    let s = session_with(&[("frmMain.frm", FRM)], vb_config());
    assert_eq!(s.member_status_at(0, at(FRM, "txtDate(0).Text", "Text")), Some(MemberStatus::Resolved));
    assert_eq!(s.member_status_at(0, at(FRM, "cmdOK.Caption", "Caption")), Some(MemberStatus::Resolved));
    // Both are members of different types; the same member used twice
    // (`cmdOK.Caption`, `Me.cmdOK.Caption`) is one library member.
    let first = at(FRM, "cmdOK.Caption = ", "Caption");
    let second = at(FRM, "Me.cmdOK.Caption", "Caption");
    let refs = s.references(0, first, false);
    assert_eq!(sites(&[FRM], &refs), vec![
        (0, first, "Caption".to_string()),
        (0, second, "Caption".to_string()),
    ]);
    // Members of a library type have no declaration in the project.
    assert_eq!(s.definition(0, first), None);
    assert_eq!(s.rename(0, first, "Title"), vec![]);
}

#[test]
fn me_reaches_controls_and_the_designer_base_type() {
    let s = session_with(&[("frmMain.frm", FRM)], vb_config());
    assert_eq!(s.declaration_at(0, at(FRM, "Me.cmdOK.Caption", "cmdOK")), Some(control(1)));
    // `Show` is not declared in the form: it is a member of `VB.Form`.
    assert_eq!(s.member_status_at(0, at(FRM, "Me.Show", "Show")), Some(MemberStatus::Resolved));
    let without = session_with(&[("frmMain.frm", FRM)], no_libraries());
    assert_eq!(
        without.member_status_at(0, at(FRM, "Me.Show", "Show")),
        Some(MemberStatus::LibraryNotLoaded("vb.form".to_string()))
    );
    assert_eq!(without.diagnostics(0), vec![]);
}

#[test]
fn unknown_member_of_a_designer_with_its_library_loaded_is_reported() {
    let main = "Sub Run()\n    frmMain.NoSuchThing\nEnd Sub\n";
    let s = session_with(&[("frmMain.frm", FRM), ("modMain.bas", main)], vb_config());
    let site = at(main, "frmMain.NoSuchThing", "NoSuchThing");
    assert_eq!(s.member_status_at(1, site), Some(MemberStatus::NoSuchMember));
    assert_eq!(common::coded(&s, main, 1, 461).len(), 1);
}

#[test]
fn form_name_is_a_default_instance_with_its_own_members_and_controls() {
    let main = "Sub Run()\n    frmMain.RefreshView\n    frmMain.cmdOK.Caption = \"z\"\n    frmMain.txtDate(1).Text = \"t\"\nEnd Sub\n";
    let s = session(&[("frmMain.frm", FRM), ("modMain.bas", main)]);
    assert_eq!(s.declaration_at(1, at(main, "frmMain.RefreshView", "frmMain")), Some(DeclId::Module { module: 0 }));
    assert_eq!(s.declaration_at(1, at(main, "frmMain.RefreshView", "RefreshView")), Some(DeclId::Proc { module: 0, idx: 0 }));
    assert_eq!(s.declaration_at(1, at(main, "frmMain.cmdOK.Caption", "cmdOK")), Some(control(1)));
    assert_eq!(s.declaration_at(1, at(main, "frmMain.txtDate(1).Text", "txtDate")), Some(control(0)));
}

#[test]
fn event_handler_is_an_ordinary_procedure() {
    let s = session(&[("frmMain.frm", FRM)]);
    // `cmdOK_Click` is procedure 1 (after RefreshView); a call to it resolves
    // to that procedure and is not linked to the control.
    let call = at(FRM, "cmdOK_Click\r\nEnd Sub", "cmdOK_Click");
    assert_eq!(s.declaration_at(0, call), Some(DeclId::Proc { module: 0, idx: 1 }));
    let handler = at(FRM, "Sub cmdOK_Click", "cmdOK_Click");
    assert_eq!(s.declaration_at(0, handler), Some(DeclId::Proc { module: 0, idx: 1 }));
    // Renaming the control touches only the control's names, not the handler.
    let edits = s.rename(0, at(FRM, "cmdOK.Caption = ", "cmdOK"), "btnOK");
    let spans: Vec<u32> = edits.iter().map(|e| e.span.start).collect();
    assert_eq!(
        spans,
        vec![
            at(FRM, "VB.CommandButton cmdOK", "cmdOK"),
            at(FRM, "cmdOK.Caption = ", "cmdOK"),
            at(FRM, "Me.cmdOK.Caption", "cmdOK"),
        ]
    );
    assert!(edits.iter().all(|e| e.new_text == "btnOK" && e.module == 0));
}

#[test]
fn user_control_of_the_project_types_a_control() {
    let ctl = "VERSION 5.00\r\nBegin VB.UserControl ctlTrip\r\nEnd\r\nAttribute VB_Name = \"ctlTrip\"\r\nPublic Sub Reload()\r\nEnd Sub\r\n";
    let frm = "VERSION 5.00\r\nBegin VB.Form frmHost\r\n   Begin Project1.ctlTrip ctlMain\r\n   End\r\nEnd\r\nAttribute VB_Name = \"frmHost\"\r\nSub T()\r\n    ctlMain.Reload\r\nEnd Sub\r\n";
    let config = SessionConfig { project_name: Some("Project1".into()), libraries: vec![] };
    let s = session_with(&[("ctlTrip.ctl", ctl), ("frmHost.frm", frm)], config);
    assert_eq!(s.declaration_at(1, at(frm, "ctlMain.Reload", "ctlMain")), Some(DeclId::Control { module: 1, idx: 0 }));
    assert_eq!(s.declaration_at(1, at(frm, "ctlMain.Reload", "Reload")), Some(DeclId::Proc { module: 0, idx: 0 }));
    // The control's type, written on its `Begin` line, is a use of the
    // UserControl's module.
    let decl = at(ctl, "\"ctlTrip\"", "\"ctlTrip\"");
    let refs = s.references(0, decl, false);
    assert_eq!(
        sites(&[ctl, frm], &refs),
        vec![(1, at(frm, "Project1.ctlTrip ctlMain", "Project1.ctlTrip"), "Project1.ctlTrip".to_string())]
    );
}

#[test]
fn control_hover_shows_name_and_type() {
    let s = session(&[("frmMain.frm", FRM)]);
    let h = s.hover(0, at(FRM, "cmdOK.Caption = ", "cmdOK")).unwrap();
    assert_eq!(h.text, "cmdOK As VB.CommandButton");
    let arr = s.hover(0, at(FRM, "txtDate(0).Text", "txtDate")).unwrap();
    assert_eq!(arr.text, "txtDate() As VB.TextBox");
}

#[test]
fn unqualified_names_in_a_designer_reach_the_members_it_inherits() {
    let lib = r#"[{"name":"VB","types":[
        {"name":"Form","kind":"class","members":[
          {"name":"Hide","kind":"method","param_count":0},
          {"name":"ScaleX","kind":"method","return_type":"Single","param_count":3},
          {"name":"PopupMenu","kind":"method","param_count":1},
          {"name":"Caption","kind":"field","return_type":"String"}]},
        {"name":"TextBox","kind":"class","members":[{"name":"Text","kind":"field","return_type":"String"}]}]}]"#;
    let frm = "VERSION 5.00\r\nBegin VB.Form frmA\r\n   Begin VB.TextBox Caption\r\n   End\r\nEnd\r\nAttribute VB_Name = \"frmA\"\r\n\
               Option Explicit\r\nPrivate Sub Go()\r\n    Hide\r\n    x = ScaleX(1, 2, 3)\r\n    PopupMenu mnuFile\r\nEnd Sub\r\n";
    let libs = Library::from_json(lib).unwrap();
    let s = session_with(&[("frmA.frm", frm)], SessionConfig { project_name: None, libraries: libs });
    let one = |m: &str| s.external_member_references("VB", "Form", m);
    assert_eq!(one("Hide").iter().map(|l| l.span.start).collect::<Vec<_>>(), vec![at(frm, "    Hide", "Hide")]);
    assert_eq!(one("ScaleX").iter().map(|l| l.span.start).collect::<Vec<_>>(), vec![at(frm, "ScaleX(1", "ScaleX")]);
    assert_eq!(one("PopupMenu").iter().map(|l| l.span.start).collect::<Vec<_>>(), vec![at(frm, "PopupMenu mnuFile", "PopupMenu")]);
    // A control of the same name as an inherited member is the control.
    assert_eq!(one("Caption"), vec![]);
    // The names that are not inherited remain undeclared (`x`, `mnuFile`).
    let undeclared: Vec<u32> = s.diagnostics(0).iter().map(|d| d.span.start).collect();
    assert_eq!(undeclared, vec![at(frm, "x = ScaleX", "x"), at(frm, "mnuFile", "mnuFile")]);
}

#[test]
fn inherited_members_are_unknown_when_the_designer_library_is_not_supplied() {
    let frm = "VERSION 5.00\r\nBegin VB.Form frmA\r\nEnd\r\nAttribute VB_Name = \"frmA\"\r\nOption Explicit\r\nPrivate Sub Go()\r\n    Hide\r\nEnd Sub\r\n";
    let s = session_with(&[("frmA.frm", frm)], no_libraries());
    assert_eq!(s.diagnostics(0).len(), 1);
}

const CTL_LIB: &str = r#"{
  "name": "R8DateTimeCtl",
  "types": [
    {"name": "RMDateTimeCtrl", "kind": "class", "members": [
      {"name": "Value", "kind": "field", "return_type": "Date"}]}
  ]
}"#;

const CTL_FRM: &str = "VERSION 5.00\r\n\
Begin VB.Form frmDates\r\n\
   Begin R8DateTimeCtlCtl.RMDateTimeCtrl txtDue\r\n\
   End\r\n\
End\r\n\
Attribute VB_Name = \"frmDates\"\r\n\
Sub T()\r\n\
    txtDue.Value = Now\r\n\
    txtDue.Missing = 1\r\n\
End Sub\r\n";

#[test]
fn control_qualified_with_library_name_and_ctl_resolves_to_that_library() {
    let config = SessionConfig { project_name: None, libraries: Library::from_json(CTL_LIB).unwrap() };
    let s = session_with(&[("frmDates.frm", CTL_FRM)], config);
    assert_eq!(s.member_status_at(0, at(CTL_FRM, "txtDue.Value", "Value")), Some(MemberStatus::Resolved));
    assert_eq!(s.member_status_at(0, at(CTL_FRM, "txtDue.Missing", "Missing")), Some(MemberStatus::NoSuchMember));
}

#[test]
fn control_qualified_with_ctl_of_a_library_that_is_not_loaded_stays_not_loaded() {
    let s = session_with(&[("frmDates.frm", CTL_FRM)], no_libraries());
    assert_eq!(
        s.member_status_at(0, at(CTL_FRM, "txtDue.Value", "Value")),
        Some(MemberStatus::LibraryNotLoaded("r8datetimectlctl".to_string()))
    );
}

const OCX_LIB: &str = r#"{
  "name": "MSComctlLib",
  "types": [
    {"name": "TreeView", "kind": "coclass", "default_interface": "ITreeView", "is_control": true},
    {"name": "ITreeView", "kind": "interface", "members": [
      {"name": "Nodes", "kind": "propertyget", "return_type": "Long"}]},
    {"name": "ImageList", "kind": "coclass", "default_interface": "IImageList"},
    {"name": "IImageList", "kind": "interface", "members": [
      {"name": "Count", "kind": "propertyget", "return_type": "Long"}]}
  ]
}"#;

const EXTENDER_LIB: &str = r#"{
  "name": "VB",
  "types": [
    {"name": "_VBControlExtender", "kind": "interface", "members": [
      {"name": "Left", "kind": "field", "return_type": "Single"},
      {"name": "Tag", "kind": "field", "return_type": "String"}]},
    {"name": "VBControlExtender", "kind": "coclass", "default_interface": "_VBControlExtender"},
    {"name": "Form", "kind": "class", "members": [
      {"name": "Show", "kind": "method", "param_count": 0}]}
  ]
}"#;

const OCX_FRM: &str = "VERSION 5.00\r\n\
Begin VB.Form frmTree\r\n\
   Begin MSComctlLib.TreeView tvwMain\r\n\
   End\r\n\
End\r\n\
Attribute VB_Name = \"frmTree\"\r\n\
Sub T()\r\n\
    Dim imlMain As MSComctlLib.ImageList\r\n\
    n = tvwMain.Nodes\r\n\
    tvwMain.Left = 5\r\n\
    tvwMain.Missing = 1\r\n\
    imlMain.Left = 5\r\n\
End Sub\r\n";

#[test]
fn control_of_a_library_has_the_extender_members_the_container_adds() {
    let mut libraries = Library::from_json(OCX_LIB).unwrap();
    libraries.extend(Library::from_json(EXTENDER_LIB).unwrap());
    let config = SessionConfig { project_name: None, libraries };
    let s = session_with(&[("frmTree.frm", OCX_FRM)], config);
    // the library's own member is found first
    assert_eq!(s.member_status_at(0, at(OCX_FRM, "tvwMain.Nodes", "Nodes")), Some(MemberStatus::Resolved));
    // an extender member of a control
    assert_eq!(s.member_status_at(0, at(OCX_FRM, "tvwMain.Left", "Left")), Some(MemberStatus::Resolved));
    // neither the library nor the extender has it
    assert_eq!(s.member_status_at(0, at(OCX_FRM, "tvwMain.Missing", "Missing")), Some(MemberStatus::NoSuchMember));
    // a class that is not a control gets none of the extender's members
    assert_eq!(s.member_status_at(0, at(OCX_FRM, "imlMain.Left", "Left")), Some(MemberStatus::NoSuchMember));
}

const UNFLAGGED_LIB: &str = r#"{
  "name": "R8TabControl",
  "types": [
    {"name": "RMTabControl", "kind": "coclass", "default_interface": "_RMTabControl"},
    {"name": "_RMTabControl", "kind": "interface", "members": [
      {"name": "AddTab", "kind": "method", "param_count": 0}]}
  ]
}"#;

const PLACED_FRM: &str = "VERSION 5.00\r\n\
Begin VB.Form frmTabs\r\n\
   Begin R8TabControl.RMTabControl tabMain\r\n\
   End\r\n\
End\r\n\
Attribute VB_Name = \"frmTabs\"\r\n\
Sub T()\r\n\
    tabMain.AddTab\r\n\
    tabMain.Left = 5\r\n\
    Dim loose As R8TabControl.RMTabControl\r\n\
    loose.Tag = \"x\"\r\n\
End Sub\r\n";

#[test]
fn library_class_placed_on_a_form_gets_the_extender_members_even_when_unflagged() {
    let mut libraries = Library::from_json(UNFLAGGED_LIB).unwrap();
    libraries.extend(Library::from_json(EXTENDER_LIB).unwrap());
    let config = SessionConfig { project_name: None, libraries };
    let s = session_with(&[("frmTabs.frm", PLACED_FRM)], config);
    assert_eq!(s.member_status_at(0, at(PLACED_FRM, "tabMain.AddTab", "AddTab")), Some(MemberStatus::Resolved));
    assert_eq!(s.member_status_at(0, at(PLACED_FRM, "tabMain.Left", "Left")), Some(MemberStatus::Resolved));
    // the class is a control class: held in a variable it has the same members
    assert_eq!(s.member_status_at(0, at(PLACED_FRM, "loose.Tag", "Tag")), Some(MemberStatus::Resolved));
}
