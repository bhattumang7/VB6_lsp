//! A class's (or form's) public members are reached through an instance, not
//! as project-wide names; standard modules' members, public types and enums
//! of any module, and members of `VB_GlobalNameSpace` classes are.

mod common;

use common::{at, session};
use vb6_core::session::DeclId;

#[test]
fn class_public_procedure_is_not_a_global_name() {
    let cls = "Public Sub Save()\nEnd Sub\n";
    let main = "Option Explicit\nSub Run()\n    Save\nEnd Sub\n";
    let s = session(&[("clsTrip.cls", cls), ("modMain.bas", main)]);
    assert_eq!(s.declaration_at(1, at(main, "    Save", "Save")), None);
    // The bare name is undefined: a call to a name that is nothing.
    let diags = s.diagnostics(1);
    assert_eq!(diags.len(), 1);
    assert_eq!(diags[0].code, 0x23);
    assert_eq!(s.references(0, at(cls, "Sub Save", "Save"), false), vec![]);
}

#[test]
fn class_public_variable_is_not_a_global_name() {
    let cls = "Public Total As Long\n";
    let main = "Option Explicit\nSub Run()\n    Total = 1\nEnd Sub\n";
    let s = session(&[("clsTrip.cls", cls), ("modMain.bas", main)]);
    assert_eq!(s.declaration_at(1, at(main, "Total = 1", "Total")), None);
    assert_eq!(s.diagnostics(1).len(), 1);
}

#[test]
fn form_public_procedure_is_not_a_global_name() {
    let frm = "VERSION 5.00\nBegin VB.Form frmMain\nEnd\nAttribute VB_Name = \"frmMain\"\nPublic Sub Refresh2()\nEnd Sub\n";
    let main = "Option Explicit\nSub Run()\n    Refresh2\nEnd Sub\n";
    let s = session(&[("frmMain.frm", frm), ("modMain.bas", main)]);
    assert_eq!(s.declaration_at(1, at(main, "    Refresh2", "Refresh2")), None);
}

#[test]
fn standard_module_members_stay_global() {
    let util = "Public Sub DoThing()\nEnd Sub\nPublic gLevel As Long\n";
    let main = "Sub Run()\n    DoThing\n    gLevel = 1\nEnd Sub\n";
    let s = session(&[("modUtil.bas", util), ("modMain.bas", main)]);
    assert_eq!(s.declaration_at(1, at(main, "    DoThing", "DoThing")), Some(DeclId::Proc { module: 0, idx: 0 }));
    assert_eq!(
        s.declaration_at(1, at(main, "gLevel = 1", "gLevel")),
        Some(DeclId::ModuleVar { module: 0, idx: 0 })
    );
}

#[test]
fn public_enum_and_type_of_a_class_stay_project_wide() {
    let cls = "Public Enum TripState\n    Active\n    Done\nEnd Enum\nPublic Type TLeg\n    Km As Long\nEnd Type\n";
    let main = "Sub Run()\n    Dim n As Long\n    n = Done\n    Dim l As TLeg\n    l.Km = 1\nEnd Sub\n";
    let s = session(&[("clsTrip.cls", cls), ("modMain.bas", main)]);
    assert_eq!(
        s.declaration_at(1, at(main, "n = Done", "Done")),
        Some(DeclId::EnumMember { module: 0, enum_idx: 0, member_idx: 1 })
    );
    assert_eq!(
        s.declaration_at(1, at(main, "l.Km", "Km")),
        Some(DeclId::TypeMember { module: 0, type_idx: 0, member_idx: 0 })
    );
}

#[test]
fn global_namespace_class_members_are_project_wide() {
    let cls = "Attribute VB_GlobalNameSpace = True\nPublic Sub Helper()\nEnd Sub\n";
    let main = "Sub Run()\n    Helper\nEnd Sub\n";
    let s = session(&[("clsGlobal.cls", cls), ("modMain.bas", main)]);
    assert_eq!(s.declaration_at(1, at(main, "    Helper", "Helper")), Some(DeclId::Proc { module: 0, idx: 0 }));
}

#[test]
fn own_class_members_resolve_unqualified_inside_the_class() {
    let cls = "Public Sub A()\nEnd Sub\nPublic Sub B()\n    A\nEnd Sub\n";
    let s = session(&[("clsTrip.cls", cls)]);
    assert_eq!(s.declaration_at(0, at(cls, "    A", "A")), Some(DeclId::Proc { module: 0, idx: 0 }));
}
