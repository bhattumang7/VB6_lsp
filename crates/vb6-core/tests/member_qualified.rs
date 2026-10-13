//! Qualified names: module and project qualifiers on members, qualified class
//! and type names in declarations, enum names as qualifiers.

mod common;

use common::{at, coded, session, session_with};
use vb6_core::sema::MemberStatus;
use vb6_core::session::{DeclId, SessionConfig};

const CLS_TRIP: &str = "Public Sub Save()\nEnd Sub\nPrivate Sub Hidden()\nEnd Sub\n";
const MOD_UTIL: &str = "Public Sub DoThing()\nEnd Sub\nPrivate Sub Secret()\nEnd Sub\nPublic gLevel As Long\n\
                        Public Enum Color\n    Red\n    Green\nEnd Enum\nPublic Type TPoint\n    X As Long\nEnd Type\n";

fn config(name: &str) -> SessionConfig {
    SessionConfig { project_name: Some(name.to_string()), libraries: vec![] }
}

fn proc(module: usize, idx: usize) -> DeclId {
    DeclId::Proc { module, idx }
}

#[test]
fn module_qualifier_resolves_the_module_and_its_member() {
    let main = "Sub Run()\n    modUtil.DoThing\n    modUtil.gLevel = 3\nEnd Sub\n";
    let s = session(&[("modUtil.bas", MOD_UTIL), ("modMain.bas", main)]);
    assert_eq!(s.declaration_at(1, at(main, "modUtil.DoThing", "DoThing")), Some(proc(0, 0)));
    assert_eq!(
        s.declaration_at(1, at(main, "modUtil.gLevel", "gLevel")),
        Some(DeclId::ModuleVar { module: 0, idx: 0 })
    );
    // The qualifier names the module itself.
    assert_eq!(
        s.declaration_at(1, at(main, "modUtil.DoThing", "modUtil")),
        Some(DeclId::Module { module: 0 })
    );
}

#[test]
fn module_qualifier_does_not_reach_private_members() {
    let main = "Sub Run()\n    modUtil.Secret\nEnd Sub\n";
    let s = session(&[("modUtil.bas", MOD_UTIL), ("modMain.bas", main)]);
    let site = at(main, "modUtil.Secret", "Secret");
    assert_eq!(s.declaration_at(1, site), None);
    assert_eq!(s.member_status_at(1, site), Some(MemberStatus::NoSuchMember));
    assert_eq!(coded(&s, main, 1, 461).len(), 1);
}

#[test]
fn module_name_from_the_vb_name_attribute() {
    let util = "Attribute VB_Name = \"modTools\"\nPublic Sub DoThing()\nEnd Sub\n";
    let main = "Sub Run()\n    modTools.DoThing\nEnd Sub\n";
    let s = session(&[("Whatever.bas", util), ("modMain.bas", main)]);
    assert_eq!(s.declaration_at(1, at(main, "modTools.DoThing", "DoThing")), Some(proc(0, 0)));
    // The module is declared where its name is written.
    let def = s.definition(1, at(main, "modTools.DoThing", "modTools")).unwrap();
    assert_eq!(def.module, 0);
    assert_eq!(def.span.start, at(util, "\"modTools\"", "\"modTools\""));
}

#[test]
fn project_qualifier_then_module_then_member() {
    let main = "Sub Run()\n    Project1.modUtil.DoThing\n    Project1.DoThing\nEnd Sub\n";
    let s = session_with(&[("modUtil.bas", MOD_UTIL), ("modMain.bas", main)], config("Project1"));
    assert_eq!(s.declaration_at(1, at(main, "Project1.modUtil.DoThing", "DoThing")), Some(proc(0, 0)));
    assert_eq!(
        s.declaration_at(1, at(main, "Project1.modUtil.DoThing", "modUtil")),
        Some(DeclId::Module { module: 0 })
    );
    assert_eq!(s.declaration_at(1, at(main, "Project1.DoThing", "DoThing")), Some(proc(0, 0)));
}

#[test]
fn qualified_class_declaration_resolves_the_class() {
    let main = "Sub Run()\n    Dim o As Project1.clsTrip\n    Set o = New Project1.clsTrip\n    o.Save\nEnd Sub\n";
    let s = session_with(&[("clsTrip.cls", CLS_TRIP), ("modMain.bas", main)], config("Project1"));
    assert_eq!(s.declaration_at(1, at(main, "o.Save", "Save")), Some(proc(0, 0)));
    // Both spellings of the type name are uses of the class.
    assert_eq!(
        s.declaration_at(1, at(main, "As Project1.clsTrip", "clsTrip")),
        Some(DeclId::Module { module: 0 })
    );
    assert_eq!(
        s.declaration_at(1, at(main, "New Project1.clsTrip", "clsTrip")),
        Some(DeclId::Module { module: 0 })
    );
    assert_eq!(s.diagnostics(1), vec![]);
}

#[test]
fn qualified_class_of_a_project_that_is_not_loaded_is_unresolved() {
    let main = "Option Explicit\nSub Run()\n    Dim o As Other.clsTrip\n    o.Save\nEnd Sub\n";
    let s = session_with(&[("clsTrip.cls", CLS_TRIP), ("modMain.bas", main)], config("Project1"));
    let site = at(main, "o.Save", "Save");
    assert_eq!(s.declaration_at(1, site), None);
    assert_eq!(s.member_status_at(1, site), Some(MemberStatus::LibraryNotLoaded("other".to_string())));
    assert_eq!(s.diagnostics(1), vec![]);
}

#[test]
fn qualified_class_that_does_not_exist_in_the_project_is_unknown() {
    let main = "Sub Run()\n    Dim o As Project1.clsMissing\n    o.Save\nEnd Sub\n";
    let s = session_with(&[("clsTrip.cls", CLS_TRIP), ("modMain.bas", main)], config("Project1"));
    assert_eq!(
        s.member_status_at(1, at(main, "o.Save", "Save")),
        Some(MemberStatus::UnknownType("clsmissing".to_string()))
    );
}

#[test]
fn module_qualified_type_resolves_its_fields() {
    let main = "Sub Run()\n    Dim p As modUtil.TPoint\n    p.X = 1\nEnd Sub\n";
    let s = session(&[("modUtil.bas", MOD_UTIL), ("modMain.bas", main)]);
    assert_eq!(
        s.declaration_at(1, at(main, "p.X", "X")),
        Some(DeclId::TypeMember { module: 0, type_idx: 0, member_idx: 0 })
    );
    assert_eq!(
        s.declaration_at(1, at(main, "As modUtil.TPoint", "TPoint")),
        Some(DeclId::Type { module: 0, idx: 0 })
    );
    assert_eq!(
        s.declaration_at(1, at(main, "As modUtil.TPoint", "modUtil")),
        Some(DeclId::Module { module: 0 })
    );
}

#[test]
fn enum_name_qualifies_its_members() {
    let main = "Sub Run()\n    Dim c As Long\n    c = Color.Green\n    c = modUtil.Color.Red\nEnd Sub\n";
    let s = session(&[("modUtil.bas", MOD_UTIL), ("modMain.bas", main)]);
    let enum_member = |idx: usize| DeclId::EnumMember { module: 0, enum_idx: 0, member_idx: idx };
    assert_eq!(s.declaration_at(1, at(main, "Color.Green", "Green")), Some(enum_member(1)));
    assert_eq!(s.declaration_at(1, at(main, "Color.Green", "Color")), Some(DeclId::Enum { module: 0, idx: 0 }));
    assert_eq!(s.declaration_at(1, at(main, "modUtil.Color.Red", "Red")), Some(enum_member(0)));
    assert_eq!(s.diagnostics(1), vec![]);
}

#[test]
fn references_to_a_class_include_every_type_position() {
    let helper = "Public Function Make() As clsTrip\nEnd Function\n";
    let main = "Private m As clsTrip\nSub Run(ByVal t As clsTrip)\n    Dim o As New clsTrip\nEnd Sub\n";
    let s = session(&[("clsTrip.cls", "Attribute VB_Name = \"clsTrip\"\n"), ("modHelper.bas", helper), ("modMain.bas", main)]);
    let decl = at("Attribute VB_Name = \"clsTrip\"\n", "\"clsTrip\"", "\"clsTrip\"");
    let refs = s.references(0, decl, false);
    let got: Vec<(usize, u32)> = refs.iter().map(|l| (l.module, l.span.start)).collect();
    let want = vec![
        (1, at(helper, "As clsTrip", "clsTrip")),
        (2, at(main, "m As clsTrip", "clsTrip")),
        (2, at(main, "t As clsTrip", "clsTrip")),
        (2, at(main, "New clsTrip", "clsTrip")),
    ];
    assert_eq!(got, want);
    
}

#[test]
fn class_name_qualifier_reaches_enum_members_not_instance_members() {
    let cls = "Public Enum TripState\n    Active\n    Done\nEnd Enum\nPublic Sub Save()\nEnd Sub\n";
    let main = "Sub Run()\n    Dim n As Long\n    n = clsTrip.Done\n    n = clsTrip.TripState.Active\n    clsTrip.Save\nEnd Sub\n";
    let s = session(&[("clsTrip.cls", cls), ("modMain.bas", main)]);
    let member = |idx: usize| DeclId::EnumMember { module: 0, enum_idx: 0, member_idx: idx };
    assert_eq!(s.declaration_at(1, at(main, "clsTrip.Done", "Done")), Some(member(1)));
    assert_eq!(s.declaration_at(1, at(main, "clsTrip.TripState.Active", "Active")), Some(member(0)));
    // A class name that is not a default instance gives no access to instance members.
    assert_eq!(s.member_status_at(1, at(main, "clsTrip.Save", "Save")), Some(MemberStatus::NoSuchMember));
}

#[test]
fn module_name_qualifies_enum_members() {
    let main = "Sub Run()\n    Dim n As Long\n    n = modUtil.Green\nEnd Sub\n";
    let s = session(&[("modUtil.bas", MOD_UTIL), ("modMain.bas", main)]);
    assert_eq!(
        s.declaration_at(1, at(main, "modUtil.Green", "Green")),
        Some(DeclId::EnumMember { module: 0, enum_idx: 0, member_idx: 1 })
    );
}

#[test]
fn implements_names_the_interface_class() {
    let iface = "Attribute VB_Name = \"IShape\"\nPublic Sub Draw()\nEnd Sub\n";
    let impl_ = "Implements IShape\nPrivate Sub IShape_Draw()\nEnd Sub\n";
    let s = session(&[("IShape.cls", iface), ("clsSquare.cls", impl_)]);
    assert_eq!(s.declaration_at(1, at(impl_, "Implements IShape", "IShape")), Some(DeclId::Module { module: 0 }));
    let refs = s.references(0, at(iface, "\"IShape\"", "\"IShape\""), false);
    assert_eq!(refs.len(), 1);
    assert_eq!((refs[0].module, refs[0].span.start), (1, at(impl_, "Implements IShape", "IShape")));
}

#[test]
fn module_hover_names_its_kind() {
    let cls = "Attribute VB_Name = \"clsTrip\"\n";
    let main = "Dim o As clsTrip\n";
    let s = session(&[("clsTrip.cls", cls), ("modMain.bas", main)]);
    assert_eq!(s.hover(1, at(main, "As clsTrip", "clsTrip")).unwrap().text, "Class clsTrip");
}
