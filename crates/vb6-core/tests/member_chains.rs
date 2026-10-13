//! Members reached through the result of a function or property, through
//! another member, through an indexed array, or through a default member.

mod common;

use common::{at, session};
use vb6_core::sema::MemberStatus;
use vb6_core::session::DeclId;

const CLS_TRIP: &str = "Public Sub Save()\nEnd Sub\nPublic Name As String\n";

fn proc(module: usize, idx: usize) -> DeclId {
    DeclId::Proc { module, idx }
}

fn var(module: usize, idx: usize) -> DeclId {
    DeclId::ModuleVar { module, idx }
}

#[test]
fn function_result_member() {
    let main = "Function GetTrip() As clsTrip\nEnd Function\nSub Run()\n    GetTrip().Save\n    GetTrip.Save\nEnd Sub\n";
    let s = session(&[("clsTrip.cls", CLS_TRIP), ("modMain.bas", main)]);
    assert_eq!(s.declaration_at(1, at(main, "GetTrip().Save", "GetTrip")), Some(proc(1, 0)));
    assert_eq!(s.declaration_at(1, at(main, "GetTrip().Save", "Save")), Some(proc(0, 0)));
    assert_eq!(s.declaration_at(1, at(main, "GetTrip.Save", "Save")), Some(proc(0, 0)));
}

#[test]
fn function_of_another_module_result_member() {
    let util = "Public Function GetTrip() As clsTrip\nEnd Function\n";
    let main = "Sub Run()\n    GetTrip().Save\nEnd Sub\n";
    let s = session(&[("clsTrip.cls", CLS_TRIP), ("modUtil.bas", util), ("modMain.bas", main)]);
    assert_eq!(s.declaration_at(2, at(main, "GetTrip().Save", "GetTrip")), Some(proc(1, 0)));
    assert_eq!(s.declaration_at(2, at(main, "GetTrip().Save", "Save")), Some(proc(0, 0)));
}

#[test]
fn property_get_result_chain() {
    let route = "Public Property Get Trip() As clsTrip\nEnd Property\n";
    let main = "Sub Run()\n    Dim oRoute As clsRoute\n    oRoute.Trip.Save\n    oRoute.Trip.Name = \"n\"\nEnd Sub\n";
    let s = session(&[("clsTrip.cls", CLS_TRIP), ("clsRoute.cls", route), ("modMain.bas", main)]);
    assert_eq!(s.declaration_at(2, at(main, "oRoute.Trip.Save", "Trip")), Some(proc(1, 0)));
    assert_eq!(s.declaration_at(2, at(main, "oRoute.Trip.Save", "Save")), Some(proc(0, 0)));
    assert_eq!(s.declaration_at(2, at(main, "oRoute.Trip.Name", "Name")), Some(var(0, 0)));
}

#[test]
fn public_object_field_chain() {
    let route = "Public Trip As clsTrip\n";
    let main = "Sub Run()\n    Dim oRoute As clsRoute\n    oRoute.Trip.Save\nEnd Sub\n";
    let s = session(&[("clsTrip.cls", CLS_TRIP), ("clsRoute.cls", route), ("modMain.bas", main)]);
    assert_eq!(s.declaration_at(2, at(main, "oRoute.Trip.Save", "Trip")), Some(var(1, 0)));
    assert_eq!(s.declaration_at(2, at(main, "oRoute.Trip.Save", "Save")), Some(proc(0, 0)));
}

#[test]
fn function_with_arguments_result_member() {
    let main = "Function Find(ByVal id As Long) As clsTrip\nEnd Function\nSub Run()\n    Find(3).Save\nEnd Sub\n";
    let s = session(&[("clsTrip.cls", CLS_TRIP), ("modMain.bas", main)]);
    assert_eq!(s.declaration_at(1, at(main, "Find(3).Save", "Save")), Some(proc(0, 0)));
}

#[test]
fn function_returning_an_array_is_indexed() {
    let main = "Function All() As clsTrip()\nEnd Function\nSub Run()\n    All()(1).Save\n    All(2).Save\nEnd Sub\n";
    let s = session(&[("clsTrip.cls", CLS_TRIP), ("modMain.bas", main)]);
    assert_eq!(s.declaration_at(1, at(main, "All()(1).Save", "Save")), Some(proc(0, 0)));
    assert_eq!(s.declaration_at(1, at(main, "All(2).Save", "Save")), Some(proc(0, 0)));
}

#[test]
fn property_returning_an_array_is_indexed() {
    let route = "Public Property Get Trips() As clsTrip()\nEnd Property\n";
    let main = "Sub Run()\n    Dim oRoute As clsRoute\n    oRoute.Trips(1).Save\nEnd Sub\n";
    let s = session(&[("clsTrip.cls", CLS_TRIP), ("clsRoute.cls", route), ("modMain.bas", main)]);
    assert_eq!(s.declaration_at(2, at(main, "oRoute.Trips(1).Save", "Trips")), Some(proc(1, 0)));
    assert_eq!(s.declaration_at(2, at(main, "oRoute.Trips(1).Save", "Save")), Some(proc(0, 0)));
}

#[test]
fn public_array_field_is_indexed() {
    let route = "Public Trips(5) As clsTrip\n";
    let main = "Sub Run()\n    Dim oRoute As clsRoute\n    oRoute.Trips(2).Save\nEnd Sub\n";
    let s = session(&[("clsTrip.cls", CLS_TRIP), ("clsRoute.cls", route), ("modMain.bas", main)]);
    assert_eq!(s.declaration_at(2, at(main, "oRoute.Trips(2).Save", "Save")), Some(proc(0, 0)));
}

#[test]
fn default_member_of_a_collection_class_types_the_call() {
    let coll = "Public Function Item(ByVal i As Long) As clsTrip\nAttribute Item.VB_UserMemId = 0\nEnd Function\n";
    let main = "Sub Run()\n    Dim oTrips As clsTrips\n    oTrips(1).Save\n    oTrips.Item(1).Save\nEnd Sub\n";
    let s = session(&[("clsTrip.cls", CLS_TRIP), ("clsTrips.cls", coll), ("modMain.bas", main)]);
    assert_eq!(s.declaration_at(2, at(main, "oTrips(1).Save", "Save")), Some(proc(0, 0)));
    assert_eq!(s.declaration_at(2, at(main, "oTrips.Item(1).Save", "Item")), Some(proc(1, 0)));
    assert_eq!(s.declaration_at(2, at(main, "oTrips.Item(1).Save", "Save")), Some(proc(0, 0)));
}

#[test]
fn class_without_a_default_member_cannot_be_indexed() {
    let main = "Sub Run()\n    Dim oTrip As clsTrip\n    oTrip(1).Save\nEnd Sub\n";
    let s = session(&[("clsTrip.cls", CLS_TRIP), ("modMain.bas", main)]);
    let site = at(main, "oTrip(1).Save", "Save");
    assert_eq!(s.declaration_at(1, site), None);
    assert_eq!(s.member_status_at(1, site), Some(MemberStatus::BaseUnresolved));
}

#[test]
fn member_of_a_late_bound_result_stays_late_bound() {
    let main = "Function Make() As Object\nEnd Function\nSub Run()\n    Make().Save\n    Make.Save\nEnd Sub\n";
    let s = session(&[("clsTrip.cls", CLS_TRIP), ("modMain.bas", main)]);
    for ctx in ["Make().Save", "Make.Save"] {
        let site = at(main, ctx, "Save");
        assert_eq!(s.declaration_at(1, site), None, "{ctx}");
        assert_eq!(s.member_status_at(1, site), Some(MemberStatus::LateBound), "{ctx}");
    }
}

#[test]
fn builtin_result_member_is_late_bound() {
    let main = "Sub Run()\n    CreateObject(\"X.Y\").Save\nEnd Sub\n";
    let s = session(&[("modMain.bas", main)]);
    assert_eq!(
        s.member_status_at(0, at(main, "CreateObject(\"X.Y\").Save", "Save")),
        Some(MemberStatus::LateBound)
    );
}

#[test]
fn set_from_a_member_result_assigns_the_type_to_the_target() {
    let route = "Public Property Get Trip() As clsTrip\nEnd Property\n";
    let main = "Sub Run()\n    Dim oRoute As clsRoute\n    Dim oTrip As clsTrip\n    Set oTrip = oRoute.Trip\n    oTrip.Save\nEnd Sub\n";
    let s = session(&[("clsTrip.cls", CLS_TRIP), ("clsRoute.cls", route), ("modMain.bas", main)]);
    assert_eq!(s.declaration_at(2, at(main, "oRoute.Trip", "Trip")), Some(proc(1, 0)));
    assert_eq!(s.declaration_at(2, at(main, "oTrip.Save", "Save")), Some(proc(0, 0)));
}
