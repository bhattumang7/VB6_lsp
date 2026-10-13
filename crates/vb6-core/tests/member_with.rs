//! Leading-dot members inside `With` resolve against the innermost object.

mod common;

use common::{at, nth, session};
use vb6_core::sema::MemberStatus;
use vb6_core::session::DeclId;

const CLS_TRIP: &str = "Public Sub Save()\nEnd Sub\nPublic Name As String\n";
const CLS_ROUTE: &str = "Public Sub Plan()\nEnd Sub\nPublic Trip As clsTrip\n";

fn proc(module: usize, idx: usize) -> DeclId {
    DeclId::Proc { module, idx }
}

#[test]
fn with_object_variable() {
    let main = "Sub Run()\n    Dim oTrip As clsTrip\n    With oTrip\n        .Save\n        .Name = \"n\"\n    End With\nEnd Sub\n";
    let s = session(&[("clsTrip.cls", CLS_TRIP), ("modMain.bas", main)]);
    assert_eq!(s.declaration_at(1, at(main, ".Save", "Save")), Some(proc(0, 0)));
    assert_eq!(s.declaration_at(1, at(main, ".Name", "Name")), Some(DeclId::ModuleVar { module: 0, idx: 0 }));
}

#[test]
fn nested_with_uses_the_innermost_object() {
    let main = "Sub Run()\n    Dim oRoute As clsRoute\n    Dim oTrip As clsTrip\n    With oRoute\n        .Plan\n\
                \x20       With oTrip\n            .Save\n        End With\n        .Plan\n    End With\nEnd Sub\n";
    let s = session(&[("clsTrip.cls", CLS_TRIP), ("clsRoute.cls", CLS_ROUTE), ("modMain.bas", main)]);
    // Outer object is clsRoute (module 1), inner is clsTrip (module 0).
    assert_eq!(s.declaration_at(2, nth(main, "Plan", 0) ), Some(proc(1, 0)));
    assert_eq!(s.declaration_at(2, at(main, ".Save", "Save")), Some(proc(0, 0)));
    // After the inner block ends, the outer object is current again.
    assert_eq!(s.declaration_at(2, nth(main, "Plan", 1)), Some(proc(1, 0)));
}

#[test]
fn with_over_a_member_chain() {
    let main = "Sub Run()\n    Dim oRoute As clsRoute\n    With oRoute.Trip\n        .Save\n    End With\nEnd Sub\n";
    let s = session(&[("clsTrip.cls", CLS_TRIP), ("clsRoute.cls", CLS_ROUTE), ("modMain.bas", main)]);
    assert_eq!(s.declaration_at(2, at(main, ".Save", "Save")), Some(proc(0, 0)));
}

#[test]
fn with_new_object() {
    let main = "Sub Run()\n    With New clsTrip\n        .Save\n    End With\nEnd Sub\n";
    let s = session(&[("clsTrip.cls", CLS_TRIP), ("modMain.bas", main)]);
    assert_eq!(s.declaration_at(1, at(main, ".Save", "Save")), Some(proc(0, 0)));
}

#[test]
fn with_function_result() {
    let main = "Function GetTrip() As clsTrip\nEnd Function\nSub Run()\n    With GetTrip()\n        .Save\n    End With\nEnd Sub\n";
    let s = session(&[("clsTrip.cls", CLS_TRIP), ("modMain.bas", main)]);
    assert_eq!(s.declaration_at(1, at(main, ".Save", "Save")), Some(proc(0, 0)));
}

#[test]
fn chained_members_after_a_with_member() {
    let main = "Sub Run()\n    Dim oRoute As clsRoute\n    With oRoute\n        .Trip.Save\n    End With\nEnd Sub\n";
    let s = session(&[("clsTrip.cls", CLS_TRIP), ("clsRoute.cls", CLS_ROUTE), ("modMain.bas", main)]);
    assert_eq!(s.declaration_at(2, at(main, ".Trip.Save", "Trip")), Some(DeclId::ModuleVar { module: 1, idx: 0 }));
    assert_eq!(s.declaration_at(2, at(main, ".Trip.Save", "Save")), Some(proc(0, 0)));
}

#[test]
fn with_over_a_late_bound_object_is_late_bound() {
    let main = "Sub Run()\n    Dim o As Object\n    With o\n        .Save\n    End With\nEnd Sub\n";
    let s = session(&[("clsTrip.cls", CLS_TRIP), ("modMain.bas", main)]);
    let site = at(main, ".Save", "Save");
    assert_eq!(s.declaration_at(1, site), None);
    assert_eq!(s.member_status_at(1, site), Some(MemberStatus::LateBound));
}

#[test]
fn leading_dot_outside_with_is_unresolved() {
    let main = "Sub Run()\n    .Save\nEnd Sub\n";
    let s = session(&[("clsTrip.cls", CLS_TRIP), ("modMain.bas", main)]);
    let site = at(main, ".Save", "Save");
    assert_eq!(s.declaration_at(1, site), None);
    assert_eq!(s.member_status_at(1, site), Some(MemberStatus::BaseUnresolved));
}

#[test]
fn parenthesised_new_takes_a_member() {
    let main = "Sub Run()\n    (New clsTrip).Save\nEnd Sub\n";
    let s = session(&[("clsTrip.cls", CLS_TRIP), ("modMain.bas", main)]);
    assert_eq!(s.declaration_at(1, at(main, ").Save", "Save")), Some(proc(0, 0)));
}

#[test]
fn with_bang_member_is_late_bound() {
    let main = "Sub Run()\n    Dim oTrip As clsTrip\n    With oTrip\n        x = !Key\n    End With\nEnd Sub\n";
    let s = session(&[("clsTrip.cls", CLS_TRIP), ("modMain.bas", main)]);
    assert_eq!(s.member_status_at(1, at(main, "!Key", "Key")), Some(MemberStatus::Bang));
}
