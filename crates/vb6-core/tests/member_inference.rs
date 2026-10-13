//! Procedure-local inference of the object type held by a `Variant`/`Object`
//! local: every `Set` assigns the same type (Inferred), different types
//! (Ambiguous), or something else may change the variable (late-bound).

mod common;

use common::{at, coded, session};
use vb6_core::sema::MemberStatus;
use vb6_core::session::DeclId;

const CLS_TRIP: &str = "Public Sub Save()\nEnd Sub\nPublic Property Get Parent() As clsTrip\nEnd Property\n";
const CLS_OTHER: &str = "Public Sub Save()\nEnd Sub\n";

fn run(body: &str) -> (vb6_core::session::Session, String) {
    let main = format!(
        "Function MakeTrip() As clsTrip\nEnd Function\nSub Take(ByRef t As Object)\nEnd Sub\nSub TakeVal(ByVal t As Object)\nEnd Sub\n\
         Sub Run(pObj As Object)\n    Dim o As Object\n    Dim v\n    Dim other As clsOther\n    Dim n As Long\n{body}\nEnd Sub\n"
    );
    let s = session(&[("clsTrip.cls", CLS_TRIP), ("clsOther.cls", CLS_OTHER), ("modMain.bas", &main)]);
    (s, main)
}

fn save(idx: usize) -> Option<DeclId> {
    Some(DeclId::Proc { module: idx, idx: 0 })
}

#[test]
fn single_assignment_infers_the_type() {
    let (s, main) = run("    Set o = New clsTrip\n    o.Save");
    let site = at(&main, "o.Save", "Save");
    assert_eq!(s.declaration_at(2, site), save(0));
    assert_eq!(s.member_status_at(2, site), Some(MemberStatus::Inferred));
    // The member's use is a reference to the declaration.
    let refs = s.references(0, at(CLS_TRIP, "Sub Save", "Save"), false);
    assert_eq!(refs.len(), 1);
    assert_eq!(refs[0].span.start, site);
}

#[test]
fn untyped_local_is_inferred_too() {
    let (s, main) = run("    Set v = MakeTrip()\n    v.Save");
    assert_eq!(s.declaration_at(2, at(&main, "v.Save", "Save")), save(0));
    assert_eq!(s.member_status_at(2, at(&main, "v.Save", "Save")), Some(MemberStatus::Inferred));
}

#[test]
fn several_assignments_of_one_type_agree() {
    let (s, main) = run("    Set o = New clsTrip\n    Set o = MakeTrip()\n    Set o = Nothing\n    o.Save");
    assert_eq!(s.member_status_at(2, at(&main, "o.Save", "Save")), Some(MemberStatus::Inferred));
}

#[test]
fn different_types_are_ambiguous_with_their_candidates() {
    let (s, main) = run("    Set o = New clsTrip\n    Set o = New clsOther\n    o.Save");
    let site = at(&main, "o.Save", "Save");
    assert_eq!(s.declaration_at(2, site), None);
    assert_eq!(
        s.member_status_at(2, site),
        Some(MemberStatus::Ambiguous(vec!["clsTrip".to_string(), "clsOther".to_string()]))
    );
    assert_eq!(s.diagnostics(2), vec![]);
}

#[test]
fn an_assignment_of_unknown_type_leaves_it_late_bound() {
    let (s, main) = run("    Set o = New clsTrip\n    Set o = CreateObject(\"X.Y\")\n    o.Save");
    assert_eq!(s.member_status_at(2, at(&main, "o.Save", "Save")), Some(MemberStatus::LateBound));
}

#[test]
fn assignment_from_another_late_bound_value_leaves_it_late_bound() {
    let (s, main) = run("    Set o = New clsTrip\n    Set o = pObj\n    o.Save");
    assert_eq!(s.member_status_at(2, at(&main, "o.Save", "Save")), Some(MemberStatus::LateBound));
}

#[test]
fn passing_by_reference_blocks_the_inference() {
    let (s, main) = run("    Set o = New clsTrip\n    Take o\n    o.Save");
    assert_eq!(s.member_status_at(2, at(&main, "o.Save", "Save")), Some(MemberStatus::LateBound));
}

#[test]
fn passing_by_value_or_to_a_builtin_keeps_it() {
    let (s, main) = run("    Set o = New clsTrip\n    TakeVal o\n    n = Len(TypeName(o))\n    o.Save");
    assert_eq!(s.member_status_at(2, at(&main, "o.Save", "Save")), Some(MemberStatus::Inferred));
    // `Take (o)` parses to the same call as `Take(o)`, which passes by
    // reference, so the parentheses are not taken as passing by value.
    let (s, main) = run("    Set o = New clsTrip\n    Take (o)\n    o.Save");
    assert_eq!(s.member_status_at(2, at(&main, "o.Save", "Save")), Some(MemberStatus::LateBound));
}

#[test]
fn for_each_over_the_variable_blocks_the_inference() {
    let (s, main) = run("    Set o = New clsTrip\n    For Each o In pObj\n    Next\n    o.Save");
    assert_eq!(s.member_status_at(2, at(&main, "o.Save", "Save")), Some(MemberStatus::LateBound));
}

#[test]
fn let_of_a_value_is_harmless_and_let_of_an_object_blocks() {
    let (s, main) = run("    Set v = New clsTrip\n    v = 5\n    v.Save");
    assert_eq!(s.member_status_at(2, at(&main, "v.Save", "Save")), Some(MemberStatus::Inferred));
    let (s, main) = run("    Set v = New clsTrip\n    v = other\n    v.Save");
    assert_eq!(s.member_status_at(2, at(&main, "v.Save", "Save")), Some(MemberStatus::LateBound));
}

#[test]
fn parameters_and_module_variables_are_not_inferred() {
    let main = "Private mObj As Object\nSub Run(pObj As Object)\n    Set pObj = New clsTrip\n    pObj.Save\n    Set mObj = New clsTrip\n    mObj.Save\nEnd Sub\n";
    let s = session(&[("clsTrip.cls", CLS_TRIP), ("modMain.bas", main)]);
    assert_eq!(s.member_status_at(1, at(main, "pObj.Save", "Save")), Some(MemberStatus::LateBound));
    assert_eq!(s.member_status_at(1, at(main, "mObj.Save", "Save")), Some(MemberStatus::LateBound));
}

#[test]
fn members_of_an_inferred_value_are_inferred_through_the_chain() {
    let (s, main) = run("    Set o = New clsTrip\n    o.Parent.Save");
    assert_eq!(s.member_status_at(2, at(&main, "o.Parent.Save", "Parent")), Some(MemberStatus::Inferred));
    assert_eq!(s.member_status_at(2, at(&main, "o.Parent.Save", "Save")), Some(MemberStatus::Inferred));
}

#[test]
fn a_missing_member_of_an_inferred_type_is_reported_once() {
    let (s, main) = run("    Set o = New clsTrip\n    o.Nope");
    let site = at(&main, "o.Nope", "Nope");
    assert_eq!(s.member_status_at(2, site), Some(MemberStatus::NoSuchMember));
    assert_eq!(coded(&s, &main, 2, 461).len(), 1);
}

#[test]
fn use_before_the_set_in_a_loop_is_inferred_flow_insensitively() {
    let (s, main) = run("    Do\n        If Not o Is Nothing Then o.Save\n        Set o = New clsTrip\n    Loop Until n = 1");
    assert_eq!(s.member_status_at(2, at(&main, "o.Save", "Save")), Some(MemberStatus::Inferred));
}

#[test]
fn summary_counts_inferred_and_ambiguous() {
    let (s, _) = run("    Set o = New clsTrip\n    o.Save\n    Set v = New clsTrip\n    Set v = New clsOther\n    v.Save");
    let sum = s.member_access_summary();
    assert_eq!((sum.inferred, sum.ambiguous), (1, 1));
}
