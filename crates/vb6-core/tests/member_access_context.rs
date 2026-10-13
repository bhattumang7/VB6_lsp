//! The accessor of a property member follows the statement it appears in:
//! `Let` for every statement that stores into the member, `Set` for
//! `Set … =`, `Get` for everything else (including a by-reference argument,
//! whose callee may or may not write: the read is what the site does).

mod common;

use common::{at, session};
use vb6_core::session::DeclId;

// Let Title (0), Get Title (1), Set Owner (2), Get Owner (3).
const CLS: &str = "Public Property Let Title(ByVal v As String)\nEnd Property\n\
Public Property Get Title() As String\nEnd Property\n\
Public Property Set Owner(ByVal o As clsC)\nEnd Property\n\
Public Property Get Owner() As clsC\nEnd Property\n";

const LET: usize = 0;
const GET: usize = 1;
const SET: usize = 2;
const GET_OWNER: usize = 3;

fn accessor(statement: &str, member: &str) -> Option<DeclId> {
    let main = format!(
        "Sub Run(o As clsC, c As Collection, n As Long)\n    Dim s As String\n    Dim other As clsC\n    {statement}\nEnd Sub\nSub Take(ByRef t As String)\nEnd Sub\n"
    );
    let s = session(&[("clsC.cls", CLS), ("modMain.bas", &main)]);
    s.declaration_at(1, at(&main, statement, member))
}

fn proc(idx: usize) -> Option<DeclId> {
    Some(DeclId::Proc { module: 0, idx })
}

#[test]
fn value_assignment_is_let() {
    assert_eq!(accessor("o.Title = \"t\"", "Title"), proc(LET));
}

#[test]
fn right_hand_side_is_get_and_left_is_let() {
    let main = "Sub Run(o As clsC)\n    o.Title = o.Title & \"x\"\nEnd Sub\n";
    let s = session(&[("clsC.cls", CLS), ("modMain.bas", main)]);
    assert_eq!(s.declaration_at(1, at(main, "o.Title = ", "Title")), proc(LET));
    assert_eq!(s.declaration_at(1, at(main, "= o.Title &", "Title")), proc(GET));
}

#[test]
fn set_assignment_is_set_and_its_value_is_get() {
    assert_eq!(accessor("Set o.Owner = other", "Owner"), proc(SET));
    assert_eq!(accessor("Set other = o.Owner", "Owner"), proc(GET_OWNER));
    assert_eq!(accessor("Set o.Owner = Nothing", "Owner"), proc(SET));
}

#[test]
fn mid_statement_target_is_let() {
    assert_eq!(accessor("Mid(o.Title, 1, 2) = \"ab\"", "Title"), proc(LET));
}

#[test]
fn lset_and_rset_targets_are_let() {
    assert_eq!(accessor("LSet o.Title = \"ab\"", "Title"), proc(LET));
    assert_eq!(accessor("RSet o.Title = \"ab\"", "Title"), proc(LET));
}

#[test]
fn input_statements_store_into_their_targets() {
    assert_eq!(accessor("Input #1, o.Title", "Title"), proc(LET));
    assert_eq!(accessor("Line Input #1, o.Title", "Title"), proc(LET));
    assert_eq!(accessor("Get #1, , o.Title", "Title"), proc(LET));
}

#[test]
fn for_loop_variable_is_let() {
    assert_eq!(accessor("For o.Title = 1 To 3: Next", "Title"), proc(LET));
}

#[test]
fn reads_are_get() {
    assert_eq!(accessor("s = o.Title", "Title"), proc(GET));
    assert_eq!(accessor("n = Len(o.Title)", "Title"), proc(GET));
    assert_eq!(accessor("s = UCase(o.Title) & \"x\"", "Title"), proc(GET));
    assert_eq!(accessor("If o.Title = \"x\" Then n = 1", "Title"), proc(GET));
}

#[test]
fn byref_argument_is_a_read() {
    assert_eq!(accessor("Take o.Title", "Title"), proc(GET));
    assert_eq!(accessor("Take (o.Title)", "Title"), proc(GET));
}

#[test]
fn only_the_outermost_member_of_a_target_is_stored_into() {
    // `o.Owner.Title = x` reads Owner (Get) and assigns Title (Let).
    let main = "Sub Run(o As clsC)\n    o.Owner.Title = \"x\"\nEnd Sub\n";
    let s = session(&[("clsC.cls", CLS), ("modMain.bas", main)]);
    assert_eq!(s.declaration_at(1, at(main, "o.Owner.Title", "Owner")), proc(GET_OWNER));
    assert_eq!(s.declaration_at(1, at(main, "o.Owner.Title", "Title")), proc(LET));
}

#[test]
fn bare_property_name_inside_the_class_follows_the_context() {
    let cls = "Private mTitle As String\nPublic Property Let Title(ByVal v As String)\n    mTitle = v\nEnd Property\n\
               Public Property Get Title() As String\n    Title = mTitle\nEnd Property\n\
               Sub Use()\n    Dim s As String\n    Title = \"t\"\n    s = Title\nEnd Sub\n";
    let s = session(&[("clsC.cls", cls)]);
    assert_eq!(s.declaration_at(0, at(cls, "Title = \"t\"", "Title")), proc(0));
    assert_eq!(s.declaration_at(0, at(cls, "s = Title", "Title")), proc(1));
}
