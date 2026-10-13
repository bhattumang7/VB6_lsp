//! Fields of user-defined types resolve to the field declaration in the
//! `Type` block, in the same module or a public `Type` of another module.

mod common;

use common::{at, session, sites};
use vb6_core::sema::MemberStatus;
use vb6_core::session::DeclId;

fn field(module: usize, type_idx: usize, member_idx: usize) -> DeclId {
    DeclId::TypeMember { module, type_idx, member_idx }
}

#[test]
fn field_of_a_same_module_type() {
    let src = "Private Type TAction\n    ID As Long\n    Name As String\nEnd Type\nSub Run()\n    Dim oAction As TAction\n    oAction.ID = 1\n    x = oAction.Name\nEnd Sub\n";
    let s = session(&[("modMain.bas", src)]);
    assert_eq!(s.declaration_at(0, at(src, "oAction.ID", "ID")), Some(field(0, 0, 0)));
    assert_eq!(s.declaration_at(0, at(src, "oAction.Name", "Name")), Some(field(0, 0, 1)));
    let def = s.definition(0, at(src, "oAction.ID", "ID")).unwrap();
    assert_eq!(def.span.start, at(src, "    ID As", "ID"));
    let refs = s.references(0, at(src, "    ID As", "ID"), false);
    assert_eq!(sites(&[src], &refs), vec![(0, at(src, "oAction.ID", "ID"), "ID".to_string())]);
}

#[test]
fn field_of_a_public_type_in_another_module() {
    let types = "Public Type TAction\n    ID As Long\n    Name As String\nEnd Type\n";
    let main = "Sub Run()\n    Dim oAction As TAction\n    oAction.Name = \"n\"\nEnd Sub\n";
    let s = session(&[("modTypes.bas", types), ("modMain.bas", main)]);
    assert_eq!(s.declaration_at(1, at(main, "oAction.Name", "Name")), Some(field(0, 0, 1)));
    let def = s.definition(1, at(main, "oAction.Name", "Name")).unwrap();
    assert_eq!(def.module, 0);
    assert_eq!(def.span.start, at(types, "    Name As", "Name"));
}

#[test]
fn private_type_of_another_module_is_not_visible() {
    let types = "Private Type TAction\n    ID As Long\nEnd Type\n";
    let main = "Sub Run()\n    Dim oAction As TAction\n    oAction.ID = 1\nEnd Sub\n";
    let s = session(&[("modTypes.bas", types), ("modMain.bas", main)]);
    let site = at(main, "oAction.ID", "ID");
    assert_eq!(s.declaration_at(1, site), None);
    assert_eq!(s.member_status_at(1, site), Some(MemberStatus::UnknownType("taction".to_string())));
}

#[test]
fn missing_field_of_a_known_type_is_reported() {
    let src = "Private Type TAction\n    ID As Long\nEnd Type\nSub Run()\n    Dim a As TAction\n    a.Nope = 1\nEnd Sub\n";
    let s = session(&[("modMain.bas", src)]);
    let site = at(src, "a.Nope", "Nope");
    assert_eq!(s.member_status_at(0, site), Some(MemberStatus::NoSuchMember));
    assert_eq!(common::coded(&s, src, 0, 461), vec![(site, "Nope".to_string())]);
}

#[test]
fn nested_type_fields() {
    let src = "Private Type TPoint\n    X As Long\nEnd Type\nPrivate Type TLine\n    A As TPoint\n    B As TPoint\nEnd Type\n\
               Sub Run()\n    Dim l As TLine\n    l.A.X = 1\n    l.B.X = 2\nEnd Sub\n";
    let s = session(&[("modMain.bas", src)]);
    assert_eq!(s.declaration_at(0, at(src, "l.A.X", "A")), Some(field(0, 1, 0)));
    assert_eq!(s.declaration_at(0, at(src, "l.A.X", "X")), Some(field(0, 0, 0)));
    assert_eq!(s.declaration_at(0, at(src, "l.B.X", "B")), Some(field(0, 1, 1)));
    let refs = s.references(0, at(src, "    X As", "X"), false);
    assert_eq!(refs.len(), 2);
}

#[test]
fn array_of_types_element_field() {
    let src = "Private Type TPoint\n    X As Long\nEnd Type\nSub Run()\n    Dim p(3) As TPoint\n    p(1).X = 5\nEnd Sub\n";
    let s = session(&[("modMain.bas", src)]);
    assert_eq!(s.declaration_at(0, at(src, "p(1).X", "X")), Some(field(0, 0, 0)));
}

#[test]
fn array_field_of_a_type_is_indexed() {
    let src = "Private Type TBag\n    Items(5) As TPoint\nEnd Type\nPrivate Type TPoint\n    X As Long\nEnd Type\n\
               Sub Run()\n    Dim b As TBag\n    b.Items(2).X = 1\nEnd Sub\n";
    let s = session(&[("modMain.bas", src)]);
    assert_eq!(s.declaration_at(0, at(src, "b.Items(2).X", "Items")), Some(field(0, 0, 0)));
    assert_eq!(s.declaration_at(0, at(src, "b.Items(2).X", "X")), Some(field(0, 1, 0)));
}

#[test]
fn redim_of_a_member_array_is_a_use_of_the_member() {
    let src = "Private Type TBag\n    Items() As Long\nEnd Type\nSub Run()\n    Dim b As TBag\n    ReDim b.Items(5)\nEnd Sub\n";
    let s = session(&[("modMain.bas", src)]);
    assert_eq!(s.declaration_at(0, at(src, "ReDim b.Items", "Items")), Some(field(0, 0, 0)));
}

#[test]
fn redim_of_a_plain_array_is_a_use_of_the_variable() {
    let src = "Sub Run()\n    Dim a() As Long\n    ReDim a(5)\nEnd Sub\n";
    let s = session(&[("modMain.bas", src)]);
    assert_eq!(
        s.declaration_at(0, at(src, "ReDim a(5)", "a")),
        Some(DeclId::Local { module: 0, proc: 0, idx: 0 })
    );
}

#[test]
fn udt_variable_of_a_class_field() {
    let cls = "Private Type TInfo\n    Code As Long\nEnd Type\nPublic Info As TInfo\n";
    let main = "Sub Run()\n    Dim o As clsC\n    o.Info.Code = 1\nEnd Sub\n";
    let s = session(&[("clsC.cls", cls), ("modMain.bas", main)]);
    // `Info` is public; its private type is not nameable here, but its
    // fields are reached through the variable.
    assert_eq!(s.declaration_at(1, at(main, "o.Info.Code", "Info")), Some(DeclId::ModuleVar { module: 0, idx: 0 }));
    assert_eq!(s.declaration_at(1, at(main, "o.Info.Code", "Code")), Some(field(0, 0, 0)));
}
