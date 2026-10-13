//! `obj.Member` resolves to the declaration it names when the declared type of
//! `obj` is known: variables, parameters, module variables and arrays of
//! them; visibility; the accessor chosen by context; late-bound bases.

mod common;

use common::{at, coded, nth, session, sites, text};
use vb6_core::sema::MemberStatus;
use vb6_core::session::{DeclId, Location};

const CLS_TRIP: &str = "\
Public Sub Save()
End Sub
Private Sub Hidden()
End Sub
Public Name As String
Public Property Let Title(ByVal v As String)
End Property
Public Property Get Title() As String
End Property
Public Property Set Owner(ByVal o As clsTrip)
End Property
Public Property Get Owner() As clsTrip
End Property
Friend Sub Audit()
End Sub
";

fn proc(module: usize, idx: usize) -> DeclId {
    DeclId::Proc { module, idx }
}

fn one(src: &str, start: u32, len: usize) -> (u32, String) {
    (start, src[start as usize..start as usize + len].to_string())
}

#[test]
fn local_variable_member_resolves_to_the_class_procedure() {
    let main = "Option Explicit\nSub Run()\n    Dim oTrip As clsTrip\n    Set oTrip = New clsTrip\n    oTrip.Save\nEnd Sub\n";
    let s = session(&[("clsTrip.cls", CLS_TRIP), ("modMain.bas", main)]);
    let site = at(main, "oTrip.Save", "Save");
    assert_eq!(s.declaration_at(1, site), Some(proc(0, 0)));
    let def = s.definition(1, site).unwrap();
    assert_eq!(def, Location { module: 0, span: def.span });
    assert_eq!(text(CLS_TRIP, def), "Save");
    assert_eq!(def.span.start, at(CLS_TRIP, "Sub Save", "Save"));
    // The declaration, queried by its own name, lists the call site.
    let refs = s.references(0, at(CLS_TRIP, "Sub Save", "Save"), false);
    assert_eq!(sites(&[CLS_TRIP, main], &refs), vec![(1, site, "Save".to_string())]);
    assert_eq!(s.diagnostics(1), vec![]);
}

#[test]
fn as_new_declaration_resolves_like_as() {
    let main = "Sub Run()\n    Dim oTrip As New clsTrip\n    oTrip.Save\nEnd Sub\n";
    let s = session(&[("clsTrip.cls", CLS_TRIP), ("modMain.bas", main)]);
    assert_eq!(s.declaration_at(1, at(main, "oTrip.Save", "Save")), Some(proc(0, 0)));
}

#[test]
fn parameter_base() {
    let main = "Sub Show(ByVal t As clsTrip)\n    t.Save\n    t.Name = \"x\"\nEnd Sub\n";
    let s = session(&[("clsTrip.cls", CLS_TRIP), ("modMain.bas", main)]);
    assert_eq!(s.declaration_at(1, at(main, "t.Save", "Save")), Some(proc(0, 0)));
    assert_eq!(
        s.declaration_at(1, at(main, "t.Name", "Name")),
        Some(DeclId::ModuleVar { module: 0, idx: 0 })
    );
}

#[test]
fn module_variable_base() {
    let main = "Private mTrip As clsTrip\nSub Run()\n    mTrip.Save\nEnd Sub\n";
    let s = session(&[("clsTrip.cls", CLS_TRIP), ("modMain.bas", main)]);
    assert_eq!(s.declaration_at(1, at(main, "mTrip.Save", "Save")), Some(proc(0, 0)));
}

#[test]
fn project_variable_of_another_module_is_a_base() {
    let globals = "Public gTrip As clsTrip\n";
    let main = "Sub Run()\n    gTrip.Save\nEnd Sub\n";
    let s = session(&[("clsTrip.cls", CLS_TRIP), ("modGlobals.bas", globals), ("modMain.bas", main)]);
    assert_eq!(s.declaration_at(2, at(main, "gTrip.Save", "gTrip")), Some(DeclId::ModuleVar { module: 1, idx: 0 }));
    assert_eq!(s.declaration_at(2, at(main, "gTrip.Save", "Save")), Some(proc(0, 0)));
}

#[test]
fn array_element_base() {
    let main = "Sub Run()\n    Dim aTrips(3) As clsTrip\n    Dim i As Long\n    aTrips(i).Save\n    aTrips(1).Name = \"n\"\nEnd Sub\n";
    let s = session(&[("clsTrip.cls", CLS_TRIP), ("modMain.bas", main)]);
    assert_eq!(s.declaration_at(1, at(main, "aTrips(i).Save", "Save")), Some(proc(0, 0)));
    assert_eq!(
        s.declaration_at(1, at(main, "aTrips(1).Name", "Name")),
        Some(DeclId::ModuleVar { module: 0, idx: 0 })
    );
}

#[test]
fn dynamic_array_parameter_element_base() {
    let main = "Sub Show(aTrips() As clsTrip)\n    aTrips(0).Save\nEnd Sub\n";
    let s = session(&[("clsTrip.cls", CLS_TRIP), ("modMain.bas", main)]);
    assert_eq!(s.declaration_at(1, at(main, "aTrips(0).Save", "Save")), Some(proc(0, 0)));
}

#[test]
fn private_member_is_not_reachable_and_is_reported() {
    let main = "Sub Run()\n    Dim oTrip As clsTrip\n    oTrip.Hidden\nEnd Sub\n";
    let s = session(&[("clsTrip.cls", CLS_TRIP), ("modMain.bas", main)]);
    let site = at(main, "oTrip.Hidden", "Hidden");
    assert_eq!(s.declaration_at(1, site), None);
    assert_eq!(s.member_status_at(1, site), Some(MemberStatus::NoSuchMember));
    assert_eq!(coded(&s, main, 1, 461), vec![one(main, site, 6)]);
    // The private declaration has no use sites.
    assert_eq!(s.references(0, at(CLS_TRIP, "Sub Hidden", "Hidden"), false), vec![]);
}

#[test]
fn missing_member_of_a_known_class_is_reported() {
    let main = "Sub Run()\n    Dim oTrip As clsTrip\n    oTrip.Nope\nEnd Sub\n";
    let s = session(&[("clsTrip.cls", CLS_TRIP), ("modMain.bas", main)]);
    let site = at(main, "oTrip.Nope", "Nope");
    assert_eq!(s.declaration_at(1, site), None);
    assert_eq!(coded(&s, main, 1, 461), vec![one(main, site, 4)]);
}

#[test]
fn friend_member_is_reachable_from_another_module() {
    let main = "Sub Run()\n    Dim oTrip As clsTrip\n    oTrip.Audit\nEnd Sub\n";
    let s = session(&[("clsTrip.cls", CLS_TRIP), ("modMain.bas", main)]);
    assert_eq!(s.declaration_at(1, at(main, "oTrip.Audit", "Audit")), Some(proc(0, 6)));
    assert_eq!(coded(&s, main, 1, 461), vec![]);
}

#[test]
fn property_accessor_follows_the_context() {
    let main = "Sub Run()\n    Dim oTrip As clsTrip\n    Dim oOwner As clsTrip\n    Dim s As String\n\
                \x20   oTrip.Title = \"t\"\n    s = oTrip.Title\n    Set oTrip.Owner = oOwner\n    Set oOwner = oTrip.Owner\nEnd Sub\n";
    let s = session(&[("clsTrip.cls", CLS_TRIP), ("modMain.bas", main)]);
    // Declared order: Let Title (2), Get Title (3), Set Owner (4), Get Owner (5).
    assert_eq!(s.declaration_at(1, at(main, "oTrip.Title = ", "Title")), Some(proc(0, 2)));
    assert_eq!(s.declaration_at(1, at(main, "s = oTrip.Title", "Title")), Some(proc(0, 3)));
    assert_eq!(s.declaration_at(1, at(main, "Set oTrip.Owner", "Owner")), Some(proc(0, 4)));
    assert_eq!(s.declaration_at(1, at(main, "= oTrip.Owner", "Owner")), Some(proc(0, 5)));
    // Every accessor is one symbol: a query on any of them lists all uses.
    let all = vec![
        (1, at(main, "oTrip.Title = ", "Title"), "Title".to_string()),
        (1, at(main, "s = oTrip.Title", "Title"), "Title".to_string()),
    ];
    let from_get = s.references(0, at(CLS_TRIP, "Property Get Title", "Title"), false);
    assert_eq!(sites(&[CLS_TRIP, main], &from_get), all);
    let from_let = s.references(0, at(CLS_TRIP, "Property Let Title", "Title"), false);
    assert_eq!(sites(&[CLS_TRIP, main], &from_let), all);
}

#[test]
fn assignment_to_a_get_only_property_names_the_missing_accessor() {
    let cls = "Public Property Get Count() As Long\nEnd Property\n";
    let main = "Sub Run()\n    Dim o As clsC\n    o.Count = 1\nEnd Sub\n";
    let s = session(&[("clsC.cls", cls), ("modMain.bas", main)]);
    let site = at(main, "o.Count", "Count");
    assert_eq!(s.declaration_at(1, site), None);
    assert_eq!(
        s.member_status_at(1, site),
        Some(MemberStatus::MissingAccessor(vb6_core::sema::Access::Let))
    );
    assert_eq!(coded(&s, main, 1, 461), vec![]);
}

#[test]
fn variant_base_is_late_bound_and_not_reported() {
    let main = "Option Explicit\nSub Run()\n    Dim v As Variant\n    v.Save\n    Dim o As Object\n    o.Save\n    Dim u\n    u.Save\nEnd Sub\n";
    let s = session(&[("clsTrip.cls", CLS_TRIP), ("modMain.bas", main)]);
    for ctx in ["v.Save", "o.Save", "u.Save"] {
        let site = at(main, ctx, "Save");
        assert_eq!(s.declaration_at(1, site), None, "{ctx}");
        assert_eq!(s.member_status_at(1, site), Some(MemberStatus::LateBound), "{ctx}");
    }
    assert_eq!(s.diagnostics(1), vec![]);
    assert_eq!(s.references(0, at(CLS_TRIP, "Sub Save", "Save"), false), vec![]);
}

#[test]
fn value_typed_base_is_not_an_object() {
    let main = "Sub Run()\n    Dim n As Long\n    n.Save\nEnd Sub\n";
    let s = session(&[("clsTrip.cls", CLS_TRIP), ("modMain.bas", main)]);
    let site = at(main, "n.Save", "Save");
    assert_eq!(s.declaration_at(1, site), None);
    assert_eq!(s.member_status_at(1, site), Some(MemberStatus::NotAnObject));
    assert_eq!(coded(&s, main, 1, 461), vec![]);
}

#[test]
fn class_that_is_not_in_the_session_is_left_unresolved() {
    let main = "Option Explicit\nSub Run()\n    Dim o As clsMissing\n    o.Save\nEnd Sub\n";
    let s = session(&[("clsTrip.cls", CLS_TRIP), ("modMain.bas", main)]);
    let site = at(main, "o.Save", "Save");
    assert_eq!(s.declaration_at(1, site), None);
    assert_eq!(s.member_status_at(1, site), Some(MemberStatus::UnknownType("clsmissing".to_string())));
    assert_eq!(s.diagnostics(1), vec![]);
}

#[test]
fn member_of_an_undeclared_name_is_not_a_second_diagnostic() {
    let main = "Option Explicit\nSub Run()\n    ghost.Save\nEnd Sub\n";
    let s = session(&[("modMain.bas", main)]);
    assert_eq!(s.member_status_at(0, at(main, "ghost.Save", "Save")), Some(MemberStatus::BaseUnresolved));
    // Exactly one diagnostic: the undeclared base.
    let diags = s.diagnostics(0);
    assert_eq!(diags.len(), 1);
    assert_eq!(diags[0].span.start, nth(main, "ghost", 0));
}

#[test]
fn me_member_resolves_in_the_class_and_private_is_not_reachable() {
    let cls = "Public Sub Refresh()\nEnd Sub\nPrivate Sub Helper()\nEnd Sub\nSub T()\n    Me.Refresh\n    Me.Helper\nEnd Sub\n";
    let s = session(&[("clsA.cls", cls)]);
    assert_eq!(s.declaration_at(0, at(cls, "Me.Refresh", "Refresh")), Some(proc(0, 0)));
    assert_eq!(s.declaration_at(0, at(cls, "Me.Helper", "Helper")), None);
    assert_eq!(s.member_status_at(0, at(cls, "Me.Helper", "Helper")), Some(MemberStatus::NoSuchMember));
}

#[test]
fn me_property_accessor_follows_the_context() {
    let cls = "Public Property Let P(ByVal v As Long)\nEnd Property\nPublic Property Get P() As Long\nEnd Property\n\
               Sub T()\n    Dim n As Long\n    Me.P = 1\n    n = Me.P\nEnd Sub\n";
    let s = session(&[("clsA.cls", cls)]);
    assert_eq!(s.declaration_at(0, at(cls, "Me.P = 1", "P")), Some(proc(0, 0)));
    assert_eq!(s.declaration_at(0, at(cls, "n = Me.P", "P")), Some(proc(0, 1)));
}

#[test]
fn cross_module_variable_keeps_its_declared_type() {
    let globals = "Public gCount As Long\n";
    let main = "Sub Run()\n    gCount.Save\nEnd Sub\n";
    let s = session(&[("clsTrip.cls", CLS_TRIP), ("modGlobals.bas", globals), ("modMain.bas", main)]);
    let site = at(main, "gCount.Save", "Save");
    assert_eq!(s.declaration_at(2, at(main, "gCount.Save", "gCount")), Some(DeclId::ModuleVar { module: 1, idx: 0 }));
    assert_eq!(s.declaration_at(2, site), None);
    assert_eq!(s.member_status_at(2, site), Some(MemberStatus::NotAnObject));
}
