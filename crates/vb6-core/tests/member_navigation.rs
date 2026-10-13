//! End to end on a multi-module session: references, definition, hover,
//! document highlights, rename and call hierarchy see member sites.

mod common;

use common::{at, nth, session, sites, text};
use vb6_core::session::{DeclId, Location};

const CLS_TRIP: &str = "Attribute VB_Name = \"clsTrip\"\n\
Public Sub Save()\nEnd Sub\n\
Public Property Get Title() As String\nEnd Property\n";

const MAIN: &str = "Sub Run()\n\
    Dim oTrip As clsTrip\n\
    Dim s As String\n\
    Set oTrip = New clsTrip\n\
    oTrip.Save\n\
    s = oTrip.Title\n\
End Sub\n\
Sub Other()\n\
    Dim oT As clsTrip\n\
    Dim p As Long\n\
    Run\n\
    oT.Save\n\
    p = AddressOf Run\n\
End Sub\n";

fn world() -> vb6_core::session::Session {
    session(&[("clsTrip.cls", CLS_TRIP), ("modMain.bas", MAIN)])
}

const SRCS: [&str; 2] = [CLS_TRIP, MAIN];

#[test]
fn references_from_the_declaration_and_from_a_use_agree() {
    let s = world();
    let decl = at(CLS_TRIP, "Sub Save", "Save");
    let want = vec![
        (1, at(MAIN, "oTrip.Save", "Save"), "Save".to_string()),
        (1, at(MAIN, "oT.Save", "Save"), "Save".to_string()),
    ];
    assert_eq!(sites(&SRCS, &s.references(0, decl, false)), want);
    assert_eq!(sites(&SRCS, &s.references(1, at(MAIN, "oTrip.Save", "Save"), false)), want);
    assert_eq!(sites(&SRCS, &s.references(1, at(MAIN, "oT.Save", "Save"), false)), want);
    // With the declaration, in project order.
    let mut with_decl = vec![(0, decl, "Save".to_string())];
    with_decl.extend(want);
    assert_eq!(sites(&SRCS, &s.references(0, decl, true)), with_decl);
}

#[test]
fn definition_of_a_member_use_is_the_member_declaration() {
    let s = world();
    let def = s.definition(1, at(MAIN, "oTrip.Save", "Save")).unwrap();
    assert_eq!(def, Location { module: 0, span: def.span });
    assert_eq!(def.span.start, at(CLS_TRIP, "Sub Save", "Save"));
    assert_eq!(text(CLS_TRIP, def), "Save");
    let title = s.definition(1, at(MAIN, "oTrip.Title", "Title")).unwrap();
    assert_eq!(title.module, 0);
    assert_eq!(title.span.start, at(CLS_TRIP, "Get Title", "Title"));
}

#[test]
fn hover_on_a_member_use_shows_the_declaration() {
    let s = world();
    let site = at(MAIN, "oTrip.Save", "Save");
    let h = s.hover(1, site).unwrap();
    assert_eq!(h.text, "Public Sub Save()");
    assert_eq!((h.span.start, h.span.len), (site, 4));
    let t = s.hover(1, at(MAIN, "oTrip.Title", "Title")).unwrap();
    assert_eq!(t.text, "Public Property Get Title() As String");
}

#[test]
fn document_highlights_cover_member_uses_in_the_module() {
    let s = world();
    let spans = s.document_highlights(1, at(MAIN, "oTrip.Save", "Save"));
    let starts: Vec<u32> = spans.iter().map(|sp| sp.start).collect();
    assert_eq!(starts, vec![at(MAIN, "oTrip.Save", "Save"), at(MAIN, "oT.Save", "Save")]);
}

#[test]
fn rename_a_class_member_edits_declaration_and_member_uses() {
    let s = world();
    let edits = s.rename(1, at(MAIN, "oTrip.Save", "Save"), "Persist");
    let got: Vec<(usize, u32, &str)> =
        edits.iter().map(|e| (e.module, e.span.start, e.new_text.as_str())).collect();
    assert_eq!(
        got,
        vec![
            (0, at(CLS_TRIP, "Sub Save", "Save"), "Persist"),
            (1, at(MAIN, "oTrip.Save", "Save"), "Persist"),
            (1, at(MAIN, "oT.Save", "Save"), "Persist"),
        ]
    );
    assert!(edits.iter().all(|e| e.span.len == 4));
}

#[test]
fn outgoing_calls_include_member_calls_and_property_reads() {
    let s = world();
    // `Run` is procedure 0 of module 1.
    let out = s.outgoing_calls_from(1, 0);
    let got: Vec<(String, usize, u32, Vec<(usize, u32)>)> = out
        .iter()
        .map(|c| {
            (
                c.callee.name.clone(),
                c.callee.location.module,
                c.callee.location.span.start,
                c.call_sites.iter().map(|l| (l.module, l.span.start)).collect(),
            )
        })
        .collect();
    assert_eq!(
        got,
        vec![
            (
                "Save".to_string(),
                0,
                at(CLS_TRIP, "Sub Save", "Save"),
                vec![(1, at(MAIN, "oTrip.Save", "Save"))]
            ),
            (
                "Title".to_string(),
                0,
                at(CLS_TRIP, "Get Title", "Title"),
                vec![(1, at(MAIN, "oTrip.Title", "Title"))]
            ),
        ]
    );
}

#[test]
fn outgoing_calls_exclude_address_of_operands() {
    let s = world();
    // `Other` is procedure 1 of module 1: a call to Run, a call to Save, and
    // `AddressOf Run`, which is not a call.
    let out = s.outgoing_calls_from(1, 1);
    let names: Vec<(&str, usize)> =
        out.iter().map(|c| (c.callee.name.as_str(), c.call_sites.len())).collect();
    assert_eq!(names, vec![("Save", 1), ("Run", 1)]);
    let run = out.iter().find(|c| c.callee.name == "Run").unwrap();
    assert_eq!(run.call_sites[0].span.start, nth(MAIN, "Run\n", 0));
}

#[test]
fn incoming_calls_name_each_caller_and_its_member_sites() {
    let s = world();
    // `Save` is procedure 0 of module 0.
    let incoming = s.incoming_calls_to(0, 0);
    let got: Vec<(String, Vec<u32>)> = incoming
        .iter()
        .map(|c| (c.caller.name.clone(), c.call_sites.iter().map(|l| l.span.start).collect()))
        .collect();
    assert_eq!(
        got,
        vec![
            ("Run".to_string(), vec![at(MAIN, "oTrip.Save", "Save")]),
            ("Other".to_string(), vec![at(MAIN, "oT.Save", "Save")]),
        ]
    );
    // Run is called once (by Other); its AddressOf use is not a call.
    let run = s.incoming_calls_to(1, 0);
    assert_eq!(run.len(), 1);
    assert_eq!(run[0].caller.name, "Other");
    assert_eq!(run[0].call_sites.len(), 1);
}

#[test]
fn call_hierarchy_by_cursor_selects_the_declaration_not_a_namesake() {
    // Two classes declare `Save`; the cursor picks the declaration.
    let other = "Public Sub Save()\nEnd Sub\n";
    let main = "Sub Run()\n    Dim a As clsA\n    Dim b As clsB\n    a.Save\n    b.Save\nEnd Sub\n";
    let s = session(&[("clsA.cls", other), ("clsB.cls", other), ("modMain.bas", main)]);
    let a = s.incoming_calls_at(0, at(other, "Sub Save", "Save"));
    let b = s.incoming_calls_at(1, at(other, "Sub Save", "Save"));
    assert_eq!(a.len(), 1);
    assert_eq!(b.len(), 1);
    assert_eq!(a[0].call_sites[0].span.start, at(main, "a.Save", "Save"));
    assert_eq!(b[0].call_sites[0].span.start, at(main, "b.Save", "Save"));
    let out = s.outgoing_calls_at(2, at(main, "Sub Run", "Run"));
    let callees: Vec<(usize, String)> =
        out.iter().map(|c| (c.callee.location.module, c.callee.name.clone())).collect();
    assert_eq!(callees, vec![(0, "Save".to_string()), (1, "Save".to_string())]);
}

#[test]
fn declaration_at_reports_the_resolved_declaration() {
    let s = world();
    assert_eq!(s.declaration_at(1, at(MAIN, "oTrip.Save", "Save")), Some(DeclId::Proc { module: 0, idx: 0 }));
    assert_eq!(s.declaration_at(0, at(CLS_TRIP, "Sub Save", "Save")), Some(DeclId::Proc { module: 0, idx: 0 }));
}
