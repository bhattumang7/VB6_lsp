//! The base of each member access, the stores into each name and the
//! arguments of each call: what a host follows to learn which object a
//! late-bound member access runs on.

mod common;

use common::{at, nth, session};
use vb6_core::frontend::ast::Span;
use vb6_core::session::{Argument, DeclId, Receiver, StaticType, Store, StoreKind};

const CLS_PART: &str =
    "Public Sub Save()\nEnd Sub\nPublic Property Get Items() As Collection\nEnd Property\n";
const MAIN: &str = "Private mPart As Object\n\
Function Make() As clsPart\nEnd Function\n\
Sub Keep(ByVal t As Object, Optional n As Long)\n    t.Save\nEnd Sub\n\
Sub Run()\n    Dim o As Object\n    Dim x As clsPart\n    Set mPart = New clsPart\n    Set o = Make()\n    mPart.Save\n    x.Items(1).Save\n    With mPart\n        .Save\n    End With\n    Keep x, n:=3\n    Set o = Nothing\nEnd Sub\n";

fn sp(start: u32, len: u32) -> Span {
    Span { start, len }
}

fn fixture() -> vb6_core::session::Session {
    session(&[("clsPart.cls", CLS_PART), ("modMain.bas", MAIN)])
}

#[test]
fn receivers_name_the_declaration_holding_the_base() {
    let s = fixture();
    let got = s.member_receivers(1);
    let m_part = at(MAIN, "    mPart.Save", "mPart");
    let items = at(MAIN, "x.Items(1)", "Items");
    let with_obj = at(MAIN, "With mPart", "mPart");
    assert_eq!(
        got,
        vec![
            Receiver {
                member: sp(at(MAIN, "t.Save", "Save"), 4),
                head: Some(sp(at(MAIN, "t.Save", "t"), 1)),
                through_call: false, root: None,
                base_type: StaticType::LateBound,
            },
            Receiver {
                member: sp(at(MAIN, "    mPart.Save", "Save"), 4),
                head: Some(sp(m_part, 5)),
                through_call: false, root: None,
                base_type: StaticType::LateBound
            },
            Receiver {
                member: sp(items, 5),
                head: Some(sp(at(MAIN, "x.Items(1)", "x"), 1)),
                through_call: false, root: None,
                base_type: StaticType::Class("clsPart".into()),
            },
            Receiver {
                member: sp(at(MAIN, "x.Items(1).Save", "Save"), 4),
                head: Some(sp(items, 5)),
                through_call: true, root: Some(sp(at(MAIN, "x.Items(1)", "x"), 1)),
                base_type: StaticType::LateBound,
            },
            Receiver {
                member: sp(at(MAIN, "        .Save", "Save"), 4),
                head: Some(sp(with_obj, 5)),
                through_call: false, root: None,
                base_type: StaticType::LateBound
            },
        ]
    );
    // The heads lead to the declarations the host joins on.
    assert_eq!(
        s.declaration_at(1, m_part),
        Some(DeclId::ModuleVar { module: 1, idx: 0 })
    );
    assert_eq!(
        s.declaration_at(1, with_obj),
        Some(DeclId::ModuleVar { module: 1, idx: 0 })
    );
    assert_eq!(
        s.declaration_at(1, at(MAIN, "t.Save", "t")),
        Some(DeclId::Param {
            module: 1,
            proc: 1,
            idx: 0
        })
    );
}

#[test]
fn stores_carry_the_static_type_of_the_value() {
    let s = fixture();
    assert_eq!(
        s.object_stores(1),
        vec![
            Store {
                kind: StoreKind::Set,
                target: sp(at(MAIN, "Set mPart", "mPart"), 5),
                value_head: None,
                value_type: StaticType::Class("clsPart".into()),
            },
            Store {
                kind: StoreKind::Set,
                target: sp(nth(MAIN, "Set o = ", 0) + 4, 1),
                value_head: Some(sp(at(MAIN, "= Make()", "Make"), 4)),
                value_type: StaticType::Class("clsPart".into()),
            },
            Store {
                kind: StoreKind::Set,
                target: sp(nth(MAIN, "Set o = ", 1) + 4, 1),
                value_head: None,
                value_type: StaticType::Nothing,
            },
        ]
    );
}

#[test]
fn for_each_stores_the_element_of_an_array() {
    let main = "Sub Run()\n    Dim a() As clsPart\n    Dim c As Collection\n    Dim v\n    For Each v In a\n    Next\n    For Each v In c\n    Next\nEnd Sub\n";
    let s = session(&[("clsPart.cls", CLS_PART), ("modMain.bas", main)]);
    assert_eq!(
        s.object_stores(1),
        vec![
            Store {
                kind: StoreKind::ForEach,
                target: sp(nth(main, "For Each v", 0) + 9, 1),
                value_head: Some(sp(at(main, "In a\n", "a"), 1)),
                value_type: StaticType::Class("clsPart".into()),
            },
            Store {
                kind: StoreKind::ForEach,
                target: sp(nth(main, "For Each v", 1) + 9, 1),
                value_head: Some(sp(at(main, "In c\n", "c"), 1)),
                value_type: StaticType::Unresolved,
            },
        ]
    );
}

#[test]
fn arguments_carry_position_name_and_type() {
    let s = fixture();
    let keep = at(MAIN, "    Keep x", "Keep");
    assert_eq!(
        s.call_arguments(1),
        vec![
            Argument {
                callee: sp(at(MAIN, "x.Items(1)", "Items"), 5),
                receiver: Some(sp(at(MAIN, "x.Items(1)", "x"), 1)),
                index: Some(0),
                name: None,
                value_head: None,
                value_type: StaticType::Value("Integer".into()),
            },
            Argument {
                callee: sp(keep, 4),
                receiver: None,
                index: Some(0),
                name: None,
                value_head: Some(sp(at(MAIN, "Keep x", "x"), 1)),
                value_type: StaticType::Class("clsPart".into())
            },
            Argument {
                callee: sp(keep, 4),
                receiver: None,
                index: None,
                name: Some("n".into()),
                value_head: None,
                value_type: StaticType::Value("Integer".into())
            },
        ]
    );
    assert_eq!(
        s.declaration_at(1, keep),
        Some(DeclId::Proc { module: 1, idx: 1 })
    );
}

#[test]
fn a_function_assigned_in_its_body_stores_into_its_return_slot() {
    let main = "Function Make() As Object\n    Set Make = New clsPart\nEnd Function\nSub Run()\n    Make().Save\nEnd Sub\n";
    let s = session(&[("clsPart.cls", CLS_PART), ("modMain.bas", main)]);
    let target = at(main, "Set Make", "Make");
    assert_eq!(
        s.object_stores(1),
        vec![Store { kind: StoreKind::Set, target: sp(target, 4), value_head: None, value_type: StaticType::Class("clsPart".into()) }]
    );
    // Inside the function its name is the return slot: a local of the procedure, not the procedure.
    assert_eq!(s.declaration_at(1, target), Some(DeclId::Local { module: 1, proc: 0, idx: 0 }));
    let call = at(main, "Make().Save", "Make");
    assert_eq!(
        s.member_receivers(1),
        vec![Receiver {
            member: sp(at(main, "Make().Save", "Save"), 4),
            head: Some(sp(call, 4)),
            through_call: true, root: None,
            base_type: StaticType::LateBound,
        }]
    );
    assert_eq!(s.declaration_at(1, call), Some(DeclId::Proc { module: 1, idx: 0 }));
}

#[test]
fn library_member_at_names_the_member_a_use_resolved_to() {
    use common::session_with;
    use vb6_core::session::{Library, SessionConfig};
    let libs = r#"[{"name":"OtLib","types":[{"name":"Cls","kind":"class","members":[{"name":"Save","kind":"method"}]}]}]"#;
    let main = "Sub Run(o As OtLib.Cls)\n    o.Save\nEnd Sub\n";
    let s = session_with(
        &[("modMain.bas", main)],
        SessionConfig { project_name: None, libraries: Library::from_json(libs).unwrap() },
    );
    assert_eq!(
        s.library_member_at(0, at(main, "o.Save", "Save")),
        Some(("otlib".to_string(), "Cls".to_string(), "save".to_string()))
    );
    assert_eq!(s.library_member_at(0, at(main, "Sub Run", "Run")), None);
}

#[test]
fn a_method_call_names_the_object_it_is_called_on() {
    let main = "Sub Run()\n    Dim c As Collection\n    Dim t As clsPart\n    c.Add t\nEnd Sub\n";
    let s = session(&[("clsPart.cls", CLS_PART), ("modMain.bas", main)]);
    let got = s.call_arguments(1);
    assert_eq!(got.len(), 1);
    assert_eq!(got[0].callee, sp(at(main, "c.Add", "Add"), 3));
    assert_eq!(got[0].receiver, Some(sp(at(main, "c.Add", "c"), 1)));
    assert_eq!(got[0].value_type, StaticType::Class("clsPart".into()));
    assert_eq!(s.library_member_at(1, got[0].callee.start).map(|m| (m.0, m.1, m.2)),
               Some(("vba".to_string(), "_Collection".to_string(), "add".to_string())));
}

#[test]
fn root_names_the_object_a_member_value_is_read_from() {
    let main = "Sub Run()\n    Dim d As Object\n    d.Item(\"k\").Save\n    d.Owner.Save\nEnd Sub\n";
    let s = session(&[("clsPart.cls", CLS_PART), ("modMain.bas", main)]);
    let got = s.member_receivers(1);
    let d = at(main, "d.Item", "d");
    assert_eq!(got.len(), 4);
    // d.Item("k").Save: member Item has root d; Save has base d.Item(...) whose root is d
    assert_eq!(got[1].root, Some(sp(d, 1)));
    assert_eq!(got[3].root, Some(sp(at(main, "d.Owner.Save", "d"), 1)));
}
