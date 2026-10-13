//! External types: described by the host as plain data, resolved through
//! coclass and interface links, queried by `(library, type, member)`.

mod common;

use common::{at, session_with, sites};
use vb6_core::sema::MemberStatus;
use vb6_core::session::{
    Library, LibraryMember, LibraryMemberKind, LibraryType, LibraryTypeKind, SessionConfig,
};

fn member(name: &str, kind: LibraryMemberKind, ret: Option<&str>) -> LibraryMember {
    LibraryMember {
        name: name.to_string(),
        kind,
        return_type: ret.map(String::from),
        is_default: false,
        param_count: None,
    }
}

fn ado() -> Library {
    Library {
        name: "ADODB".to_string(),
        types: vec![
            LibraryType {
                name: "_Connection".into(),
                kind: LibraryTypeKind::Interface,
                default_interface: None,
                interfaces: vec![],
                extends: vec!["_ADO".into()],
                is_app_object: false,
                members: vec![
                    member("Open", LibraryMemberKind::Method, None),
                    member("Execute", LibraryMemberKind::Method, Some("Recordset")),
                    member("State", LibraryMemberKind::PropertyGet, Some("Long")),
                    member("ConnectionString", LibraryMemberKind::PropertyGet, Some("String")),
                    member("ConnectionString", LibraryMemberKind::PropertyLet, None),
                ],
            },
            LibraryType {
                name: "_ADO".into(),
                kind: LibraryTypeKind::Interface,
                default_interface: None,
                interfaces: vec![],
                extends: vec![],
                is_app_object: false,
                members: vec![member("Properties", LibraryMemberKind::PropertyGet, Some("Object"))],
            },
            LibraryType {
                name: "Connection".into(),
                kind: LibraryTypeKind::Coclass,
                default_interface: Some("_Connection".into()),
                interfaces: vec![],
                extends: vec![],
                is_app_object: false,
                members: vec![],
            },
            LibraryType {
                name: "Recordset".into(),
                kind: LibraryTypeKind::Class,
                default_interface: None,
                interfaces: vec![],
                extends: vec![],
                is_app_object: false,
                members: vec![
                    member("EOF", LibraryMemberKind::PropertyGet, Some("Boolean")),
                    member("MoveNext", LibraryMemberKind::Method, None),
                    LibraryMember { is_default: true, ..member("Fields", LibraryMemberKind::PropertyGet, Some("Field")) },
                ],
            },
            LibraryType {
                name: "Field".into(),
                kind: LibraryTypeKind::Class,
                default_interface: None,
                interfaces: vec![],
                extends: vec![],
                is_app_object: false,
                members: vec![member("Value", LibraryMemberKind::Field, Some("Variant"))],
            },
        ],
    }
}

fn app() -> Library {
    Library {
        name: "VB".to_string(),
        types: vec![LibraryType {
            name: "App".into(),
            kind: LibraryTypeKind::Class,
            default_interface: None,
            interfaces: vec![],
            extends: vec![],
            is_app_object: true,
            members: vec![member("Path", LibraryMemberKind::PropertyGet, Some("String"))],
        }],
    }
}

fn config() -> SessionConfig {
    SessionConfig { project_name: None, libraries: vec![ado(), app()] }
}

#[test]
fn json_round_trip() {
    let libs = vec![ado(), app()];
    let json = Library::to_json(&libs).unwrap();
    assert_eq!(Library::from_json(&json).unwrap(), libs);
    // A single library object is accepted as well as an array.
    let single = serde_json_single(&ado());
    assert_eq!(Library::from_json(&single).unwrap(), vec![ado()]);
}

fn serde_json_single(l: &Library) -> String {
    let arr = Library::to_json(std::slice::from_ref(l)).unwrap();
    arr[1..arr.len() - 1].to_string()
}

#[test]
fn json_defaults_describe_a_minimal_member() {
    let json = r#"{"name":"L","types":[{"name":"T","kind":"class","members":[{"name":"M","kind":"method"}]}]}"#;
    let libs = Library::from_json(json).unwrap();
    assert_eq!(libs[0].types[0].members[0], member("M", LibraryMemberKind::Method, None));
    assert_eq!(libs[0].types[0].default_interface, None);
    assert!(!libs[0].types[0].is_app_object);
}

#[test]
fn member_of_a_declared_library_class_resolves_through_the_coclass_and_extends_chain() {
    let main = "Sub Run()\n    Dim cn As ADODB.Connection\n    cn.Open\n    x = cn.Properties\nEnd Sub\n";
    let s = session_with(&[("modMain.bas", main)], config());
    // `Open` is declared by the default interface `_Connection`.
    assert_eq!(s.member_status_at(0, at(main, "cn.Open", "Open")), Some(MemberStatus::Resolved));
    // `Properties` comes from `_ADO`, which `_Connection` extends.
    assert_eq!(s.member_status_at(0, at(main, "cn.Properties", "Properties")), Some(MemberStatus::Resolved));
    assert_eq!(s.diagnostics(0), vec![]);
}

#[test]
fn library_member_references_are_keyed_by_library_type_and_member() {
    let main = "Sub A()\n    Dim cn As ADODB.Connection\n    cn.Open\nEnd Sub\nSub B()\n    Dim c2 As New ADODB.Connection\n    c2.Open\n    c2.Execute\nEnd Sub\n";
    let s = session_with(&[("modMain.bas", main)], config());
    let open = s.external_member_references("adodb", "Connection", "open");
    assert_eq!(
        sites(&[main], &open),
        vec![
            (0, at(main, "cn.Open", "Open"), "Open".to_string()),
            (0, at(main, "c2.Open", "Open"), "Open".to_string()),
        ]
    );
    // Asked for by the interface that declares it, the same sites.
    assert_eq!(s.external_member_references("ADODB", "_Connection", "Open"), open);
    // A different member of the same type is a different key.
    let exec = s.external_member_references("ADODB", "Connection", "Execute");
    assert_eq!(sites(&[main], &exec), vec![(0, at(main, "c2.Execute", "Execute"), "Execute".to_string())]);
    // Unknown names have no sites.
    assert_eq!(s.external_member_references("ADODB", "Connection", "Nope"), vec![]);
    assert_eq!(s.external_member_references("Nope", "Connection", "Open"), vec![]);
}

#[test]
fn references_at_a_library_member_use_lists_the_same_sites() {
    let main = "Sub A()\n    Dim cn As ADODB.Connection\n    cn.Open\n    cn.Open\nEnd Sub\n";
    let s = session_with(&[("modMain.bas", main)], config());
    let first = at(main, "    cn.Open\n    cn.Open", "Open");
    let refs = s.references(0, first, false);
    assert_eq!(refs.len(), 2);
    assert_eq!(s.document_highlights(0, first).len(), 2);
}

#[test]
fn property_accessors_pick_by_context_and_share_the_member_name() {
    let main = "Sub A()\n    Dim cn As ADODB.Connection\n    cn.ConnectionString = \"x\"\n    s = cn.ConnectionString\nEnd Sub\n";
    let s = session_with(&[("modMain.bas", main)], config());
    assert_eq!(s.member_status_at(0, at(main, "cn.ConnectionString = ", "ConnectionString")), Some(MemberStatus::Resolved));
    // The write and the read are different members (Let, Get) of one name;
    // asking by name lists both.
    let both = s.external_member_references("ADODB", "Connection", "ConnectionString");
    assert_eq!(both.len(), 2);
}

#[test]
fn assigning_a_get_only_library_property_names_the_missing_accessor() {
    let main = "Sub A()\n    Dim cn As ADODB.Connection\n    cn.State = 1\nEnd Sub\n";
    let s = session_with(&[("modMain.bas", main)], config());
    assert_eq!(
        s.member_status_at(0, at(main, "cn.State", "State")),
        Some(MemberStatus::MissingAccessor(vb6_core::sema::Access::Let))
    );
}

#[test]
fn unknown_member_of_a_loaded_library_type_is_reported() {
    let main = "Sub A()\n    Dim cn As ADODB.Connection\n    cn.Nope\nEnd Sub\n";
    let s = session_with(&[("modMain.bas", main)], config());
    assert_eq!(s.member_status_at(0, at(main, "cn.Nope", "Nope")), Some(MemberStatus::NoSuchMember));
    assert_eq!(common::coded(&s, main, 0, 461).len(), 1);
}

#[test]
fn result_types_flow_through_library_members() {
    let main = "Sub A()\n    Dim cn As ADODB.Connection\n    cn.Execute.MoveNext\n    x = cn.Execute.EOF\n    y = cn.Execute.Fields(0).Value\nEnd Sub\n";
    let s = session_with(&[("modMain.bas", main)], config());
    assert_eq!(s.member_status_at(0, at(main, "cn.Execute.MoveNext", "MoveNext")), Some(MemberStatus::Resolved));
    assert_eq!(s.member_status_at(0, at(main, "cn.Execute.EOF", "EOF")), Some(MemberStatus::Resolved));
    // `Fields` is a property with unspecified parameters, so its call is not
    // taken to index its result; the declared result type is `Field`.
    assert_eq!(s.member_status_at(0, at(main, "Fields(0).Value", "Value")), Some(MemberStatus::Resolved));
}

#[test]
fn default_member_of_a_library_type_types_the_call() {
    let main = "Sub A()\n    Dim rs As ADODB.Recordset\n    x = rs(0).Value\nEnd Sub\n";
    let s = session_with(&[("modMain.bas", main)], config());
    assert_eq!(s.member_status_at(0, at(main, "rs(0).Value", "Value")), Some(MemberStatus::Resolved));
}

#[test]
fn application_object_global_resolves_its_members() {
    let main = "Option Explicit\nSub A()\n    Dim s As String\n    s = App.Path\nEnd Sub\n";
    let s = session_with(&[("modMain.bas", main)], config());
    assert_eq!(s.member_status_at(0, at(main, "App.Path", "Path")), Some(MemberStatus::Resolved));
    assert_eq!(s.diagnostics(0), vec![]);
    let uses = s.external_member_references("VB", "App", "Path");
    assert_eq!(sites(&[main], &uses), vec![(0, at(main, "App.Path", "Path"), "Path".to_string())]);
}

#[test]
fn library_name_qualifies_a_type_name_in_an_expression() {
    let main = "Sub A()\n    Dim cn As ADODB.Connection\n    Set cn = New ADODB.Connection\n    cn.Open\nEnd Sub\n";
    let s = session_with(&[("modMain.bas", main)], config());
    assert_eq!(s.member_status_at(0, at(main, "cn.Open", "Open")), Some(MemberStatus::Resolved));
}

#[test]
fn library_that_is_not_supplied_leaves_its_members_unresolved() {
    let main = "Option Explicit\nSub A()\n    Dim cn As ADODB.Connection\n    cn.Open\nEnd Sub\n";
    let s = session_with(&[("modMain.bas", main)], SessionConfig::default());
    assert_eq!(
        s.member_status_at(0, at(main, "cn.Open", "Open")),
        Some(MemberStatus::LibraryNotLoaded("adodb".to_string()))
    );
    assert_eq!(s.diagnostics(0), vec![]);
    assert_eq!(s.external_member_references("ADODB", "Connection", "Open"), vec![]);
}

#[test]
fn hover_on_a_library_member_is_rendered_from_the_library_description() {
    let main = "Sub A()\n    Dim cn As ADODB.Connection\n    x = cn.State\nEnd Sub\n";
    let s = session_with(&[("modMain.bas", main)], config());
    let h = s.hover(0, at(main, "cn.State", "State")).unwrap();
    assert_eq!(h.text, "Property Get ADODB._Connection.State As Long");
}
