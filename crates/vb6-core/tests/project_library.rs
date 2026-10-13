//! The type library a project compiles to: which modules and members VB6
//! exposes, how each is described, and a second project resolving through it.

mod common;

use common::{at, session_with};
use vb6_core::sema::MemberStatus;
use vb6_core::session::{
    Instancing, Library, LibraryMember, LibraryMemberKind, LibraryParam, LibraryType,
    LibraryTypeKind, ProjectKind, SessionConfig,
};

/// A class module file: header block with `MultiUse`, then the attributes.
fn class(name: &str, multi_use: i32, exposed: bool, creatable: bool, global: bool, body: &str) -> String {
    let b = |v: bool| if v { "True" } else { "False" };
    format!(
        "VERSION 1.0 CLASS\nBEGIN\n  MultiUse = {multi_use}  'True\n  Persistable = 0  'NotPersistable\nEND\n\
         Attribute VB_Name = \"{name}\"\nAttribute VB_GlobalNameSpace = {}\nAttribute VB_Creatable = {}\n\
         Attribute VB_PredeclaredId = False\nAttribute VB_Exposed = {}\n{body}",
        b(global),
        b(creatable),
        b(exposed)
    )
}

fn public_class(name: &str, body: &str) -> String {
    class(name, -1, true, true, false, body)
}

fn project(name: &str) -> SessionConfig {
    SessionConfig { project_name: Some(name.into()), ..SessionConfig::default() }
}

fn export(files: &[(&str, &str)], kind: ProjectKind, config: SessionConfig) -> Library {
    session_with(files, config).project_library(kind)
}

fn param(name: &str, ty: Option<&str>) -> LibraryParam {
    LibraryParam { name: name.into(), type_name: ty.map(String::from), by_ref: true, ..Default::default() }
}

fn byval(name: &str, ty: Option<&str>) -> LibraryParam {
    LibraryParam { by_ref: false, ..param(name, ty) }
}

fn method(name: &str, ret: Option<&str>, params: Vec<LibraryParam>) -> LibraryMember {
    LibraryMember {
        name: name.into(),
        kind: LibraryMemberKind::Method,
        return_type: ret.map(String::from),
        param_count: Some(params.len()),
        params,
        ..Default::default()
    }
}

fn accessor(name: &str, kind: LibraryMemberKind, ret: Option<&str>, params: Vec<LibraryParam>) -> LibraryMember {
    LibraryMember { kind, ..method(name, ret, params) }
}

fn field(name: &str, ty: Option<&str>) -> LibraryMember {
    LibraryMember { name: name.into(), kind: LibraryMemberKind::Field, return_type: ty.map(String::from), ..Default::default() }
}

fn class_type(name: &str, instancing: Option<Instancing>, members: Vec<LibraryMember>) -> LibraryType {
    LibraryType { name: name.into(), kind: LibraryTypeKind::Class, instancing, members, ..Default::default() }
}

// ── Which modules are exposed ─────────────────────────────────────────────────

#[test]
fn a_standard_exe_exposes_nothing() {
    let cls = public_class("clsA", "Public Sub Run()\nEnd Sub\n");
    let lib = export(&[("clsA.cls", &cls)], ProjectKind::Exe, project("App"));
    assert_eq!(lib, Library { name: "App".into(), types: vec![] });
}

#[test]
fn only_exposed_classes_of_an_activex_dll_are_in_the_library() {
    let public = public_class("clsPublic", "Public Sub Run()\nEnd Sub\n");
    let private = class("clsPrivate", -1, false, true, false, "Public Sub Run()\nEnd Sub\n");
    let module = "Attribute VB_Name = \"modUtil\"\nPublic Sub Helper()\nEnd Sub\n";
    let form = "VERSION 5.00\nBegin VB.Form frmMain\n   Caption = \"Main\"\nEnd\nAttribute VB_Name = \"frmMain\"\nPublic Sub Run()\nEnd Sub\n";
    let lib = export(
        &[("clsPublic.cls", &public), ("clsPrivate.cls", &private), ("modUtil.bas", module), ("frmMain.frm", form)],
        ProjectKind::OleDll,
        project("Lib"),
    );
    assert_eq!(
        lib,
        Library {
            name: "Lib".into(),
            types: vec![class_type("clsPublic", Some(Instancing::MultiUse), vec![method("Run", None, vec![])])],
        }
    );
}

#[test]
fn instancing_follows_creatable_global_namespace_and_multi_use() {
    let cases = [
        (class("c", -1, true, false, false, ""), Instancing::PublicNotCreatable),
        (class("c", -1, true, true, false, ""), Instancing::MultiUse),
        (class("c", 0, true, true, false, ""), Instancing::SingleUse),
        (class("c", -1, true, true, true, ""), Instancing::GlobalMultiUse),
        (class("c", 0, true, true, true, ""), Instancing::GlobalSingleUse),
    ];
    for (src, instancing) in cases {
        let lib = export(&[("c.cls", &src)], ProjectKind::OleExe, project("P"));
        let global = matches!(instancing, Instancing::GlobalMultiUse | Instancing::GlobalSingleUse);
        assert_eq!(
            lib.types,
            vec![LibraryType { is_global: global, ..class_type("c", Some(instancing), vec![]) }],
            "{instancing:?}"
        );
    }
}

#[test]
fn a_public_usercontrol_is_exposed_by_a_control_project_only() {
    let ctl = "VERSION 5.00\nBegin VB.UserControl ucGrid\n   ClientHeight = 3600\nEnd\n\
               Attribute VB_Name = \"ucGrid\"\nAttribute VB_GlobalNameSpace = False\nAttribute VB_Creatable = True\n\
               Attribute VB_PredeclaredId = False\nAttribute VB_Exposed = True\nPublic Sub Refresh()\nEnd Sub\n";
    let expected = vec![class_type("ucGrid", None, vec![method("Refresh", None, vec![])])];
    assert_eq!(export(&[("ucGrid.ctl", ctl)], ProjectKind::Control, project("Ctl")).types, expected);
    assert_eq!(export(&[("ucGrid.ctl", ctl)], ProjectKind::OleDll, project("Ctl")).types, vec![]);
}

#[test]
fn a_private_usercontrol_is_not_exposed() {
    let ctl = "VERSION 5.00\nBegin VB.UserControl ucInner\n   ClientHeight = 3600\nEnd\n\
               Attribute VB_Name = \"ucInner\"\nAttribute VB_Exposed = False\nPublic Sub Refresh()\nEnd Sub\n";
    assert_eq!(export(&[("ucInner.ctl", ctl)], ProjectKind::Control, project("Ctl")).types, vec![]);
}

#[test]
fn a_public_userdocument_is_exposed_by_an_activex_dll() {
    let dob = "VERSION 5.00\nBegin VB.UserDocument docMain\n   ClientHeight = 3600\nEnd\n\
               Attribute VB_Name = \"docMain\"\nAttribute VB_Exposed = True\nPublic Sub Show()\nEnd Sub\n";
    assert_eq!(
        export(&[("docMain.dob", dob)], ProjectKind::OleDll, project("Doc")).types,
        vec![class_type("docMain", None, vec![method("Show", None, vec![])])]
    );
    assert_eq!(export(&[("docMain.dob", dob)], ProjectKind::Control, project("Doc")).types, vec![]);
}

// ── Members ─────────────────────────────────────────────────────────────────

#[test]
fn private_and_friend_members_are_not_exposed() {
    let cls = public_class(
        "clsA",
        "Public Sub Visible()\nEnd Sub\nPrivate Sub Hidden()\nEnd Sub\nFriend Sub InProject()\nEnd Sub\n\
         Friend Property Get Count() As Long\nEnd Property\nSub NoModifier()\nEnd Sub\n\
         Private m_n As Long\nDim m_s As String\n",
    );
    let lib = export(&[("clsA.cls", &cls)], ProjectKind::OleDll, project("P"));
    assert_eq!(
        lib.types[0].members,
        vec![method("Visible", None, vec![]), method("NoModifier", None, vec![])]
    );
}

#[test]
fn procedures_describe_every_parameter() {
    let cls = public_class(
        "clsA",
        "Public Function Find(ByVal Id As Long, Name As String, ByRef Rows() As Variant, Optional ByVal Limit As Integer = -1, \
         Optional Label = \"say \"\"hi\"\"\", ParamArray Extra() As Variant) As Boolean\nEnd Function\n\
         Public Sub Untyped(a, Optional b)\nEnd Sub\n",
    );
    let lib = export(&[("clsA.cls", &cls)], ProjectKind::OleDll, project("P"));
    assert_eq!(
        lib.types[0].members,
        vec![
            method(
                "Find",
                Some("Boolean"),
                vec![
                    byval("Id", Some("Long")),
                    param("Name", Some("String")),
                    param("Rows", Some("Variant()")),
                    LibraryParam { optional: true, default: Some("-1".into()), ..byval("Limit", Some("Integer")) },
                    LibraryParam { optional: true, default: Some("\"say \"\"hi\"\"\"".into()), ..param("Label", None) },
                    LibraryParam { param_array: true, ..param("Extra", Some("Variant()")) },
                ],
            ),
            method("Untyped", None, vec![param("a", None), LibraryParam { optional: true, ..param("b", None) }]),
        ]
    );
}

#[test]
fn default_values_are_written_as_vb_source() {
    let cls = public_class(
        "clsA",
        "Public Enum eMode\n  emFast = 1\nEnd Enum\n\
         Public Sub S(Optional a As Boolean = True, Optional b As eMode = emFast, Optional c As Date = #1/31/2000#, \
         Optional d As Date = #1/31/2000 2:05:09 PM#, Optional e As Double = 1.5, Optional f As Long = &H10, \
         Optional g As Long = 2 * 3, Optional h As Variant = Null)\nEnd Sub\n",
    );
    let lib = export(&[("clsA.cls", &cls)], ProjectKind::OleDll, project("P"));
    let defaults: Vec<Option<String>> = lib.types[0].members[0].params.iter().map(|p| p.default.clone()).collect();
    assert_eq!(
        defaults,
        vec![
            Some("True".into()),
            Some("emFast".into()),
            Some("#1/31/2000#".into()),
            Some("#1/31/2000 2:05:09 PM#".into()),
            Some("1.5".into()),
            Some("16".into()),
            Some("2 * 3".into()),
            Some("Null".into()),
        ]
    );
}

#[test]
fn properties_variables_and_type_characters() {
    let cls = public_class(
        "clsA",
        "Public Caption As String\nPublic Items() As Long\nPublic Owner As clsA\nPublic Anything\n\
         Public Property Get Value() As Long\nEnd Property\n\
         Public Property Let Value(ByVal v As Long)\nEnd Property\n\
         Public Property Set Parent(ByVal p As Object)\nEnd Property\n\
         Public Function Title$()\nEnd Function\n",
    );
    let lib = export(&[("clsA.cls", &cls)], ProjectKind::OleDll, project("P"));
    assert_eq!(
        lib.types[0].members,
        vec![
            field("Caption", Some("String")),
            field("Items", Some("Long()")),
            field("Owner", Some("clsA")),
            field("Anything", None),
            accessor("Value", LibraryMemberKind::PropertyGet, Some("Long"), vec![]),
            accessor("Value", LibraryMemberKind::PropertyLet, None, vec![byval("v", Some("Long"))]),
            accessor("Parent", LibraryMemberKind::PropertySet, None, vec![byval("p", Some("Object"))]),
            method("Title", Some("String"), vec![]),
        ]
    );
}

#[test]
fn member_attributes_give_the_default_member_dispids_and_flags() {
    let cls = public_class(
        "colItems",
        "Public Property Get Item(ByVal Index As Long) As Variant\nAttribute Item.VB_UserMemId = 0\nEnd Property\n\
         Public Function NewEnum() As IUnknown\nAttribute NewEnum.VB_UserMemId = -4\nAttribute NewEnum.VB_MemberFlags = \"40\"\nEnd Function\n\
         Public Count As Long\nAttribute Count.VB_VarUserMemId = 1610809344\n",
    );
    let lib = export(&[("colItems.cls", &cls)], ProjectKind::OleDll, project("P"));
    assert_eq!(
        lib.types[0].members,
        vec![
            LibraryMember {
                is_default: true,
                dispid: Some(0),
                ..accessor("Item", LibraryMemberKind::PropertyGet, Some("Variant"), vec![byval("Index", Some("Long"))])
            },
            // IUnknown is a type of the bundled stdole library.
            LibraryMember { dispid: Some(-4), member_flags: Some(0x40), ..method("NewEnum", Some("stdole.IUnknown"), vec![]) },
            LibraryMember { dispid: Some(1_610_809_344), ..field("Count", Some("Long")) },
        ]
    );
}

#[test]
fn events_are_members_with_their_parameters() {
    let cls = public_class(
        "clsA",
        "Public Event Changed(ByVal OldValue As Long, NewValue As Long)\nEvent Closed()\n",
    );
    let lib = export(&[("clsA.cls", &cls)], ProjectKind::OleDll, project("P"));
    let event = |name: &str, params: Vec<LibraryParam>| LibraryMember {
        kind: LibraryMemberKind::Event,
        ..method(name, None, params)
    };
    assert_eq!(
        lib.types[0].members,
        vec![
            event("Changed", vec![byval("OldValue", Some("Long")), param("NewValue", Some("Long"))]),
            event("Closed", vec![]),
        ]
    );
}

#[test]
fn implemented_interfaces_are_listed_qualified_by_their_library() {
    let shape = class("IShape", -1, true, false, false, "Public Sub Draw()\nEnd Sub\n");
    let circle = public_class(
        "clsCircle",
        "Implements IShape\nImplements ADODB.Connection\nImplements Recordset\nPrivate Sub IShape_Draw()\nEnd Sub\n",
    );
    let lib = export(&[("IShape.cls", &shape), ("clsCircle.cls", &circle)], ProjectKind::OleDll, project("Geo"));
    assert_eq!(
        lib.types[1],
        LibraryType {
            interfaces: vec!["IShape".into(), "ADODB.Connection".into(), "Recordset".into()],
            ..class_type("clsCircle", Some(Instancing::MultiUse), vec![])
        }
    );
}

#[test]
fn a_type_of_a_referenced_library_is_qualified_with_it() {
    let ado = Library {
        name: "ADODB".into(),
        types: vec![LibraryType { name: "Recordset".into(), kind: LibraryTypeKind::Class, ..Default::default() }],
    };
    let mut config = project("Data");
    config.libraries.push(ado);
    let cls = public_class(
        "clsQuery",
        "Implements Recordset\nPublic Function Open(ByVal rs As Recordset) As Recordset\nEnd Function\n\
         Public Function Other() As adodb.Recordset\nEnd Function\n",
    );
    let lib = export(&[("clsQuery.cls", &cls)], ProjectKind::OleDll, config);
    assert_eq!(
        lib.types[0],
        LibraryType {
            interfaces: vec!["ADODB.Recordset".into()],
            ..class_type(
                "clsQuery",
                Some(Instancing::MultiUse),
                vec![
                    method("Open", Some("ADODB.Recordset"), vec![byval("rs", Some("ADODB.Recordset"))]),
                    method("Other", Some("adodb.Recordset"), vec![]),
                ],
            )
        }
    );
}

// ── Enums and types ─────────────────────────────────────────────────────────

#[test]
fn public_enums_and_types_of_exposed_classes_follow_their_class() {
    let cls = public_class(
        "clsShop",
        "Public Enum eColor\n  ecRed\n  ecGreen = 5\n  ecBlue\nEnd Enum\nPrivate Enum eHidden\n  ehA\nEnd Enum\n\
         Public Type tItem\n  Name As String\n  Codes(1 To 3) As Integer\n  Label As String * 8\n  Extra\nEnd Type\n\
         Private Type tInternal\n  X As Long\nEnd Type\n",
    );
    let lib = export(&[("clsShop.cls", &cls)], ProjectKind::OleDll, project("Shop"));
    let constant = |name: &str, value: i64| LibraryMember {
        name: name.into(),
        kind: LibraryMemberKind::Constant,
        return_type: Some("Long".into()),
        value: Some(value),
        ..Default::default()
    };
    assert_eq!(
        lib.types,
        vec![
            class_type("clsShop", Some(Instancing::MultiUse), vec![]),
            LibraryType {
                name: "tItem".into(),
                kind: LibraryTypeKind::Record,
                members: vec![
                    field("Name", Some("String")),
                    field("Codes", Some("Integer()")),
                    field("Label", Some("String")),
                    field("Extra", None),
                ],
                ..Default::default()
            },
            LibraryType {
                name: "eColor".into(),
                kind: LibraryTypeKind::Enum,
                members: vec![constant("ecRed", 0), constant("ecGreen", 5), constant("ecBlue", 6)],
                ..Default::default()
            },
        ]
    );
}

#[test]
fn enums_and_types_of_standard_modules_and_private_classes_are_not_exposed() {
    let module = "Attribute VB_Name = \"modTypes\"\nPublic Enum eMode\n  emA\nEnd Enum\nPublic Type tPoint\n  X As Long\nEnd Type\n";
    let private = class("clsPrivate", -1, false, true, false, "Public Enum eKind\n  ekA\nEnd Enum\n");
    let lib = export(&[("modTypes.bas", module), ("clsPrivate.cls", &private)], ProjectKind::OleDll, project("P"));
    assert_eq!(lib.types, vec![]);
}

// ── Round trip and use from another project ─────────────────────────────────

#[test]
fn the_exported_library_round_trips_through_json() {
    let cls = public_class(
        "clsA",
        "Public Event Changed(ByVal v As Long)\nPublic Enum eA\n  eaOne = 1\nEnd Enum\n\
         Public Property Get Item(Optional ByVal i As Long = 2) As Variant\nAttribute Item.VB_UserMemId = 0\nAttribute Item.VB_MemberFlags = \"200\"\nEnd Property\n",
    );
    let lib = export(&[("clsA.cls", &cls)], ProjectKind::OleDll, project("P"));
    let json = Library::to_json(std::slice::from_ref(&lib)).unwrap();
    assert_eq!(Library::from_json(&json).unwrap(), vec![lib]);
}

fn consumer(files: &[(&str, &str)], producer: &Library) -> vb6_core::session::Session {
    let mut config = project("Client");
    config.libraries.push(producer.clone());
    session_with(files, config)
}

#[test]
fn another_project_resolves_classes_members_enums_and_records_through_the_library() {
    let cls = public_class(
        "clsOrder",
        "Public Enum eState\n  esOpen = 1\nEnd Enum\nPublic Type tLine\n  Qty As Long\nEnd Type\n\
         Public Event Shipped()\nPublic Total As Currency\n\
         Public Function AddLine(ByVal Qty As Long) As tLine\nEnd Function\n",
    );
    let lib = export(&[("clsOrder.cls", &cls)], ProjectKind::OleDll, project("Orders"));
    let client = "Sub T()\n    Dim o As Orders.clsOrder\n    Set o = New Orders.clsOrder\n    \
                  x = o.AddLine(1).Qty\n    y = o.Total\n    z = esOpen\n    w = o.Shipped\nEnd Sub\n";
    let s = consumer(&[("modClient.bas", client)], &lib);
    let status = |ctx: &str, member: &str| s.member_status_at(0, at(client, ctx, member));
    assert_eq!(status("o.AddLine(1)", "AddLine"), Some(MemberStatus::Resolved));
    assert_eq!(status("(1).Qty", "Qty"), Some(MemberStatus::Resolved));
    assert_eq!(status("o.Total", "Total"), Some(MemberStatus::Resolved));
    // An event is handled by the client, never read as a member.
    assert_eq!(status("o.Shipped", "Shipped"), Some(MemberStatus::NoSuchMember));
    assert_eq!(s.definition(0, at(client, "z = esOpen", "esOpen")), None, "a library constant has no source in the client");
    let diags: Vec<(u32, u32)> = s.diagnostics(0).iter().map(|d| (d.code, d.span.start)).collect();
    assert_eq!(diags, vec![(461, at(client, "o.Shipped", "Shipped"))], "only reading the event is an error");
}

#[test]
fn another_project_implements_an_interface_of_the_library() {
    let shape = class("IShape", -1, true, false, false, "Public Sub Draw()\nEnd Sub\n");
    let lib = export(&[("IShape.cls", &shape)], ProjectKind::OleDll, project("Geo"));
    let client = public_class("clsSquare", "Implements Geo.IShape\nPrivate Sub IShape_Draw()\nEnd Sub\n");
    let s = consumer(&[("clsSquare.cls", &client)], &lib);
    assert!(s.diagnostics(0).is_empty(), "{:?}", s.diagnostics(0));
}

#[test]
fn date_defaults_cover_day_zero_times_alone_and_dates_before_1900() {
    let cls = public_class(
        "clsA",
        "Public Sub S(Optional a As Date = #12/30/1899#, Optional b As Date = #3:00:00 PM#, \
         Optional c As Date = #1/1/1800#, Optional d As Date = #12/31/9999 11:59:59 PM#, \
         Optional e As Date = #2/29/2000 12:00:01 AM#)\nEnd Sub\n",
    );
    let lib = export(&[("clsA.cls", &cls)], ProjectKind::OleDll, project("P"));
    let defaults: Vec<Option<String>> = lib.types[0].members[0].params.iter().map(|p| p.default.clone()).collect();
    assert_eq!(
        defaults,
        vec![
            Some("#12:00:00 AM#".into()),
            Some("#3:00:00 PM#".into()),
            Some("#1/1/1800#".into()),
            Some("#12/31/9999 11:59:59 PM#".into()),
            Some("#2/29/2000 12:00:01 AM#".into()),
        ]
    );
}
