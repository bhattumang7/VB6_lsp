//! Names that resolve to a referenced project's library, and finding their declarations in the
//! referenced project's own session.

use vb6_engine::frontend::ast::Span;
use vb6_engine::session::{
    ExternalTarget, Library, LibraryMemberKind, Location, ProjectKind, Session, SessionConfig,
};

const CLASS: &str = "VERSION 1.0 CLASS\nBEGIN\n  MultiUse = -1  'True\nEND\nAttribute VB_Name = \"clsThing\"\n\
Attribute VB_GlobalNameSpace = False\nAttribute VB_Creatable = True\nAttribute VB_PredeclaredId = False\n\
Attribute VB_Exposed = True\nPublic Sub Hello()\nEnd Sub\nPublic Property Get Size() As Long\nEnd Property\n\
Public Property Let Size(ByVal v As Long)\nEnd Property\nPublic Total As Long\nPrivate Hidden As Long\n\
Public Enum Colour\n    Red = 1\n    Blue = 2\nEnd Enum\nPublic Type Pair\n    Left As Long\n    Right As Long\nEnd Type\n";

const APP: &str = "Option Explicit\nSub Run()\n    Dim t As Lib.clsThing\n    t.Hello\n    t.Size = 1\n    t.Total = 2\n    Dim c As Long\n    c = Lib.Blue\nEnd Sub\n";

fn library_session() -> Session {
    Session::with_config(
        vec![("file:///lib/clsThing.cls".to_string(), CLASS.as_bytes().to_vec())],
        SessionConfig { project_name: Some("Lib".into()), ..Default::default() },
    )
}

fn app_session(lib: &Session) -> Session {
    let mut libraries = Library::bundled();
    libraries.push(lib.project_library(ProjectKind::OleDll));
    Session::with_config(
        vec![("file:///app/modMain.bas".to_string(), APP.as_bytes().to_vec())],
        SessionConfig { project_name: Some("App".into()), libraries, conditional_compilation: None },
    )
}

fn offset_of(source: &str, needle: &str, nth: usize) -> u32 {
    source.match_indices(needle).nth(nth).expect("needle present").0 as u32
}

fn span_of(source: &str, needle: &str, nth: usize) -> Span {
    Span { start: offset_of(source, needle, nth), len: needle.len() as u32 }
}

#[test]
fn a_member_of_a_referenced_class_is_an_external_target() {
    let lib = library_session();
    let app = app_session(&lib);
    let module = app.module_of("file:///app/modMain.bas").unwrap();

    let hello = app.external_target_at(module, offset_of(APP, "Hello", 0));
    let size = app.external_target_at(module, offset_of(APP, "Size", 0));
    let total = app.external_target_at(module, offset_of(APP, "Total", 0));

    assert_eq!(
        hello,
        Some(ExternalTarget { library: "Lib".into(), type_name: "clsThing".into(), member: Some(("Hello".into(), LibraryMemberKind::Method)) })
    );
    assert_eq!(
        size,
        Some(ExternalTarget { library: "Lib".into(), type_name: "clsThing".into(), member: Some(("Size".into(), LibraryMemberKind::PropertyLet)) })
    );
    assert_eq!(
        total,
        Some(ExternalTarget { library: "Lib".into(), type_name: "clsThing".into(), member: Some(("Total".into(), LibraryMemberKind::Field)) })
    );
}

#[test]
fn an_enumeration_member_of_a_referenced_project_is_an_external_target() {
    let lib = library_session();
    let app = app_session(&lib);
    let module = app.module_of("file:///app/modMain.bas").unwrap();

    let blue = app.external_target_at(module, offset_of(APP, "Blue", 0));

    assert_eq!(
        blue,
        Some(ExternalTarget { library: "Lib".into(), type_name: "Colour".into(), member: Some(("Blue".into(), LibraryMemberKind::Constant)) })
    );
}

#[test]
fn a_class_name_is_an_external_target_without_a_member() {
    let lib = library_session();
    let app = app_session(&lib);
    let module = app.module_of("file:///app/modMain.bas").unwrap();

    let class = app.external_target_at(module, offset_of(APP, "clsThing", 0));

    assert_eq!(class, Some(ExternalTarget { library: "Lib".into(), type_name: "clsThing".into(), member: None }));
}

#[test]
fn a_name_of_this_session_or_of_a_runtime_library_is_not_an_external_target() {
    let lib = library_session();
    let app = app_session(&lib);
    let module = app.module_of("file:///app/modMain.bas").unwrap();

    assert_eq!(app.external_target_at(module, offset_of(APP, "Run", 0)), None);
    assert_eq!(app.external_target_at(module, offset_of(APP, "Dim", 0)), None);
    assert_eq!(app.external_target_at(module, 100_000), None);
}

#[test]
fn a_class_is_declared_by_the_start_of_its_file() {
    let lib = library_session();

    assert_eq!(lib.declaration_of("clsThing", None), Some(Location { module: 0, span: Span { start: 0, len: 0 } }));
    assert_eq!(lib.declaration_of("CLSTHING", None), Some(Location { module: 0, span: Span { start: 0, len: 0 } }));
}

#[test]
fn a_member_is_declared_by_its_name_and_a_property_accessor_by_its_kind() {
    let lib = library_session();

    assert_eq!(
        lib.declaration_of("clsThing", Some(("hello", LibraryMemberKind::Method))),
        Some(Location { module: 0, span: span_of(CLASS, "Hello", 0) })
    );
    assert_eq!(
        lib.declaration_of("clsThing", Some(("Size", LibraryMemberKind::PropertyGet))),
        Some(Location { module: 0, span: span_of(CLASS, "Size", 0) })
    );
    assert_eq!(
        lib.declaration_of("clsThing", Some(("Size", LibraryMemberKind::PropertyLet))),
        Some(Location { module: 0, span: span_of(CLASS, "Size", 1) })
    );
    // An accessor kind the property lacks falls back to the accessor it has first.
    assert_eq!(
        lib.declaration_of("clsThing", Some(("Size", LibraryMemberKind::PropertySet))),
        Some(Location { module: 0, span: span_of(CLASS, "Size", 0) })
    );
    assert_eq!(
        lib.declaration_of("clsThing", Some(("Total", LibraryMemberKind::Field))),
        Some(Location { module: 0, span: span_of(CLASS, "Total", 0) })
    );
}

#[test]
fn enumerations_and_records_are_declared_by_their_names_and_members() {
    let lib = library_session();

    assert_eq!(lib.declaration_of("Colour", None), Some(Location { module: 0, span: span_of(CLASS, "Colour", 0) }));
    assert_eq!(
        lib.declaration_of("Colour", Some(("Blue", LibraryMemberKind::Constant))),
        Some(Location { module: 0, span: span_of(CLASS, "Blue", 0) })
    );
    assert_eq!(lib.declaration_of("Pair", None), Some(Location { module: 0, span: span_of(CLASS, "Pair", 0) }));
    assert_eq!(
        lib.declaration_of("Pair", Some(("Right", LibraryMemberKind::Field))),
        Some(Location { module: 0, span: span_of(CLASS, "Right", 0) })
    );
}

#[test]
fn what_the_project_does_not_declare_publicly_is_not_found() {
    let lib = library_session();

    assert_eq!(lib.declaration_of("clsThing", Some(("Hidden", LibraryMemberKind::Field))), None);
    assert_eq!(lib.declaration_of("clsThing", Some(("Missing", LibraryMemberKind::Method))), None);
    assert_eq!(lib.declaration_of("Colour", Some(("Green", LibraryMemberKind::Constant))), None);
    assert_eq!(lib.declaration_of("Pair", Some(("Top", LibraryMemberKind::Field))), None);
    assert_eq!(lib.declaration_of("clsOther", None), None);
}
