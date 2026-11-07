//! Reading registered and on-disk type libraries into engine libraries.

use std::path::PathBuf;
use std::sync::Arc;
use vb6_engine::session::{LibraryMember, LibraryMemberKind, LibraryParam, LibraryType, LibraryTypeKind};
use vb6_typelib::{load_library, TypeLibraryError, TypeLibraryRequest};

const OLE_AUTOMATION: &str = "{00020430-0000-0000-C000-000000000046}";
const ADODB_27: &str = "{EF53050B-882E-4776-B643-EDA472E8E3F2}";

fn request(guid: &str, major: u16, minor: u16, paths: Vec<PathBuf>) -> TypeLibraryRequest {
    TypeLibraryRequest { guid: guid.to_string(), major, minor, lcid: 0, paths, description: String::new() }
}

fn constant(name: &str, value: i64) -> LibraryMember {
    LibraryMember { name: name.into(), kind: LibraryMemberKind::Constant, value: Some(value), ..Default::default() }
}

fn getter(name: &str, ty: &str, dispid: i32, is_default: bool) -> LibraryMember {
    LibraryMember {
        name: name.into(),
        kind: LibraryMemberKind::PropertyGet,
        return_type: Some(ty.into()),
        is_default,
        param_count: Some(0),
        dispid: Some(dispid),
        ..Default::default()
    }
}

fn find<'a>(library: &'a vb6_engine::session::Library, name: &str) -> &'a LibraryType {
    library.types.iter().find(|t| t.name == name).unwrap_or_else(|| panic!("no type {name}"))
}

#[cfg(windows)]
mod windows_only {
    use super::*;

    fn stdole_file() -> PathBuf {
        PathBuf::from(std::env::var("SystemRoot").unwrap()).join("System32").join("stdole2.tlb")
    }

    #[test]
    fn a_library_is_read_from_the_file_the_project_names() {
        let library = load_library(&request(OLE_AUTOMATION, 2, 0, vec![stdole_file()])).unwrap();

        assert_eq!(library.name, "stdole");
        let names: Vec<&str> = library.types.iter().map(|t| t.name.as_str()).collect();
        assert_eq!(
            names,
            vec![
                "GUID", "DISPPARAMS", "EXCEPINFO", "IUnknown", "IDispatch", "IEnumVARIANT", "OLE_TRISTATE", "IFont",
                "Font", "StdFont", "IPicture", "Picture", "StdPicture", "LoadPictureConstants", "StdFunctions",
                "FontEvents",
            ]
        );
    }

    #[test]
    fn an_enumeration_has_its_constants_with_their_values() {
        let library = load_library(&request(OLE_AUTOMATION, 2, 0, vec![stdole_file()])).unwrap();

        assert_eq!(
            find(&library, "OLE_TRISTATE"),
            &LibraryType {
                name: "OLE_TRISTATE".into(),
                kind: LibraryTypeKind::Enum,
                members: vec![constant("Unchecked", 0), constant("Checked", 1), constant("Gray", 2)],
                ..Default::default()
            }
        );
    }

    #[test]
    fn a_coclass_lists_its_interfaces_and_the_events_of_its_source_interface() {
        let library = load_library(&request(OLE_AUTOMATION, 2, 0, vec![stdole_file()])).unwrap();

        assert_eq!(
            find(&library, "StdFont"),
            &LibraryType {
                name: "StdFont".into(),
                kind: LibraryTypeKind::Coclass,
                default_interface: Some("Font".into()),
                interfaces: vec!["IFont".into()],
                members: vec![LibraryMember {
                    name: "FontChanged".into(),
                    kind: LibraryMemberKind::Event,
                    param_count: Some(1),
                    params: vec![LibraryParam {
                        name: "PropertyName".into(),
                        type_name: Some("String".into()),
                        ..Default::default()
                    }],
                    dispid: Some(9),
                    ..Default::default()
                }],
                ..Default::default()
            }
        );
    }

    #[test]
    fn a_dispatch_interface_has_its_variables_as_fields_and_extends_its_base() {
        let library = load_library(&request(OLE_AUTOMATION, 2, 0, vec![stdole_file()])).unwrap();

        let field = |name: &str, ty: &str, dispid: i32, is_default: bool| LibraryMember {
            name: name.into(),
            kind: LibraryMemberKind::Field,
            return_type: Some(ty.into()),
            is_default,
            dispid: Some(dispid),
            ..Default::default()
        };
        assert_eq!(
            find(&library, "Font"),
            &LibraryType {
                name: "Font".into(),
                kind: LibraryTypeKind::Interface,
                extends: vec!["IDispatch".into()],
                members: vec![
                    field("Name", "String", 0, true),
                    field("Size", "Currency", 2, false),
                    field("Bold", "Boolean", 3, false),
                    field("Italic", "Boolean", 4, false),
                    field("Underline", "Boolean", 5, false),
                    field("Strikethrough", "Boolean", 6, false),
                    field("Weight", "Integer", 7, false),
                    field("Charset", "Integer", 8, false),
                ],
                ..Default::default()
            }
        );
    }

    #[test]
    fn a_registered_library_is_found_by_its_guid_and_version() {
        let library = load_library(&request(ADODB_27, 2, 7, Vec::new())).unwrap();

        assert_eq!(library.name, "ADODB");
        assert_eq!(library.types.len(), 94);
    }

    #[test]
    fn a_registered_library_serves_a_request_for_an_older_minor_version() {
        let library = load_library(&request(ADODB_27, 2, 0, Vec::new())).unwrap();

        assert_eq!(library.name, "ADODB");
    }

    #[test]
    fn an_enumeration_of_a_registered_library_has_its_values() {
        let library = load_library(&request(ADODB_27, 2, 7, Vec::new())).unwrap();

        assert_eq!(
            find(&library, "CursorTypeEnum"),
            &LibraryType {
                name: "CursorTypeEnum".into(),
                kind: LibraryTypeKind::Enum,
                members: vec![
                    constant("adOpenUnspecified", -1),
                    constant("adOpenForwardOnly", 0),
                    constant("adOpenKeyset", 1),
                    constant("adOpenDynamic", 2),
                    constant("adOpenStatic", 3),
                ],
                ..Default::default()
            }
        );
    }

    #[test]
    fn properties_carry_their_types_dispatch_ids_and_the_default_member() {
        let library = load_library(&request(ADODB_27, 2, 7, Vec::new())).unwrap();

        assert_eq!(
            find(&library, "Error"),
            &LibraryType {
                name: "Error".into(),
                kind: LibraryTypeKind::Interface,
                extends: vec!["stdole.IDispatch".into()],
                members: vec![
                    getter("Number", "Long", 1, false),
                    getter("Source", "String", 2, false),
                    getter("Description", "String", 0, true),
                    getter("HelpFile", "String", 3, false),
                    getter("HelpContext", "Long", 4, false),
                    getter("SQLState", "String", 5, false),
                    getter("NativeError", "Long", 6, false),
                ],
                ..Default::default()
            }
        );
    }

    #[test]
    fn methods_carry_their_parameters_and_the_restricted_ones_are_left_out() {
        let library = load_library(&request(ADODB_27, 2, 7, Vec::new())).unwrap();

        let errors = find(&library, "Errors");
        let shown: Vec<(&str, LibraryMemberKind)> = errors.members.iter().map(|m| (m.name.as_str(), m.kind)).collect();
        assert_eq!(
            shown,
            vec![
                ("Count", LibraryMemberKind::PropertyGet),
                ("Refresh", LibraryMemberKind::Method),
                ("Item", LibraryMemberKind::PropertyGet),
                ("Clear", LibraryMemberKind::Method),
            ]
        );
        assert_eq!(
            errors.members[2],
            LibraryMember {
                name: "Item".into(),
                kind: LibraryMemberKind::PropertyGet,
                return_type: Some("Error".into()),
                is_default: true,
                param_count: Some(1),
                params: vec![LibraryParam {
                    name: "Index".into(),
                    type_name: Some("Variant".into()),
                    ..Default::default()
                }],
                dispid: Some(0),
                ..Default::default()
            }
        );
    }

    #[test]
    fn optional_parameters_keep_their_default_values_and_by_reference_passing() {
        let library = load_library(&request(ADODB_27, 2, 7, Vec::new())).unwrap();

        let fields = find(&library, "Fields");
        let append = fields.members.iter().find(|m| m.name == "Append").unwrap();
        assert_eq!(
            append.params,
            vec![
                LibraryParam { name: "Name".into(), type_name: Some("String".into()), ..Default::default() },
                LibraryParam { name: "Type".into(), type_name: Some("DataTypeEnum".into()), ..Default::default() },
                LibraryParam {
                    name: "DefinedSize".into(),
                    type_name: Some("Long".into()),
                    optional: true,
                    default: Some("0".into()),
                    ..Default::default()
                },
                LibraryParam {
                    name: "Attrib".into(),
                    type_name: Some("FieldAttributeEnum".into()),
                    optional: true,
                    default: Some("-1".into()),
                    ..Default::default()
                },
                LibraryParam { name: "FieldValue".into(), type_name: Some("Variant".into()), optional: true, ..Default::default() },
            ]
        );

        let connection = find(&library, "Connection");
        let info_message = connection.members.iter().find(|m| m.name == "InfoMessage").unwrap();
        assert_eq!(
            info_message.params,
            vec![
                LibraryParam { name: "pError".into(), type_name: Some("Error".into()), ..Default::default() },
                LibraryParam {
                    name: "adStatus".into(),
                    type_name: Some("EventStatusEnum".into()),
                    by_ref: true,
                    ..Default::default()
                },
                LibraryParam { name: "pConnection".into(), type_name: Some("_Connection".into()), ..Default::default() },
            ]
        );
        assert_eq!(info_message.kind, LibraryMemberKind::Event);
        assert!(!info_message.is_default);
    }

    #[test]
    fn a_library_read_twice_is_shared() {
        let first = load_library(&request(ADODB_27, 2, 7, Vec::new())).unwrap();
        let second = load_library(&request(ADODB_27.to_ascii_lowercase().as_str(), 2, 7, Vec::new())).unwrap();

        assert!(Arc::ptr_eq(&first, &second));
    }

    #[test]
    fn a_library_that_is_neither_registered_nor_on_disk_is_reported_with_where_it_was_sought() {
        let missing = PathBuf::from(r"C:\no\such\dir\missing.tlb");
        let mut r = request("{DEADBEEF-0000-4000-8000-0000000000AA}", 1, 0, vec![missing.clone()]);
        r.description = "Missing Lib".into();

        let err = load_library(&r).unwrap_err();

        assert_eq!(
            err,
            TypeLibraryError::NotFound {
                what: "Missing Lib {DEADBEEF-0000-4000-8000-0000000000AA} 1.0".into(),
                tried: vec![r"C:\no\such\dir\missing.tlb".into()],
                reason: String::new(),
            }
        );
        assert_eq!(
            err.to_string(),
            r"Type library Missing Lib {DEADBEEF-0000-4000-8000-0000000000AA} 1.0 is not registered and was not found at C:\no\such\dir\missing.tlb"
        );
    }

    #[test]
    fn a_file_holding_another_library_is_not_taken_for_the_one_asked_for() {
        let file = stdole_file();
        let err = load_library(&request("{DEADBEEF-0000-4000-8000-0000000000AB}", 1, 0, vec![file.clone()])).unwrap_err();

        assert_eq!(
            err,
            TypeLibraryError::NotFound {
                what: "{DEADBEEF-0000-4000-8000-0000000000AB} 1.0".into(),
                tried: vec![file.display().to_string()],
                reason: format!("{} holds a different library", file.display()),
            }
        );
    }

    #[test]
    fn a_file_that_is_not_a_type_library_is_reported() {
        let dir = std::env::temp_dir().join(format!("vb6lsp-notlib-{}", std::process::id()));
        std::fs::create_dir_all(&dir).unwrap();
        let file = dir.join("plain.tlb");
        std::fs::write(&file, b"this is not a type library").unwrap();

        let err = load_library(&request("{DEADBEEF-0000-4000-8000-0000000000AC}", 1, 0, vec![file.clone()])).unwrap_err();
        std::fs::remove_dir_all(&dir).unwrap();

        match err {
            TypeLibraryError::NotFound { what, tried, reason } => {
                assert_eq!(what, "{DEADBEEF-0000-4000-8000-0000000000AC} 1.0");
                assert_eq!(tried, vec![file.display().to_string()]);
                assert!(reason.starts_with(&format!("{}: ", file.display())), "{reason}");
            }
            other => panic!("{other:?}"),
        }
    }
}

#[cfg(not(windows))]
#[test]
fn without_a_registry_a_library_the_project_names_no_file_for_is_not_found() {
    let err = load_library(&request(OLE_AUTOMATION, 2, 0, Vec::new())).unwrap_err();

    assert_eq!(
        err,
        TypeLibraryError::NotFound {
            what: "{00020430-0000-0000-C000-000000000046} 2.0".into(),
            tried: Vec::new(),
            reason: String::new(),
        }
    );
    assert_eq!(
        err.to_string(),
        "Type library {00020430-0000-0000-C000-000000000046} 2.0 is not registered and the project names no file for it"
    );
}
