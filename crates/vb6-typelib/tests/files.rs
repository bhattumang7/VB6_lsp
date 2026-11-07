//! Reading a type library from the files a project names.

use std::path::PathBuf;
use vb6_engine::session::{LibraryMember, LibraryMemberKind, LibraryParam, LibraryType, LibraryTypeKind};
use vb6_typelib::files::{find_existing, read_from_files, with_native_separators};
use vb6_typelib::{TypeLibraryError, TypeLibraryRequest};

fn temp_dir(name: &str) -> PathBuf {
    let dir = std::env::temp_dir().join(format!("vb6lsp-files-{name}-{}", std::process::id()));
    let _ = std::fs::remove_dir_all(&dir);
    std::fs::create_dir_all(&dir).unwrap();
    dir
}

#[test]
fn separators_follow_the_running_platform() {
    if cfg!(windows) {
        assert_eq!(with_native_separators(r"..\Lib\x.tlb"), r"..\Lib\x.tlb");
        assert_eq!(with_native_separators("a/b"), "a/b");
    } else {
        assert_eq!(with_native_separators(r"..\Lib\x.tlb"), "../Lib/x.tlb");
        assert_eq!(with_native_separators("a/b"), "a/b");
    }
}

#[test]
fn a_path_is_found_whatever_the_case_of_its_parts() {
    let dir = temp_dir("case");
    std::fs::create_dir_all(dir.join("Lib").join("Sub")).unwrap();
    let file = dir.join("Lib").join("Sub").join("File.TLB");
    std::fs::write(&file, b"x").unwrap();

    let found = find_existing(&dir.join("lib").join("SUB").join("file.tlb"));
    let exact = find_existing(&file);
    let absent = find_existing(&dir.join("lib").join("sub").join("other.tlb"));
    let absent_dir = find_existing(&dir.join("nothing").join("file.tlb"));
    std::fs::remove_dir_all(&dir).unwrap();

    // On a file system that ignores case the path asked for already exists and is kept as written.
    assert_eq!(found.map(|p| p.to_string_lossy().to_lowercase()), Some(file.to_string_lossy().to_lowercase()));
    assert_eq!(exact, Some(file));
    assert_eq!(absent, None);
    assert_eq!(absent_dir, None);
}

#[test]
fn a_file_that_does_not_exist_is_reported_with_where_it_was_sought() {
    let missing = temp_dir("missing").join("none.tlb");
    let request = TypeLibraryRequest {
        guid: "{DEADBEEF-0000-4000-8000-0000000000AA}".into(),
        major: 1,
        minor: 0,
        lcid: 0,
        paths: vec![missing.clone()],
        description: "Missing".into(),
    };

    let err = read_from_files(&request).unwrap_err();

    assert_eq!(
        err,
        TypeLibraryError::NotFound {
            what: "Missing {DEADBEEF-0000-4000-8000-0000000000AA} 1.0".into(),
            tried: vec![missing.display().to_string()],
            reason: String::new(),
        }
    );
}

#[test]
fn a_file_that_is_not_a_type_library_is_reported() {
    let dir = temp_dir("plain");
    let file = dir.join("plain.tlb");
    std::fs::write(&file, b"this is not a type library").unwrap();
    let request = TypeLibraryRequest {
        guid: "{DEADBEEF-0000-4000-8000-0000000000AB}".into(),
        major: 1,
        minor: 0,
        lcid: 0,
        paths: vec![file.clone()],
        description: String::new(),
    };

    let err = read_from_files(&request).unwrap_err();
    std::fs::remove_dir_all(&dir).unwrap();

    assert_eq!(
        err,
        TypeLibraryError::NotFound {
            what: "{DEADBEEF-0000-4000-8000-0000000000AB} 1.0".into(),
            tried: vec![file.display().to_string()],
            reason: format!("{}: not a type library or an executable image", file.display()),
        }
    );
}

#[cfg(windows)]
mod windows_only {
    use super::*;

    const OLE_AUTOMATION: &str = "{00020430-0000-0000-C000-000000000046}";

    fn stdole() -> PathBuf {
        PathBuf::from(std::env::var("SystemRoot").unwrap()).join("System32").join("stdole2.tlb")
    }

    fn request(guid: &str, path: PathBuf) -> TypeLibraryRequest {
        TypeLibraryRequest { guid: guid.into(), major: 2, minor: 0, lcid: 0, paths: vec![path], description: String::new() }
    }

    #[test]
    fn the_library_of_a_file_is_described_with_its_types_in_order() {
        let (library, missing) = read_from_files(&request(OLE_AUTOMATION, stdole())).unwrap();

        assert_eq!(library.name, "stdole");
        assert_eq!(missing, Vec::<String>::new());
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
    fn an_enumeration_has_its_constants_and_a_coclass_its_interfaces_and_events() {
        let (library, _) = read_from_files(&request(OLE_AUTOMATION, stdole())).unwrap();
        let find = |name: &str| library.types.iter().find(|t| t.name == name).unwrap().clone();

        let constant = |name: &str, value: i64| LibraryMember {
            name: name.into(),
            kind: LibraryMemberKind::Constant,
            value: Some(value),
            ..Default::default()
        };
        assert_eq!(
            find("OLE_TRISTATE"),
            LibraryType {
                name: "OLE_TRISTATE".into(),
                kind: LibraryTypeKind::Enum,
                members: vec![constant("Unchecked", 0), constant("Checked", 1), constant("Gray", 2)],
                ..Default::default()
            }
        );
        assert_eq!(
            find("StdFont"),
            LibraryType {
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
    fn a_dispatch_interface_derives_from_the_automation_interface_and_has_fields() {
        let (library, _) = read_from_files(&request(OLE_AUTOMATION, stdole())).unwrap();
        let font = library.types.iter().find(|t| t.name == "Font").unwrap();

        let field = |name: &str, ty: &str, dispid: i32, is_default: bool| LibraryMember {
            name: name.into(),
            kind: LibraryMemberKind::Field,
            return_type: Some(ty.into()),
            is_default,
            dispid: Some(dispid),
            ..Default::default()
        };
        assert_eq!(
            font,
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
    fn a_vtable_interface_names_the_value_a_property_assignment_takes_after_its_getter() {
        let (library, _) = read_from_files(&request(OLE_AUTOMATION, stdole())).unwrap();
        let ifont = library.types.iter().find(|t| t.name == "IFont").unwrap();

        let name_let = ifont.members.iter().find(|m| m.name == "Name" && m.kind == LibraryMemberKind::PropertyLet).unwrap();

        assert_eq!(
            name_let.params,
            vec![LibraryParam { name: "pname".into(), type_name: Some("String".into()), ..Default::default() }]
        );
    }

    #[test]
    fn a_file_holding_another_library_is_not_taken_for_the_one_asked_for() {
        let file = stdole();

        let err = read_from_files(&request("{DEADBEEF-0000-4000-8000-0000000000AC}", file.clone())).unwrap_err();

        assert_eq!(
            err,
            TypeLibraryError::NotFound {
                what: "{DEADBEEF-0000-4000-8000-0000000000AC} 2.0".into(),
                tried: vec![file.display().to_string()],
                reason: format!("{} holds a different library", file.display()),
            }
        );
    }

    #[test]
    fn a_resource_index_after_the_file_name_selects_the_library_of_a_dll() {
        let ado = PathBuf::from(r"C:\Program Files (x86)\Common Files\System\ado\msado15.dll");
        let bytes = std::fs::read(&ado).unwrap();
        let image = vb6_typelib::pe::type_library_bytes(&bytes, None).unwrap();
        let guid = format!("{{{}}}", vb6_typelib::msft::library_guid(image).unwrap());
        let guid = guid.as_str();

        let plain = read_from_files(&request(guid, ado.clone())).unwrap().0;
        let indexed = read_from_files(&request(guid, PathBuf::from(format!("{}\\1", ado.display())))).unwrap().0;

        assert_eq!(plain.name, "ADODB");
        assert_eq!(indexed, plain);
    }
}
