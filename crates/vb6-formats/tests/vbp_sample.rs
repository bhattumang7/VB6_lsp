//! The bundled sample project parsed end to end: every part of the parsed file is checked.

use std::path::{Path, PathBuf};

use vb6_formats::vbp::{
    CompatibilityMode, CompilationType, OptimizationType, ProjectType, StartMode, ThreadingModel,
    TypeLibReference, VbpFile,
};

fn sample() -> (PathBuf, VbpFile) {
    let path = Path::new(env!("CARGO_MANIFEST_DIR"))
        .join("..")
        .join("..")
        .join("tests")
        .join("fixtures")
        .join("vb6_sample")
        .join("Project1.vbp");
    let vbp = VbpFile::parse(&path).expect("the sample project parses");
    (path, vbp)
}

#[test]
fn the_sample_projects_identity_and_members() {
    let (path, vbp) = sample();
    let dir = path.parent().unwrap();

    assert_eq!(vbp.project_type, ProjectType::Exe);
    assert_eq!(vbp.name, "Project1");
    assert_eq!(vbp.startup.as_deref(), Some("Form1"));

    assert_eq!(vbp.forms.len(), 1);
    let form = &vbp.forms[0];
    assert_eq!(form.name, "");
    assert_eq!(form.relative_path, PathBuf::from("Form1.frm"));
    assert_eq!(form.absolute_path, dir.join("Form1.frm"));
    assert_eq!(form.line, 2);

    assert!(vbp.modules.is_empty());
    assert!(vbp.classes.is_empty());
    assert!(vbp.user_controls.is_empty());
    assert!(vbp.property_pages.is_empty());
    assert!(vbp.user_documents.is_empty());
    assert!(vbp.designers.is_empty());
    assert!(vbp.related_documents.is_empty());
}

#[test]
fn the_sample_projects_references_and_objects() {
    let (_, vbp) = sample();

    assert_eq!(vbp.references.len(), 1);
    match &vbp.references[0] {
        TypeLibReference::Compiled { guid, major, minor, lcid, path, description, line } => {
            assert_eq!(guid, "00020430-0000-0000-C000-000000000046");
            assert_eq!((*major, *minor, *lcid), (2, 0, 0));
            assert_eq!(path, &PathBuf::from(r"..\..\..\..\..\Windows\SysWOW64\stdole2.tlb"));
            assert_eq!(description, "OLE Automation");
            assert_eq!(*line, 3);
        }
        other => panic!("expected a compiled library reference, got {other:?}"),
    }

    assert_eq!(vbp.objects.len(), 2);
    assert_eq!(vbp.objects[0].guid, "0ECD9B60-23AA-11D0-B351-00A0C9055D8E");
    assert_eq!(vbp.objects[0].version, "6.0");
    assert_eq!(vbp.objects[0].filename, "MShflxgd.ocx");
    assert_eq!(vbp.objects[0].line, 4);
    assert_eq!(vbp.objects[1].guid, "65E121D4-0C60-11D2-A9FC-0000F8754DA1");
    assert_eq!(vbp.objects[1].version, "2.0");
    assert_eq!(vbp.objects[1].filename, "MSChrt20.ocx");
    assert_eq!(vbp.objects[1].line, 5);
}

#[test]
fn the_sample_projects_settings() {
    let (_, vbp) = sample();

    assert_eq!(
        (vbp.version_info.major, vbp.version_info.minor, vbp.version_info.revision),
        (1, 0, 0)
    );
    assert_eq!(vbp.version_info.auto_increment, 0);

    assert_eq!(vbp.compilation.compilation_type, CompilationType::NativeCode);
    assert_eq!(vbp.compilation.optimization_type, OptimizationType::FavorFastCode);
    assert!(!vbp.compilation.favor_pentium_pro);
    assert!(!vbp.compilation.code_view_debug_info);
    assert!(!vbp.compilation.no_aliasing);
    assert!(!vbp.compilation.bounds_check);
    assert!(!vbp.compilation.overflow_check);
    assert!(!vbp.compilation.floating_point_check);
    assert!(!vbp.compilation.fdiv_check);
    assert!(!vbp.compilation.unrounded_fp);

    assert_eq!(vbp.threading.start_mode, StartMode::StandAlone);
    assert!(!vbp.threading.unattended);
    assert!(!vbp.threading.retained);
    assert_eq!(vbp.threading.max_threads, 1);
    assert_eq!(vbp.threading.thread_per_object, Some(0));
    assert_eq!(vbp.threading.threading_model, ThreadingModel::ApartmentThreaded);

    assert_eq!(vbp.compatibility.mode, CompatibilityMode::NoCompatibility);
    assert_eq!(vbp.compatibility.compatible_exe, None);
    assert!(!vbp.compatibility.server_support_files);
    assert!(vbp.custom_sections.is_empty());
}
