use super::*;

#[test]
fn test_parse_simple_vbp() {
    let content = r#"
Type=Exe
Name="TestProject"
Module=ModMain; ModMain.bas
Module=ModUtils; Utils\ModUtils.bas
Class=clsDatabase; clsDatabase.cls
Form=frmMain.frm
Startup="Sub Main"
ExeName32="TestProject.exe"
"#;

    let vbp = VbpFile::parse_content(Path::new("C:\\Projects\\Test.vbp"), content).unwrap();

    assert_eq!(vbp.project_type, ProjectType::Exe);
    assert_eq!(vbp.name, "TestProject");
    assert_eq!(vbp.modules.len(), 2);
    assert_eq!(vbp.classes.len(), 1);
    assert_eq!(vbp.forms.len(), 1);
    assert_eq!(vbp.startup, Some("Sub Main".to_string()));
    assert_eq!(vbp.exe_name, Some("TestProject.exe".to_string()));
}

#[test]
fn test_project_type_parsing() {
    assert_eq!(ProjectType::from_str("Exe"), ProjectType::Exe);
    assert_eq!(ProjectType::from_str("OleDll"), ProjectType::OleDll);
    assert_eq!(ProjectType::from_str("OLEEXE"), ProjectType::OleExe);
    assert_eq!(ProjectType::from_str("Control"), ProjectType::Control);
}

#[test]
fn test_version_info_parsing() {
    let content = r#"
Type=Exe
Name="TestProject"
MajorVer=1
MinorVer=2
RevisionVer=3
AutoIncrementVer=1
VersionCompanyName="Test Company"
VersionFileDescription="Test Description"
VersionLegalCopyright="Copyright 2024"
VersionProductName="Test Product"
"#;

    let vbp = VbpFile::parse_content(Path::new("C:\\Projects\\Test.vbp"), content).unwrap();

    assert_eq!(vbp.version_info.major, 1);
    assert_eq!(vbp.version_info.minor, 2);
    assert_eq!(vbp.version_info.revision, 3);
    assert_eq!(vbp.version_info.auto_increment, 1);
    assert_eq!(
        vbp.version_info.company_name,
        Some("Test Company".to_string())
    );
    assert_eq!(
        vbp.version_info.file_description,
        Some("Test Description".to_string())
    );
}

#[test]
fn test_compilation_settings_parsing() {
    let content = r#"
Type=Exe
Name="TestProject"
CompilationType=0
OptimizationType=1
FavorPentiumPro(tm)=-1
CodeViewDebugInfo=-1
NoAliasing=-1
BoundsCheck=-1
OverflowCheck=-1
FlPointCheck=-1
FDIVCheck=-1
UnroundedFP=-1
CondComp="DEBUG=1"
"#;

    let vbp = VbpFile::parse_content(Path::new("C:\\Projects\\Test.vbp"), content).unwrap();

    assert_eq!(vbp.compilation.compilation_type, CompilationType::NativeCode);
    assert_eq!(
        vbp.compilation.optimization_type,
        OptimizationType::FavorFastCode
    );
    assert!(vbp.compilation.favor_pentium_pro);
    assert!(vbp.compilation.code_view_debug_info);
    assert!(vbp.compilation.no_aliasing);
    assert!(vbp.compilation.bounds_check);
    assert!(vbp.compilation.overflow_check);
    assert!(vbp.compilation.floating_point_check);
    assert!(vbp.compilation.fdiv_check);
    assert!(vbp.compilation.unrounded_fp);
    assert_eq!(
        vbp.compilation.conditional_compile,
        Some("DEBUG=1".to_string())
    );
}

#[test]
fn test_threading_settings_parsing() {
    let content = r#"
Type=Exe
Name="TestProject"
StartMode=1
Unattended=-1
Retained=-1
ThreadPerObject=0
MaxNumberOfThreads=4
ThreadingModel=1
"#;

    let vbp = VbpFile::parse_content(Path::new("C:\\Projects\\Test.vbp"), content).unwrap();

    assert_eq!(vbp.threading.start_mode, StartMode::Automation);
    assert!(vbp.threading.unattended);
    assert!(vbp.threading.retained);
    assert_eq!(vbp.threading.thread_per_object, Some(0));
    assert_eq!(vbp.threading.max_threads, 4);
    assert_eq!(
        vbp.threading.threading_model,
        ThreadingModel::ApartmentThreaded
    );
}

#[test]
fn test_custom_section_parsing() {
    let content = r#"
Type=Exe
Name="TestProject"

[MS Transaction Server]
AutoRefresh=1
"#;

    let vbp = VbpFile::parse_content(Path::new("C:\\Projects\\Test.vbp"), content).unwrap();

    let mts_section = vbp.get_custom_section("MS Transaction Server");
    assert!(mts_section.is_some());
    let mts = mts_section.unwrap();
    assert_eq!(mts.get("AutoRefresh"), Some(&"1".to_string()));
}

#[test]
fn test_compatibility_settings_parsing() {
    let content = r#"
Type=Exe
Name="TestProject"
CompatibleMode="2"
CompatibleEXE32="C:\Projects\MyApp.exe"
NoControlUpgrade=1
RemoveUnusedControlInfo=-1
ServerSupportFiles=-1
"#;

    let vbp = VbpFile::parse_content(Path::new("C:\\Projects\\Test.vbp"), content).unwrap();

    assert_eq!(vbp.compatibility.mode, CompatibilityMode::Binary);
    assert_eq!(
        vbp.compatibility.compatible_exe,
        Some(PathBuf::from("C:\\Projects\\MyApp.exe"))
    );
    assert!(!vbp.compatibility.upgrade_controls); // NoControlUpgrade=1 means don't upgrade
    assert!(vbp.compatibility.remove_unused_control_info);
    assert!(vbp.compatibility.server_support_files);
}

#[test]
fn test_get_subproject_references() {
    let content = r#"
Type=Exe
Name="TestProject"
Reference=*\G{00020430-0000-0000-C000-000000000046}#2.0#0#C:\Windows\System32\stdole2.tlb#OLE Automation
Reference=*\ACommonLib.vbp
Reference=*\AUtils.vbp
"#;

    let vbp = VbpFile::parse_content(Path::new("C:\\Projects\\Test.vbp"), content).unwrap();

    let subprojects = vbp.get_subproject_references();
    assert_eq!(subprojects.len(), 2);

    let compiled = vbp.get_compiled_references();
    assert_eq!(compiled.len(), 1);
}
