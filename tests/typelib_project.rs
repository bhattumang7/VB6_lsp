//! A project's type-library and control-library references reach the analysis.

mod common;

use common::TempDir;
use vb6_lsp::project::{ReferenceKind, ReferenceStatus, Selection};

fn no_overlay(_: &str) -> Option<Vec<u8>> {
    None
}

const USES_ADO: &str = "Option Explicit\nSub Run()\n    Dim rs As ADODB.Recordset\n    Set rs = New ADODB.Recordset\n    rs.MoveNext\n    rs.NoSuchMethod\n    NoSuchProc\nEnd Sub\n";

fn diagnostics(s: &Selection) -> Vec<String> {
    let session = s.units[0].session.read().unwrap();
    let main = session.module_of(&s.units[0].members[0].key).unwrap();
    session.diagnostics(main).iter().map(|d| d.message().unwrap_or_default()).collect()
}

#[cfg(windows)]
mod windows_only {
    use super::*;

    const ADODB_REFERENCE: &str = "Reference=*\\G{EF53050B-882E-4776-B643-EDA472E8E3F2}#2.7#0#C:\\Program Files (x86)\\Common Files\\System\\ado\\msado15.dll#Microsoft ActiveX Data Objects 2.7 Library\n";

    #[test]
    fn a_referenced_type_library_is_loaded_into_the_projects_analysis() {
        let dir = TempDir::new("ado");
        let vbp = dir.write(
            "App.vbp",
            &format!("Type=Exe\nModule=modMain; modMain.bas\n{ADODB_REFERENCE}Name=\"App\"\n"),
        );
        dir.write("modMain.bas", USES_ADO);

        let s = Selection::load(&vbp, &no_overlay).unwrap();

        assert_eq!(s.problems, vec![]);
        assert_eq!(
            s.units[0].references.iter().map(|r| (r.kind, r.status)).collect::<Vec<_>>(),
            vec![(ReferenceKind::TypeLibrary, ReferenceStatus::Loaded)]
        );
        assert_eq!(s.units[0].type_libraries.iter().map(|l| l.name.as_str()).collect::<Vec<_>>(), vec!["ADODB"]);
        assert_eq!(
            diagnostics(&s),
            vec!["Method or data member not found".to_string(), "Sub or Function not defined".to_string()]
        );
    }

    #[test]
    fn the_same_code_without_the_reference_leaves_the_library_unknown_so_its_members_are_not_checked() {
        let dir = TempDir::new("noado");
        let vbp = dir.write("App.vbp", "Type=Exe\nModule=modMain; modMain.bas\nName=\"App\"\n");
        dir.write("modMain.bas", USES_ADO);

        let s = Selection::load(&vbp, &no_overlay).unwrap();

        assert_eq!(s.units[0].type_libraries.len(), 0);
        assert_eq!(diagnostics(&s), vec!["Sub or Function not defined".to_string()]);
    }

    #[test]
    fn a_control_library_is_found_through_the_registry_and_loaded() {
        let dir = TempDir::new("object");
        let vbp = dir.write(
            "App.vbp",
            "Type=Exe
Module=modMain; modMain.bas
Object={EF53050B-882E-4776-B643-EDA472E8E3F2}#2.7#0; msado15.dll
Name=\"App\"
",
        );
        dir.write("modMain.bas", "Option Explicit
Sub Run()
    Dim rs As ADODB.Recordset
    rs.MoveNext
    rs.NoSuchMember
End Sub
");

        let s = Selection::load(&vbp, &no_overlay).unwrap();

        assert_eq!(s.problems, vec![]);
        assert_eq!(
            s.units[0].references.iter().map(|r| (r.kind, r.status)).collect::<Vec<_>>(),
            vec![(ReferenceKind::Control, ReferenceStatus::Loaded)]
        );
        assert_eq!(s.units[0].type_libraries.iter().map(|l| l.name.as_str()).collect::<Vec<_>>(), vec!["ADODB"]);
        assert_eq!(diagnostics(&s), vec!["Method or data member not found".to_string()]);
    }
}
