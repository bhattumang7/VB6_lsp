//! The Session parses each file as the module kind its extension implies, so
//! object-module-only statements are accepted in class modules and designers
//! and rejected (with their exact span) in standard modules.

use vb6_core::frontend::parser::ModuleKind;
use vb6_core::session::{module_kind_for_path, Session};

/// "Only valid in object module".
const ERR_OBJECT_MODULE_ONLY: u32 = 0xdee1;

/// `(code, start, len, label)` of every diagnostic of the only module in a
/// one-file session.
fn diags(path: &str, src: &str) -> Vec<(u32, u32, u32, Option<&'static str>)> {
    let s = Session::from_sources(vec![(path.to_string(), src.as_bytes().to_vec())]);
    s.diagnostics(0).iter().map(|d| (d.code, d.span.start, d.span.len, d.label)).collect()
}

#[test]
fn extension_selects_module_kind() {
    assert_eq!(module_kind_for_path("M.bas"), ModuleKind::Standard);
    assert_eq!(module_kind_for_path("C.cls"), ModuleKind::Class);
    assert_eq!(module_kind_for_path("F.frm"), ModuleKind::Form);
    assert_eq!(module_kind_for_path("U.ctl"), ModuleKind::Form);
    assert_eq!(module_kind_for_path("P.pag"), ModuleKind::Form);
    assert_eq!(module_kind_for_path("D.dob"), ModuleKind::Form);
    assert_eq!(module_kind_for_path("R.dsr"), ModuleKind::Form);
}

#[test]
fn extension_match_ignores_case_and_directories() {
    assert_eq!(module_kind_for_path("dir.bas/C.CLS"), ModuleKind::Class);
    assert_eq!(module_kind_for_path("dir.cls/M.Bas"), ModuleKind::Standard);
    assert_eq!(module_kind_for_path("Forms/Main.FRM"), ModuleKind::Form);
}

#[test]
fn unknown_or_missing_extension_is_standard() {
    assert_eq!(module_kind_for_path("NoExtension"), ModuleKind::Standard);
    assert_eq!(module_kind_for_path("Notes.txt"), ModuleKind::Standard);
    assert_eq!(module_kind_for_path(""), ModuleKind::Standard);
}

#[test]
fn implements_in_class_module_is_accepted() {
    let src = "Option Explicit\nImplements IShape\n";
    assert_eq!(diags("Circle.cls", src), vec![]);
}

#[test]
fn event_in_class_module_is_accepted() {
    let src = "Public Event Changed(ByVal Value As Long)\n";
    assert_eq!(diags("Model.cls", src), vec![]);
}

#[test]
fn implements_in_every_designer_kind_is_accepted() {
    let src = "Implements IShape\n";
    for path in ["Main.frm", "Grid.ctl", "General.pag", "Doc.dob", "Report.dsr"] {
        assert_eq!(diags(path, src), vec![], "{path}");
    }
}

#[test]
fn implements_in_standard_module_is_rejected_at_the_keyword() {
    let src = "Option Explicit\nImplements IShape\n";
    let start = src.find("Implements").unwrap() as u32;
    assert_eq!(
        diags("Shapes.bas", src),
        vec![(ERR_OBJECT_MODULE_ONLY, start, "Implements".len() as u32, None)]
    );
}

#[test]
fn access_modified_event_in_class_module_is_accepted() {
    for src in ["Public Event Changed()\n", "Private Event Changed()\n"] {
        assert_eq!(diags("Model.cls", src), vec![], "{src:?}");
    }
}

#[test]
fn with_events_in_class_and_form_is_accepted() {
    for src in ["Private WithEvents m_Timer As Clock\n", "Public WithEvents m_Timer As Clock\n",
                "Dim WithEvents m_Timer As Clock\n"] {
        assert_eq!(diags("Model.cls", src), vec![], "{src:?}");
        assert_eq!(diags("Main.frm", src), vec![], "{src:?}");
    }
}

#[test]
fn with_events_in_standard_module_is_rejected_at_the_keyword() {
    for src in ["Private WithEvents m_Timer As Clock\n", "Dim WithEvents m_Timer As Clock\n"] {
        let start = src.find("WithEvents").unwrap() as u32;
        assert_eq!(
            diags("Timers.bas", src),
            vec![(ERR_OBJECT_MODULE_ONLY, start, "WithEvents".len() as u32, None)],
            "{src:?}"
        );
    }
}

#[test]
fn class_only_statements_inside_conditional_compilation_are_checked() {
    let src = "#If Win32 Then\nImplements IShape\nEvent Changed()\n#End If\n";
    assert_eq!(diags("Circle.cls", src), vec![]);
    let imp = src.find("Implements").unwrap() as u32;
    let ev = src.find("Event").unwrap() as u32;
    assert_eq!(
        diags("Shapes.bas", src),
        vec![
            (ERR_OBJECT_MODULE_ONLY, imp, "Implements".len() as u32, None),
            (ERR_OBJECT_MODULE_ONLY, ev, "Event".len() as u32, None),
        ]
    );
}

#[test]
fn access_modified_event_in_standard_module_is_rejected_at_the_keyword() {
    let src = "Public Event Changed()\n";
    let start = src.find("Event").unwrap() as u32;
    assert_eq!(
        diags("Events.bas", src),
        vec![(ERR_OBJECT_MODULE_ONLY, start, "Event".len() as u32, None)]
    );
}
