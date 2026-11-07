use super::*;

const SAMPLE: &str = r#"VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmMain
   Caption         =   "Main Window"
   Icon            =   "frmMain.frx":0000
   Begin VB.ListBox lstItems
      List            =   "frmMain.frx":0C10
      ItemData        =   "frmMain.frx":0C9A
   End
   Begin MSComctlLib.Toolbar tbMain
      _ExtentX        =   100
      Caption         =   $"frmMain.frx":1000
   End
   BeginProperty Font
      Name            =   "MS Sans Serif"
      Size            =   8.25
   EndProperty
End
Attribute VB_Name = "frmMain"
Private Sub Form_Load()
End Sub
"#;

fn reference(file: &str, offset: u32, dollar: bool) -> Option<FrxRef> {
    Some(FrxRef { file: file.to_string(), offset, dollar })
}

fn prop(name: &str, value: &str, frx: Option<FrxRef>) -> DesignerProp {
    DesignerProp { name: name.to_string(), value: value.to_string(), frx }
}

#[test]
fn the_whole_designer_of_a_form() {
    let expected = FormDesigner {
        objects: vec![ObjectEntry {
            clsid: "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}".to_string(),
            version: "2.0".to_string(),
            file: "MSCOMCTL.OCX".to_string(),
        }],
        root: Some(DesignerControl {
            type_name: "VB.Form".to_string(),
            name: "frmMain".to_string(),
            properties: vec![
                prop("Caption", "\"Main Window\"", None),
                prop("Icon", "\"frmMain.frx\" : 0000", reference("frmMain.frx", 0, false)),
            ],
            children: vec![
                DesignerControl {
                    type_name: "VB.ListBox".to_string(),
                    name: "lstItems".to_string(),
                    properties: vec![
                        prop("List", "\"frmMain.frx\" : 0C10", reference("frmMain.frx", 0xC10, false)),
                        prop(
                            "ItemData",
                            "\"frmMain.frx\" : 0C9A",
                            reference("frmMain.frx", 0xC9A, false),
                        ),
                    ],
                    children: Vec::new(),
                },
                DesignerControl {
                    type_name: "MSComctlLib.Toolbar".to_string(),
                    name: "tbMain".to_string(),
                    properties: vec![
                        prop("_ExtentX", "100", None),
                        prop(
                            "Caption",
                            "$\"frmMain.frx\" : 1000",
                            reference("frmMain.frx", 0x1000, true),
                        ),
                    ],
                    children: Vec::new(),
                },
            ],
        }),
    };
    assert_eq!(parse_designer(SAMPLE), expected);
}

#[test]
fn the_library_of_a_control_type() {
    let d = parse_designer(SAMPLE);
    let root = d.root.unwrap();
    assert_eq!(root.library(), Some("VB"));
    assert!(root.is_intrinsic());
    assert_eq!(root.children[1].library(), Some("MSComctlLib"));
    assert!(!root.children[1].is_intrinsic());

    let unqualified = DesignerControl {
        type_name: "Form".to_string(),
        name: "f".to_string(),
        properties: Vec::new(),
        children: Vec::new(),
    };
    assert_eq!(unqualified.library(), None);
    assert!(!unqualified.is_intrinsic());
}

#[test]
fn the_resource_references_are_found_with_their_control_paths() {
    let refs = parse_designer(SAMPLE).resource_refs();
    let summary: Vec<_> = refs
        .iter()
        .map(|r| {
            (
                r.control_path.as_str(),
                r.control_type.as_str(),
                r.property.as_str(),
                r.frx.offset,
                r.frx.dollar,
            )
        })
        .collect();
    assert_eq!(
        summary,
        vec![
            ("frmMain", "VB.Form", "Icon", 0, false),
            ("frmMain.lstItems", "VB.ListBox", "List", 0xC10, false),
            ("frmMain.lstItems", "VB.ListBox", "ItemData", 0xC9A, false),
            ("frmMain.tbMain", "MSComctlLib.Toolbar", "Caption", 0x1000, true),
        ]
    );
}

#[test]
fn a_trailing_comment_is_not_part_of_a_value() {
    let src = "VERSION 5.00\nBegin VB.Form f\n   Appearance = 0 'Flat\n   BorderStyle = 1 'Fixed Single\nEnd\n";
    let root = parse_designer(src).root.unwrap();
    assert_eq!(root.properties, vec![prop("Appearance", "0", None), prop("BorderStyle", "1", None)]);
}

#[test]
fn a_property_group_contributes_only_its_references_under_a_dotted_path() {
    let src = "VERSION 5.00\r\nBegin MSComctlLib.ImageList il\r\n   BeginProperty Images {D4A2-1}\r\n      NumListImages = 2\r\n      BeginProperty ListImage1 {D4A2-2}\r\n         Picture = \"f.frx\":0010\r\n         Key = \"a\"\r\n      EndProperty\r\n      BeginProperty ListImage2 {D4A2-2}\r\n         Picture = \"f.frx\":0030\r\n      EndProperty\r\n   EndProperty\r\n   BeginProperty Font\r\n      Name = \"Arial\"\r\n   EndProperty\r\n   Caption = \"x\"\r\nEnd\r\n";
    let root = parse_designer(src).root.unwrap();
    assert_eq!(
        root.properties,
        vec![
            prop(
                "Images.ListImage1.Picture",
                "\"f.frx\" : 0010",
                reference("f.frx", 0x10, false)
            ),
            prop(
                "Images.ListImage2.Picture",
                "\"f.frx\" : 0030",
                reference("f.frx", 0x30, false)
            ),
            prop("Caption", "\"x\"", None),
        ]
    );
}

#[test]
fn nothing_after_the_designer_block_becomes_a_control() {
    let src = "VERSION 5.00\nBegin VB.Form f\nEnd\nAttribute VB_Name = \"f\"\nPrivate Sub Go()\n    x = 1\nEnd Sub\n";
    let d = parse_designer(src);
    let root = d.root.unwrap();
    assert_eq!(root.name, "f");
    assert!(root.properties.is_empty());
    assert!(root.children.is_empty());
}

#[test]
fn an_object_line_without_a_companion_file_has_an_empty_file() {
    let src = "VERSION 5.00\nObject = \"{ABCD0000-0000-0000-0000-000000000000}#1.5#0\"\nBegin VB.Form f\nEnd\n";
    assert_eq!(
        parse_designer(src).objects,
        vec![ObjectEntry {
            clsid: "{ABCD0000-0000-0000-0000-000000000000}".to_string(),
            version: "1.5".to_string(),
            file: String::new(),
        }]
    );
}

#[test]
fn an_object_line_without_a_version_has_an_empty_version() {
    let src = "VERSION 5.00\nObject = \"{ABCD0000-0000-0000-0000-000000000000}\"; \"X.OCX\"\nBegin VB.Form f\nEnd\n";
    assert_eq!(
        parse_designer(src).objects,
        vec![ObjectEntry {
            clsid: "{ABCD0000-0000-0000-0000-000000000000}".to_string(),
            version: String::new(),
            file: "X.OCX".to_string(),
        }]
    );
}

#[test]
fn an_object_line_with_no_class_id_is_skipped() {
    let src = "VERSION 5.00\nObject = \"#2.0#0\"; \"X.OCX\"\nBegin VB.Form f\nEnd\n";
    assert_eq!(parse_designer(src).objects, Vec::new());
}

#[test]
fn several_object_lines_keep_their_order() {
    let src = "VERSION 5.00\nObject = \"{A}#1.0#0\"; \"A.OCX\"\nObject = \"{B}#2.0#0\"; \"B.OCX\"\nBegin VB.Form f\nEnd\n";
    let files: Vec<_> = parse_designer(src).objects.iter().map(|o| o.file.clone()).collect();
    assert_eq!(files, vec!["A.OCX".to_string(), "B.OCX".to_string()]);
}

#[test]
fn a_class_module_has_no_root_control() {
    let src = "VERSION 1.0 CLASS\nAttribute VB_Name = \"C\"\nPublic Sub M()\nEnd Sub\n";
    assert_eq!(parse_designer(src), FormDesigner::default());
}

#[test]
fn a_source_the_form_parser_cannot_read_gives_an_empty_designer() {
    for src in [
        "",
        "Begin VB.Form f\nEnd\n",
        "VERSION 5.00\nBegin VB.Form f\n   Caption = \"x\"\n",
        "VERSION 5.00\nObject = \nBegin VB.Form f\nEnd\n",
    ] {
        assert_eq!(parse_designer(src), FormDesigner::default(), "{src:?}");
    }
}

#[test]
fn lines_ended_by_a_carriage_return_alone_are_read() {
    let src = "VERSION 5.00\rBegin VB.Form f\r   Icon = \"f.frx\":0000\rEnd\r";
    let root = parse_designer(src).root.unwrap();
    assert_eq!(
        root.properties,
        vec![prop("Icon", "\"f.frx\" : 0000", reference("f.frx", 0, false))]
    );
}
