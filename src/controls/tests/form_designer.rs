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

    #[test]
    fn parses_object_header() {
        let d = parse_designer(SAMPLE);
        assert_eq!(d.objects.len(), 1);
        assert_eq!(d.objects[0].clsid, "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}");
        assert_eq!(d.objects[0].version, "2.0");
        assert_eq!(d.objects[0].file, "MSCOMCTL.OCX");
    }

    #[test]
    fn parses_control_tree() {
        let d = parse_designer(SAMPLE);
        let root = d.root.as_ref().expect("root form");
        assert_eq!(root.type_name, "VB.Form");
        assert_eq!(root.name, "frmMain");
        assert!(root.is_intrinsic());
        // children: lstItems + tbMain (BeginProperty Font is not a control)
        assert_eq!(root.children.len(), 2);
        assert_eq!(root.children[0].name, "lstItems");
        assert_eq!(root.children[1].type_name, "MSComctlLib.Toolbar");
    }

    #[test]
    fn collects_frx_references() {
        let d = parse_designer(SAMPLE);
        let refs = d.resource_refs();
        // Icon, List, ItemData, Toolbar Caption ($)
        assert_eq!(refs.len(), 4);
        let icon = refs.iter().find(|r| r.property == "Icon").unwrap();
        assert_eq!(icon.control_path, "frmMain");
        assert_eq!(icon.frx.offset, 0);
        let cap = refs.iter().find(|r| r.property == "Caption").unwrap();
        assert!(cap.frx.dollar);
        assert_eq!(cap.frx.offset, 0x1000);
        assert_eq!(cap.control_path, "frmMain.tbMain");
    }

    #[test]
    fn font_group_not_treated_as_reference() {
        let d = parse_designer(SAMPLE);
        // The BeginProperty Font block must not leak Name/Size as form properties.
        let root = d.root.unwrap();
        assert!(!root.properties.iter().any(|p| p.name == "Size"));
    }

    #[test]
    fn stops_at_code_section() {
        let d = parse_designer(SAMPLE);
        // Only one root; nothing from the code section becomes a control.
        assert!(d.root.is_some());
    }
