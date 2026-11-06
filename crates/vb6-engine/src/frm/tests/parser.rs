    use super::*;

    const MINIMAL_FRM: &str = r#"VERSION 5.00
Begin VB.Form Form1
   Caption         =   "Hello"
   ClientHeight    =   3600
   BackColor       =   &H80000005&
End
"#;

    #[test]
    fn parse_minimal_form() {
        let f = parse_frm(MINIMAL_FRM).unwrap();
        assert_eq!(f.version, "5.00");
        assert_eq!(f.kind, FileKind::Form);
        let root = f.root.unwrap();
        assert_eq!(root.control_type, "VB.Form");
        assert_eq!(root.name, "Form1");
        assert_eq!(root.properties.len(), 3);
        assert_eq!(root.properties[0].name, "Caption");
        match &root.properties[0].value {
            PropValue::Scalar(s) => assert_eq!(s, "\"Hello\""),
            _ => panic!("expected scalar"),
        }
    }

    #[test]
    fn parse_nested_controls() {
        let src = r#"VERSION 5.00
Begin VB.Form Form1
   Caption = "Test"
   Begin VB.CommandButton cmdOK
      Caption = "OK"
      Height  = 375
   End
End
"#;
        let f = parse_frm(src).unwrap();
        let root = f.root.unwrap();
        assert_eq!(root.children.len(), 1);
        assert_eq!(root.children[0].name, "cmdOK");
    }

    #[test]
    fn parse_object_ref() {
        // In .frm files the progid is in outer double-quotes
        let src = "VERSION 5.00\r\nObject = \"{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0\"; \"MSCOMCTL.OCX\"\r\nBegin VB.Form Form1\r\nEnd\r\n";
        let f = parse_frm(src).unwrap();
        assert_eq!(f.objects.len(), 1);
        assert!(f.objects[0].progid.contains("831FDD16"));
        assert_eq!(f.objects[0].filename, Some("MSCOMCTL.OCX".into()));
    }

    #[test]
    fn parse_begin_property() {
        let src = r#"VERSION 5.00
Begin VB.Form Form1
   BeginProperty Font
      Name = "MS Sans Serif"
      Size = 8.25
   EndProperty
End
"#;
        let f = parse_frm(src).unwrap();
        let root = f.root.unwrap();
        assert_eq!(root.properties.len(), 1);
        assert_eq!(root.properties[0].name, "Font");
        assert_eq!(root.properties[0].kind, PropKind::BeginProperty { guid: None });
        match &root.properties[0].value {
            PropValue::Bag(sub) => {
                assert_eq!(sub.len(), 2);
                assert_eq!(sub[0].name, "Name");
            }
            _ => panic!("expected bag"),
        }
    }

    #[test]
    fn parse_class_file() {
        let src = r#"VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
END
Attribute VB_Name = "Class1"
"#;
        let f = parse_frm(src).unwrap();
        assert_eq!(f.kind, FileKind::Class);
        assert_eq!(f.attributes.len(), 1);
        assert_eq!(f.attributes[0].name, "VB_Name");
        assert_eq!(f.attributes[0].value, "Class1");
    }
