    use super::*;

    #[test]
    fn lists_form_controls() {
        let src = "VERSION 5.00\n\
                   Begin VB.Form Form1\n\
                   Caption = \"Hi\"\n\
                   Begin VB.CommandButton cmdOK\n\
                   End\n\
                   End\n\
                   Attribute VB_Name = \"Form1\"\n";
        let controls = form_controls(src);
        assert_eq!(controls.len(), 2);
        assert_eq!(controls[0].name, "Form1");
        assert_eq!(controls[0].control_type, "VB.Form");
        assert_eq!(controls[0].depth, 0);
        assert_eq!(controls[1].name, "cmdOK");
        assert_eq!(controls[1].depth, 1);
        // Span points at the control name in source.
        let s = controls[1].span;
        assert_eq!(&src[s.start as usize..(s.start + s.len) as usize], "cmdOK");
    }

    #[test]
    fn lists_form_controls_in_a_file_with_carriage_return_line_breaks() {
        let src = "VERSION 5.00\rBegin VB.Form Form1\rCaption = \"Hi\"\rBegin VB.CommandButton cmdOK\rEnd\rEnd\rAttribute VB_Name = \"Form1\"\r";
        let controls = form_controls(src);
        assert_eq!(controls.len(), 2);
        assert_eq!((controls[0].name.as_str(), controls[0].control_type.as_str(), controls[0].depth), ("Form1", "VB.Form", 0));
        assert_eq!((controls[1].name.as_str(), controls[1].control_type.as_str(), controls[1].depth), ("cmdOK", "VB.CommandButton", 1));
        let s = controls[1].span;
        assert_eq!(&src[s.start as usize..(s.start + s.len) as usize], "cmdOK");
        let t = controls[1].type_span;
        assert_eq!(&src[t.start as usize..(t.start + t.len) as usize], "VB.CommandButton");
    }
