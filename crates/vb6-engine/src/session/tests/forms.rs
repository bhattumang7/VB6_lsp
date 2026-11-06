    use super::*;
    use crate::frx::records::PictureRecord;
    use crate::frx::records::StdFontRecord;

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

    /// Build a minimal `.frx` containing a single record at offset 0.
    fn frx_with(record: &FrxRecord) -> Vec<u8> {
        let mut out = Vec::new();
        record.write(&mut out);
        out
    }

    #[test]
    fn describes_picture_reference() {
        // A 4-byte BMP-magic "image".
        let pic = FrxRecord::Picture(PictureRecord { clsid: None, data: &[0x42, 0x4D, 0x01, 0x02] });
        let frx = frx_with(&pic);
        let desc = describe_frx_reference("Icon", "\"Form1.frx\":0000", &frx).unwrap();
        assert_eq!(desc, "Picture: BMP (4 bytes)");
    }

    #[test]
    fn describes_font_reference() {
        let font = FrxRecord::Font(StdFontRecord {
            charset: 0,
            flags: 0x01, // italic
            weight: 700, // bold
            size_times_10k: 82_500,
            name: b"MS Sans Serif".to_vec(),
        });
        let frx = frx_with(&font);
        let desc = describe_frx_reference("Font", "\"Form1.frx\":0000", &frx).unwrap();
        assert_eq!(desc, "Font: MS Sans Serif 8.25pt, bold italic");
    }

    #[test]
    fn non_frx_value_is_none() {
        assert!(describe_frx_reference("Caption", "\"plain text\"", &[]).is_none());
    }
