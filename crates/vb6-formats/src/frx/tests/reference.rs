use super::*;

#[test]
fn parse_both_forms() {
    let a = parse_frx_reference("\"frmMain.frx\":0000").unwrap();
    assert_eq!(a.file, "frmMain.frx");
    assert_eq!(a.offset, 0);
    assert!(!a.dollar);

    let b = parse_frx_reference("$\"frmMain.frx\":0BCA").unwrap();
    assert_eq!(b.offset, 0x0BCA);
    assert!(b.dollar);

    let c = parse_frx_reference("\"Ctl.ctx\":00AC").unwrap();
    assert_eq!(c.offset, 0xAC);

    assert!(parse_frx_reference("Just a caption").is_none());
}

#[test]
fn parse_rejects_malformed_references() {
    // No offset digits.
    assert!(parse_frx_reference("\"a.frx\":").is_none());
    assert!(parse_frx_reference("\"a.frx\":zz").is_none());
    // Missing colon, missing closing quote, wrong extension.
    assert!(parse_frx_reference("\"a.frx\" 0001").is_none());
    assert!(parse_frx_reference("\"a.frx:0001").is_none());
    assert!(parse_frx_reference("\"a.txt\":0001").is_none());
    // Offset does not fit in 32 bits.
    assert!(parse_frx_reference("\"a.frx\":123456789").is_none());
}

#[test]
fn parse_tolerates_spacing_case_and_trailing_text() {
    let r = parse_frx_reference("  $ \"A.FRX\" : 00ff trailing").unwrap();
    assert_eq!(r, FrxRef { file: "A.FRX".to_string(), offset: 0xFF, dollar: true });
}

#[test]
fn kind_mapping() {
    assert_eq!(kind_for_property("Icon", false), PropKind::Picture);
    assert_eq!(kind_for_property("Font", false), PropKind::Font);
    assert_eq!(kind_for_property("Caption", true), PropKind::StringLong);
    assert_eq!(kind_for_property("Caption", false), PropKind::StringShort);
    assert_eq!(kind_for_property("List", false), PropKind::List);
    assert_eq!(kind_for_property("ItemData", false), PropKind::ItemData);
    assert_eq!(kind_for_property("PropertyPages", false), PropKind::PropertyPages);
    assert_eq!(kind_for_property("_GridInfo", false), PropKind::OcxBag);
    assert_eq!(kind_for_property("Images.ListImage1.Picture", false), PropKind::Picture);
    assert_eq!(kind_for_property("TabPicture(0)", false), PropKind::Picture);
}

#[test]
fn kind_mapping_unlisted_names_follow_the_dollar_flag() {
    assert_eq!(kind_for_property("Whatever", true), PropKind::StringLong);
    assert_eq!(kind_for_property("Whatever", false), PropKind::OcxBag);
    assert_eq!(kind_for_property("Bands", true), PropKind::OcxBag);
    assert_eq!(kind_for_property("ToolTipText", true), PropKind::StringLong);
    assert_eq!(kind_for_property("ToolTipText", false), PropKind::StringShort);
}
