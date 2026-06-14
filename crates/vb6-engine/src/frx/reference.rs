/// A reference to a blob inside a companion `.frx` / `.ctx` file,
/// as written on a `.frm` property line.
///
/// Examples:
/// ```text
/// Icon    =   "frmMain.frx":0000
/// Caption =   $"frmMain.frx":0BCA
/// ```
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct FrxRef {
    /// Companion file name, e.g. `frmMain.frx`.
    pub file: String,
    /// Byte offset into the companion file.
    pub offset: u32,
    /// `true` for the `$"...":N` form (4-byte-prefixed long string).
    pub dollar: bool,
}

/// What kind of FRX blob a property reference points to.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum PropKind {
    Picture,
    Font,
    /// Short string: 1-byte length prefix (no `$` in `.frm`).
    StringShort,
    /// Long string: 4-byte length prefix (`$"..."` in `.frm`).
    StringLong,
    List,
    ItemData,
    PropertyPages,
    /// Unknown / vendor control bag — keep opaque.
    OcxBag,
}

/// Parse an FRX reference from a `.frm` property value string.
///
/// Accepts both forms: `"file.frx":NNNN` and `$"file.frx":NNNN`.
pub fn parse_frx_reference(value: &str) -> Option<FrxRef> {
    let v = value.trim();
    let (dollar, rest) = if let Some(r) = v.strip_prefix('$') {
        (true, r.trim_start())
    } else {
        (false, v)
    };
    let rest = rest.strip_prefix('"')?;
    let (file, after) = rest.split_once('"')?;
    let after = after.trim_start();
    let off_str = after.strip_prefix(':')?.trim();
    let hex: String = off_str.chars().take_while(|c| c.is_ascii_hexdigit()).collect();
    if hex.is_empty() {
        return None;
    }
    let offset = u32::from_str_radix(&hex, 16).ok()?;
    let fl = file.to_ascii_lowercase();
    if !fl.ends_with(".frx") && !fl.ends_with(".ctx") {
        return None;
    }
    Some(FrxRef { file: file.to_string(), offset, dollar })
}

/// Map a designer property name (plus the `$` flag from the `.frm` line) to
/// the expected blob kind so the parser knows which deserializer to invoke.
pub fn kind_for_property(name: &str, dollar: bool) -> PropKind {
    // Take the last path segment (e.g. "Images.ListImage1.Picture" → "Picture")
    // and strip a trailing "(N)" index (e.g. "TabPicture(0)" → "TabPicture").
    let base = name.rsplit('.').next().unwrap_or(name);
    let base = base.split('(').next().unwrap_or(base);
    match base.trim().to_ascii_lowercase().as_str() {
        "picture" | "icon" | "image" | "mouseicon" | "dragicon" | "toolboxbitmap"
        | "disabledpicture" | "downpicture" | "maskpicture" | "tabpicture" => PropKind::Picture,
        "font" | "mousefont" => PropKind::Font,
        "list" => PropKind::List,
        "itemdata" => PropKind::ItemData,
        "propertypages" => PropKind::PropertyPages,
        "_gridinfo" | "bands" | "column" | "sortkey" | "fmtcondition" | "formatstyle"
        | "template" | "printerproperties" | "initbuttons" | "bindings"
        | "initlistimages" => PropKind::OcxBag,
        "caption" | "text" | "textrtf" | "tag" | "tooltiptext" | "title" => {
            if dollar { PropKind::StringLong } else { PropKind::StringShort }
        }
        _ => {
            if dollar { PropKind::StringLong } else { PropKind::OcxBag }
        }
    }
}

#[cfg(test)]
mod tests {
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
}
