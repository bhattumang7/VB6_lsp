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
    let (dollar, rest) = split_dollar_flag(value.trim());
    let (file, after) = rest.strip_prefix('"')?.split_once('"')?;
    if !is_companion_file(file) {
        return None;
    }
    let offset = parse_hex_prefix(after.trim_start().strip_prefix(':')?.trim())?;
    Some(FrxRef { file: file.to_string(), offset, dollar })
}

/// Separate a leading `$` flag from the rest of a property value.
fn split_dollar_flag(v: &str) -> (bool, &str) {
    match v.strip_prefix('$') {
        Some(r) => (true, r.trim_start()),
        None => (false, v),
    }
}

/// Whether `file` names an FRX or CTX companion file (any case).
fn is_companion_file(file: &str) -> bool {
    let lower = file.to_ascii_lowercase();
    lower.ends_with(".frx") || lower.ends_with(".ctx")
}

/// The value of the leading run of hex digits, `None` when there is none or
/// it does not fit in 32 bits.
fn parse_hex_prefix(text: &str) -> Option<u32> {
    let end = text.find(|c: char| !c.is_ascii_hexdigit()).unwrap_or(text.len());
    if end == 0 {
        return None;
    }
    u32::from_str_radix(&text[..end], 16).ok()
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
            string_or(dollar, PropKind::StringShort)
        }
        _ => string_or(dollar, PropKind::OcxBag),
    }
}

/// The long-string kind for the `$` form, otherwise `short`.
fn string_or(dollar: bool, short: PropKind) -> PropKind {
    if dollar { PropKind::StringLong } else { short }
}

#[cfg(test)]
#[path = "tests/reference.rs"]
mod tests;
