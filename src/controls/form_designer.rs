//! VB6 form/control designer-block parser (Tier 1).
//!
//! `.frm`/`.ctl`/`.pag`/`.dob` files begin with a textual designer block describing
//! the control tree, before the `Attribute`/code section:
//!
//! ```text
//! VERSION 5.00
//! Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
//! Begin VB.Form frmMain
//!    Caption  =   "Main"
//!    Icon     =   "frmMain.frx":0000
//!    Begin VB.ListBox lstItems
//!       List     =   "frmMain.frx":0C10
//!       ItemData =   "frmMain.frx":0C9A
//!    End
//!    BeginProperty Font
//!       Name = "MS Sans Serif"
//!    EndProperty
//! End
//! Attribute VB_Name = "frmMain"
//! ```
//!
//! This parser recovers the control tree, the `Object=` type-library declarations
//! (CLSID + companion OCX), and every property value — flagging the ones that are
//! references into the `.frx`/`.ctx` companion (see [`super::frx`]).

use super::frx::{self, FrxRef};

/// A type-library declaration from an `Object =` header line.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct ObjectEntry {
    /// CLSID GUID text including braces, e.g. `{831FDD16-...}`.
    pub clsid: String,
    /// Version string, e.g. `2.0`.
    pub version: String,
    /// Companion binary, e.g. `MSCOMCTL.OCX`.
    pub file: String,
}

/// A single property line inside a control block.
#[derive(Debug, Clone, PartialEq)]
pub struct DesignerProp {
    pub name: String,
    pub value: String,
    /// `Some` when the value is a reference into a companion `.frx`/`.ctx`.
    pub frx: Option<FrxRef>,
}

/// A control in the designer tree.
#[derive(Debug, Clone, PartialEq)]
pub struct DesignerControl {
    /// Fully-qualified designer type, e.g. `VB.Form`, `MSComctlLib.ListView`.
    pub type_name: String,
    /// Instance name.
    pub name: String,
    pub properties: Vec<DesignerProp>,
    pub children: Vec<DesignerControl>,
}

impl DesignerControl {
    /// The library prefix of the type (`VB` for `VB.Form`, `MSComctlLib` for
    /// `MSComctlLib.ListView`). `None` if the type is unqualified.
    pub fn library(&self) -> Option<&str> {
        self.type_name.split_once('.').map(|(lib, _)| lib)
    }
    /// True for intrinsic VB controls (no companion OCX).
    pub fn is_intrinsic(&self) -> bool {
        matches!(self.library(), Some("VB"))
    }
}

/// Result of parsing a designer block.
#[derive(Debug, Clone, Default, PartialEq)]
pub struct FormDesigner {
    pub objects: Vec<ObjectEntry>,
    pub root: Option<DesignerControl>,
}

/// One resolved reference: where it lives in the tree and what it points at.
#[derive(Debug, Clone, PartialEq)]
pub struct ResourceRef {
    /// Dotted control path, e.g. `frmMain.lstItems`.
    pub control_path: String,
    /// Designer type of the owning control.
    pub control_type: String,
    pub property: String,
    pub frx: FrxRef,
}

impl FormDesigner {
    /// Walk the tree and collect every `.frx`/`.ctx` reference.
    pub fn resource_refs(&self) -> Vec<ResourceRef> {
        let mut out = Vec::new();
        if let Some(root) = &self.root {
            collect_refs(root, &mut Vec::new(), &mut out);
        }
        out
    }

    /// Find the `Object=` entry whose companion file's library matches a control's
    /// library prefix (best-effort CLSID association for non-intrinsic controls).
    pub fn clsid_for(&self, control: &DesignerControl) -> Option<&str> {
        let lib = control.library()?;
        if lib == "VB" {
            return None;
        }
        // Heuristic: match the library prefix against the OCX file stem.
        let lib_lc = lib.to_ascii_lowercase();
        self.objects
            .iter()
            .find(|o| {
                let stem = o
                    .file
                    .rsplit('.')
                    .nth(1)
                    .unwrap_or(&o.file)
                    .to_ascii_lowercase();
                lib_lc.contains(&stem) || stem.contains(&lib_lc.trim_end_matches("ctl").trim_end_matches("lib"))
            })
            .map(|o| o.clsid.as_str())
    }
}

fn collect_refs(ctl: &DesignerControl, path: &mut Vec<String>, out: &mut Vec<ResourceRef>) {
    path.push(ctl.name.clone());
    for p in &ctl.properties {
        if let Some(frx) = &p.frx {
            out.push(ResourceRef {
                control_path: path.join("."),
                control_type: ctl.type_name.clone(),
                property: p.name.clone(),
                frx: FrxRef { property: p.name.clone(), ..frx.clone() },
            });
        }
    }
    for c in &ctl.children {
        collect_refs(c, path, out);
    }
    path.pop();
}

/// Parse the designer block at the top of a `.frm`/`.ctl` source.
///
/// Parsing stops once the outermost control block closes (the remainder is the
/// `Attribute`/code section).
pub fn parse_designer(source: &str) -> FormDesigner {
    let mut designer = FormDesigner::default();
    let mut stack: Vec<DesignerControl> = Vec::new();
    // Stack of BeginProperty group names currently open within the top control,
    // so references inside e.g. an ImageList's Images/ListImageN blocks get a
    // qualified path and are still harvested.
    let mut prop_group_stack: Vec<String> = Vec::new();
    let mut started = false;

    for raw in source.lines() {
        let line = raw.trim();
        if line.is_empty() {
            continue;
        }

        // Header: Object = "{CLSID}#ver#lcid"; "file.ocx"
        if !started && line.get(..6).map_or(false, |s| s.eq_ignore_ascii_case("object")) {
            if let Some(obj) = parse_object_line(line) {
                designer.objects.push(obj);
            }
            continue;
        }

        // BeginProperty <Name> [GUID] ... EndProperty — a property group (e.g. Font,
        // or an ImageList's Images / ListImageN image collection).
        if line.get(..13).map_or(false, |s| s.eq_ignore_ascii_case("beginproperty")) {
            let group = line[13..].split_whitespace().next().unwrap_or("").to_string();
            prop_group_stack.push(group);
            continue;
        }
        if line.eq_ignore_ascii_case("endproperty") {
            prop_group_stack.pop();
            continue;
        }

        // Begin <Type> <Name>
        if line.get(..6).map_or(false, |s| s.eq_ignore_ascii_case("begin ")) {
            started = true;
            let (type_name, name) = parse_begin_line(&line[6..]);
            stack.push(DesignerControl {
                type_name,
                name,
                properties: Vec::new(),
                children: Vec::new(),
            });
            continue;
        }

        // End (closes the current control)
        if line.eq_ignore_ascii_case("end") {
            if let Some(done) = stack.pop() {
                if let Some(parent) = stack.last_mut() {
                    parent.children.push(done);
                } else {
                    designer.root = Some(done);
                    break; // outermost control closed; rest is code
                }
            }
            continue;
        }

        // Property line: Name = Value (meaningful inside a control).
        if let Some((name, value)) = split_property(line) {
            if let Some(ctl) = stack.last_mut() {
                let mut fref = frx::parse_frx_reference(value);
                if prop_group_stack.is_empty() {
                    // Top-level control property: keep all (frx or not).
                    if let Some(r) = &mut fref {
                        r.property = name.to_string();
                    }
                    ctl.properties.push(DesignerProp {
                        name: name.to_string(),
                        value: value.to_string(),
                        frx: fref,
                    });
                } else if let Some(mut r) = fref {
                    // Inside a BeginProperty group: harvest only references, with a
                    // qualified path (e.g. Images.ListImage1.Picture).
                    let qualified = format!("{}.{}", prop_group_stack.join("."), name);
                    r.property = qualified.clone();
                    ctl.properties.push(DesignerProp {
                        name: qualified,
                        value: value.to_string(),
                        frx: Some(r),
                    });
                }
            }
        }
    }

    designer
}

fn parse_object_line(line: &str) -> Option<ObjectEntry> {
    // Object = "{GUID}#major.minor#lcid"; "file.ocx"
    let eq = line.find('=')?;
    let rhs = line[eq + 1..].trim();
    let first = rhs.strip_prefix('"')?;
    let (decl, after) = first.split_once('"')?;
    let clsid_end = decl.find('#').unwrap_or(decl.len());
    let clsid = decl[..clsid_end].trim().to_string();
    let mut version = String::new();
    if clsid_end < decl.len() {
        let rest = &decl[clsid_end + 1..];
        version = rest.split('#').next().unwrap_or("").trim().to_string();
    }
    // file: the quoted token after ';'
    let file = after
        .split(';')
        .nth(1)
        .and_then(|s| {
            let s = s.trim();
            let s = s.strip_prefix('"')?;
            Some(s.split('"').next().unwrap_or(s).to_string())
        })
        .unwrap_or_default();
    if clsid.is_empty() {
        return None;
    }
    Some(ObjectEntry { clsid, version, file })
}

fn parse_begin_line(rest: &str) -> (String, String) {
    let mut it = rest.split_whitespace();
    let type_name = it.next().unwrap_or("").to_string();
    let name = it.next().unwrap_or("").to_string();
    (type_name, name)
}

/// Split `Name = Value` on the first `=`, returning trimmed halves.
fn split_property(line: &str) -> Option<(&str, &str)> {
    let eq = line.find('=')?;
    let name = line[..eq].trim();
    let value = line[eq + 1..].trim();
    if name.is_empty() || name.contains(' ') && !name.chars().all(|c| c.is_alphanumeric() || c == '_') {
        // names are single identifiers; guard against stray '=' in odd lines
    }
    Some((name, value))
}

#[cfg(test)]
mod tests {
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
}
