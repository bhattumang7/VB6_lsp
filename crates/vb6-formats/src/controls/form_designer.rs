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
//! This module recovers the control tree, the `Object=` type-library declarations
//! (CLSID + companion OCX), and every property value — flagging the ones that are
//! references into the `.frx`/`.ctx` companion (see [`crate::frx`]).

use crate::frm::{parse_frm, BeginBlock, ObjectRef, PropValue, Property};
use crate::frx::{self, FrxRef};

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

}

fn collect_refs(ctl: &DesignerControl, path: &mut Vec<String>, out: &mut Vec<ResourceRef>) {
    path.push(ctl.name.clone());
    for p in &ctl.properties {
        if let Some(frx) = &p.frx {
            out.push(ResourceRef {
                control_path: path.join("."),
                control_type: ctl.type_name.clone(),
                property: p.name.clone(),
                frx: frx.clone(),
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
/// This is a view of what the [`.frm` parser](crate::frm) reads: the control tree, the `Object =`
/// declarations, and the properties that point into a companion file. A property value is
/// its text without a trailing comment. A source the `.frm` parser cannot read (no `VERSION`
/// header, an unterminated block, a malformed `Object =` line) gives an empty designer.
pub fn parse_designer(source: &str) -> FormDesigner {
    let Ok(frm) = parse_frm(source) else {
        return FormDesigner::default();
    };
    FormDesigner {
        objects: frm.objects.iter().filter_map(object_entry).collect(),
        root: frm.root.as_ref().map(control),
    }
}

/// `{CLSID}#version#lcid` and the companion file of an `Object =` line; `None` when the
/// declaration has no CLSID.
fn object_entry(object: &ObjectRef) -> Option<ObjectEntry> {
    let mut fields = object.progid.split('#');
    let clsid = fields.next().unwrap_or("").trim().to_string();
    if clsid.is_empty() {
        return None;
    }
    let version = fields.next().unwrap_or("").trim().to_string();
    let file = object.filename.clone().unwrap_or_default();
    Some(ObjectEntry { clsid, version, file })
}

fn control(block: &BeginBlock) -> DesignerControl {
    let mut properties = Vec::new();
    for property in &block.properties {
        collect_property(property, &mut Vec::new(), &mut properties);
    }
    DesignerControl {
        type_name: block.control_type.clone(),
        name: block.name.clone(),
        properties,
        children: block.children.iter().map(control).collect(),
    }
}

/// Record a property of a control. Every plain property of the control is kept; inside a
/// `BeginProperty` group only the references into a companion file are, under their dotted
/// path (`Images.ListImage1.Picture`).
fn collect_property(property: &Property, groups: &mut Vec<String>, out: &mut Vec<DesignerProp>) {
    match &property.value {
        PropValue::Scalar(value) => {
            let frx = frx::parse_frx_reference(value);
            if groups.is_empty() {
                out.push(DesignerProp { name: property.name.clone(), value: value.clone(), frx });
            } else if frx.is_some() {
                let name = format!("{}.{}", groups.join("."), property.name);
                out.push(DesignerProp { name, value: value.clone(), frx });
            }
        }
        PropValue::Bag(members) => {
            groups.push(property.name.clone());
            for member in members {
                collect_property(member, groups, out);
            }
            groups.pop();
        }
    }
}

#[cfg(test)]
#[path = "tests/form_designer.rs"]
mod tests;
