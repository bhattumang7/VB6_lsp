//! Description of an external type library, supplied to a session by its host.
//!
//! A project references libraries (the VB runtime, ADO, OCX controls, other
//! DLLs) whose types are not VB6 source and so are not in the session. The host
//! describes them with this plain, JSON-serialisable model, produced however it
//! likes (source analysis, a one-time type-library dump). The session reads no
//! type-library format itself.
//!
//! Names are matched case-insensitively. A coclass exposes the members of its
//! default interface, then of the other interfaces it lists; an interface also
//! exposes the members of every interface it `extends`.

use serde::{Deserialize, Serialize};

/// One external library.
#[derive(Debug, Clone, PartialEq, Eq, Serialize, Deserialize)]
pub struct Library {
    /// The name VB6 code qualifies the library's types with (`ADODB`, `VB`).
    pub name: String,
    pub types: Vec<LibraryType>,
}

/// The kind of a library type.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Default, Serialize, Deserialize)]
#[serde(rename_all = "lowercase")]
pub enum LibraryTypeKind {
    /// A creatable class or a plain object type.
    #[default]
    Class,
    Interface,
    /// A coclass: its members come from `default_interface` and `interfaces`.
    Coclass,
    /// A module of global functions and constants.
    Module,
    /// An enumeration: its members are global constants.
    Enum,
    /// A user-defined type (a record): its members are its fields.
    Record,
}

/// How a class of an ActiveX project is exposed, as VB6's `Instancing`
/// property sets it.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Serialize, Deserialize)]
pub enum Instancing {
    /// Public, but only the component creates instances.
    PublicNotCreatable,
    /// Creatable; each instance comes from a new copy of the component.
    SingleUse,
    /// As `SingleUse`, and the members are global names of the client.
    GlobalSingleUse,
    /// Creatable; one copy of the component serves every instance.
    MultiUse,
    /// As `MultiUse`, and the members are global names of the client.
    GlobalMultiUse,
}

/// The kind of project a `.vbp` builds (its `Type=` line).
#[derive(Debug, Clone, Copy, PartialEq, Eq, Serialize, Deserialize)]
pub enum ProjectKind {
    /// A Standard EXE: nothing of it is visible to other projects.
    Exe,
    /// An ActiveX DLL.
    OleDll,
    /// An ActiveX EXE.
    OleExe,
    /// An ActiveX control (OCX).
    Control,
}

impl ProjectKind {
    /// The kind a `.vbp` `Type=` value names (`Exe`, `OleDll`, `OleExe`,
    /// `Control`), case-insensitively.
    pub fn from_vbp_type(value: &str) -> Option<Self> {
        match value.trim().to_ascii_lowercase().as_str() {
            "exe" => Some(ProjectKind::Exe),
            "oledll" => Some(ProjectKind::OleDll),
            "oleexe" => Some(ProjectKind::OleExe),
            "control" => Some(ProjectKind::Control),
            _ => None,
        }
    }
}

/// A class, interface, coclass, module, enumeration or record of a library.
#[derive(Debug, Clone, PartialEq, Eq, Default, Serialize, Deserialize)]
pub struct LibraryType {
    pub name: String,
    pub kind: LibraryTypeKind,
    /// For a coclass, the interface its members come from.
    #[serde(default)]
    pub default_interface: Option<String>,
    /// For a coclass, the further interfaces it implements, searched after the
    /// default interface.
    #[serde(default)]
    pub interfaces: Vec<String>,
    /// For an interface, the interfaces it extends.
    #[serde(default)]
    pub extends: Vec<String>,
    /// The library provides a global instance named like the type (`App`,
    /// `Screen`, `Err`).
    #[serde(default)]
    pub is_app_object: bool,
    /// The type is the library's global object: its members are project-wide
    /// names (`App`, `Screen`, `LoadPicture`), as the members of a module are.
    #[serde(default)]
    pub is_global: bool,
    /// The type gains members at run time (a form with its controls and
    /// public procedures): a member the library does not describe is looked
    /// up when the program runs instead of being rejected.
    #[serde(default)]
    pub extensible: bool,
    /// The type is an ActiveX control (`TYPEFLAG_FCONTROL`): the container adds the extender's
    /// members (`Left`, `Top`, `Tag`, ...) to the ones the library describes.
    #[serde(default, skip_serializing_if = "std::ops::Not::not")]
    pub is_control: bool,
    /// For a class of an ActiveX project, how it is exposed (`None` when the
    /// library does not say).
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub instancing: Option<Instancing>,
    #[serde(default)]
    pub members: Vec<LibraryMember>,
}

/// The kind of a library member.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Default, Serialize, Deserialize)]
#[serde(rename_all = "lowercase")]
pub enum LibraryMemberKind {
    /// A `Sub` or `Function`.
    #[default]
    Method,
    PropertyGet,
    PropertyLet,
    PropertySet,
    /// A data member readable and writable.
    Field,
    /// A constant (an enum member or a module constant).
    Constant,
    /// An event the type raises: not a member a client calls, but one it
    /// handles (`WithEvents`).
    Event,
}

/// A parameter of a library member.
#[derive(Debug, Clone, PartialEq, Eq, Default, Serialize, Deserialize)]
pub struct LibraryParam {
    pub name: String,
    /// The declared type, as [`LibraryMember::return_type`] spells types;
    /// `None` when undeclared (a Variant).
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub type_name: Option<String>,
    /// Passed by reference.
    #[serde(default)]
    pub by_ref: bool,
    #[serde(default)]
    pub optional: bool,
    /// The default value of an optional parameter, as VB source text.
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub default: Option<String>,
    /// A `ParamArray`.
    #[serde(default)]
    pub param_array: bool,
}

/// A member of a library type.
#[derive(Debug, Clone, PartialEq, Eq, Default, Serialize, Deserialize)]
pub struct LibraryMember {
    pub name: String,
    pub kind: LibraryMemberKind,
    /// The type name of the result: a type of the same library, or qualified
    /// (`Library.Type`). `None` for a `Sub`; `Object`, `Variant` and
    /// `IDispatch` are late-bound.
    #[serde(default)]
    pub return_type: Option<String>,
    /// The member is the type's default member (DISPID 0): `o(1)` is
    /// `o.Item(1)`.
    #[serde(default)]
    pub is_default: bool,
    /// The number of declared parameters of a procedure member. `None` when
    /// the host does not describe them; a call's arguments are then never
    /// taken to index the result.
    #[serde(default)]
    pub param_count: Option<usize>,
    /// The parameters, when the host describes them; `param_count` is then
    /// their number.
    #[serde(default, skip_serializing_if = "Vec::is_empty")]
    pub params: Vec<LibraryParam>,
    /// The member's dispatch id when the library assigns one explicitly
    /// (`0` for the default member).
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub dispid: Option<i32>,
    /// Type-library member flags (hidden, restricted, ...), as the library
    /// sets them.
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub member_flags: Option<u32>,
    /// The value of a constant (an enum member), when the library gives it.
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub value: Option<i64>,
}

/// The runtime libraries every VB6 project sees, described by the engine.
const VBA_JSON: &str = include_str!("../../data/vba.json");
const VB_JSON: &str = include_str!("../../data/vb.json");
const VBRUN_JSON: &str = include_str!("../../data/vbrun.json");
const STDOLE_JSON: &str = include_str!("../../data/stdole.json");

impl Library {
    /// The libraries the engine describes itself, which every VB6 project
    /// sees: `VBA` (strings, conversion, math, date and file functions,
    /// constants, `Err`, `Collection`), `VB` (`Form`, `UserControl`, the
    /// intrinsic controls, `App`, `Screen`, `Printer`, the global procedures),
    /// `VBRUN` (the runtime objects) and `stdole` (`StdFont`, `StdPicture`).
    pub fn bundled() -> Vec<Library> {
        [VBA_JSON, VB_JSON, VBRUN_JSON, STDOLE_JSON]
            .iter()
            .flat_map(|text| {
                Library::from_json(text).expect("bundled library data is valid JSON")
            })
            .collect()
    }

    /// Parse one library, or an array of libraries, from JSON.
    pub fn from_json(text: &str) -> Result<Vec<Library>, serde_json::Error> {
        let value: serde_json::Value = serde_json::from_str(text)?;
        if value.is_array() {
            serde_json::from_value(value)
        } else {
            serde_json::from_value(value).map(|l| vec![l])
        }
    }

    /// Render libraries as a JSON array.
    pub fn to_json(libraries: &[Library]) -> Result<String, serde_json::Error> {
        serde_json::to_string(libraries)
    }
}

/// What a session is told about its surroundings beyond the source files.
///
/// The default describes a project that references only the runtime libraries
/// the engine bundles (see [`Library::bundled`]); name no libraries at all
/// with `libraries: Vec::new()`.
#[derive(Debug, Clone, PartialEq, Eq, Serialize, Deserialize)]
pub struct SessionConfig {
    /// The project's name (`Project1`), which qualifies its classes and
    /// modules (`Project1.clsTrip`).
    #[serde(default)]
    pub project_name: Option<String>,
    /// The libraries the project references.
    #[serde(default)]
    pub libraries: Vec<Library>,
}

impl Default for SessionConfig {
    fn default() -> Self {
        SessionConfig { project_name: None, libraries: Library::bundled() }
    }
}
