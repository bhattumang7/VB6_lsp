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
#[derive(Debug, Clone, Copy, PartialEq, Eq, Serialize, Deserialize)]
#[serde(rename_all = "lowercase")]
pub enum LibraryTypeKind {
    /// A creatable class or a plain object type.
    Class,
    Interface,
    /// A coclass: its members come from `default_interface` and `interfaces`.
    Coclass,
    /// A module of global functions and constants.
    Module,
    /// An enumeration: its members are global constants.
    Enum,
}

/// A class, interface, coclass, module or enumeration of a library.
#[derive(Debug, Clone, PartialEq, Eq, Serialize, Deserialize)]
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
    /// A generic object type: a value declared of this type has members that
    /// are only looked up when the program runs, as with `Object`.
    #[serde(default)]
    pub late_bound: bool,
    #[serde(default)]
    pub members: Vec<LibraryMember>,
}

/// The kind of a library member.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Serialize, Deserialize)]
#[serde(rename_all = "lowercase")]
pub enum LibraryMemberKind {
    /// A `Sub` or `Function`.
    Method,
    PropertyGet,
    PropertyLet,
    PropertySet,
    /// A data member readable and writable.
    Field,
    /// A constant (an enum member or a module constant).
    Constant,
}

/// A member of a library type.
#[derive(Debug, Clone, PartialEq, Eq, Serialize, Deserialize)]
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
