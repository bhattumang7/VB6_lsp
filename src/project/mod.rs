//! Opening a VB6 project or project group: path identity, the loaded project model, and a
//! snapshot of it for editors.

pub mod info;
pub mod model;
pub mod paths;

pub use model::{
    LoadError, MemberFile, Overlay, Problem, ReferenceInfo, ReferenceKind, ReferenceStatus,
    Selection, Severity, Unit,
};
