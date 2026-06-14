/// VB6 `.vbp` project file parser.
///
/// VBP files are line-oriented `Key=Value` text files (no sections).
/// Some values are quoted (`"..."`), others are bare integers or paths.
///
/// The component list (Form, Module, Class, UserControl, etc.) comes first,
/// one entry per line.  Project-level settings follow.
pub mod parser;

pub use parser::{
    Module, ModuleKind, OcxObject, ProjectFile, Reference, ResFile, VbpError,
    parse_vbp,
};
