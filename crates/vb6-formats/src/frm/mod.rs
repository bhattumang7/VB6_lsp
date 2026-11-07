/// VB6 `.frm` / `.cls` / `.ctl` / `.dob` / `.pag` text-file parser.
///
/// All VB6 designer files share the same top-level structure:
///   VERSION <major>.<minor>
///   Object = "{progid}"; "name"   (zero or more — OCX references)
///   Attribute VB_Name = "ModuleName"
///   Begin <TypeName> <ControlName>
///       <properties>
///       Begin <TypeName> <ControlName>   (nested controls)
///           <properties>
///       End
///   End
pub mod lexer;
pub mod parser;

pub use parser::{
    Attribute, BeginBlock, FrmError, FrmFile, ObjectRef, PropKind, PropValue, Property,
    parse_frm,
};
