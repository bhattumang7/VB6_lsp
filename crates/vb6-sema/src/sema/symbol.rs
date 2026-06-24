//! Symbol table entries produced by the binder.

use crate::frontend::ast::{ProcKind, Span};
use crate::frontend::diagnostics::Diagnostics;
use crate::sema::types::VbaType;

/// Flags describing how a parameter is passed.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Default)]
pub struct ParamFlags {
    /// Parameter is declared with `ByVal` (copy semantics).
    pub by_val: bool,
    /// Parameter is declared with `ByRef` (reference semantics; VBA default).
    pub by_ref: bool,
    /// Parameter is `Optional`.
    pub optional: bool,
    /// Parameter is a dynamic array (`name()`).
    pub is_array: bool,
    /// Parameter is `ParamArray`.
    pub param_array: bool,
}

impl ParamFlags {
    pub fn from_bits(flags: u16) -> Self {
        ParamFlags {
            optional:    flags & 0x01 != 0,
            by_val:      flags & 0x02 != 0,
            by_ref:      flags & 0x04 != 0,
            is_array:    flags & 0x08 != 0,
            param_array: flags & 0x20 != 0,
        }
    }
}

/// A bound procedure parameter.
#[derive(Debug, Clone)]
pub struct BoundParam {
    /// Interned name index (scanner sym_id).
    pub sym_id: u32,
    /// Declared type (Variant if no `As` clause).
    pub vba_type: VbaType,
    pub flags: ParamFlags,
    /// Source span of the parameter-name identifier (for LSP go-to-definition).
    pub name_span: Span,
}

/// A bound variable (module-level or local).
#[derive(Debug, Clone)]
pub struct BoundVar {
    /// Interned name index (scanner sym_id).
    pub sym_id: u32,
    /// Declared type (Variant if no `As` clause).
    pub vba_type: VbaType,
    /// True for `Const` declarations.
    pub is_const: bool,
    /// For a `Const` whose initializer folds to an integer, its value (so the
    /// code generator can emit a folded literal at each use site). `None` for
    /// non-const declarations and consts whose value is not an integer constant.
    pub const_value: Option<i64>,
    /// True for `Static` locals.
    pub is_static: bool,
    /// True if declared with `Public` or `Global`.
    pub is_public: bool,
    /// Source span of the variable-name identifier (for LSP go-to-definition).
    pub name_span: Span,
}

/// A member of a user-defined `Type ... End Type`.
#[derive(Debug, Clone)]
pub struct BoundTypeMember {
    pub sym_id: u32,
    pub vba_type: VbaType,
    /// Source span of the member-name identifier (for LSP go-to-definition).
    pub name_span: Span,
}

/// A bound `Type ... End Type` declaration.
#[derive(Debug, Clone)]
pub struct BoundTypeDecl {
    pub sym_id: u32,
    pub members: Vec<BoundTypeMember>,
    pub is_public: bool,
    /// Source span of the type-name identifier (for LSP go-to-definition).
    pub name_span: Span,
}

/// A member of an `Enum ... End Enum`.
#[derive(Debug, Clone)]
pub struct BoundEnumMember {
    pub sym_id: u32,
    /// Resolved constant value.
    pub value: i64,
    /// Source span of the member-name identifier (for LSP go-to-definition).
    pub name_span: Span,
}

/// A bound `Enum ... End Enum` declaration.
#[derive(Debug, Clone)]
pub struct BoundEnumDecl {
    pub sym_id: u32,
    pub members: Vec<BoundEnumMember>,
    pub is_public: bool,
    /// Source span of the enum-name identifier (for LSP go-to-definition).
    pub name_span: Span,
}

/// A bound procedure (Sub, Function, or Property).
#[derive(Debug, Clone)]
pub struct BoundProc {
    pub sym_id: u32,
    pub kind: ProcKind,
    pub params: Vec<BoundParam>,
    /// Return type (Variant for Subs and untyped Functions).
    pub ret_type: VbaType,
    /// Local variable declarations found in the body.
    pub locals: Vec<BoundVar>,
    /// NodeId of the body block node.
    pub body: u32,
    pub is_public: bool,
    /// Source span of the procedure-name identifier (for LSP go-to-definition).
    pub name_span: Span,
}

/// Resolution of a single name-reference occurrence.
#[derive(Debug, Clone, PartialEq, Eq)]
pub enum NameResolution {
    /// Resolved to a local variable in procedure `proc_idx`, local `local_idx`.
    Local { proc_idx: usize, local_idx: usize },
    /// Resolved to a parameter in procedure `proc_idx`, param `param_idx`.
    Param { proc_idx: usize, param_idx: usize },
    /// Resolved to a module-level variable at `var_idx`.
    ModuleVar(usize),
    /// Resolved to a module-level procedure at `proc_idx`.
    Proc(usize),
    /// Resolved to an enum member.
    EnumMember { enum_idx: usize, member_idx: usize },
    /// Name matches a known VB6 built-in.
    Builtin,
    /// Resolved to a public declaration in another module of the project.
    ///
    /// Produced by the project-level cross-module resolution pass (the session
    /// layer), not by single-module [`bind`](crate::sema::bind), which leaves
    /// such names [`Unresolved`](NameResolution::Unresolved).
    External { module: usize, decl: ExternalDecl },
    /// Could not be resolved (unresolved forward reference or external name).
    Unresolved,
}

/// Which kind of public declaration an [`NameResolution::External`] points at,
/// and its index within that module's corresponding `BoundModule` vector.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum ExternalDecl {
    /// Index into the module's `procs`.
    Proc(usize),
    /// Index into the module's `module_vars`.
    Var(usize),
    /// Index into the module's `type_decls`.
    Type(usize),
    /// Index into the module's `enum_decls`.
    Enum(usize),
    /// A member of a public enum (enum members are project-scoped constants).
    EnumMember { enum_idx: usize, member_idx: usize },
}

/// The fully-bound representation of one VBA module.
#[derive(Debug, Default, Clone)]
pub struct BoundModule {
    /// All procedures in declaration order.
    pub procs: Vec<BoundProc>,
    /// Module-level variable declarations.
    pub module_vars: Vec<BoundVar>,
    /// User-defined type declarations.
    pub type_decls: Vec<BoundTypeDecl>,
    /// Enum declarations.
    pub enum_decls: Vec<BoundEnumDecl>,
    /// Name resolution for each `NameRef` node, keyed by `NodeId.0`.
    pub resolutions: std::collections::HashMap<u32, NameResolution>,
    /// Inferred type for each expression node, keyed by `NodeId.0`.
    pub types: std::collections::HashMap<u32, VbaType>,
    /// Semantic diagnostics detectable from this module alone (e.g. duplicate
    /// declaration in scope). Project-scoped diagnostics (e.g. "Variable not
    /// defined", which requires cross-module knowledge) are added by the session
    /// after cross-module resolution.
    pub diagnostics: Diagnostics,
    /// True if the module declares `Option Explicit` (drives the project-level
    /// undeclared-variable check).
    pub option_explicit: bool,
}
