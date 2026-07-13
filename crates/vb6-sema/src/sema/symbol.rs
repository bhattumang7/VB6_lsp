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
    /// For a `Const` whose initializer is a non-integer literal (String, Double,
    /// Single, Currency, Date, Boolean), the literal itself — so the code generator
    /// can fold it at each use site. Integer-valued consts use `const_value`; this
    /// is `None` for those and for non-const declarations.
    pub const_lit: Option<crate::frontend::ast::AstLit>,
    /// For a fixed-length string (`As String * n`), the declared length `n`. The
    /// type stays [`VbaType::String`] (a 4-byte pointer slot) but the assignment
    /// copy is length-aware, so the code generator needs the length.
    pub fixed_string_len: Option<u16>,
    /// For a fixed-size array, the number of declared dimensions (the SAFEARRAY
    /// descriptor size and the element-access opcode depend on it). `None` for
    /// non-arrays and dynamic (`Dim a()`) arrays.
    pub array_dims: Option<u16>,
    /// True for `Static` locals.
    pub is_static: bool,
    /// True if declared with `Public` or `Global`.
    pub is_public: bool,
    /// Source span of the variable-name identifier (for LSP go-to-definition).
    pub name_span: Span,
}

/// A known EXTERNAL class module's ordered member list, supplied by the
/// caller (a project/fixture-level compiler that has already bound the
/// class's own module) for cross-module member resolution — `Dim o As New
/// ClassName` / `o.Field`. Sema has no concept of a "class module" or a
/// project beyond this: a class's `Public` fields bind exactly like a
/// Standard module's (`BoundModule::module_vars`); this struct is just that
/// declaration-ordered list, handed back in for a DIFFERENT module's
/// `resolve_member_type` to consult.
///
/// `members` MUST be in strict source-declaration order across every kind
/// (fields, property accessors, procedures) — VB6's class vtable-slot
/// numbering is a single running counter over the whole class's declaration
/// sequence, not per-kind (confirmed live: a `Property Let` declared before
/// its own `Property Get` gets the earlier slot; a Sub/Function declared
/// between two fields still advances the counter positionally). See
/// `resolve_class_field`'s slot-numbering rule.
#[derive(Debug, Clone, Default)]
pub struct ExternalClass {
    pub members: Vec<ClassMemberSlot>,
}

/// Which accessor an explicit `Property` declaration is.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum AccessorKind {
    Get,
    Let,
    Set,
}

/// One class-member declaration that consumes one or more class vtable
/// dispatch slots, in the order it appears in the class's source. Vtable
/// layout is assigned by walking `ExternalClass::members` in order and
/// summing each entry's slot width (see `resolve_class_field`).
#[derive(Debug, Clone)]
pub enum ClassMemberSlot {
    /// A `Public` field. Synthesizes an implicit `Get`+`Let` (value type,
    /// two slots) or `Get`+`Let`+`Set` (object/Variant type, three slots) —
    /// oracle-confirmed: the middle `Let` slot exists for an object/Variant
    /// field even though a caller only ever calls `Get` or `Set` on it.
    Field {
        name: String,
        vba_type: VbaType,
        /// True for an object/reference-assignable type (Object or a class
        /// type) or Variant — these synthesize a `Set` accessor too.
        is_object: bool,
    },
    /// One explicit `Property Get`/`Let`/`Set` accessor (one slot). Multiple
    /// entries with the same `name` (possibly non-contiguous in the list —
    /// VB6 does not require a property's accessors to be declared together)
    /// together form one logical property.
    PropertyAccessor {
        name: String,
        vba_type: VbaType,
        kind: AccessorKind,
    },
    /// A `Public Sub`/`Function` declaration (one slot), called through the
    /// vtable like a Property accessor but with its own argument-staging
    /// convention (each argument — including ByVal — is staged into its own
    /// addressable temp before the call, right-to-left, byte-exact against
    /// `argcount_probe`/`funcarg_probe`; see the `vb6-class-vtable-slot-rule`
    /// memory note).
    Method {
        name: String,
        /// Return type (`Variant` for a `Sub`, matching `BoundProc::ret_type`'s
        /// convention for an untyped Function — a `Sub` is never used as a
        /// value so this is only consulted when calling a `Function`).
        ret_type: VbaType,
        /// Each parameter's declared type and `ByVal`-ness (`true` = ByVal,
        /// `false` = ByRef/default). Optional/ParamArray/array parameters are
        /// out of scope, matching the intra-module call path's own
        /// exact-arg-count restriction.
        params: Vec<(VbaType, bool)>,
    },
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

/// A single-argument intrinsic emitted as a dedicated, argument-type-indexed
/// opcode (`Len`/`Abs`/`Sgn`/`Int`/`Fix`).
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum UnaryIntrinsic {
    Len,
    Abs,
    Sgn,
    Int,
    Fix,
}

/// How a single argument of a String-returning runtime call ([`BuiltinCall::RtcString`])
/// is passed.
#[derive(Debug, Clone, PartialEq, Eq)]
pub enum RtcArg {
    /// Pushed by value, coerced to this type (e.g. `Chr`'s `Long` argument, or the
    /// start position of `Left`/`Mid`).
    ByVal(VbaType),
    /// Boxed into a hidden 16-byte temp tagged with the argument's runtime VARTYPE
    /// (e.g. the String argument of `UCase`/`Left`/`Mid`, or the Variant argument
    /// of `Str`/`Hex`/`Oct`). The argument must be a simple variable reference.
    Boxed,
    /// An omitted Optional Variant parameter: no source argument is supplied. The
    /// compiler materialises a hidden Missing variant temp (VT_ERROR /
    /// DISP_E_PARAMNOTFOUND) and passes it by address. The temp is freed alongside
    /// the result temp.
    MissingVariant,
}

/// How a built-in (intrinsic) call is emitted by the code generator.
#[derive(Debug, Clone, PartialEq)]
pub enum BuiltinCall {
    /// A type-conversion intrinsic (`CInt`/`CLng`/`CStr`/…): its single argument
    /// is converted to the given type, reusing the assignment-conversion opcodes.
    Convert(VbaType),
    /// A dedicated-opcode unary intrinsic; the opcode is selected by the argument
    /// type at lowering time.
    Unary(UnaryIntrinsic),
    /// A single-argument runtime-library call with a numeric result (`Asc`, `Sqr`,
    /// `Val`): emitted as a runtime call whose opcode is selected by the result
    /// type. `arg` is the argument type (for the size-based push), `ret` the result.
    RtcNumeric { arg: VbaType, ret: VbaType },
    /// A runtime-library call returning a String (`Chr`/`Space`/`UCase`/`Left`/
    /// `Mid`/`Str`/…). Each parameter is pushed either by value or boxed into a
    /// hidden 16-byte temp (see [`RtcArg`]); the result is produced into a final
    /// hidden string temp and moved to the target. `args` describes the
    /// passing mode of each parameter, in source order.
    RtcString { args: Vec<RtcArg> },
    /// `InStr` in its 2- or 3-argument form, returning a Long. Emitted as a
    /// dedicated opcode with four operands pushed in order — start (Long), string1,
    /// string2, compare-mode (Long) — where an omitted leading start defaults to
    /// literal 1 and the compare-mode defaults to literal 0 (`Option Compare
    /// Binary`). `three_arg` is true when an explicit start is supplied.
    Instr { three_arg: bool },
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
    /// Classification of each intrinsic (built-in) call, keyed by the `Call`
    /// node's `NodeId.0`. Lets the code generator emit the right form for a
    /// built-in without a name table.
    pub builtins: std::collections::HashMap<u32, BuiltinCall>,
    /// Semantic diagnostics detectable from this module alone (e.g. duplicate
    /// declaration in scope). Project-scoped diagnostics (e.g. "Variable not
    /// defined", which requires cross-module knowledge) are added by the session
    /// after cross-module resolution.
    pub diagnostics: Diagnostics,
    /// True if the module declares `Option Explicit` (drives the project-level
    /// undeclared-variable check).
    pub option_explicit: bool,
    /// Type-sym → matched external class (see [`ExternalClass`]), for every
    /// `UserType` reference the binder resolved against `bind_with_classes`'s
    /// `classes` table. Lets codegen recognize a `VbaType::UserDefined(sym)`
    /// declaration as a class instance (not a same-module UDT) by sym_id,
    /// with no scanner/interner access of its own.
    pub class_field_info: std::collections::HashMap<u32, ExternalClass>,
    /// Fully-resolved vtable dispatch info for each class-member
    /// `MemberAccess` node whose base resolves to a known external class
    /// instance, keyed by the `MemberAccess` node's OWN `NodeId.0` — not the
    /// member name's sym_id, which repeats across every occurrence of that
    /// name in source. Codegen has no scanner/interner to re-derive which
    /// class member a sym_id names, so sema resolves each access SITE
    /// individually here (by walking the matched `ExternalClass::members` in
    /// declaration order and summing slot widths — see [`ResolvedClassMember`]).
    pub class_member_slots: std::collections::HashMap<u32, ResolvedClassMember>,
}

/// Fully-resolved class-member vtable dispatch info for one `MemberAccess`
/// access site (see `BoundModule::class_member_slots`).
#[derive(Debug, Clone)]
pub struct ResolvedClassMember {
    /// Vtable byte offset of the `Get` accessor, when the member has one.
    pub get_slot: Option<u16>,
    /// Vtable byte offset of the `Let` accessor, when the member has one.
    pub let_slot: Option<u16>,
    /// Vtable byte offset of the `Set` accessor, when the member has one.
    pub set_slot: Option<u16>,
    /// Vtable byte offset of a `Sub`/`Function` method, when the member is one.
    pub method_slot: Option<u16>,
    /// The method's return type and each parameter's `(type, is_by_val)`, set
    /// only when `method_slot` is `Some` — codegen has no scanner/interner to
    /// re-derive this from the member name, so sema hands it over directly
    /// (mirrors how the slot itself is resolved once here, not per use).
    pub method_ret_type: Option<VbaType>,
    pub method_params: Vec<(VbaType, bool)>,
    /// True for an explicit `Property` accessor; false for a plain field. A
    /// property's `Let`/`Set` call stages its argument into a temp frame
    /// slot first (`0x59 <offset>`); a field's store does not.
    pub is_property: bool,
}
