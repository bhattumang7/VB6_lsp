//! Project-level context for member resolution.
//!
//! A single-module [`bind`](crate::sema::bind) knows only its own declarations.
//! [`bind_in_project`](crate::sema::bind_in_project) additionally takes a
//! [`ProjectScope`]: the host's view of every other module (and of any
//! external type library), so that `obj.Member` can be resolved against the
//! declared type of `obj`. Sema itself has no notion of a project beyond this
//! trait.

use crate::sema::symbol::{NameResolution, TypeName};
use crate::sema::types::VbaType;

/// Which accessor a member use needs: `Get` when read (or called), `Let` when
/// the target of `=`, `Set` when the target of `Set … =`.
///
/// `Let` is every statement that stores into the member: `=`, `LSet`, `RSet`,
/// the string target of `Mid(…) =`, the variable of `For` and `For Each`, and
/// the targets of `Input #`, `Line Input #` and `Get #`. A member passed as an
/// argument is `Get`: the site reads it, whether or not a `ByRef` callee later
/// writes through the reference. Only the outermost member of a target is
/// stored into (`a.b.c = 1` stores into `c` and reads `a` and `b`).
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
pub enum Access {
    Get,
    Let,
    Set,
}

/// What a bare name used as the base of a member access denotes when it is not
/// a value.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
pub enum Qualifier {
    /// A module of the project, named by its module name.
    Module(usize),
    /// The project itself, named by its project name.
    Project,
    /// A loaded external library (index into the host's library table).
    Library(usize),
    /// An enum type name: `Color.Red`.
    Enum { module: usize, idx: usize },
}

/// The static type of an expression, as far as the project can tell.
#[derive(Debug, Clone, PartialEq, Eq)]
pub enum SemType {
    /// A non-object value type (`Long`, `String`, …, never `Variant`/`Object`).
    Value(VbaType),
    /// `Variant`, `Object`, `Any`, or an untyped declaration: members resolve
    /// at run time and are never guessed.
    LateBound,
    /// A user-defined type declared in `module`, at `idx` of its type list.
    Udt { module: usize, idx: usize },
    /// An instance of a class, form or other designer module.
    Class { module: usize },
    /// An instance of a type of an external library.
    Lib { lib: usize, ty: usize },
    Array(Box<SemType>),
    /// A name that stands for a module, project, library or enum.
    Namespace(Qualifier),
    /// A named type that is not in the project (unqualified), by lowercase name.
    Unknown(String),
    /// A type qualified by a project or library that is not loaded.
    LibraryNotLoaded(String),
    /// An expression whose own resolution failed.
    Unresolved,
}

impl SemType {
    /// The element type when this is an array.
    pub fn element(&self) -> Option<&SemType> {
        match self {
            SemType::Array(inner) => Some(inner),
            _ => None,
        }
    }
}

/// Why a `MemberAccess` node is, or is not, resolved to a declaration.
#[derive(Debug, Clone, PartialEq, Eq)]
pub enum MemberStatus {
    /// The member resolved; the node has an entry in `resolutions`.
    Resolved,
    /// The base is `Variant`, `Object` or untyped: bound at run time.
    LateBound,
    /// `base!key`: a default-member lookup by name, bound at run time.
    Bang,
    /// The base is a value type (`Long`, `String`, an array, …).
    NotAnObject,
    /// The base's declared type is not a type of the project (unqualified name).
    UnknownType(String),
    /// The base's type, or the member's, lives in a library that is not loaded.
    LibraryNotLoaded(String),
    /// The base's class is known and has no reachable member of this name.
    NoSuchMember,
    /// The member exists but not with the accessor the use needs.
    MissingAccessor(Access),
    /// The base expression itself did not resolve.
    BaseUnresolved,
    /// The member resolved against a type inferred for the base: a procedure
    /// local of type `Variant`/`Object` that every `Set` in the procedure
    /// assigns an object of one and the same type, and that nothing else
    /// stores into or hands out by reference.
    Inferred,
    /// The base is such a local, but its `Set`s assign objects of different
    /// types: the member is not resolved. The candidates are the type names,
    /// in the order first assigned.
    Ambiguous(Vec<String>),
}

/// The parameter list of a procedure a name resolved to.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum Arity {
    /// The procedure declares exactly this many parameters.
    Known(usize),
    /// The procedure's parameters are not described (an external procedure
    /// whose library data omits them).
    Unknown,
}

/// A resolved member (or project-level name): what a use site points at and the
/// static type of the result.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct MemberHit {
    pub resolution: NameResolution,
    pub sem: SemType,
    /// The parameter list when the member is a procedure; `None` for a
    /// variable, field, control or constant.
    pub callable: Option<Arity>,
}

/// Outcome of a member lookup.
#[derive(Debug, Clone, PartialEq, Eq)]
pub enum MemberLookup {
    Found(MemberHit),
    /// The base has no member of this name reachable from the use site.
    NoSuchMember,
    /// The name exists but not as the accessor the use needs.
    MissingAccessor(Access),
    /// The base type's library is not loaded, so its members are unknown.
    LibraryNotLoaded(String),
}

/// The host's view of the project: every module's declarations and any loaded
/// external libraries. All `from` arguments are the module being bound.
pub trait ProjectScope {
    /// The type a declared type name denotes, seen from module `from`.
    fn sem_of_type_name(&self, from: usize, ty: &TypeName) -> SemType;

    /// The declaration a type name refers to: a `Type`, an `Enum`, a class
    /// module, or a library type; `Unresolved` when it names none.
    fn resolve_type_decl(&self, from: usize, ty: &TypeName) -> NameResolution;

    /// What the qualifier of a qualified type name (`Project.Class`,
    /// `Module.Type`, `Library.Type`) names; `Unresolved` when it names none.
    fn resolve_qualifier(&self, from: usize, qualifier: &str) -> NameResolution;

    /// How a type is written in source (`clsTrip`, `ADODB.Connection`).
    fn type_name(&self, ty: &SemType) -> String;

    /// The static type of `Me` in module `from`.
    fn self_type(&self, from: usize) -> SemType;

    /// A name that belongs to module `from`'s own scope but not to its parsed
    /// declarations (the controls of a form).
    fn lookup_module_scope(&self, from: usize, name: &str) -> Option<MemberHit>;

    /// A project-wide name: a public declaration of another module, a module
    /// name, the project name, or a library's global object.
    fn lookup_global(&self, from: usize, name: &str) -> Option<MemberHit>;

    /// Member `name` of a value of type `base`, accessed with `access`.
    fn lookup_member(&self, from: usize, base: &SemType, name: &str, access: Access) -> MemberLookup;

    /// The default member of a value of type `base` (`oTrips(1)`).
    fn default_member(&self, from: usize, base: &SemType) -> Option<MemberHit>;
}
