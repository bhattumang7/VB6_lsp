//! Semantic analysis: name resolution, type annotation, symbol tables.
//!
//! Entry point: [`bind`].

pub mod binder;
pub mod builtins;
pub mod project;
pub mod symbol;
pub mod types;

pub use binder::{
    bind, bind_in_project, bind_with_classes, unaccounted_member_accesses, unbound_namerefs,
};
pub use project::{
    Access, Arity, MemberHit, MemberLookup, MemberStatus, ProjectScope, Qualifier, SemType,
};
pub use symbol::{
    AccessorKind, BoundEnumDecl, BoundEnumMember, BoundModule, BoundParam, BoundProc,
    BoundTypeDecl, BoundTypeMember, BoundVar, BuiltinCall, ClassMemberSlot, DeclTypeNames,
    ExternalClass, ExternalDecl, NameResolution, ParamFlags, ResolvedClassMember, RtcArg, TypeName,
    UnaryIntrinsic,
};
pub use types::VbaType;
