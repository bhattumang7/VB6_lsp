//! Semantic analysis: name resolution, type annotation, symbol tables.
//!
//! Entry point: [`bind`].

pub mod binder;
pub mod builtins;
pub mod symbol;
pub mod types;

pub use binder::{bind, unbound_namerefs};
pub use symbol::{
    BoundEnumDecl, BoundEnumMember, BoundModule, BoundParam, BoundProc, BoundTypeDecl,
    BoundTypeMember, BoundVar, BuiltinCall, ExternalDecl, NameResolution, ParamFlags,
};
pub use types::VbaType;
