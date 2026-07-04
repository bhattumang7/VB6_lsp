//! Semantic analysis: name resolution, type annotation, symbol tables.
//!
//! Entry point: [`bind`].

pub mod binder;
pub mod builtins;
pub mod symbol;
pub mod types;

pub use binder::{bind, bind_with_classes, unbound_namerefs};
pub use symbol::{
    BoundEnumDecl, BoundEnumMember, BoundModule, BoundParam, BoundProc, BoundTypeDecl,
    BoundTypeMember, BoundVar, BuiltinCall, ExternalClass, ExternalDecl, NameResolution,
    ParamFlags, RtcArg, UnaryIntrinsic,
};
pub use types::VbaType;
