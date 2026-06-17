//! Bridge from the vb6-sema binder's symbol model to the codegen reference
//! emitter.
//!
//! vb6-sema resolves a name to a [`vb6_sema::VbaType`] and a storage kind
//! ([`vb6_sema::NameResolution`]); it does **not** compute frame offsets (those
//! come from [`crate::bind::ProcFrame`], which reproduces VB6's exact frame
//! layout). This module maps a `VbaType` onto the two codegen quantities the
//! reference path needs:
//!
//! * the **frame type-context** ([`type_ctx`]) — drives `ProcFrame`'s slot
//!   sizing/alignment;
//! * the **value class** ([`value_class`]) — the `nType` index
//!   [`crate::Emitter::emit_reference`] feeds to the load/store opcode formula.
//!
//! The value class is one of the quantities the reference resolver
//! (`EbResolveIdentRef`) produces inside the real compiler; here we supply it
//! directly from the declared type. Only the types whose load/store flow
//! through `EbEmitExpression2`'s simple offset path are mapped — Single, Double,
//! String, Date, Object, Variant, and UDTs resolve through the value-class
//! expression branch that is not yet ported, so [`value_class`] returns `None`
//! for them rather than guessing.

use vb6_sema::sema::VbaType;

use crate::bind::{LocalVar, ProcFrame};
use crate::emit::Emitter;

/// A declared type whose local load/store the bridge cannot yet emit (no
/// confirmed simple load/store opcode — e.g. String/Byte use runtime-helper
/// call sequences, and Date/Variant/Object/UDT are not yet mapped).
#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub struct UnsupportedType;

/// Map a `VbaType` to the frame type-context consumed by
/// [`crate::bind::ProcFrame::declare_local`].
///
/// Returns `None` for types whose frame sizing is not yet confirmed (Date,
/// Variant, Decimal, arrays) — see `frame_size_of_ctx` in [`crate::bind`].
pub fn type_ctx(t: &VbaType) -> Option<usize> {
    Some(match t {
        VbaType::Object => 0,
        VbaType::Integer | VbaType::Boolean | VbaType::Byte => 1,
        VbaType::Long => 2,
        VbaType::Single => 3,
        VbaType::Double => 4,
        VbaType::String => 5,
        VbaType::Currency => 6,
        VbaType::Date | VbaType::Variant | VbaType::Decimal => return None,
        VbaType::UserDefined(_) | VbaType::Array(_) => return None,
    })
}

/// The load/store type-context for a type that has a confirmed simple
/// (single-opcode) typed load and store. This is the same indexing as
/// [`type_ctx`] but restricted to the numeric primitives whose load/store
/// opcodes are oracle-confirmed (`RT_LOAD_BY_CTX` / `RT_STORE_BY_CTX`):
/// Integer→1, Long→2, Single→3, Double→4, Currency→6.
///
/// Returns `None` for String/Byte (which assign via runtime-helper sequences,
/// not a single load/store opcode) and for Boolean/Date/Object/Variant/UDT
/// (not yet confirmed) — the bridge reports those as [`UnsupportedType`] rather
/// than emit an unverified opcode.
pub fn load_store_ctx(t: &VbaType) -> Option<usize> {
    Some(match t {
        VbaType::Integer => 1,
        VbaType::Long => 2,
        VbaType::Single => 3,
        VbaType::Double => 4,
        VbaType::Currency => 6,
        _ => return None,
    })
}

/// Emit a typed load of a local variable at `frame_offset`, opcode chosen from
/// `ty`. Errors with [`UnsupportedType`] for a type without a confirmed simple
/// load opcode.
pub fn emit_local_load(
    emitter: &mut Emitter,
    ty: &VbaType,
    frame_offset: i16,
) -> Result<(), UnsupportedType> {
    let ctx = load_store_ctx(ty).ok_or(UnsupportedType)?;
    emitter.emit_typed_load(ctx, frame_offset);
    Ok(())
}

/// Emit a typed store of a local variable (mirror of [`emit_local_load`]).
pub fn emit_local_store(
    emitter: &mut Emitter,
    ty: &VbaType,
    frame_offset: i16,
) -> Result<(), UnsupportedType> {
    let ctx = load_store_ctx(ty).ok_or(UnsupportedType)?;
    emitter.emit_var_store(ctx, frame_offset);
    Ok(())
}

/// Allocate a procedure's local frame from the binder's locals, taken in
/// declaration order. The returned `Vec<LocalVar>` is indexed by the binder's
/// `local_idx` (as in [`vb6_sema::sema::NameResolution::Local`]), so a resolved
/// local maps directly to its frame slot.
///
/// Returns `Err(UnsupportedType)` if any local has a type whose frame size is
/// not yet confirmed ([`type_ctx`] is `None`) — the whole frame layout would be
/// wrong past that point, so we refuse rather than guess.
pub fn frame_from_local_types(types: &[VbaType]) -> Result<Vec<LocalVar>, UnsupportedType> {
    let mut frame = ProcFrame::new();
    let mut out = Vec::with_capacity(types.len());
    for ty in types {
        let ctx = type_ctx(ty).ok_or(UnsupportedType)?;
        out.push(frame.declare_anon(ctx));
    }
    Ok(out)
}

/// Emit a load of the local variable resolved to `local_idx`, given the proc's
/// declared local types (declaration order) and its allocated frame slots (from
/// [`frame_from_local_types`]). This is the resolution→emit path for a
/// [`vb6_sema::sema::NameResolution::Local`].
pub fn emit_resolved_local_load(
    emitter: &mut Emitter,
    local_idx: usize,
    types: &[VbaType],
    slots: &[LocalVar],
) -> Result<(), UnsupportedType> {
    emit_local_load(emitter, &types[local_idx], slots[local_idx].frame_offset)
}

/// Emit a store of the local resolved to `local_idx` (mirror of
/// [`emit_resolved_local_load`]).
pub fn emit_resolved_local_store(
    emitter: &mut Emitter,
    local_idx: usize,
    types: &[VbaType],
    slots: &[LocalVar],
) -> Result<(), UnsupportedType> {
    emit_local_store(emitter, &types[local_idx], slots[local_idx].frame_offset)
}

#[cfg(test)]
#[path = "tests/bridge_tests.rs"]
mod tests;
