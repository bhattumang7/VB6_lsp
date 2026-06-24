//! Bridge from the vb6-sema binder's symbol model to the codegen reference
//! emitter.
//!
//! vb6-sema resolves a name to a [`vb6_sema::VbaType`] and a storage kind
//! (`NameResolution`); it does **not** compute frame offsets (those come from
//! [`crate::bind::ProcFrame`] / [`crate::bind::ParamFrame`], which reproduce
//! VB6's exact frame layout). This module maps a `VbaType` onto the two codegen
//! quantities the reference path needs:
//!
//! * the **frame type-context** ([`type_ctx`]) — drives `ProcFrame`'s slot
//!   sizing/alignment;
//! * the **load/store context** ([`load_store_ctx`]) — selects the oracle-
//!   confirmed load/store opcode from the `RT_LOAD_BY_CTX` / `RT_STORE_BY_CTX`
//!   tables in [`crate::emit`].
//!
//! Three storage classes are bridged:
//!
//! * **Locals** (`NameResolution::Local`): negative frame offsets, same opcodes
//!   for ByVal and ByRef.
//! * **Parameters** (`NameResolution::Param`): positive frame offsets starting
//!   at +12; ByVal uses the same opcodes as locals; ByRef uses opcode+0x14.
//! * **Module globals** (`NameResolution::ModuleVar`): 4-byte operand
//!   `[module_desc][field_offset]`, opcodes = local_opcode+0x28.

use vb6_sema::sema::VbaType;

use crate::bind::{LocalVar, ParamFrame, ParamVar, ProcFrame};
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
        // Boolean is stored exactly as Integer (2-byte, Integer-class storage and
        // load/store opcodes — same node tag 6).
        VbaType::Integer | VbaType::Boolean => 1,
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

// ── Parameter bridge ──────────────────────────────────────────────────────────

/// Emit a ByVal parameter load.  ByVal parameters use the same opcodes as
/// locals but have positive frame offsets (first param at +12).
pub fn emit_byval_param_load(
    emitter: &mut Emitter,
    ty: &VbaType,
    frame_offset: i16,
) -> Result<(), UnsupportedType> {
    let ctx = load_store_ctx(ty).ok_or(UnsupportedType)?;
    emitter.emit_typed_load(ctx, frame_offset);
    Ok(())
}

/// Emit a ByVal parameter store.
pub fn emit_byval_param_store(
    emitter: &mut Emitter,
    ty: &VbaType,
    frame_offset: i16,
) -> Result<(), UnsupportedType> {
    let ctx = load_store_ctx(ty).ok_or(UnsupportedType)?;
    emitter.emit_var_store(ctx, frame_offset);
    Ok(())
}

/// Emit a ByRef parameter load.  ByRef parameter opcodes are
/// `RT_LOAD_BY_CTX[ctx] + 0x14` (oracle-confirmed for Long: 0x6c→0x80).
pub fn emit_byref_param_load(
    emitter: &mut Emitter,
    ty: &VbaType,
    frame_offset: i16,
) -> Result<(), UnsupportedType> {
    let ctx = load_store_ctx(ty).ok_or(UnsupportedType)?;
    emitter.emit_byref_load(ctx, frame_offset);
    Ok(())
}

/// Emit a ByRef parameter store.  ByRef parameter store opcodes are
/// `RT_STORE_BY_CTX[ctx] + 0x14` (oracle-confirmed for Long: 0x71→0x85).
pub fn emit_byref_param_store(
    emitter: &mut Emitter,
    ty: &VbaType,
    frame_offset: i16,
) -> Result<(), UnsupportedType> {
    let ctx = load_store_ctx(ty).ok_or(UnsupportedType)?;
    emitter.emit_byref_store(ctx, frame_offset);
    Ok(())
}

/// Allocate a procedure's parameter frame from the binder's parameter list,
/// taken in declaration order (left-to-right).  The returned `Vec<ParamVar>`
/// is indexed by the binder's `param_idx`.
///
/// Returns `Err(UnsupportedType)` if any parameter has a type whose frame size
/// is not yet confirmed.
pub fn param_frame_from_types(
    types: &[VbaType],
    byref_flags: &[bool],
) -> Result<Vec<ParamVar>, UnsupportedType> {
    debug_assert_eq!(types.len(), byref_flags.len());
    let mut frame = ParamFrame::new();
    let mut out = Vec::with_capacity(types.len());
    for (ty, &byref) in types.iter().zip(byref_flags.iter()) {
        let ctx = type_ctx(ty).ok_or(UnsupportedType)?;
        out.push(frame.declare_anon_param(ctx, byref));
    }
    Ok(out)
}

/// Emit a load of the parameter resolved to `param_idx`.
pub fn emit_resolved_param_load(
    emitter: &mut Emitter,
    param_idx: usize,
    types: &[VbaType],
    slots: &[ParamVar],
) -> Result<(), UnsupportedType> {
    let ty = &types[param_idx];
    let slot = &slots[param_idx];
    if slot.byref {
        emit_byref_param_load(emitter, ty, slot.frame_offset)
    } else {
        emit_byval_param_load(emitter, ty, slot.frame_offset)
    }
}

/// Emit a store to the parameter resolved to `param_idx`.
pub fn emit_resolved_param_store(
    emitter: &mut Emitter,
    param_idx: usize,
    types: &[VbaType],
    slots: &[ParamVar],
) -> Result<(), UnsupportedType> {
    let ty = &types[param_idx];
    let slot = &slots[param_idx];
    if slot.byref {
        emit_byref_param_store(emitter, ty, slot.frame_offset)
    } else {
        emit_byval_param_store(emitter, ty, slot.frame_offset)
    }
}

// ── Module global bridge ──────────────────────────────────────────────────────

/// Emit a module-level global variable load.  Opcodes are
/// `RT_LOAD_BY_CTX[ctx] + 0x28` (oracle-confirmed: Integer=0x93, Long=0x94,
/// Double=0x97).  `module_desc` is the compiled module-object descriptor (the
/// 2-byte value the compiled form assigns to this module); `field_offset` is
/// the byte offset of this variable within the module's global data block.
pub fn emit_global_var_load(
    emitter: &mut Emitter,
    ty: &VbaType,
    module_desc: u16,
    field_offset: u16,
) -> Result<(), UnsupportedType> {
    let ctx = load_store_ctx(ty).ok_or(UnsupportedType)?;
    emitter.emit_global_load(ctx, module_desc, field_offset);
    Ok(())
}

/// Emit a module-level global variable store (mirror of
/// [`emit_global_var_load`]).  Opcodes are `RT_STORE_BY_CTX[ctx] + 0x28`
/// (oracle-confirmed: Integer=0x98, Long=0x99, Double=0x9c).
pub fn emit_global_var_store(
    emitter: &mut Emitter,
    ty: &VbaType,
    module_desc: u16,
    field_offset: u16,
) -> Result<(), UnsupportedType> {
    let ctx = load_store_ctx(ty).ok_or(UnsupportedType)?;
    emitter.emit_global_store(ctx, module_desc, field_offset);
    Ok(())
}

#[cfg(test)]
#[path = "tests/bridge_tests.rs"]
mod tests;
