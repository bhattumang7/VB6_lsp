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

use crate::emit::{Emitter, RefDescriptor};

/// A declared type that the reference emitter cannot yet lower (its load/store
/// goes through a not-yet-ported branch of `EbEmitExpression2`).
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

/// Map a `VbaType` to the VB6 value class used as `emit_reference`'s `nType`.
///
/// Returns `None` for types whose resolved reference does not flow through
/// `EbEmitExpression2`'s simple offset path (Single/Double/String/Date/Object/
/// Variant/UDT/array) — those need the value-class expression branch, not yet
/// ported. The mapped classes (Integer 6, Long 8, Currency 12) are
/// oracle-confirmed via the load/store byte vectors.
pub fn value_class(t: &VbaType) -> Option<i32> {
    Some(match t {
        VbaType::Integer => 6,
        VbaType::Long => 8,
        VbaType::Currency => 0xc,
        _ => return None,
    })
}

/// Emit a typed load of a local variable: a kind-1 reference at `frame_offset`,
/// with the value class derived from `ty`. Errors with [`UnsupportedType`] for
/// a type whose load is not yet lowerable.
pub fn emit_local_load(
    emitter: &mut Emitter,
    ty: &VbaType,
    frame_offset: i16,
) -> Result<(), UnsupportedType> {
    let class = value_class(ty).ok_or(UnsupportedType)?;
    let desc = RefDescriptor {
        kind: 1,
        operand: frame_offset as u16,
        word6: 0,
    };
    emitter.emit_reference(&desc, 1, 0, class); // nOp 1 = value load
    Ok(())
}

/// Emit a typed store of a local variable (mirror of [`emit_local_load`],
/// nOp 4 = store).
pub fn emit_local_store(
    emitter: &mut Emitter,
    ty: &VbaType,
    frame_offset: i16,
) -> Result<(), UnsupportedType> {
    let class = value_class(ty).ok_or(UnsupportedType)?;
    let desc = RefDescriptor {
        kind: 1,
        operand: frame_offset as u16,
        word6: 0,
    };
    emitter.emit_reference(&desc, 4, 0, class); // nOp 4 = store
    Ok(())
}

#[cfg(test)]
#[path = "tests/bridge_tests.rs"]
mod tests;
