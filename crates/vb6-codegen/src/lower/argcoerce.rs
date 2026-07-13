//! Structural port of `EbEmitArgCoerce` (VBA6.DLL `@0fabc1b5`, decompiled at
//! `vba6_part0002.c:5906`, 1210 bytes) — the per-argument coercion body every
//! call site (intra-module, class-method/vtable, Property) ultimately routes
//! through. **Multi-session port, IN PROGRESS.** See the
//! `vb6-class-vtable-slot-rule` memory note's "WORD-FORM PORT" section for
//! the full status, the callee inventory, and the continuation plan — this
//! header only orients a reader of the code itself.
//!
//! ## Why this isn't (and can't be) a byte-for-byte transliteration
//!
//! The original operates directly on VBA6.DLL's internal compiler memory
//! layout: raw pointer arithmetic into a parameter-descriptor table
//! (`*piVar2 + 0x19 + argSlot`, `pProcDesc + RT_TYPE_OFFSET[..] * 2`, …) and a
//! **word-form expression node** whose exact field layout this port assumes
//! is the SAME 10-word/40-byte record already modeled by [`crate::node::
//! RawNode`] (confirmed consistent across everything else read this session:
//! `word[0]` low16 = opcode, high16 = type tag; `word[1]` low16 = flags;
//! etc.). What this codebase does NOT have — and would need, as a
//! prerequisite to a truly byte-for-byte port — is a Rust model of VBA6.DLL's
//! *parameter-descriptor table* memory layout (the structure `pProcDesc`
//! points into, indexed by `argSlot`). This codebase instead already models
//! "a method's parameters" as `Vec<(VbaType, bool)>` (`ClassMemberSlot::
//! Method.params` / `BoundParam`) — a semantic, not byte-layout, model.
//!
//! So this port translates EbEmitArgCoerce's *decision tree* (the branches:
//! is the source a literal needing synthesis, is it wrapped in a `0x70`/
//! `0x71` indirection node, which of the ~13 `local_18` (`EbGetExpressionType2`
//! result) cases applies, …) into terms of this codebase's own types
//! (`ExprNode`, `VbaType`, `RawNode`/`NodeArena`), rather than reproducing
//! the raw offset arithmetic on structures that don't exist here. Every
//! branch that bottoms out in a not-yet-ported callee is an explicit,
//! named [`LowerError`] gate citing the exact C function/address needed —
//! never a guessed byte.

use super::*;

/// Which of `EbEmitArgCoerce`'s callees a given call path needs, ported or
/// not. Kept as an explicit enum (rather than a bare string) so a future
/// continuation session can `grep` this file for `Callee::` and get the
/// complete, precise inventory of what remains.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub(super) enum UnportedCallee {
    /// `EbParseLiteralValue @0fabdcf7` (398B, `vba6_part0002.c:7910`) —
    /// synthesizes a literal-value node for an omitted-Optional argument's
    /// default value (the `pArgNode == 0` branch). Needed for `Optional`
    /// support; see `argoptional_probe` in the memory note for the already-
    /// confirmed oracle behavior (default value substituted as a literal at
    /// the call site) this callee is presumed to implement.
    ParseLiteralValue,
    /// `EbValidateArgumentType @0faecb32` (783B, `vba6_part0003.c:14997`) —
    /// the `local_c == 0x1b` (array-type-node) path's validator.
    ValidateArgumentType,
    /// `EbCheckDispatchProperty @0fac1458` (132B, `vba6_part0002.c:11890`) —
    /// gates the `local_10 != 0` (source was a `0x71`-wrapped node) path for
    /// `local_18 == 7`.
    CheckDispatchProperty,
    /// `EbResolveTypeBinding2 @0fabde33` (1044B, `vba6_part0002.c:8018`) —
    /// the `local_18 == 7` (by far the most common — plain value/reference
    /// type) path's core binding resolver. The single highest-value callee
    /// to port next: covers the ordinary scalar/Object/String argument case
    /// this session's oracle-grounded `lower_class_method_call` already
    /// handles empirically — porting this would let the two implementations
    /// be cross-checked against each other.
    ResolveTypeBinding2,
    /// `EbProcessArrayNode @0fafe90e` (339B, `vba6_part0003.c:41957`) — the
    /// `local_18 == 8` (array-type argument) path.
    ProcessArrayNode,
    /// `EbResolveSetNode @0facb5cc` (567B, `vba6_part0002.c:24580`) — the
    /// `local_18 == 0xc` (Variant/object-slot) path's non-`0x71`-wrapped case.
    ResolveSetNode,
    /// `EbCheckSetBinding @0fac09f2` (81B, `vba6_part0002.c:11206`) +
    /// `EbEmitPropertyExpr @0fac0a2d` (960B, `vba6_part0002.c:11224`) — the
    /// `local_18 == 9` (`0x1d`-taggable) shared tail every `0x1d`/object-slot
    /// path funnels through.
    SetBindingAndPropertyExpr,
    /// `EbBuildTypeCoercion @0fabdb91` (355B, `vba6_part0002.c:7819`) — the
    /// final explicit-conversion wrap applied when `flags & 2 == 0`.
    BuildTypeCoercion,
    /// `EbCompileTypeExpression2 @0fac6df0` (582B, `vba6_part0002.c:18715`)
    /// — the `flags & 0x200` ("compile as a type expression") alternate tail.
    CompileTypeExpression2,
    /// `EbGetTypePropertyFromWalk @0fabcb68` (46B, `vba6_part0002.c:6457`) +
    /// `EbProcessOperatorTree @0faed45b` (402B, `vba6_part0003.c:15829`) —
    /// the final operator-tree post-pass (`LAB_0fabc2e6`'s tail), run on
    /// every path's result before returning.
    FinalOperatorTreeWalk,
}

impl UnportedCallee {
    pub(super) fn as_lower_error(self) -> LowerError {
        // A dedicated variant per callee would let a caller distinguish
        // *which* piece of EbEmitArgCoerce blocked a given call — deferred
        // until a second caller (beyond `lower_class_method_call`) actually
        // needs that distinction; `UnsupportedNode` is the correct external
        // signal either way (this call cannot be lowered byte-exact yet).
        LowerError::UnsupportedNode
    }
}

/// `local_18` in the original — `EbGetExpressionType2`'s result, an already-
/// ported classifier (see `resolver.rs`'s existing port). Named here to keep
/// this file's branch structure legible against the original's `if
/// (local_18 == N)` chain without re-deriving the meaning of each `N` inline.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
#[allow(dead_code)] // populated as each branch below is ported
enum ExprTypeClass {
    /// `local_18 == 4`: a plain/simple expression — routes to
    /// `EbEmitExpression4` (already read this session, `vba6_part0001.c:
    /// 6111`, itself only partially traced).
    Simple = 4,
    /// `local_18 == 5`: routes straight to `EbEmitExprWithTypeAlt(0x9ccc,
    /// ...)` — an error/diagnostic path (type mismatch), not a value-
    /// producing path; safe to treat as a hard reject.
    Diagnostic5 = 5,
    /// `local_18 == 7`: the common case (plain value/reference argument) —
    /// see [`UnportedCallee::ResolveTypeBinding2`].
    Common = 7,
    /// `local_18 == 8`: array-typed argument — see
    /// [`UnportedCallee::ProcessArrayNode`].
    Array = 8,
    /// `local_18 == 9` / `0xc`: Variant/object-slot argument — see
    /// [`UnportedCallee::ResolveSetNode`] /
    /// [`UnportedCallee::SetBindingAndPropertyExpr`].
    ObjectSlot9 = 9,
    ObjectSlotC = 0xc,
}

/// Placeholder entry point — NOT YET CALLED from any live lowering path.
/// Ported so far: the function's overall shape and the literal-vs-wrapped-
/// node discrimination at its head (`LAB_0fabc209` and earlier, `vba6_part
/// 0002.c:5927-5959`). Everything past `EbResolveAttributePointer` bottoms
/// out in an [`UnportedCallee`] gate. Continuing this port means, in order
/// of value: (1) `ResolveTypeBinding2` (the common case), (2)
/// `FinalOperatorTreeWalk` (needed by every path including the common one),
/// (3) the remaining rarer branches.
#[allow(dead_code, unused_variables)]
pub(super) fn eb_emit_arg_coerce(
    ctx: &LowerCtx,
    arg_id: NodeId,
    param_ty: &VbaType,
    flags: u32,
    expr_arena: &ExprArena,
    arena: &mut NodeArena,
) -> Result<NodeRef, LowerError> {
    // `pArgNode == 0` in the original: no already-evaluated node was handed
    // in, meaning this argument was OMITTED (a missing Optional) and must be
    // synthesized from its declared default value. This codebase has no
    // default-value expression modeled on `BoundParam` yet (see the memory
    // note's Optional section) — `Optional` support wholly depends on this
    // gate closing first, in EITHER this port or the sema layer.
    // (No `is_omitted` signal reaches this function yet; this stands as the
    // named landing point for when one does.)
    let _ = UnportedCallee::ParseLiteralValue;

    // `local_10` in the original: whether the source node was wrapped in a
    // `0x71` node (vs. `0x70`, unwrapped identically) — an indirection this
    // codebase's `ExprNode` tree does not carry (its parser resolves this at
    // a different stage). Treated as always-unwrapped (`local_10 = 0`) for
    // every argument shape this codebase's front end can produce; revisit
    // if a `0x71`-sourced case is ever found to matter byte-wise.

    // `EbResolveAttributePointer` + the `local_c == 0x1b` check (an array-
    // type declaration node, distinct from an array-typed VALUE argument):
    // this codebase's `VbaType::Array(_)` is the semantic equivalent signal.
    if matches!(param_ty, VbaType::Array(_)) {
        return Err(UnportedCallee::ValidateArgumentType.as_lower_error());
    }

    // `local_18 = EbGetExpressionType2(...)`: this codebase's closest
    // existing analog is `vba_type_to_node_tag`/`load_store_ctx`'s type
    // classification, but neither is a faithful port of
    // `EbGetExpressionType2` itself (see `resolver.rs` for what IS ported of
    // that family) — mapping `param_ty` to the exact `local_18` class
    // (4/5/7/8/9/0xc/…) needs `EbGetExpressionType2` traced specifically for
    // a CALL-ARGUMENT context (its behavior already differs by caller
    // elsewhere in this codebase's grounding work this session). Every
    // param type this port could plausibly classify bottoms out at
    // `ResolveTypeBinding2` (the `local_18 == 7` common case) or a rarer
    // gate; since the classification itself isn't ported, gate uniformly.
    Err(UnportedCallee::ResolveTypeBinding2.as_lower_error())
}

/// **PORTED, verified.** `EbGetTypeCode2` (VBA6.DLL `@0faaf420`, decompiled
/// at `vba6_part0001.c:46418`, 23 bytes) — a genuine leaf: a 32-byte table
/// lookup (`RT_ARG_COERCE_TYPE_CODE`, confirmed-extent, `tables.rs`) with one
/// special case. Used inside `EbEmitArgCoerce` (`vba6_part0002.c:5997,
/// 6012`) to map a resolved type-class index to a coercion type-code, and
/// inside `EbResolveTypeBinding2` similarly — not yet CALLED from
/// `eb_emit_arg_coerce` above (that requires `EbGetExpressionType2`'s
/// call-argument-context classification to be traced first, to know what
/// index to pass in), but this piece itself is complete and trustworthy.
pub(super) fn eb_get_type_code2(index: u8) -> u8 {
    if index == 0x13 {
        0
    } else {
        crate::tables::RT_ARG_COERCE_TYPE_CODE[index as usize]
    }
}

/// **VERIFIED via live TTD, not (only) static reconstruction.** In the
/// class-method-call context, `EbEmitArgCoerce`'s `pProcDesc` parameter is
/// NOT a per-call structure — it's a *fixed address inside VBA6.DLL's own
/// image*, confirmed identical (`0x0fab5b74`) across four different argument
/// types/modes in one live trace (`argtype_probe`, breakpoint at
/// `EbEmitArgCoerce`'s entry, `0fabc1b5`). That address is exactly
/// `RT_CALL_CONV_RECORDS`'s record index 2 (`0x0fab5b74 - 0x0fab5b38 ==
/// 2 * 0x1e`) — the SAME already-extracted call-convention dispatch table
/// used by the intra-module call path, selected here because a class/vtable
/// method call is `RT_CALL_KIND_CLASS` kind 8 → class 2.
///
/// `EbEmitArgCoerce`'s `uVar6` computation reads one byte from this record
/// (`pProcDesc[RT_TYPE_OFFSET[iVar5] * 2] & 0x1f`) and feeds it through
/// [`eb_get_type_code2`]. Statically, this looked like it could read out of
/// the record's 30-byte bound for several relevant type tags (Integer=2,
/// Object=9, Boolean=11 all map through `RT_TYPE_OFFSET`'s sentinel class,
/// `19 * 2 == 38 > 30`) — genuinely ambiguous from the table alone. Directly
/// observed instead (same trace, breakpoint at `EbGetTypeCode2`'s entry,
/// `0faaf420`): for all four argument cases this session has grounded
/// (Integer ByRef, Integer ByVal, `String`, `Object`), the SECOND
/// `EbGetTypeCode2` call (this one) is invoked with argument `0x13` every
/// time — landing on `EbGetTypeCode2`'s own special case, so `uVar6 == 0`
/// for every one of them. This resolves the ambiguity empirically rather
/// than by continuing to guess at which intermediate value is "really"
/// reachable; the mechanism producing `0x13` specifically (rather than a
/// literal `RT_TYPE_OFFSET`-driven byte read) is still not fully traced —
/// `iVar5`'s own source (`EbGetTypeCode2((&DAT_0fab4b78)[local_c])`, then
/// possibly overridden to `0x10`/`8`) needs `DAT_0fab4b78`
/// (`OPERAND_TYPECLASS`) to be confirmed-extent (currently only a 40-byte
/// prefix of an 88-byte manifest entry looks like genuine table data; the
/// rest resembles code/pointers) before it can be ported rather than
/// observed. Returns `Some(0)` for the four grounded cases; `None`
/// (unverified) for anything else — do NOT extend this to other types
/// without a matching live observation.
pub(super) fn arg_coerce_type_code_for_grounded_case(param_ty: &VbaType) -> Option<u8> {
    match param_ty {
        VbaType::Integer | VbaType::String | VbaType::Object => Some(eb_get_type_code2(0x13)),
        _ => None,
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn eb_get_type_code2_special_case_0x13_returns_zero() {
        // The table's own byte at index 0x13 is 0x14 — the special case
        // must override it, not read through to the table.
        assert_eq!(crate::tables::RT_ARG_COERCE_TYPE_CODE[0x13], 0x14);
        assert_eq!(eb_get_type_code2(0x13), 0);
    }

    #[test]
    fn eb_get_type_code2_table_passthrough() {
        assert_eq!(eb_get_type_code2(0x00), 0x05);
        assert_eq!(eb_get_type_code2(0x08), 0x16);
        assert_eq!(eb_get_type_code2(0x1f), 0x07);
    }

    #[test]
    fn arg_coerce_type_code_matches_ttd_observation() {
        assert_eq!(arg_coerce_type_code_for_grounded_case(&VbaType::Integer), Some(0));
        assert_eq!(arg_coerce_type_code_for_grounded_case(&VbaType::String), Some(0));
        assert_eq!(arg_coerce_type_code_for_grounded_case(&VbaType::Object), Some(0));
        assert_eq!(arg_coerce_type_code_for_grounded_case(&VbaType::Long), None);
    }
}
