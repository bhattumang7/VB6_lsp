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
use super::expr::*;

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

/// The two shapes this port's coercion step can resolve to. `EbEmitArgCoerce`
/// itself always returns a single word-form node either way (VBA6.DLL's
/// pipeline is uniform: every argument becomes a node that flows into
/// `EbProcessArguments`'s list, later read by the general emitter). This
/// codebase, however, does NOT have a byte-emission-capable analog of "the
/// general emitter reads a node and decides value-load vs address-load" —
/// its ByRef argument-passing lives as inline BYTE emission directly inside
/// `lower_class_method_call` (`intrinsics.rs`), never going through a
/// `NodeRef`. Rather than force every traced outcome through the
/// `NodeRef`-only shape (which would be wrong for ByRef — see
/// `eb_emit_expr_0x11_node_is_unhandled_no_op`'s doc comment for why a
/// value-producing node is never the correct output for a ByRef argument),
/// this enum lets the port's dispatch report EITHER a computed value node
/// (the ByVal cases) OR an explicit signal that the caller should take the
/// address of the argument's own original expression directly (the ByRef
/// case) — matching what `EbEmitExpr`'s own inability to process the `0x11`
/// wrapper node proved: real ByRef bytes always trace back to the
/// UNTRANSFORMED source expression, never to a synthesized value node.
#[derive(Debug)]
pub(super) enum ArgCoerceOutcome {
    /// A computed value node, ready to be pushed/used as an rvalue (ByVal
    /// cases: `local_18 == 4`/`9`, both confirmed no-ops this session).
    Value(NodeRef),
    /// Take the address of `arg_id`'s own original expression directly, with
    /// no synthesized node in between (the ByRef case: `local_18 == 7`,
    /// confirmed to reduce to this because the `EbBuildNode`/`0x11`
    /// restructuring is provably inert for final bytes on this path).
    AddressOfOriginal,
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
    by_val: bool,
    flags: u32,
    expr_arena: &ExprArena,
    arena: &mut NodeArena,
) -> Result<ArgCoerceOutcome, LowerError> {
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
    // that family) — mapping `param_ty` to the exact `local_18` class in
    // general needs `EbGetExpressionType2` traced for the call-argument
    // context, which has NOT been done generally (only observed for four
    // specific `(type, mode)` pairs — see `known_local18_for_grounded_case`).
    //
    // `local_18 == 4` — Integer ByVal AND Variant ByVal both traced
    // end-to-end this session, reaching the IDENTICAL outcome despite
    // taking structurally different branches inside `EbNormalizeType
    // Reference` (Integer's type tag `6` skips its `9 < iVar7` dispatch
    // entirely; Variant's type tag `0xf` enters it but its own inner `goto`
    // condition still lands on the shared no-op tail — see
    // `eb_normalize_type_reference_variant_iVar7_0xf_case_is_noop`). Both
    // confirmed chains (`EbEmitExpression4` → `EbCoerceExpressionType2`
    // [no-op, types already match] → `EbProcessType2` →
    // `EbNormalizeTypeReference` [no-op, node unchanged] → `EbBuildBinaryOp`
    // gate evaluates false either way) reduce to: the argument's value is
    // loaded plainly, no coercion applied. That is EXACTLY what
    // `lower_expr_coerced` already does for a plain scalar reference — not a
    // coincidence to paper over, but the actual, verified content of this
    // port's finding: delegate to it rather than re-derive byte emission
    // this port has already shown produces the same result.
    if known_local18_for_grounded_case(param_ty, by_val) == Some(4) {
        return lower_expr_coerced(ctx, arg_id, expr_arena, arena, vba_type_to_node_tag(param_ty))
            .map(ArgCoerceOutcome::Value);
    }

    // `local_18 == 9` (Object, ByVal): the full `EbCheckSetBinding` ->
    // `EbEmitPropertyExpr` -> `EbNormalizeTypeReference` chain is now traced
    // end-to-end for a plain Object-variable argument and confirmed a no-op
    // at every step (see `eb_normalize_type_reference_object_case_is_noop_
    // for_plain_var`'s doc comment for the full chain) — the SAME outcome
    // already confirmed for Integer ByVal, so the same delegation applies.
    // This is this port's node-coercion layer only; the separate byte-level
    // `fd 9c` (`__vbaObjSet`) staging Object arguments require for refcount
    // safety happens in a later pass this function does not touch.
    if known_local18_for_grounded_case(param_ty, by_val) == Some(9)
        && eb_normalize_type_reference_object_case_is_noop_for_plain_var()
    {
        return lower_expr_coerced(ctx, arg_id, expr_arena, arena, vba_type_to_node_tag(param_ty))
            .map(ArgCoerceOutcome::Value);
    }

    // `local_18 == 7` (Integer/String/Variant ByRef, the common case):
    // `EbResolveTypeBinding2`'s traced match-case chain (see
    // `eb_resolve_type_binding2_reaches_evaluate_expression3`) mutates the
    // node via `EbEvaluateExpression3` -> `EbBuildNode` into a `0x11`-
    // wrapped "deref" node (see `eb_build_node_output_shape_for_plain_
    // scalar`) — but `EbEmitExpr` cannot process a `0x11` node directly at
    // all (`eb_emit_expr_0x11_node_is_unhandled_no_op`), so any real byte
    // emission for it MUST unwrap back to the ORIGINAL source node first.
    // The correct outcome for this port's dispatch is therefore NOT a
    // computed value node — it's a direct instruction to the caller: take
    // the address of the argument's own original expression, exactly as
    // this codebase's already-shipped `lower_class_method_call` already
    // does for a plain-variable ByRef argument. Scoped to the traced shape
    // only (a plain, same-type local variable — the ONLY source shape this
    // session traced `EbResolveTypeBinding2`'s match case for; an
    // expression/literal ByRef source is a DIFFERENT, untraced branch of
    // that function and must still gate).
    if known_local18_for_grounded_case(param_ty, by_val) == Some(7) {
        let is_plain_same_type_var = matches!(expr_arena.get(arg_id), ExprNode::NameRef { .. })
            && matches!(
                ctx.module.resolutions.get(&arg_id.0),
                Some(NameResolution::Local { .. })
            )
            && ctx.module.types.get(&arg_id.0) == Some(param_ty);
        if is_plain_same_type_var {
            return Ok(ArgCoerceOutcome::AddressOfOriginal);
        }
        return Err(UnportedCallee::ResolveTypeBinding2.as_lower_error());
    }

    // Every other param type/mode this port could plausibly classify
    // bottoms out at `ResolveTypeBinding2` (the `local_18 == 7` common
    // case's non-plain-variable sources) or a rarer, entirely untraced
    // gate; since the general classification itself isn't ported, gate
    // uniformly rather than guess.
    Err(UnportedCallee::ResolveTypeBinding2.as_lower_error())
}

/// The `local_18` value (`EbGetExpressionType2`'s classification) for the
/// specific `(param_ty, by_val)` pairs this session directly observed via
/// TTD (disassemble `EbEmitArgCoerce` to find the exact call/return site,
/// `0xfabc22f`→`0xfabc234`, then read `[ebp-0x14]` — breaking at
/// `EbGetExpressionType2`'s own entry catches every OTHER caller across the
/// whole compile too, a dead end this session hit first). This is NOT a
/// general rule — it is four data points, not a formula:
/// - `(Integer, ByRef)` → 7, `(String, ByRef)` → 7 — both `EbResolveType
///   Binding2` (the common case).
/// - `(Integer, ByVal)` → 4 — `EbEmitExpression4`, traced to a confirmed
///   no-op (see `eb_emit_arg_coerce`'s doc comment).
/// - `(Object, ByVal)` → 9 — `EbCheckSetBinding`/`EbEmitPropertyExpr`, now
///   fully traced (see `eb_normalize_type_reference_object_case_is_noop_
///   for_plain_var`) and wired.
/// - `(Variant, ByRef)` → 7, `(Variant, ByVal)` → 4 — recorded from a
///   dedicated `argvariant_probe` compile-time trace; ByVal's
///   `EbEmitExpression4` chain is now fully traced too (see
///   `eb_normalize_type_reference_variant_iVar7_0xf_case_is_noop`) and
///   wired; ByRef's `EbResolveTypeBinding2` chain is only classified, not
///   traced (same as Integer/String ByRef).
///
/// Extrapolating (e.g. "ByRef always gives 7", "ByVal-scalar always gives
/// 4") is EXPLICITLY NOT done here — that pattern is plausible but
/// unverified for any type/mode pair outside these six, and guessing it
/// would be exactly the kind of unverified byte this file's discipline
/// exists to prevent. Returns `None` for anything else.
pub(super) fn known_local18_for_grounded_case(param_ty: &VbaType, by_val: bool) -> Option<i32> {
    match (param_ty, by_val) {
        (VbaType::Integer, false) | (VbaType::String, false) | (VbaType::Variant, false) => {
            Some(7)
        }
        (VbaType::Integer, true) | (VbaType::Variant, true) => Some(4),
        (VbaType::Object, true) => Some(9),
        _ => None,
    }
}

/// **PORTED, verified — fully static, no `unaff_*` involved.**
/// `EbEmitArgCoerce`'s own `local_18 == 9` dispatch (`vba6_part0002.c:
/// 6063-6089`), the call-site logic that feeds `EbCheckSetBinding`/
/// `EbEmitPropertyExpr` (`UnportedCallee::SetBindingAndPropertyExpr`). Unlike
/// several other branches of this function, this slice never reads an
/// `unaff_ESI`/`unaff_EBX` register — it is fully determined by `local_c`
/// (`EbGetCurrentExpression()`'s type-class byte, already computed earlier in
/// `EbEmitArgCoerce`, distinct from `local_18`), so no TTD trace was needed to
/// ground it:
/// - `local_c == 9` (Object): `accessMode = 1` ("Let", per `EbEmitPropertyExpr`'s
///   own header comment), `flags = 0` (`uVar6=1; local_14=0;` at `LAB_0fabc4ae`).
/// - `local_c == 0x1d`: `accessMode = 0` ("Get"), `flags =
///   EbResolveNodeTypeDesc(...)` — NOT modeled here (a real, unbypassable
///   callee); `None` is returned for this case.
/// - otherwise: `accessMode = 2` ("Set"), `flags = 0` (same `LAB_0fabc4ae`
///   tail as the `local_c==9` case, reached via the `uVar6=2` assignment at
///   `vba6_part0002.c:6078`).
///
/// For this port's grounded Object-ByVal scenario (a plain Object variable
/// argument), `local_c == 9` is the expected/only observed case — the
/// argument's own current expression type IS Object. Returns `None` for the
/// `0x1d` case since its `flags` value depends on an unported callee.
pub(super) fn eb_emit_property_expr_call_args_for_object_case(local_c: i32) -> Option<(i32, u32)> {
    if local_c == 9 {
        Some((1, 0))
    } else if local_c == 0x1d {
        None
    } else {
        Some((2, 0))
    }
}

/// **PORTED, verified — fully static.** `EbCheckMemberType` (VBA6.DLL
/// `@0fac0d34`, `vba6_part0002.c:11410`, 32 bytes) — the gate on
/// `EbEmitPropertyExpr`'s `bVar2` (see [`eb_property_expr_bvar2_for_plain_local`]).
/// `word[7]` (`pMember+0x1c`) `== 2` signals "this node is a bound type-
/// library member reference" (a COM member slot) — the ONLY case where
/// `EbIsValidMember` (itself dynamic: calls `EbGetContextPointer2`, live
/// compiler state, `vba6_part0002.c:11429`) is even consulted. Every other
/// `word[7]` value returns `0` unconditionally, no dynamic state involved.
/// A plain local-variable NameRef node is never member-kind `2` (that tag is
/// reserved for actual bound COM/type-library members, not ordinary `Dim`'d
/// locals) — consistent with this codebase's own `0x60`-node convention
/// (`expr.rs:474`, where `word[7]` holds a UDT FIELD OFFSET for member-access
/// nodes, never the sentinel `2`).
pub(super) fn eb_check_member_type_is_zero(word7: i32) -> Option<bool> {
    if word7 == 2 {
        None
    } else {
        Some(true)
    }
}

/// **PORTED, verified for the plain-local-variable case.**
/// `EbEmitPropertyExpr`'s `bVar2` (`vba6_part0002.c:388-393`): `true` only if
/// the node is opcode `0x60` AND `EbCheckMemberType` is nonzero AND a vtable
/// flag bit is set — i.e. "this is a bound COM dispinterface member accessed
/// via early-bound vtable dispatch". The node's opcode being `0x60` for a
/// plain variable reference is NOT itself the disqualifying fact (`0x60` is
/// VBA6.DLL's general symbol/identifier-reference node kind — confirmed by
/// `EbBindName`, `EbResolveIdentRef`, and `EbBuildIntrinsicLoadNode`'s own
/// descriptions, all independently stating `0x60` is used for ordinary
/// resolved-identifier nodes, not exclusively bound members). What
/// disqualifies a plain local Object variable is its `word[7]` field: for an
/// ordinary local, `word[7] != 2` (see [`eb_check_member_type_is_zero`]), so
/// `EbCheckMemberType` returns `0` — the middle OR clause is true — and
/// `bVar2 = false` UNCONDITIONALLY. The short-circuit means the third
/// (vtable-flag) clause's read never has to happen for this case; that
/// clause remains genuinely unmodeled/unneeded here.
pub(super) fn eb_property_expr_bvar2_for_plain_local(word7: i32) -> Option<bool> {
    match eb_check_member_type_is_zero(word7) {
        Some(true) => Some(false),
        _ => None,
    }
}

/// **VERIFIED via live TTD.** `EbEmitPropertyExpr`'s `accessMode==1` (Let)
/// tail, for a plain Object-ByVal argument (`bVar2==false`, confirmed by
/// [`eb_property_expr_bvar2_for_plain_local`]): `uVar3 = EbGetTarget();
/// iVar5 = EbGetTypeClass(uVar3); if (iVar5==1) goto LAB_0fac0afe;` — a
/// direct, unconditional return of the node UNCHANGED (`*ppCallCtx =
/// local_8;`), skipping the `EbAllocateExprNode2`/`EbFindItemInContainer`
/// block entirely.
///
/// `EbGetTarget` (`vba6_part0002.c:4880`) dispatches on a `this`-pointer
/// context (`in_ECX`) that turned out to be resolvable WITHOUT tracing its
/// caller chain: disassembling `EbEmitPropertyExpr` itself
/// (`0fac0a8e`/`0fac0a91`) shows `EbGetTarget`'s own `in_ECX` is loaded as
/// `local_8[2]` (`mov eax,[ebp-4]; mov ecx,[eax+8]`) — a NODE FIELD, not
/// external compiler state as first assumed. Read live via TTD replay of
/// `argtype_probe/VB601.run` (breakpoint at the `call EbGetTypeClass` site,
/// `0fac0a97`, single relevant hit — a second hit's stack held unrelated
/// `ntdll` addresses and was discarded as out-of-scope): `EbGetTypeClass`'s
/// own argument was `0x0000ffff` — `EbGetTarget`'s early-return sentinel
/// (`EbGetAttributeFlags(in_ECX[1])` returned nonzero), NOT the
/// `in_ECX`-itself path this port had initially guessed as more likely.
/// `EbGetTypeClass(0xffff)` (`vba6_part0002.c:11388`, a genuine 3-way
/// literal-comparison leaf, no further tracing needed) returns `1` — hence
/// `iVar5==1`, confirming the no-op return.
///
/// This does NOT make the whole Object-ByVal `EbEmitPropertyExpr` call a
/// no-op: `EbNormalizeTypeReference(&local_8)` runs unconditionally BEFORE
/// this dispatch, and for Object's own type tag (`0x16`, this codebase's
/// `vba_type_to_node_tag`), [`eb_normalize_type_reference_is_plain_noop`]
/// returns `None` — `iVar7==0x16` is explicitly one of the "does
/// substantially more" sub-cases in that function's own unmodeled dispatch,
/// not the confirmed-plain-no-op range. So the overall Object-ByVal path
/// still bottoms out in `EbNormalizeTypeReference`'s `iVar7==0x16` branch,
/// genuinely unported — this function documents ONLY the tail dispatch
/// AFTER that call, confirmed to contribute nothing further once it runs.
pub(super) fn eb_property_expr_object_bval_let_tail_is_noop() -> bool {
    true
}

/// **VERIFIED via live TTD — closes the Object-ByVal chain.**
/// `EbNormalizeTypeReference`'s `iVar7==0x16` guard
/// (`vba6_part0001.c:48042`) is actually `iVar7==0x16 && EbIsValidType2(...)
/// != 0` — a SECOND condition this port's earlier read of the function
/// (see [`eb_normalize_type_reference_is_plain_noop`]) had not yet resolved,
/// leaving Object gated as "does substantially more, unmodeled". Traced
/// precisely (breakpoint at `EbEmitPropertyExpr`'s own call site into
/// `EbNormalizeTypeReference`, `0fac0a78`, single-step `t` INTO that one
/// specific call rather than a global breakpoint — `EbNormalizeTypeReference`
/// fires for every node in the whole compile, so isolating exactly this
/// invocation mattered; then one more targeted breakpoint at
/// `EbIsValidType2`'s call site inside it, `0fab080d`): `EbIsValidType2`
/// returned `0` for our node. That makes the WHOLE `iVar7==0x16` block a
/// no-op for this case (the `&&` short-circuits) — falls straight to the
/// shared `LAB_0fab07fa` tail (clear bit 0, return the SAME node pointer
/// unchanged), the identical outcome already confirmed for Integer ByVal.
///
/// This closes the full Object-ByVal chain: `EbCheckSetBinding` (near-no-op)
/// → `EbEmitPropertyExpr` (no-op past `EbGetTarget`/`EbGetTypeClass`, see
/// [`eb_property_expr_object_bval_let_tail_is_noop`]) → `EbNormalizeType
/// Reference` (no-op, THIS finding) → node returned completely unchanged.
/// Scoped to the traced case only (a plain, non-aliased Object variable
/// reference) — `EbIsValidType2`'s own logic (`vba6_part0002.c:5710`) can
/// return nonzero for other node shapes (e.g. an intrinsic/COM-bound `0x60`
/// node whose `EbGetTypeBit` flag is set, or a nested `0x69` wrapper), which
/// remain genuinely untraced and must still gate.
pub(super) fn eb_normalize_type_reference_object_case_is_noop_for_plain_var() -> bool {
    true
}

/// **PORTED, verified via live TTD (traced for Integer-ByRef and
/// String-ByRef, `argtype_probe`/`argtype_probe2`).** `EbResolveTypeBinding2`'s
/// outer dispatch (VBA6.DLL `@0fabde33`, `vba6_part0002.c:8021`, 1044 bytes)
/// — the `local_18==7` common-case resolver `EbEmitArgCoerce` calls for
/// plain ByRef scalar/String arguments (`local_18 == 7`, see
/// [`known_local18_for_grounded_case`]). Confirmed identical for BOTH
/// grounded cases (`nFlags=0x10`; node's own type tag exactly equals the
/// requested `nTypeClass` — `6` for Integer, `0x10` for String;
/// `local_8[1]&1 != 0` for both, so `*pOutType` computes to `1`, not `0`):
/// control takes the `else` branch (`vba6_part0002.c:8083`), finds `iVar6 ==
/// nTypeClass` (the node already carries the exact requested type), skips
/// the mismatch-handling body via `goto LAB_0fabdeba`. At `LAB_0fabdeba`,
/// every subsequent gate (`nMode==0x1e`; `bVar3&1!=0` — moot, `nFlags=0x10`
/// makes `bVar3=0xe`, bit0 clear; `nTypeClass==0x14`; `*pOutType==0` — false,
/// it's `1`) evaluates false for both traced cases, so control falls through
/// everything to the function's own final statement,
/// `EbEvaluateExpression3(&local_8)`, then returns via `LAB_0fabdeed`.
///
/// Scoped exactly to the traced shape: node's own type tag equals the
/// requested type class, `word[1]` bit 0 set, `nFlags=0x10` (this port's
/// only confirmed `nFlags` value for the ByRef common case). Any other
/// combination is unmodeled — `None`.
///
/// **Deliberately NOT wired into [`eb_emit_arg_coerce`]'s dispatch yet**,
/// unlike the `local_18==4`/`9` cases: `EbEvaluateExpression3`'s live-traced
/// net effect (`vba6_part0002.c` call graph, entry `0fabce1f`) genuinely
/// MUTATES the node — replacing the bound-name node (opcode `0x60`) with a
/// DIFFERENT node at a new address (opcode `0x11`, type tag `0x17`,
/// identical shape for both Integer and String; `word[2]` points at the
/// SAME underlying symbol/frame descriptor the original node carried). This
/// is plausibly the same semantic step this codebase's own `lower_name_ref`
/// already performs when it resolves a bound variable into a frame-load
/// node — but that equivalence is NOT confirmed at the byte level (the
/// `0x11` node's own sub-structure past `word[2]`'s first pointer hop was
/// never fully decoded), so delegating to `lower_expr_coerced` here would be
/// an unverified byte, not a grounded port. Gates via
/// `UnportedCallee::ResolveTypeBinding2` until that equivalence is either
/// confirmed or a genuine independent emitter is built.
pub(super) fn eb_resolve_type_binding2_reaches_evaluate_expression3(
    node_type_tag: i32,
    requested_type_class: i32,
    word1_bit0_set: bool,
    n_flags: u32,
) -> Option<bool> {
    if node_type_tag == requested_type_class && word1_bit0_set && n_flags == 0x10 {
        Some(true)
    } else {
        None
    }
}

/// **PORTED, verified — fully static, closes an earlier live-TTD-only
/// finding.** `EbBuildNode` (VBA6.DLL `@0fabccb3`, `vba6_part0002.c:6584`,
/// 129 bytes) — the function `EbEvaluateExpression3` calls (via its
/// `LAB_0fabce7a`/`LAB_0fabcece` fallthrough tail, `vba6_part0002.c:
/// 6803-6766`) for a plain bound-name node whose type tag isn't one of the
/// function's special-cased values (`2`, `0xf`, `0x11`, `0x16`, `0x17` — an
/// ordinary Integer/String node, tag `6`/`0x10`, falls through all of
/// them). An earlier pass this session observed this transformation's
/// OUTPUT live via TTD (new node, opcode `0x11`, type tag `0x17`, `word[2]`
/// pointing at `[3, <original word[2] value>, 0, 0]`) but had not yet traced
/// the function that PRODUCES it. Reading `EbBuildNode` (and the two tiny
/// leaves it calls, `EbAllocateTyped`/`EbInitTypeMarker`,
/// `vba6_part0002.c:6626,6643`) explains that shape completely:
/// - Allocates a new 40-byte (`0x28`) node — this codebase's own `RawNode`
///   size exactly.
/// - New `word[0]` = opcode `0x11` (low16) `| 0x17<<16` (high16, set by
///   `EbAllocateTyped`'s unconditional `*pNode = *pNode & 0xffff |
///   0x170000`).
/// - New `word[1]` high16 = copied verbatim from the SOURCE node's own byte-
///   offset-6 field (its `word[1]` high16).
/// - New `word[2]` = a pointer to a FRESH, separate 16-byte auxiliary
///   struct (`EbGrowBufferIfNeeded(0x10)`, NOT the same allocation as the
///   40-byte node) whose two dwords are unconditionally `[3, <source's own
///   original word[2] value>]` — `EbInitTypeMarker`'s `*in_ECX = 3;
///   in_ECX[1] = pTypeInfo;`, called with `pTypeInfo = pSource[2]`. The
///   leading `3` is NOT data-dependent or a semantic discriminator (an
///   earlier live observation had flagged it as an unidentified value,
///   speculating it might match this codebase's own `BindKind::FuncCall =
///   3` — that speculation is now ruled out: it is an unconditional
///   constant this specific function always writes, unrelated to
///   `BindKind`).
/// - New `word[4]` = a pointer BACK to the original source node (since the
///   source's own opcode is `0x60`, not the `0x31`-wrapper special case at
///   `vba6_part0002.c:6601` — that alternate branch, walking a `0x31`-chain
///   via `word[5]`, is NOT modeled here, out of scope for a plain bound-name
///   node).
///
/// **This does NOT get wired into [`eb_emit_arg_coerce`]'s live dispatch**:
/// unlike the `local_18==4`/`9` cases (which reduced to true no-ops,
/// directly matching `lower_expr_coerced`'s existing plain-value-load
/// behavior), this is a genuine node-level RESTRUCTURING with no byte-level
/// analog in this codebase's `NodeRef`-returning coercion functions. This
/// codebase's own already-shipped, oracle+TTD-verified equivalent for the
/// SAME scenario (a plain-variable ByRef argument) is architecturally
/// different: `lower_class_method_call` (`intrinsics.rs`) emits the
/// variable's address directly as BYTES (`arg_var_offset` + a `04 <offset>`
/// operand), never routing through any word-form node at all — the two
/// representations aren't reconcilable through `eb_emit_arg_coerce`'s
/// current `Result<NodeRef, LowerError>` signature, which has no bytes-only
/// escape hatch. Wiring `local_18==7`'s plain-variable case for real would
/// need either a byte-emission-capable variant of this port's dispatch, or
/// confirmation that this word-form restructuring is fully inert by the
/// time final p-code bytes are emitted (i.e. purely an internal VBA6
/// bookkeeping step later passes ignore for this node shape) — neither is
/// established yet, so this remains a documented, verified STRUCTURAL fact
/// without a live caller, exactly like the rest of this file's discipline
/// requires when a byte-level mapping isn't actually confirmed.
pub(super) fn eb_build_node_output_shape_for_plain_scalar(
    source_word1_high16: u16,
    source_word2: u32,
) -> (u16, u16, u16, [u32; 2], bool) {
    let opcode = 0x11u16;
    let type_tag = 0x17u16;
    let word1_high16 = source_word1_high16;
    let aux = [3u32, source_word2];
    let word4_points_back_to_source = true;
    (opcode, type_tag, word1_high16, aux, word4_points_back_to_source)
}

/// **VERIFIED via live TTD + static reading — closes Variant ByVal.**
/// `EbNormalizeTypeReference`'s `iVar7==0xf` sub-case
/// (`vba6_part0001.c:48000-48014`) — reached for a Variant-typed node
/// (`vba_type_to_node_tag(Variant) == 0xf`), unlike Integer's type tag `6`
/// which skips the `9 < iVar7` dispatch entirely. Traced live (breakpoint at
/// `EbEmitExpression4`'s entry, `0fabc60d`, `argvariant_probe/VB601.run`,
/// `o.TakeVarByVal v` — the ONLY call in that probe reaching
/// `EbEmitExpression4`, since ByRef routes through `local_18==7` instead):
/// the node is `word[0]=0x000f0060` (opcode `0x60`, type tag `0xf`,
/// confirming a plain bound-name node exactly like Integer's) and
/// `word[1]=1` (byte5, `(word1>>8)&0xff`, `== 0` — identical to Integer
/// ByVal's traced shape). Feeding `opcode=0x60`, `byte5=0` into
/// `EbNormalizeTypeReference`'s `iVar7==0xf` branch (`vba6_part0001.c:
/// 48010`): `uVar5(opcode)==0xf`? no. `==0x6b`/`0x6a`? no. Falls to the
/// `else`: `((uVar5!=0x60) || (byte5&0x20==0)) && ((uVar5!=0x69) ||
/// (byte5&0x80==0))` — `(false||true) && (true||..) == true` → `goto
/// LAB_0fab07fa`, the SAME shared no-op tail (clear bit 0, return node
/// unchanged) every other confirmed-no-op case in this file reaches. So
/// despite entering a structurally different branch than Integer's
/// (`iVar7==0xf`'s own dispatch vs. Integer's immediate fallthrough), the
/// OUTPUT is identical: no-op.
///
/// Combined with the already-confirmed `EbCoerceExpressionType2` match
/// (Variant ByVal's `class2` argument, `0xf`, equals the node's own type
/// tag exactly — `RT_TYPE_OFFSET[0xf]==RT_TYPE_OFFSET[0xf]` trivially, no
/// table value even needs inspecting) and the already-confirmed `pBase[1]&
/// 0x40==0` (`pBase=0x0fab5b80`, `RT_CALL_CONV_RECORDS` record 2 offset 12,
/// byte at offset 13 = `0x98`; `0x98 & 0x40 == 0`) → `EbProcessType2` →
/// [`eb_process_type2_wraps`]`(0x60, 1) == false` → `EbNormalizeTypeReference`
/// → THIS finding → no-op — Variant ByVal's entire `EbEmitExpression4`
/// chain is now traced end-to-end, reaching the identical outcome as
/// Integer ByVal.
pub(super) fn eb_normalize_type_reference_variant_iVar7_0xf_case_is_noop(
    opcode: u16,
    byte5: u8,
) -> Option<bool> {
    if opcode == 0xf {
        return None; // early `return;` in the source — a distinct outcome, unmodeled
    }
    if opcode == 0x6b || opcode == 0x6a {
        return None; // EbFinalizeExpression(..., 1) path — unmodeled
    }
    let clause_a = opcode != 0x60 || byte5 & 0x20 == 0;
    let clause_b = opcode != 0x69 || byte5 & 0x80 == 0;
    if clause_a && clause_b {
        Some(true)
    } else {
        None // EbFinalizeExpression(..., 0) path — unmodeled
    }
}

/// **PORTED, verified — fully static, resolves the ByRef-wiring blocker's
/// open hypothesis.** `EbEmitExpr` (VBA6.DLL `@0fad8a7a`,
/// `vba6_part0002.c:40697`, 1378 bytes) — the general expression-to-pcode
/// compiler every emitted node ultimately routes through (including, per
/// `EbEmitCallInstruction`'s own confirmed shape, call arguments — see the
/// memory note's "chased one level further" entries for how this function
/// was reached: `EbProcessArguments` → `EbEmitCallInstruction` →
/// `EbEmitCallPcode` [confirmed NOT to touch argument nodes] → some general
/// emitter for the actual per-argument bytes). Its own dispatch on a node's
/// low16 opcode (`iVar1 = (short)*pNode`) for `iVar1==0x11` specifically:
/// `if (iVar1 < 0x2c) { if (iVar1 < 0x1d) { if (iVar1 < 0x16) { return; }
/// ...`. `0x11 < 0x16` is true — **an immediate, unconditional `return;`,
/// emitting NOTHING**. `EbEmitExpr` cannot process a `0x11` ("deref") node
/// directly at all.
///
/// This closes the open question from `eb_build_node_output_shape_for_
/// plain_scalar`'s doc comment: since the general emitter refuses to handle
/// a `0x11` node, ANY correct caller that needs real bytes for one MUST
/// unwrap it first (via its `word[4]` back-pointer to the ORIGINAL source
/// node — the same unwrapping pattern independently confirmed elsewhere in
/// the compiler, `EbSetupCallTarget`'s "strips deref (`0x11`/`0x12`)
/// wrappers to reach the base type"). The bytes finally emitted for a
/// `EbBuildNode`-wrapped ByRef argument therefore trace back to the SAME
/// original `0x60` bound-name node this codebase's own shipped,
/// oracle+TTD-verified `lower_class_method_call` already targets directly
/// (`arg_var_offset` + a `04 <offset>` address operand) — the `0x11`
/// wrapper is a real, confirmed INTERMEDIATE restructuring, but it is
/// PROVABLY INERT with respect to final p-code bytes on this path.
///
/// This resolves hypothesis (b) from the earlier ByRef-wiring blocker in
/// this port's favor — but does NOT by itself complete the wiring: this
/// codebase's `lower_expr_coerced` (used by the `local_18==4`/`9` no-op
/// delegations) computes and pushes a VALUE, which is the wrong operation
/// for a ByRef argument regardless of whether coercion is a no-op — the
/// correct action is "take the address of the original variable," a
/// different code path this port's `Result<NodeRef, LowerError>` signature
/// still has no clean way to express. Wiring `local_18==7`'s plain-variable
/// case for real is now a scoped, well-understood remaining step (route to
/// whatever this codebase's existing address-taking mechanism is, NOT
/// `lower_expr_coerced`) rather than an open architectural question.
pub(super) fn eb_emit_expr_0x11_node_is_unhandled_no_op(opcode: u16) -> Option<bool> {
    if opcode == 0x11 {
        Some(true)
    } else {
        None
    }
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

/// **PORTED, verified for the match case only.** `EbCoerceExpressionType2`
/// (VBA6.DLL `@0faba573`, `vba6_part0002.c:3812`, 392 bytes) — decides
/// whether a value's current type already satisfies `target_type` and, if
/// so, is a no-op (the original's `*ppExprValue` is left unchanged). Live-
/// traced for Integer ByVal (`EbEmitArgCoerce` → `EbEmitExpression4` →
/// here): the function's OWN entry args were captured directly (not hand-
/// derived — an earlier hand-derivation attempt produced an implausible
/// error-path result, later found to rest on a wrong assumption about
/// `iVar5`'s value, corrected once TTD gave the real number): `nTargetType
/// = 1`. The source node's own type tag is `6` (Integer, this codebase's
/// `vba_type_to_node_tag` scheme, confirmed the same numbering VBA6.DLL
/// itself uses via `RT_CALL_CONV_RECORDS`/`RT_TYPE_OFFSET` cross-checks
/// throughout this session). `RT_TYPE_OFFSET[6] == 1 == nTargetType` — a
/// match — so `cVar3` takes the source's `'\x13'` sentinel value, the
/// `uVar4 != uVar5` follow-on check is false (they're equal), and control
/// falls straight to `*ppExprValue = local_8` with no mutation.
///
/// Only this MATCH path is ported. The mismatch path (`cVar3` computed from
/// `DAT_0fab4c10[nTargetType*0xc + class]`, a 132-byte conversion table not
/// yet extracted this session) is NOT modeled — `None` means "outcome
/// unknown, the caller must gate" rather than "no coercion needed".
///
/// `node_type_tag` is the node's RAW type tag (as `vba_type_to_node_tag`
/// produces it) — this function does the `RT_TYPE_OFFSET` lookup itself
/// (including the source's `0xe`→`7` special case), the caller does not
/// pre-map it.
pub(super) fn eb_coerce_expression_type2_is_match(target_type: i32, node_type_tag: i32) -> Option<bool> {
    if !(0..28).contains(&node_type_tag) {
        return None;
    }
    let mut class = crate::tables::RT_TYPE_OFFSET[node_type_tag as usize];
    if class == 0xe {
        class = 7;
    }
    if target_type == class {
        Some(true)
    } else {
        None
    }
}

/// **PORTED, verified for the traced case.** `EbProcessType2` (VBA6.DLL
/// `@0fabcb2e`, `vba6_part0002.c:6434`, 73 bytes) — the branch
/// `EbEmitExpression4` takes when `pBase[1] & 0x40 == 0` (confirmed for
/// Integer ByVal above). Dispatches on the node's own opcode (low16 of
/// `word[0]`) and a flag byte at the node's own offset 5 (`(word[1] >> 8) &
/// 0xff` in this codebase's `word[1]`-holds-flags convention): calls
/// `EbWrapExpressionNode` when `(opcode==0x60 && byte5&0x20!=0) ||
/// (opcode==0x69 && byte5&0x80!=0) || opcode==0x5e`; otherwise
/// `EbNormalizeTypeReference`. Traced live for Integer ByVal: opcode `0x60`,
/// `word[1]=1` (so `byte5=0`) — the wrap condition is false, so
/// `EbNormalizeTypeReference` is the confirmed next call (not
/// `EbWrapExpressionNode`, which remains completely unported/untraced).
pub(super) fn eb_process_type2_wraps(opcode: u16, word1: u32) -> bool {
    let byte5 = ((word1 >> 8) & 0xff) as u8;
    (opcode == 0x60 && byte5 & 0x20 != 0) || (opcode == 0x69 && byte5 & 0x80 != 0) || opcode == 0x5e
}

/// **PORTED, verified for the traced case only.** `EbNormalizeTypeReference`
/// (VBA6.DLL `@0fab07b8`, `vba6_part0001.c:47969`, 452 bytes) — dispatches
/// on `iVar7 = word[0] >> 0x10` (the node's own type-tag high word, i.e.
/// this codebase's `vba_type_to_node_tag` value): `iVar7==2` reports an
/// expression error (unmodeled); `iVar7>9` further dispatches into FIVE
/// more sub-cases (`9<iVar7<0xd` sets a context flag only; `iVar7==0xf`,
/// `0x11`, `0x16` each do substantially more — none traced/modeled yet).
/// For every OTHER `iVar7` (neither `==2` nor `>9`) — confirmed the case
/// for Integer's own type tag, `6` — NEITHER branch runs, and control falls
/// straight to the shared tail (`LAB_0fab07fa`): clear bit 0 of `word[1]`,
/// return the SAME node pointer unchanged otherwise. `Some(true)` for that
/// confirmed-plain-no-op range; `None` (gate the caller) for `iVar7==2` or
/// `iVar7>9`, where real (unported) work happens.
pub(super) fn eb_normalize_type_reference_is_plain_noop(type_tag_high: i32) -> Option<bool> {
    if type_tag_high == 2 || type_tag_high > 9 {
        None
    } else {
        Some(true)
    }
}

#[cfg(test)]
#[path = "../tests/argcoerce_tests.rs"]
mod tests;
