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
    // that family) — mapping `param_ty` to the exact `local_18` class in
    // general needs `EbGetExpressionType2` traced for the call-argument
    // context, which has NOT been done generally (only observed for four
    // specific `(type, mode)` pairs — see `known_local18_for_grounded_case`).
    //
    // For the ONE case fully traced end-to-end this session — Integer
    // ByVal, `local_18 == 4` — the confirmed chain (`EbEmitExpression4` →
    // `EbCoerceExpressionType2` [no-op, types already match] →
    // `EbProcessType2` → `EbNormalizeTypeReference` [no-op, node unchanged]
    // → `EbBuildBinaryOp` gate evaluates false either way) reduces to: the
    // argument's value is loaded plainly, no coercion applied. That is
    // EXACTLY what `lower_expr_coerced` already does for a plain Integer
    // reference — not a coincidence to paper over, but the actual, verified
    // content of this port's finding: delegate to it rather than re-derive
    // byte emission this port has already shown produces the same result.
    if known_local18_for_grounded_case(param_ty, by_val) == Some(4) {
        return lower_expr_coerced(ctx, arg_id, expr_arena, arena, vba_type_to_node_tag(param_ty));
    }

    // Every other param type/mode this port could plausibly classify
    // bottoms out at `ResolveTypeBinding2` (the `local_18 == 7` common
    // case, traced deep but not byte-complete — see the memory note) or a
    // rarer, entirely untraced gate; since the general classification
    // itself isn't ported, gate uniformly rather than guess.
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
/// - `(Object, ByVal)` → 9 — `EbCheckSetBinding`/`EbEmitPropertyExpr`
///   (`UnportedCallee::SetBindingAndPropertyExpr`), not traced.
///
/// Extrapolating (e.g. "ByRef always gives 7", "ByVal-scalar always gives
/// 4") is EXPLICITLY NOT done here — that pattern is plausible but
/// unverified for any type/mode pair outside these four, and guessing it
/// would be exactly the kind of unverified byte this file's discipline
/// exists to prevent. Returns `None` for anything else.
pub(super) fn known_local18_for_grounded_case(param_ty: &VbaType, by_val: bool) -> Option<i32> {
    match (param_ty, by_val) {
        (VbaType::Integer, false) | (VbaType::String, false) => Some(7),
        (VbaType::Integer, true) => Some(4),
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
