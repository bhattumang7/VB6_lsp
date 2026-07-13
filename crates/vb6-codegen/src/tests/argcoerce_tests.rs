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

#[test]
fn eb_coerce_expression_type2_is_noop_when_types_already_match() {
    // Integer ByVal (TakeIntByVal(i)): node type tag 6, nTargetType 1 —
    // RT_TYPE_OFFSET[6] == 1 == nTargetType, a match, so the ported
    // function must report "no coercion needed" (matching the live
    // trace: cVar3 == 0x13, falls straight through to `*ppExprValue =
    // local_8`, unchanged).
    assert_eq!(eb_coerce_expression_type2_is_match(1, 6), Some(true));
}

#[test]
fn eb_process_type2_wraps_matches_traced_integer_case() {
    // Integer ByVal's node: opcode 0x60, word[1] = 1 (byte@offset5 = 0).
    // 0x60 with byte@offset5 & 0x20 == 0 -> the wrap condition is false
    // -> EbNormalizeTypeReference is called next (confirmed: that IS
    // what the live trace's control flow implies — EbProcessType2 was
    // the branch taken, and EbWrapExpressionNode was never observed).
    assert_eq!(eb_process_type2_wraps(0x60, 1), false);
}

#[test]
fn eb_normalize_type_reference_noop_for_integer_class() {
    // iVar7 (type tag high) = 6 for our Integer node: not ==2, not >9,
    // so neither special-case branch runs — falls straight to
    // LAB_0fab07fa, clearing word[1] bit 0 and returning the same node.
    assert_eq!(eb_normalize_type_reference_is_plain_noop(6), Some(true));
    // iVar7==2 and iVar7>9 are NOT plain no-ops (error-report / multi-
    // case dispatch) — not modeled, must gate.
    assert_eq!(eb_normalize_type_reference_is_plain_noop(2), None);
    assert_eq!(eb_normalize_type_reference_is_plain_noop(10), None);
}

#[test]
fn eb_emit_property_expr_call_args_object_case() {
    // local_c == 9 (Object): accessMode=1 (Let), flags=0 — fully static,
    // no TTD needed (this slice of EbEmitArgCoerce reads no unaff_* regs).
    assert_eq!(eb_emit_property_expr_call_args_for_object_case(9), Some((1, 0)));
    // local_c == 0x1d: flags comes from EbResolveNodeTypeDesc, an unported
    // callee — must gate, not guess flags=0.
    assert_eq!(eb_emit_property_expr_call_args_for_object_case(0x1d), None);
    // otherwise: accessMode=2 (Set), flags=0.
    assert_eq!(eb_emit_property_expr_call_args_for_object_case(0), Some((2, 0)));
    assert_eq!(eb_emit_property_expr_call_args_for_object_case(5), Some((2, 0)));
}

#[test]
fn eb_check_member_type_is_zero_for_non_member_word7() {
    // word[7] == 2 is the "bound type-library member" sentinel; anything
    // else means EbIsValidMember (dynamic) is never consulted.
    assert_eq!(eb_check_member_type_is_zero(2), None);
    assert_eq!(eb_check_member_type_is_zero(0), Some(true));
    assert_eq!(eb_check_member_type_is_zero(5), Some(true));
}

#[test]
fn eb_property_expr_bvar2_false_for_plain_local() {
    // A plain local Object variable's node is never member-kind 2, so
    // EbCheckMemberType returns 0 and bVar2 is unconditionally false.
    assert_eq!(eb_property_expr_bvar2_for_plain_local(0), Some(false));
    // word7==2 requires the dynamic EbIsValidMember path — must gate.
    assert_eq!(eb_property_expr_bvar2_for_plain_local(2), None);
}

#[test]
fn eb_property_expr_object_bval_let_tail_is_noop_confirmed() {
    // TTD-confirmed (argtype_probe/VB601.run, EbGetTypeClass's own arg was
    // 0xffff -> iVar5==1 -> immediate no-op return in EbEmitPropertyExpr's
    // accessMode==1 branch).
    assert!(eb_property_expr_object_bval_let_tail_is_noop());
}

#[test]
fn eb_normalize_type_reference_object_case_confirmed_noop() {
    // TTD-confirmed (argtype_probe/VB601.run: EbIsValidType2 returned 0 for
    // our node at the precise call site inside EbNormalizeTypeReference,
    // isolated via a single-step-into from EbEmitPropertyExpr's own call).
    assert!(eb_normalize_type_reference_object_case_is_noop_for_plain_var());
}

#[test]
fn eb_resolve_type_binding2_reaches_evaluate_expression3_for_traced_byref_cases() {
    // Integer ByRef: node type tag 6 == requested nTypeClass 6.
    assert_eq!(
        eb_resolve_type_binding2_reaches_evaluate_expression3(6, 6, true, 0x10),
        Some(true)
    );
    // String ByRef: node type tag 0x10 == requested nTypeClass 0x10.
    assert_eq!(
        eb_resolve_type_binding2_reaches_evaluate_expression3(0x10, 0x10, true, 0x10),
        Some(true)
    );
    // Mismatched type tag, word1 bit0 clear, or a different nFlags are all
    // outside the traced shape — must gate, not guess.
    assert_eq!(
        eb_resolve_type_binding2_reaches_evaluate_expression3(6, 0x10, true, 0x10),
        None
    );
    assert_eq!(
        eb_resolve_type_binding2_reaches_evaluate_expression3(6, 6, false, 0x10),
        None
    );
    assert_eq!(
        eb_resolve_type_binding2_reaches_evaluate_expression3(6, 6, true, 0),
        None
    );
}

#[test]
fn eb_build_node_output_shape_matches_earlier_live_observation() {
    // Cross-checks the fully static EbBuildNode/EbAllocateTyped/
    // EbInitTypeMarker reading against the earlier live TTD observation
    // (0x0380008c was the source's own word[2] value in that trace) — the
    // static derivation and the live read must agree.
    let (opcode, type_tag, word1_high16, aux, points_back) =
        eb_build_node_output_shape_for_plain_scalar(0, 0x0380008c);
    assert_eq!(opcode, 0x11);
    assert_eq!(type_tag, 0x17);
    assert_eq!(word1_high16, 0);
    assert_eq!(aux, [3, 0x0380008c]);
    assert!(points_back);
}

#[test]
fn known_local18_matches_the_six_ttd_observed_pairs() {
    assert_eq!(known_local18_for_grounded_case(&VbaType::Integer, false), Some(7));
    assert_eq!(known_local18_for_grounded_case(&VbaType::String, false), Some(7));
    assert_eq!(known_local18_for_grounded_case(&VbaType::Integer, true), Some(4));
    assert_eq!(known_local18_for_grounded_case(&VbaType::Object, true), Some(9));
    assert_eq!(known_local18_for_grounded_case(&VbaType::Variant, false), Some(7));
    assert_eq!(known_local18_for_grounded_case(&VbaType::Variant, true), Some(4));
    // NOT extrapolated to untested pairs — e.g. String ByVal was never
    // observed, so this must stay None rather than guess "4" by analogy
    // with Integer ByVal.
    assert_eq!(known_local18_for_grounded_case(&VbaType::String, true), None);
    assert_eq!(known_local18_for_grounded_case(&VbaType::Object, false), None);
    assert_eq!(known_local18_for_grounded_case(&VbaType::Long, false), None);
}

#[test]
fn eb_normalize_type_reference_variant_case_confirmed_noop() {
    // TTD-confirmed (argvariant_probe/VB601.run: node word0=0x000f0060,
    // word1=1 -> opcode=0x60, byte5=0 -> the iVar7==0xf branch's own inner
    // goto condition is true -> shared no-op tail).
    assert_eq!(
        eb_normalize_type_reference_variant_iVar7_0xf_case_is_noop(0x60, 0),
        Some(true)
    );
    // opcode==0xf is a distinct early-return outcome, not modeled.
    assert_eq!(eb_normalize_type_reference_variant_iVar7_0xf_case_is_noop(0xf, 0), None);
    // opcode==0x6b/0x6a route through EbFinalizeExpression(...,1), not modeled.
    assert_eq!(eb_normalize_type_reference_variant_iVar7_0xf_case_is_noop(0x6b, 0), None);
    assert_eq!(eb_normalize_type_reference_variant_iVar7_0xf_case_is_noop(0x6a, 0), None);
    // opcode==0x60 with byte5 bit 0x20 set fails clause_a -> EbFinalizeExpression(...,0), not modeled.
    assert_eq!(
        eb_normalize_type_reference_variant_iVar7_0xf_case_is_noop(0x60, 0x20),
        None
    );
}
