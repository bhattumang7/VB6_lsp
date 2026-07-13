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
