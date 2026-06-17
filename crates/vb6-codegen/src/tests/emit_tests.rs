use super::*;
use crate::node::NodeArena;
use crate::tables::{RT_LOAD_BY_CTX, RT_STORE_BY_CTX};

/// Build a bound-symbol node whose frame offset (stored in type_info =
/// high 16 bits of word[4]) is `offset`.
fn sym(arena: &mut NodeArena, offset: i16) -> NodeRef {
    arena.alloc(NodeArena::node(0, 0, (offset as u16 as u32) << 16, 0, 0, 0))
}

/// Build a variable-load node (type 0x74) with the given type context and
/// a bound symbol at `offset`.
fn var_load(arena: &mut NodeArena, type_ctx: u16, offset: i16) -> NodeRef {
    let s = sym(arena, offset);
    arena.alloc(NodeArena::node(0x74, 0, s.0, type_ctx as u32, 0, 0))
}

/// Build a variable-load node carrying a VB6 type tag in `word[0]` high half.
/// Comparison emission selects its opcode from the *LHS operand's* type tag
/// (`EbEmitBinaryOperation2` comparison branch), so operands of a comparison
/// must carry their type tag, not just the load-time type context.
fn var_load_typed(arena: &mut NodeArena, vb_type: u16, type_ctx: u16, offset: i16) -> NodeRef {
    let s = sym(arena, offset);
    arena.alloc(NodeArena::node(0x74, vb_type, s.0, type_ctx as u32, 0, 0))
}

fn emit(arena: &NodeArena, root: NodeRef) -> Vec<u8> {
    let mut e = Emitter::new(arena);
    e.emit_expr(root, 0);
    e.into_bytes()
}

// ── Table sanity ──────────────────────────────────────────────────────────

#[test]
fn rt_load_table_confirmed_entries() {
    assert_eq!(RT_LOAD_BY_CTX[1], 0x6b, "Integer load");
    assert_eq!(RT_LOAD_BY_CTX[2], 0x6c, "Long load");
    assert_eq!(RT_LOAD_BY_CTX[3], 0x6e, "Single load");
    assert_eq!(RT_LOAD_BY_CTX[4], 0x6f, "Double load");
    assert_eq!(RT_LOAD_BY_CTX[6], 0x6d, "Currency load");
}

#[test]
fn rt_store_table_confirmed_entries() {
    assert_eq!(RT_STORE_BY_CTX[1], 0x70, "Integer store");
    assert_eq!(RT_STORE_BY_CTX[2], 0x71, "Long store");
    assert_eq!(RT_STORE_BY_CTX[3], 0x73, "Single store");
    assert_eq!(RT_STORE_BY_CTX[4], 0x74, "Double store");
    assert_eq!(RT_STORE_BY_CTX[6], 0x72, "Currency store");
}

// ── Variable loads ────────────────────────────────────────────────────────

#[test]
fn double_load_emits_6f_and_frame_offset() {
    // From the empirical probe: Double local `a` is at frame offset -140
    // (0xff74 as i16). Runtime: opcode 0x6f + LE i16 0xff74.
    let mut a = NodeArena::new();
    let v = var_load(&mut a, 4, 0xff74u16 as i16); // typeCtx 4 = Double
    assert_eq!(emit(&a, v), &[0x6f, 0x74, 0xff]);
}

#[test]
fn double_load_second_local_probe_offset() {
    // Double local `b` at frame offset -148 (0xff6c).
    let mut a = NodeArena::new();
    let v = var_load(&mut a, 4, 0xff6cu16 as i16);
    assert_eq!(emit(&a, v), &[0x6f, 0x6c, 0xff]);
}

#[test]
fn long_load_emits_6c_and_frame_offset() {
    let mut a = NodeArena::new();
    let v = var_load(&mut a, 2, -8i16); // typeCtx 2 = Long, offset -8
    assert_eq!(emit(&a, v), &[0x6c, 0xf8, 0xff]);
}

#[test]
fn integer_load_emits_6b_and_frame_offset() {
    let mut a = NodeArena::new();
    let v = var_load(&mut a, 1, -4i16); // typeCtx 1 = Integer, offset -4
    assert_eq!(emit(&a, v), &[0x6b, 0xfc, 0xff]);
}

#[test]
fn single_load_emits_6e_and_frame_offset() {
    let mut a = NodeArena::new();
    let v = var_load(&mut a, 3, -4i16); // typeCtx 3 = Single, offset -4
    assert_eq!(emit(&a, v), &[0x6e, 0xfc, 0xff]);
}

#[test]
fn currency_load_emits_6d_and_frame_offset() {
    let mut a = NodeArena::new();
    let v = var_load(&mut a, 6, -8i16); // typeCtx 6 = Currency, offset -8
    assert_eq!(emit(&a, v), &[0x6d, 0xf8, 0xff]);
}

#[test]
fn node_0x76_routes_to_same_var_load() {
    // Node type 0x76 uses the same body as 0x74.
    let mut a = NodeArena::new();
    let s = sym(&mut a, 0xff74u16 as i16);
    let v = a.alloc(NodeArena::node(0x76, 0, s.0, 4, 0, 0)); // typeCtx 4
    assert_eq!(emit(&a, v), &[0x6f, 0x74, 0xff]);
}

// ── Variable stores ───────────────────────────────────────────────────────

#[test]
fn emit_var_store_double_emits_74_and_offset() {
    let arena = NodeArena::new();
    let mut e = Emitter::new(&arena);
    e.emit_var_store(4, 0xff5cu16 as i16); // Double r at 0xff5c (-164)
    assert_eq!(e.into_bytes(), &[0x74, 0x5c, 0xff]);
}

#[test]
fn emit_var_store_long_emits_71_and_offset() {
    let arena = NodeArena::new();
    let mut e = Emitter::new(&arena);
    e.emit_var_store(2, -4i16); // Long
    assert_eq!(e.into_bytes(), &[0x71, 0xfc, 0xff]);
}

// ── Positive frame offsets (unusual but valid) ────────────────────────────

#[test]
fn positive_frame_offset_encodes_correctly() {
    // Positive offsets are unusual (locals are negative) but must round-trip.
    let mut a = NodeArena::new();
    let v = var_load(&mut a, 4, 8i16);
    assert_eq!(emit(&a, v), &[0x6f, 0x08, 0x00]);
}

// ── No-op branches ────────────────────────────────────────────────────────

#[test]
fn op_zero_emits_nothing() {
    let mut a = NodeArena::new();
    let n = a.alloc(NodeArena::node(0x0, 0, 0, 0, 0, 0));
    assert_eq!(emit(&a, n), &[]);
}

#[test]
fn op_in_0x2c_to_0x35_emits_nothing() {
    let mut a = NodeArena::new();
    let n = a.alloc(NodeArena::node(0x30, 0, 0, 0, 0, 0));
    assert_eq!(emit(&a, n), &[]);
}

// ── Binary-op opcode dispatch ─────────────────────────────────────────────
//
// Each test builds two Long-typed variable-load nodes (as stand-ins for the
// two operands), wraps them in a binary-op node, and verifies the exact
// byte sequence: [lhs-load:3 bytes] [rhs-load:3 bytes] [op-byte(s)].
//
// Opcode derivation (verified against RT_BINOP_BASE / RT_TYPE_OFFSET /
// RT_OPCODE_BYTE tables extracted from the DLL binary):
//
//   AND (op 0x23=35): base=RT_BINOP_BASE[35]=0x0021=33, Long offset=2,
//     n_opc=35, RT_OPCODE_BYTE[35]=0xc4  → 1 byte 0xc4
//
//   OR  (op 0x21=33): base=RT_BINOP_BASE[33]=0x0019=25, Long offset=2,
//     n_opc=27, RT_OPCODE_BYTE[27]=0xc5  → 1 byte 0xc5
//
//   XOR (op 0x22=34): base=RT_BINOP_BASE[34]=0x0011=17, Long offset=2,
//     n_opc=19, RT_OPCODE_BYTE[19]=0xfb  → 2 bytes [0xfb, 0x13]
//
//   ADD (op 0x16=22): base=RT_BINOP_BASE[22]=0x008e=142, Long offset=2,
//     n_opc=144, RT_OPCODE_BYTE[144]=0xaa  → 1 byte 0xaa
//
//   ADD (op 0x16=22): Integer offset=1,
//     n_opc=143, RT_OPCODE_BYTE[143]=0xa9  → 1 byte 0xa9
//
//   EQ  (op 0x6):  base=RT_BINOP_BASE[6]=0x00be=190, Long offset=2,
//     n_opc=192, RT_OPCODE_BYTE[192]=0xc3  → 1 byte 0xc3

/// Helper: allocate a minimal Long-load node at the given frame offset.
fn long_load(a: &mut NodeArena, offset: i16) -> NodeRef {
    var_load(a, 2, offset) // typeCtx 2 = Long
}

/// Helper: allocate a binary-op node with opcode `op`, result type_tag
/// `type_tag`, and the given LHS / RHS children.
fn binop(a: &mut NodeArena, op: u16, type_tag: u16, lhs: NodeRef, rhs: NodeRef) -> NodeRef {
    a.alloc(NodeArena::node(op, type_tag, lhs.0, rhs.0, 0, 0))
}

#[test]
fn and_long_emits_c4() {
    let mut a = NodeArena::new();
    let lhs = long_load(&mut a, -8);
    let rhs = long_load(&mut a, -12);
    let n = binop(&mut a, 0x23, 8, lhs, rhs); // AND, Long result
    assert_eq!(
        emit(&a, n),
        &[0x6c, 0xf8, 0xff, 0x6c, 0xf4, 0xff, 0xc4]
    );
}

#[test]
fn or_long_emits_c5() {
    let mut a = NodeArena::new();
    let lhs = long_load(&mut a, -8);
    let rhs = long_load(&mut a, -12);
    let n = binop(&mut a, 0x21, 8, lhs, rhs); // OR, Long result
    assert_eq!(
        emit(&a, n),
        &[0x6c, 0xf8, 0xff, 0x6c, 0xf4, 0xff, 0xc5]
    );
}

#[test]
fn xor_long_emits_fb_13() {
    let mut a = NodeArena::new();
    let lhs = long_load(&mut a, -8);
    let rhs = long_load(&mut a, -12);
    let n = binop(&mut a, 0x22, 8, lhs, rhs); // XOR, Long result
    // n_opc=19=0x13 → rt_byte=0xfb → extended form [0xfb, 0x13]
    assert_eq!(
        emit(&a, n),
        &[0x6c, 0xf8, 0xff, 0x6c, 0xf4, 0xff, 0xfb, 0x13]
    );
}

#[test]
fn add_long_emits_aa() {
    let mut a = NodeArena::new();
    let lhs = long_load(&mut a, -8);
    let rhs = long_load(&mut a, -12);
    let n = binop(&mut a, 0x16, 8, lhs, rhs); // ADD, Long result
    assert_eq!(
        emit(&a, n),
        &[0x6c, 0xf8, 0xff, 0x6c, 0xf4, 0xff, 0xaa]
    );
}

#[test]
fn add_integer_emits_a9() {
    let mut a = NodeArena::new();
    let lhs = var_load(&mut a, 1, -4); // Integer
    let rhs = var_load(&mut a, 1, -6);
    let n = binop(&mut a, 0x16, 6, lhs, rhs); // ADD, Integer result
    assert_eq!(
        emit(&a, n),
        &[0x6b, 0xfc, 0xff, 0x6b, 0xfa, 0xff, 0xa9]
    );
}

// ── Comparison operators (bound opcodes 0x26–0x2b) ─────────────────────────
//
// These route through EbEmitBinaryOperation2's comparison branch
// (RT_DISPATCH_FLAG[op] & 0x10 != 0), which selects the typed opcode from the
// *LHS operand's* type tag — not the comparison node's own type tag. Bound
// opcodes: eq=0x26, ne=0x27, le=0x28, ge=0x29, lt=0x2a, gt=0x2b.
//
// Oracle ground truth (re_lab/pcode_lab/cmp_survey.py), Long operands at
// frame offsets a=-136 (0xff78), b=-140 (0xff74):
//   eq=0xc7, ne=0xcc, le=0xd6, ge=0xe0, lt=0xd1, gt=0xdb

/// Long-typed comparison operand: VB6 type tag 8, load context 2.
fn long_cmp_operand(a: &mut NodeArena, offset: i16) -> NodeRef {
    var_load_typed(a, 8, 2, offset)
}

/// Build a comparison node (own type_tag 0 — irrelevant for non-UDT operands).
fn cmp(a: &mut NodeArena, op: u16, lhs: NodeRef, rhs: NodeRef) -> NodeRef {
    a.alloc(NodeArena::node(op, 0, lhs.0, rhs.0, 0, 0))
}

#[test]
fn eq_long_emits_c7() {
    let mut a = NodeArena::new();
    let lhs = long_cmp_operand(&mut a, 0xff78u16 as i16);
    let rhs = long_cmp_operand(&mut a, 0xff74u16 as i16);
    let n = cmp(&mut a, 0x26, lhs, rhs);
    assert_eq!(emit(&a, n), &[0x6c, 0x78, 0xff, 0x6c, 0x74, 0xff, 0xc7]);
}

#[test]
fn ne_long_emits_cc() {
    let mut a = NodeArena::new();
    let lhs = long_cmp_operand(&mut a, 0xff78u16 as i16);
    let rhs = long_cmp_operand(&mut a, 0xff74u16 as i16);
    let n = cmp(&mut a, 0x27, lhs, rhs);
    assert_eq!(emit(&a, n), &[0x6c, 0x78, 0xff, 0x6c, 0x74, 0xff, 0xcc]);
}

#[test]
fn le_long_emits_d6() {
    let mut a = NodeArena::new();
    let lhs = long_cmp_operand(&mut a, 0xff78u16 as i16);
    let rhs = long_cmp_operand(&mut a, 0xff74u16 as i16);
    let n = cmp(&mut a, 0x28, lhs, rhs);
    assert_eq!(emit(&a, n), &[0x6c, 0x78, 0xff, 0x6c, 0x74, 0xff, 0xd6]);
}

#[test]
fn ge_long_emits_e0() {
    let mut a = NodeArena::new();
    let lhs = long_cmp_operand(&mut a, 0xff78u16 as i16);
    let rhs = long_cmp_operand(&mut a, 0xff74u16 as i16);
    let n = cmp(&mut a, 0x29, lhs, rhs);
    assert_eq!(emit(&a, n), &[0x6c, 0x78, 0xff, 0x6c, 0x74, 0xff, 0xe0]);
}

#[test]
fn lt_long_emits_d1() {
    let mut a = NodeArena::new();
    let lhs = long_cmp_operand(&mut a, 0xff78u16 as i16);
    let rhs = long_cmp_operand(&mut a, 0xff74u16 as i16);
    let n = cmp(&mut a, 0x2a, lhs, rhs);
    assert_eq!(emit(&a, n), &[0x6c, 0x78, 0xff, 0x6c, 0x74, 0xff, 0xd1]);
}

#[test]
fn gt_long_emits_db() {
    let mut a = NodeArena::new();
    let lhs = long_cmp_operand(&mut a, 0xff78u16 as i16);
    let rhs = long_cmp_operand(&mut a, 0xff74u16 as i16);
    let n = cmp(&mut a, 0x2b, lhs, rhs);
    assert_eq!(emit(&a, n), &[0x6c, 0x78, 0xff, 0x6c, 0x74, 0xff, 0xdb]);
}

#[test]
fn eq_integer_emits_c6() {
    // Integer operands (VB6 type tag 6, load context 1) → eq opcode 0xc6.
    let mut a = NodeArena::new();
    let lhs = var_load_typed(&mut a, 6, 1, -4);
    let rhs = var_load_typed(&mut a, 6, 1, -6);
    let n = cmp(&mut a, 0x26, lhs, rhs);
    assert_eq!(emit(&a, n), &[0x6b, 0xfc, 0xff, 0x6b, 0xfa, 0xff, 0xc6]);
}

#[test]
fn eq_double_emits_c8() {
    // Double operands (VB6 type tag 11, load context 4) → eq opcode 0xc8.
    let mut a = NodeArena::new();
    let lhs = var_load_typed(&mut a, 11, 4, 0xff74u16 as i16);
    let rhs = var_load_typed(&mut a, 11, 4, 0xff6cu16 as i16);
    let n = cmp(&mut a, 0x26, lhs, rhs);
    assert_eq!(emit(&a, n), &[0x6f, 0x74, 0xff, 0x6f, 0x6c, 0xff, 0xc8]);
}

// ── Literal emission ──────────────────────────────────────────────────────

/// Build an integer-literal node (op=1) with the given type_tag and value.
fn int_lit(a: &mut NodeArena, type_tag: u16, val: i32) -> NodeRef {
    a.alloc(NodeArena::node(1, type_tag, val as u32, 0, 0, 0))
}

/// Build a float-literal node (op=3) with the given type_tag.
/// The 8-byte f64 value is stored in word[4]/word[5].
fn float_lit(a: &mut NodeArena, type_tag: u16, value: f64) -> NodeRef {
    let bits = value.to_bits();
    let lo = bits as u32;
    let hi = (bits >> 32) as u32;
    a.alloc(NodeArena::node(3, type_tag, lo, hi, 0, 0))
}

/// Build a Currency-literal node (op=2). The value is the raw i64×10000
/// Currency representation, stored in word[4]/word[5].
fn currency_lit(a: &mut NodeArena, raw_val: i64) -> NodeRef {
    let lo = raw_val as u32;
    let hi = (raw_val >> 32) as u32;
    a.alloc(NodeArena::node(2, 0, lo, hi, 0, 0))
}

// op=1, type_tag 6 (Integer), small value (-128..127):
//   n_opc=0x41a=1050 → rt_byte=0xf4 < 0xfb → emit [0xf4, value_byte]
#[test]
fn integer_small_lit_emits_f4_and_byte() {
    let mut a = NodeArena::new();
    let n = int_lit(&mut a, 6, 5);
    assert_eq!(emit(&a, n), &[0xf4, 0x05]);
}

#[test]
fn integer_small_lit_negative_emits_signed_byte() {
    let mut a = NodeArena::new();
    let n = int_lit(&mut a, 6, -3);
    assert_eq!(emit(&a, n), &[0xf4, 0xfd]); // -3 as u8 = 0xfd
}

// op=1, type_tag 6 (Integer), large value (> 127):
//   n_opc=0x3b5=949 → rt_byte=0xf3 < 0xfb → emit [0xf3] + i16 LE
#[test]
fn integer_large_lit_emits_f3_and_i16() {
    let mut a = NodeArena::new();
    let n = int_lit(&mut a, 6, 300);
    assert_eq!(emit(&a, n), &[0xf3, 0x2c, 0x01]); // 300 = 0x012c
}

// op=1, type_tag 8 (Long):
//   n_opc=0x3b8=952 → rt_byte=0xf5 < 0xfb → emit [0xf5] + i32 LE
#[test]
fn long_lit_emits_f5_and_i32() {
    let mut a = NodeArena::new();
    let n = int_lit(&mut a, 8, 0x00012345);
    assert_eq!(emit(&a, n), &[0xf5, 0x45, 0x23, 0x01, 0x00]);
}

// op=2 (Currency literal):
//   n_opc=0x3bb=955 → rt_byte=0xf6 < 0xfb → emit [0xf6] + 8 bytes raw
#[test]
fn currency_lit_emits_f6_and_8_bytes() {
    let mut a = NodeArena::new();
    // Currency 1.0 is stored as 10000 (i64 × 10000 scale)
    let n = currency_lit(&mut a, 10_000);
    let mut expected = vec![0xf6];
    expected.extend_from_slice(&10_000_i64.to_le_bytes());
    assert_eq!(emit(&a, n), expected.as_slice());
}

// op=3, type_tag=10 (Single), non-assign context (call_ctx=0):
//   n_opc=0x3b9=953 → rt_byte=0xf5 → emit [0xf5] + 4-byte f32 LE
#[test]
fn single_lit_non_assign_emits_f5_and_f32() {
    let mut a = NodeArena::new();
    let n = float_lit(&mut a, 10, 1.5_f64);
    let mut expected = vec![0xf5];
    expected.extend_from_slice(&(1.5_f32).to_bits().to_le_bytes());
    assert_eq!(emit(&a, n), expected.as_slice());
}

// op=3, type_tag=10 (Single), assign context (call_ctx=2):
//   n_opc=0x3ba=954 → rt_byte=0xf9 → emit [0xf9] + 4-byte f32 LE
#[test]
fn single_lit_assign_ctx_emits_f9_and_f32() {
    let mut a = NodeArena::new();
    let n = float_lit(&mut a, 10, 2.0_f64);
    let mut e = Emitter::new(&a);
    e.emit_expr(n, 2); // call_ctx=2 = assign context
    let mut expected = vec![0xf9];
    expected.extend_from_slice(&(2.0_f32).to_bits().to_le_bytes());
    assert_eq!(e.into_bytes(), expected.as_slice());
}

// op=3, type_tag=11 (Double), non-assign (call_ctx=0):
//   n_opc=0x3bc=956 → rt_byte=0xf6 → emit [0xf6] + 8-byte f64 LE
#[test]
fn double_lit_non_assign_emits_f6_and_f64() {
    let mut a = NodeArena::new();
    let val = 3.14_f64;
    let n = float_lit(&mut a, 11, val);
    let mut expected = vec![0xf6];
    expected.extend_from_slice(&val.to_bits().to_le_bytes());
    assert_eq!(emit(&a, n), expected.as_slice());
}

// op=4 (String literal), null string (word[1] bit 15 set):
//   EbEmitValue2(0x3b8) + EbEmitDword(0) → [0xf5, 0x00, 0x00, 0x00, 0x00]
#[test]
fn string_null_lit_emits_f5_and_four_zeros() {
    let mut a = NodeArena::new();
    let mut n = NodeArena::node(4, 0, 0, 0, 0, 0);
    // *(byte *)((int)pNode + 5) & 0x80 checks bit 15 of w[1] (byte at node+5
    // is the high byte of the low halfword of w[1], i.e. (w[1] >> 8) & 0xff;
    // bit 0x80 of that byte = bit 15 of w[1]).
    n.w[1] = 1 << 15;
    let r = a.alloc(n);
    assert_eq!(emit(&a, r), &[0xf5, 0x00, 0x00, 0x00, 0x00]);
}

// op=3, type_tag=12 (Date), assign context (call_ctx=2):
//   n_opc=0x3bd=957 → rt_byte=0xfa → emit [0xfa] + 8-byte f64 LE
#[test]
fn date_lit_assign_ctx_emits_fa_and_f64() {
    let mut a = NodeArena::new();
    let val = 44_926.0_f64; // some date serial
    let n = float_lit(&mut a, 12, val);
    let mut e = Emitter::new(&a);
    e.emit_expr(n, 2); // assign context
    let mut expected = vec![0xfa];
    expected.extend_from_slice(&val.to_bits().to_le_bytes());
    assert_eq!(e.into_bytes(), expected.as_slice());
}

// ── Op 0x36 (Like operator) ───────────────────────────────────────────────
//
// EbEmitStatement case 0x36: emit LHS(call_ctx=1), RHS(call_ctx=1), then
// emit_value2(0xd2).  RT_OPCODE_BYTE[0xd2=210]=0xfb ≥ 0xfb → extended form
// [0xfb, 0xd2].
//
// We use small integer literals as operands because emit_int_literal ignores
// call_ctx — the key behaviour under test is the Like opcode itself.

#[test]
fn op_0x36_like_emits_operands_then_fb_d2() {
    // Like lhs rhs → [lhs:2] [rhs:2] [0xfb, 0xd2]
    let mut a = NodeArena::new();
    let lhs = int_lit(&mut a, 6, 5); // Integer small: [0xf4, 0x05]
    let rhs = int_lit(&mut a, 6, 3); // Integer small: [0xf4, 0x03]
    let n = a.alloc(NodeArena::node(0x36, 0, lhs.0, rhs.0, 0, 0));
    assert_eq!(
        emit(&a, n),
        &[0xf4, 0x05, 0xf4, 0x03, 0xfb, 0xd2]
    );
}

// ── Op 0x24 (Is operator) ─────────────────────────────────────────────────
//
// EbEmitStatement case 0x24: explicit Is path — NOT routed through
// EbEmitBinaryOperation2.  Both operands use call_ctx=1; the comparison
// opcode depends on the node's type_tag:
//
//   type_tag=0x10 → emit_value2(0xf0):   RT_OPCODE_BYTE[0xf0]=0x2a  → [0x2a]
//   type_tag=2    → emit_value2(0x18b):  RT_OPCODE_BYTE[0x18b]=0xfc → [0xfc, 0x8b]
//   type_tag=10   → emit_value2(0x189) only when outer call_ctx is 1 or 3:
//                   RT_OPCODE_BYTE[0x189]=0x37 → [0x37]
//   type_tag=0xb/0xc → emit_value2(0x18a): RT_OPCODE_BYTE[0x18a]=0x39 → [0x39]
//   other type_tag → no opcode (EbValidateTypeOperation returns 0)
//
// Operands are integer literals (call_ctx is ignored by emit_int_literal).

#[test]
fn is_op_type_tag_0x10_emits_operands_then_2a() {
    // Is with type_tag=0x10 → RT_OPCODE_BYTE[0xf0]=0x2a
    let mut a = NodeArena::new();
    let lhs = int_lit(&mut a, 6, 5); // [0xf4, 0x05]
    let rhs = int_lit(&mut a, 6, 3); // [0xf4, 0x03]
    let n = a.alloc(NodeArena::node(0x24, 0x10, lhs.0, rhs.0, 0, 0));
    assert_eq!(
        emit(&a, n),
        &[0xf4, 0x05, 0xf4, 0x03, 0x2a]
    );
}

#[test]
fn is_op_type_tag_2_emits_operands_then_fc_8b() {
    // Is with type_tag=2 (String) → emit_value2(0x18b) → [0xfc, 0x8b]
    let mut a = NodeArena::new();
    let lhs = int_lit(&mut a, 6, 1); // [0xf4, 0x01]
    let rhs = int_lit(&mut a, 6, 2); // [0xf4, 0x02]
    let n = a.alloc(NodeArena::node(0x24, 2, lhs.0, rhs.0, 0, 0));
    assert_eq!(
        emit(&a, n),
        &[0xf4, 0x01, 0xf4, 0x02, 0xfc, 0x8b]
    );
}

#[test]
fn is_op_type_tag_single_outer_call_ctx_1_emits_operands_then_37() {
    // Is with type_tag=10 (Single) and outer call_ctx=1 → emit_value2(0x189)
    // RT_OPCODE_BYTE[0x189]=0x37 < 0xfb → single byte [0x37]
    let mut a = NodeArena::new();
    let lhs = int_lit(&mut a, 6, 4); // [0xf4, 0x04]
    let rhs = int_lit(&mut a, 6, 7); // [0xf4, 0x07]
    let n = a.alloc(NodeArena::node(0x24, 10, lhs.0, rhs.0, 0, 0));
    let mut e = Emitter::new(&a);
    e.emit_expr(n, 1); // outer call_ctx=1
    assert_eq!(e.into_bytes(), &[0xf4, 0x04, 0xf4, 0x07, 0x37]);
}

#[test]
fn is_op_type_tag_single_outer_call_ctx_0_emits_operands_only() {
    // Is with type_tag=10 (Single) and outer call_ctx=0 → no extra opcode
    // (EbValidateTypeOperation requires nTypeFlags 1 or 3)
    let mut a = NodeArena::new();
    let lhs = int_lit(&mut a, 6, 4); // [0xf4, 0x04]
    let rhs = int_lit(&mut a, 6, 7); // [0xf4, 0x07]
    let n = a.alloc(NodeArena::node(0x24, 10, lhs.0, rhs.0, 0, 0));
    assert_eq!(emit(&a, n), &[0xf4, 0x04, 0xf4, 0x07]);
}

#[test]
fn is_op_type_tag_double_outer_call_ctx_1_emits_operands_then_39() {
    // Is with type_tag=0xb (Double) and outer call_ctx=1 → emit_value2(0x18a)
    // RT_OPCODE_BYTE[0x18a]=0x39 < 0xfb → single byte [0x39]
    let mut a = NodeArena::new();
    let lhs = int_lit(&mut a, 6, 1); // [0xf4, 0x01]
    let rhs = int_lit(&mut a, 6, 2); // [0xf4, 0x02]
    let n = a.alloc(NodeArena::node(0x24, 0xb, lhs.0, rhs.0, 0, 0));
    let mut e = Emitter::new(&a);
    e.emit_expr(n, 1); // outer call_ctx=1
    assert_eq!(e.into_bytes(), &[0xf4, 0x01, 0xf4, 0x02, 0x39]);
}

#[test]
fn is_op_unrecognised_type_tag_emits_operands_only() {
    // type_tag=7 (not in the switch) → EbValidateTypeOperation returns 0,
    // no comparison opcode is emitted after the operands.
    let mut a = NodeArena::new();
    let lhs = int_lit(&mut a, 6, 9); // [0xf4, 0x09]
    let rhs = int_lit(&mut a, 6, 8); // [0xf4, 0x08]
    let n = a.alloc(NodeArena::node(0x24, 7, lhs.0, rhs.0, 0, 0));
    assert_eq!(emit(&a, n), &[0xf4, 0x09, 0xf4, 0x08]);
}

// ── RT_OPCODE_BYTE spot checks for Is/Like opcodes ────────────────────────

#[test]
fn rt_opcode_byte_is_like_table_entries() {
    use crate::tables::RT_OPCODE_BYTE;
    assert_eq!(RT_OPCODE_BYTE[0xd2], 0xfb, "Like (0xd2) → extended form");
    assert_eq!(RT_OPCODE_BYTE[0xf0], 0x2a, "Is Object (0xf0) → 0x2a");
    assert_eq!(RT_OPCODE_BYTE[0x18b], 0xfc, "Is String (0x18b) → extended form");
    assert_eq!(RT_OPCODE_BYTE[0x189], 0x37, "Is Single (0x189) → 0x37");
    assert_eq!(RT_OPCODE_BYTE[0x18a], 0x39, "Is Double/Currency (0x18a) → 0x39");
}
