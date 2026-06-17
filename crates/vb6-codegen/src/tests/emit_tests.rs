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

// ── emit_reference 0x4000 branch (EbEmitExpression2 LAB_0fab3b03) ─────────

fn ref_bytes(desc: &crate::emit::RefDescriptor, n_op: i32, f_flags: u32, n_type: i32) -> Vec<u8> {
    let arena = NodeArena::new();
    let mut e = Emitter::new(&arena);
    e.emit_reference(desc, n_op, f_flags, n_type);
    e.into_bytes()
}

fn local_desc(offset: i16) -> crate::emit::RefDescriptor {
    crate::emit::RefDescriptor { kind: 1, operand: offset as u16, word6: 0, word8: 0 }
}

#[test]
fn emit_reference_4000_object_ntype_emits_0x3e() {
    // EbEmitExpression2 LAB_0fab3b03: f_flags & 0x4000, nType == 0x10 (Object/Dispatch).
    // u_var6 = 0x23f; RT_OPCODE_BYTE[0x23f] = 0x3e (< 0xfb).
    // emit_opcode2(0x23f, 0): emit_value2 → [0x3e], emit_word(0) → [0x00, 0x00].
    use crate::tables::RT_OPCODE_BYTE;
    assert_eq!(RT_OPCODE_BYTE[0x23f], 0x3e);
    let desc = local_desc(0);
    let bytes = ref_bytes(&desc, 1, 0x4000, 0x10);
    assert_eq!(bytes, &[0x3e, 0x00, 0x00]);
}

#[test]
fn emit_reference_4000_other_ntype_emits_0x262_opcode() {
    // EbEmitExpression2 LAB_0fab3b03: f_flags & 0x4000, nType == 5 (Single).
    // u_var6 = 0x262; emit_opcode2(0x262, 0).
    // RT_OPCODE_BYTE[0x262] = the value at that index (table-confirmed).
    use crate::tables::RT_OPCODE_BYTE;
    let expected_opcode = RT_OPCODE_BYTE[0x262];
    let desc = local_desc(0);
    let bytes = ref_bytes(&desc, 1, 0x4000, 5);
    // emit_value2(0x262): if < 0xfb → [byte], else → [byte, low8].
    // Then emit_word(0) → [0x00, 0x00].
    if expected_opcode < 0xfb {
        assert_eq!(bytes, &[expected_opcode, 0x00, 0x00]);
    } else {
        assert_eq!(bytes, &[expected_opcode, (0x262u16 & 0xff) as u8, 0x00, 0x00]);
    }
}

#[test]
fn emit_reference_4000_nop2_object_also_emits_0x3e() {
    // nOp==2 falls through to the same 0x4000 block as nOp==1 (verified in C).
    let desc = local_desc(0);
    let bytes = ref_bytes(&desc, 2, 0x4000, 0x10);
    assert_eq!(bytes, &[0x3e, 0x00, 0x00]);
}

#[test]
fn emit_reference_kind1_nop1_long_normal_load() {
    // Baseline (no 0x4000): kind==1, nOp==1, nType==8 (Long vb-type).
    // nType 8 → RT_TYPE_OFFSET[8] = ? Long is the standard Integer load path.
    // Actually nType in EbEmitExpression2 is the *internal* VB6 type, not vb-type.
    // For a simple local-load (no 0x4000 flag), kind==1 → u_var7=0x1e0.
    // nType=8 → RT_TYPE_OFFSET[8]; load result: u_var7 | off.
    // We test the real oracle path: Long local at -136 → [0x6c, 0x78, 0xff].
    // That goes through emit_var_load (0x74 node), not emit_reference.
    // Here we test emit_reference kind==1, nOp==1, nType==2 (integer vb-type → offset 1).
    use crate::tables::RT_TYPE_OFFSET;
    let off = RT_TYPE_OFFSET[2] as i32;   // Integer offset
    use crate::tables::RT_OPCODE_BYTE;
    let u_var7 = 0x1e0i32;
    let expected_idx = if off == 10 { u_var7 | 4 } else if off == 9 { u_var7 | 1 } else { u_var7 | off };
    let expected_opcode = RT_OPCODE_BYTE[expected_idx as usize];
    let desc = crate::emit::RefDescriptor { kind: 1, operand: 0xff7au16, word6: 0, word8: 0 };
    let bytes = ref_bytes(&desc, 1, 0, 2);
    // emit_opcode2(expected_idx, 0xff7a):
    // emit_value2(expected_idx) → [expected_opcode] (if < 0xfb)
    // emit_word(0xff7a) → [0x7a, 0xff]
    if expected_opcode < 0xfb {
        assert_eq!(bytes, &[expected_opcode, 0x7a, 0xff]);
    } else {
        assert_eq!(
            bytes,
            &[expected_opcode, (expected_idx & 0xff) as u8, 0x7a, 0xff]
        );
    }
}

#[test]
fn emit_reference_kind2_byref_promotes_to_nop2() {
    // kind==2 (argument) with word6 bit 0 set (ByRef) → n_op promoted to 2.
    // With nOp==2 and no flags, the opcode index is from the nOp==2 arm.
    // We verify the output differs from the nOp==1 arm (the ByRef adjustment).
    // kind==2 u_var7=0x210; nOp==1 (no byref): u_var7|off; nOp==2 (byref): different index.
    use crate::tables::RT_TYPE_OFFSET;
    use crate::tables::RT_OPCODE_BYTE;
    let u_var7 = 0x210i32;
    let off = RT_TYPE_OFFSET[2] as i32;
    let nop1_idx = if off == 10 { u_var7 | 4 } else if off == 9 { u_var7 | 1 } else { u_var7 | off };
    // nOp2 path: off stays, then +6 if off==3||off==4; otherwise same|u_var7.
    let mut nop2_idx = if off == 10 { 4 } else if off == 9 { 1 } else { off };
    let nop2_base = nop2_idx | u_var7;
    if nop2_idx == 3 || nop2_idx == 4 { nop2_idx = nop2_base + 6; } else { nop2_idx = nop2_base; }
    // byref (word6 bit 0 = 1) → n_op becomes 2.
    let desc_byref = crate::emit::RefDescriptor { kind: 2, operand: 0x000c, word6: 1, word8: 0 };
    let desc_byval = crate::emit::RefDescriptor { kind: 2, operand: 0x000c, word6: 0, word8: 0 };
    let b_byref = ref_bytes(&desc_byref, 1, 0, 2);
    let b_byval = ref_bytes(&desc_byval, 1, 0, 2);
    // The byref descriptor produces nOp==2 output, byval produces nOp==1.
    let op_byref = RT_OPCODE_BYTE[nop2_idx as usize];
    let op_byval = RT_OPCODE_BYTE[nop1_idx as usize];
    assert_eq!(b_byval[0], op_byval);
    assert_eq!(b_byref[0], op_byref);
}

// ── emit_assign_op (case 0xe, op-kind 0) ─────────────────────────────────────
//
// These tests exercise case 0xe of emit_expr (the assignment dispatch) by
// building synthetic assignment nodes and verifying the output byte sequence.
//
// Node layout for case 0xe:
//   word[0] = (lhs_kind << 16) | 0xe
//   word[1] = flags (0 = op-kind 0, no 0x4000 Set flag, no byte-5 0x80 flag)
//   word[4] = NodeRef of the RHS child
//
// A "null-emit RHS" node has opcode=0 (hits the EbEmitStatement guard → return 0
// immediately) but carries the desired type_tag in its high word so that
// emit_assign_op reads the correct rhs_kind.
//
// Convention: index 0 in NodeArena is reserved as the null-pointer sentinel.
// All real nodes are allocated from index 1 onward.

/// Allocate a null-sentinel dummy at index 0, then the RHS and assignment nodes.
/// Returns (arena, assign_node_ref).
fn assign_node(lhs_kind: u16, rhs_kind: u16) -> (NodeArena, NodeRef) {
    let mut a = NodeArena::new();
    let _null = a.alloc(NodeArena::node(0, 0, 0, 0, 0, 0)); // idx 0 = null sentinel
    let rhs = a.alloc(NodeArena::node(0, rhs_kind, 0, 0, 0, 0)); // idx 1: null-emit RHS
    let assign = a.alloc(NodeArena::node(0xe, lhs_kind, rhs.0, 0, 0, 0)); // idx 2
    (a, assign)
}

#[test]
fn assign_currency_lhs_rhs_kind_0xb_emits_0xf2() {
    // Currency LHS (kind=0xc) + RHS kind 0xb (Boolean/Variant):
    // EbEmitAssignOp special-cases Currency LHS: rhs_kind 0xb → emit_value2(0x147).
    // RT_OPCODE_BYTE[0x147=327]: row 40 (line 198 of tables), col 7 = 0xf2 < 0xfb → [0xf2].
    let (a, n) = assign_node(0xc, 0xb);
    assert_eq!(emit(&a, n), &[0xf2]);
}

#[test]
fn assign_currency_lhs_rhs_kind_0xf_emits_fc4f() {
    // Currency LHS (kind=0xc) + RHS kind 0xf (Object/Variant): emit_value2(0x14f).
    // RT_OPCODE_BYTE[0x14f=335]: row 41 (line 199), col 7 = 0xfc → [0xfc, 0x4f].
    let (a, n) = assign_node(0xc, 0xf);
    assert_eq!(emit(&a, n), &[0xfc, 0x4f]);
}

#[test]
fn assign_variant_group_no_flag_emits_nothing() {
    // LHS kind 0xb and RHS kind 0xb: both in {10,0xb,0xc} (Variant/Boolean group).
    // Byte-5 flag 0x80 is clear (word[1]=0) → EbEmitAssignOp returns immediately.
    // Trailing EbValidateTypeOperation(0xb, 0, context=0): context=0 ≠ 3 and ≠ 1
    // → returns 1, no bytes emitted.
    let (a, n) = assign_node(0xb, 0xb);
    assert_eq!(emit(&a, n), &[]);
}

#[test]
fn assign_default_numeric_kind5_emits_fc0c() {
    // LHS kind=5 and RHS kind=5: falls to the default numeric path.
    // RT_TYPE_KIND_CLASS[5]=0 → RT_ASSIGN_BASE_OPCODE[0]=0x10c.
    // rhs_class = RT_TYPE_KIND_CLASS[5] = 0 → emit_value2(0 + 0x10c = 0x10c = 268).
    // RT_OPCODE_BYTE[268]: row 33 (line 191 of tables), col 4 = 0xfc → [0xfc, 0x0c].
    let (a, n) = assign_node(5, 5);
    assert_eq!(emit(&a, n), &[0xfc, 0x0c]);
}

#[test]
fn assign_default_numeric_kind6_same_emits_fc15() {
    // LHS kind=6, RHS kind=6: RT_TYPE_KIND_CLASS[6]=1 → RT_ASSIGN_BASE_OPCODE[1]=0x114.
    // rhs_class=1 → emit_value2(1 + 0x114 = 0x115 = 277).
    // RT_OPCODE_BYTE[277]: row 34 (line 192 of tables), col 5 = 0xfc → [0xfc, 0x15].
    // (kind 6 is not in the Variant group {10,0xb,0xc} so the default numeric path runs.)
    let (a, n) = assign_node(6, 6);
    assert_eq!(emit(&a, n), &[0xfc, 0x15]);
}

// ── traverse_node_tree ────────────────────────────────────────────────────────
//
// EbTraverseNodeTree walks a singly-linked list (opcodes 0x37/0x33) and emits
// each child statement.  Emission order is LAST-TO-FIRST: the function recurses
// on the sibling (word[5]) before emitting the current child (word[4]), so for a
// list [A, B] the byte order is: B's bytes, then A's bytes.

fn global_long_load(a: &mut NodeArena, field_offset: u16) -> NodeRef {
    // Opcode 0x77: emit_global_node_load → [0x94, module_desc_lo, module_desc_hi, field_lo, field_hi]
    // module_desc=8 (default), field_offset as given.
    let packed = 0x0008u32 | ((field_offset as u32) << 16);
    a.alloc(NodeArena::node(0x77, 0, packed, 2, 0, 0)) // type_ctx=Long=2
}

#[test]
fn traverse_single_list_node_emits_child() {
    // List node (0x37) with one child (global Long at field=0), no sibling.
    // Expected: child bytes [0x94, 0x08, 0x00, 0x00, 0x00].
    let mut a = NodeArena::new();
    let _null = a.alloc(NodeArena::node(0, 0, 0, 0, 0, 0)); // null sentinel at idx 0
    let child = global_long_load(&mut a, 0);
    let list = a.alloc(NodeArena::node(0x37, 0, child.0, 0, 0, 0)); // word[5]=0 = no sibling
    let mut e = Emitter::new(&a);
    e.traverse_node_tree(list, 0);
    assert_eq!(e.into_bytes(), &[0x94, 0x08, 0x00, 0x00, 0x00]);
}

#[test]
fn traverse_two_element_list_emits_last_to_first() {
    // List [A=field0, B=field4]:
    //   list_A → word[4]=child_A(field=0), word[5]=list_B
    //   list_B → word[4]=child_B(field=4), word[5]=0
    // EbTraverseNodeTree recurses on sibling first → emits B then A.
    // Expected: [0x94,0x08,0x00,0x04,0x00, 0x94,0x08,0x00,0x00,0x00]
    let mut a = NodeArena::new();
    let _null = a.alloc(NodeArena::node(0, 0, 0, 0, 0, 0)); // null sentinel
    let child_a = global_long_load(&mut a, 0);  // field=0
    let child_b = global_long_load(&mut a, 4);  // field=4
    let list_b = a.alloc(NodeArena::node(0x37, 0, child_b.0, 0, 0, 0));
    let list_a = a.alloc(NodeArena::node(0x37, 0, child_a.0, list_b.0, 0, 0));
    let mut e = Emitter::new(&a);
    e.traverse_node_tree(list_a, 0);
    assert_eq!(
        e.into_bytes(),
        &[0x94, 0x08, 0x00, 0x04, 0x00, 0x94, 0x08, 0x00, 0x00, 0x00]
    );
}

#[test]
fn traverse_opcode_0x33_list_also_works() {
    // Opcode 0x33 is the other list opcode; behavior is identical to 0x37.
    let mut a = NodeArena::new();
    let _null = a.alloc(NodeArena::node(0, 0, 0, 0, 0, 0));
    let child = global_long_load(&mut a, 0);
    let list = a.alloc(NodeArena::node(0x33, 0, child.0, 0, 0, 0));
    let mut e = Emitter::new(&a);
    e.traverse_node_tree(list, 0);
    assert_eq!(e.into_bytes(), &[0x94, 0x08, 0x00, 0x00, 0x00]);
}

#[test]
fn traverse_non_list_node_emits_it_directly() {
    // A non-list node passed to traverse_node_tree is emitted directly
    // (it is NOT a list node — the opcode check doesn't match 0x37/0x33).
    let mut a = NodeArena::new();
    let _null = a.alloc(NodeArena::node(0, 0, 0, 0, 0, 0));
    let g = global_long_load(&mut a, 0);
    let mut e = Emitter::new(&a);
    e.traverse_node_tree(g, 0);
    assert_eq!(e.into_bytes(), &[0x94, 0x08, 0x00, 0x00, 0x00]);
}
