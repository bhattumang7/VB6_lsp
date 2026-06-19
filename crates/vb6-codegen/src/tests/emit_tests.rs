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
/// (the binary-operation comparison branch), so operands of a comparison
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
// RT_OPCODE_BYTE tables):
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
// These route through the binary-operation comparison branch
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
//   emit 0x3b8 then a zero dword → [0xf5, 0x00, 0x00, 0x00, 0x00]
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
// case 0x36: emit LHS(call_ctx=1), RHS(call_ctx=1), then
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
// case 0x24: explicit Is path — NOT routed through the general
// binary-operation path.  Both operands use call_ctx=1; the comparison
// opcode depends on the node's type_tag:
//
//   type_tag=0x10 → emit_value2(0xf0):   RT_OPCODE_BYTE[0xf0]=0x2a  → [0x2a]
//   type_tag=2    → emit_value2(0x18b):  RT_OPCODE_BYTE[0x18b]=0xfc → [0xfc, 0x8b]
//   type_tag=10   → emit_value2(0x189) only when outer call_ctx is 1 or 3:
//                   RT_OPCODE_BYTE[0x189]=0x37 → [0x37]
//   type_tag=0xb/0xc → emit_value2(0x18a): RT_OPCODE_BYTE[0x18a]=0x39 → [0x39]
//   other type_tag → no opcode (per-type validation emits nothing)
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
    // (per-type validation requires type-flags 1 or 3)
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
    // type_tag=7 (not in the switch) → per-type validation emits nothing,
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

// ── emit_reference 0x4000 branch ─────────────────────────────────────────

fn ref_bytes(desc: &crate::emit::RefDescriptor, n_op: i32, f_flags: u32, n_type: i32) -> Vec<u8> {
    let arena = NodeArena::new();
    let mut e = Emitter::new(&arena);
    e.emit_reference(desc, n_op, f_flags, n_type);
    e.into_bytes()
}

fn local_desc(offset: i16) -> crate::emit::RefDescriptor {
    crate::emit::RefDescriptor { kind: 1, operand: offset as u16, word6: 0, word8: 0, flags1: 0 }
}

#[test]
fn emit_reference_4000_object_ntype_emits_0x3e() {
    // 0x4000 path: f_flags & 0x4000, nType == 0x10 (Object/Dispatch).
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
    // 0x4000 path: f_flags & 0x4000, nType == 5 (Single).
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
    // Actually nType in emit_reference is the *internal* VB6 type, not vb-type.
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
    let desc = crate::emit::RefDescriptor { kind: 1, operand: 0xff7au16, word6: 0, word8: 0, flags1: 0 };
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
    let desc_byref = crate::emit::RefDescriptor { kind: 2, operand: 0x000c, word6: 1, word8: 0, flags1: 0 };
    let desc_byval = crate::emit::RefDescriptor { kind: 2, operand: 0x000c, word6: 0, word8: 0, flags1: 0 };
    let b_byref = ref_bytes(&desc_byref, 1, 0, 2);
    let b_byval = ref_bytes(&desc_byval, 1, 0, 2);
    // The byref descriptor produces nOp==2 output, byval produces nOp==1.
    let op_byref = RT_OPCODE_BYTE[nop2_idx as usize];
    let op_byval = RT_OPCODE_BYTE[nop1_idx as usize];
    assert_eq!(b_byval[0], op_byval);
    assert_eq!(b_byref[0], op_byref);
}

// ── value-emitter operator-reference kinds (8/9/0xb) ─────────────────────────
// (uses the `opc2` opcode+operand helper defined later in this module)

#[test]
fn emit_reference_kind8_nop5_emits_opcode_then_two_words() {
    // kind 8, nOp 5: opcode 0x3ca (nOp==5 ⇒ +0), operand = word6, then word8
    // and the +10 operand word; the finalize tail returns cleanly at nOp 5.
    let desc = crate::emit::RefDescriptor {
        kind: 8, operand: 0x3333, word6: 0x1111, word8: 0x2222, flags1: 0,
    };
    let bytes = ref_bytes(&desc, 5, 0, 0);
    let mut want = opc2(0x3ca, 0x1111);
    want.extend_from_slice(&0x2222u16.to_le_bytes());
    want.extend_from_slice(&0x3333u16.to_le_bytes());
    assert_eq!(bytes, want);
}

#[test]
fn emit_reference_kind8_nop1_uses_0x3cb() {
    // nOp != 5 ⇒ opcode 0x3cb. nOp 1 also drives the finalize tail into the
    // gated re-entry, so only the leading opcode is asserted here via nOp 5
    // vs nOp 6-without-flag is unavailable; assert the opcode index instead.
    let desc = crate::emit::RefDescriptor {
        kind: 8, operand: 0, word6: 0x44, word8: 0, flags1: 0,
    };
    // nOp 6 with flags1 bit 0x04 clear → tail returns cleanly.
    let bytes = ref_bytes(&desc, 6, 0, 0);
    let mut want = opc2(0x3cb, 0x44);
    want.extend_from_slice(&0u16.to_le_bytes());
    want.extend_from_slice(&0u16.to_le_bytes());
    assert_eq!(bytes, want);
}

#[test]
fn emit_reference_kind9_nop5_emits_0x18e() {
    // kind 9, nOp 5: opcode (nOp==5)+0x18d = 0x18e, operand = word6.
    let desc = crate::emit::RefDescriptor {
        kind: 9, operand: 0, word6: 0x55, word8: 0, flags1: 0,
    };
    let bytes = ref_bytes(&desc, 5, 0, 0);
    assert_eq!(bytes, opc2(0x18e, 0x55));
}

#[test]
fn emit_reference_kind9_variant_normalizes_nop1_to_5() {
    // n_type 0x12 with nOp 1 normalizes nOp→5, so opcode (nOp==5)+0x18d=0x18e.
    let desc = crate::emit::RefDescriptor {
        kind: 9, operand: 0, word6: 0x66, word8: 0, flags1: 0,
    };
    let bytes = ref_bytes(&desc, 1, 0, 0x12);
    assert_eq!(bytes, opc2(0x18e, 0x66));
}

#[test]
fn emit_reference_kindb_nop5_emits_0x407_then_word6() {
    // kind 0xb, nOp 5: opcode (nOp==5)+0x406 = 0x407, operand = word8, then word6.
    let desc = crate::emit::RefDescriptor {
        kind: 0xb, operand: 0, word6: 0x77, word8: 0x88, flags1: 0,
    };
    let bytes = ref_bytes(&desc, 5, 0, 0);
    let mut want = opc2(0x407, 0x88);
    want.extend_from_slice(&0x77u16.to_le_bytes());
    assert_eq!(bytes, want);
}

// ── value-emitter typed-store conversion (0x8000 path) ───────────────────────

#[test]
fn emit_reference_store_conversion_uses_conv_table() {
    use crate::tables::{EXPR_STORE_CONV, RT_TYPE_OFFSET};
    // kind 1 (base 0x1e0), nOp 4 (store), flag 0x8000 set, 0x20 / 0x40 clear.
    // n_type 8 → RT_TYPE_OFFSET[8] = 2 (a valid conversion class).
    let n_type = 8;
    assert_eq!(RT_TYPE_OFFSET[n_type as usize], 2);
    let desc = crate::emit::RefDescriptor {
        kind: 1, operand: 0x0010, word6: 0, word8: 0, flags1: 0,
    };
    // f_flags = 0x8000: inv12 = 1, inv11 = 1 → sub = 3.
    let conv = EXPR_STORE_CONV[0][3] as i32; // 0x01
    let expected_idx = (conv + 0x10 + 0x1e0) as usize;
    let bytes = ref_bytes(&desc, 4, 0x8000, n_type);
    assert_eq!(bytes, opc2(expected_idx, 0x0010));

    // f_flags = 0x9000 (adds 0x1000): inv12 = 0, inv11 = 1 → sub = 1.
    let conv1 = EXPR_STORE_CONV[0][1] as i32;
    let idx1 = (conv1 + 0x10 + 0x1e0) as usize;
    let bytes1 = ref_bytes(&desc, 4, 0x9000, n_type);
    assert_eq!(bytes1, opc2(idx1, 0x0010));
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
// A "null-emit RHS" node has opcode=0 (hits the emitter dispatch guard → return 0
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
    // The assignment store path special-cases Currency LHS: rhs_kind 0xb → emit 0x147.
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
    // Flag byte bit 0x80 clear (word[1]=0) → store path returns immediately.
    // Trailing per-type validation(0xb, 0, context=0): context=0 ≠ 3 and ≠ 1
    // → returns 1, no bytes emitted.
    let (a, n) = assign_node(0xb, 0xb);
    assert_eq!(emit(&a, n), &[]);
}

#[test]
fn assign_default_numeric_kind5_emits_fc0c() {
    // dest kind=5, src kind=5: generic store path.
    // base = RT_ASSIGN_STORE_OPCODE[RT_TYPE_OFFSET[5]=0] = 0x10c;
    // adjust(src=5) = RT_TYPE_OFFSET[5] = 0 → emit 0x10c.
    // RT_OPCODE_BYTE[0x10c] = 0xfc → [0xfc, 0x0c].
    let (a, n) = assign_node(5, 5);
    assert_eq!(emit(&a, n), &[0xfc, 0x0c]);
}

#[test]
fn assign_default_numeric_kind6_same_emits_fc15() {
    // dest kind=6, src kind=6: base = RT_ASSIGN_STORE_OPCODE[RT_TYPE_OFFSET[6]=1] = 0x114;
    // adjust(src=6) = 1 → emit 0x115. RT_OPCODE_BYTE[0x115] = 0xfc → [0xfc, 0x15].
    // (kind 6 is not in the Variant group {10,0xb,0xc} so the generic path runs.)
    let (a, n) = assign_node(6, 6);
    assert_eq!(emit(&a, n), &[0xfc, 0x15]);
}

#[test]
fn assign_currency_src_into_long_emits_0x3c8() {
    // src kind 0xc (Currency, src_hi 0xc0000), dest kind 0x10 (not in {10,b,c},
    // not object/currency dest) → direct opcode 0x3c8.
    let (a, n) = assign_node(0x10, 0xc);
    assert_eq!(emit(&a, n), v2(0x3c8).as_slice());
}

#[test]
fn assign_single_src_into_kind5_emits_0x138() {
    // src kind 3 (src_hi 0x30000), dest kind 5 → direct opcode 0x138.
    let (a, n) = assign_node(5, 3);
    assert_eq!(emit(&a, n), v2(0x138).as_slice());
}

#[test]
fn assign_single_src_into_kind10_emits_0x3c7() {
    // src kind 3 (src_hi 0x30000), dest kind 0x10 → direct opcode 0x3c7.
    let (a, n) = assign_node(0x10, 3);
    assert_eq!(emit(&a, n), v2(0x3c7).as_slice());
}

#[test]
fn assign_single_src_into_kind6_emits_nothing() {
    // src kind 3 (src_hi 0x30000), dest kind 6 → no store opcode.
    let (a, n) = assign_node(6, 3);
    assert_eq!(emit(&a, n), &[]);
}


// ── traverse_node_tree ────────────────────────────────────────────────────────
//
// traverse_node_tree walks a singly-linked list (opcodes 0x37/0x33) and emits
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
    // traverse_node_tree recurses on sibling first → emits B then A.
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

// ── Newly-covered node-emitter cases ───────────────────────────────────────
//
// Each of these cases is fully self-contained — it only emits children and a
// fixed opcode selected by the case. The expected bytes are [child bytes…]
// followed by the case's tail opcode encoded exactly as emit_value2 would
// (`v2` below mirrors that 1-or-2-byte encoding from the extracted table).
// `emit_value2`'s own byte encoding is verified exactly elsewhere; these tests
// pin which opcode index each case selects and the child-emission order.

/// Predict the bytes emit_value2(n) emits: 1 byte if `< 0xfb`, else the
/// escape byte followed by `n as u8`.
fn v2(n: usize) -> Vec<u8> {
    use crate::tables::RT_OPCODE_BYTE;
    let b = RT_OPCODE_BYTE[n];
    if b < 0xfb {
        vec![b]
    } else {
        vec![b, n as u8]
    }
}

/// Bytes a `global_long_load` child emits at field offset 0 / 4.
const GL0: [u8; 5] = [0x94, 0x08, 0x00, 0x00, 0x00];
const GL4: [u8; 5] = [0x94, 0x08, 0x00, 0x04, 0x00];

/// A node whose `word[5]` points at `child` (the list/traversal operand used by
/// the `0x3a/0x3b/0x4d/0x4f/0x50/0x53/0x56..0x59` cases).
fn w5_node(a: &mut NodeArena, op: u16, child: NodeRef) -> NodeRef {
    a.alloc(NodeArena::node(op, 0, 0, child.0, 0, 0))
}

/// Build expected bytes: `GL0` followed by the tail opcode `n`.
fn gl0_then(n: usize) -> Vec<u8> {
    let mut v = GL0.to_vec();
    v.extend(v2(n));
    v
}

#[test]
fn case_0x33_traverses_node_list() {
    // word[4] = child, word[5] = 0 (no sibling). traverse_node_tree emits child.
    let mut a = NodeArena::new();
    let child = global_long_load(&mut a, 0);
    let n = a.alloc(NodeArena::node(0x33, 0, child.0, 0, 0, 0));
    assert_eq!(emit(&a, n), GL0.as_slice());
}

#[test]
fn case_0x37_process_linked_list_forward_order() {
    // Two-element forward list [A=field0, B=field4]:
    //   list_a → word[4]=child_a, word[5]=list_b
    //   list_b → word[4]=child_b, word[5]=0
    // process_linked_list emits in FORWARD order: A then B (unlike traverse).
    let mut a = NodeArena::new();
    let child_a = global_long_load(&mut a, 0);
    let child_b = global_long_load(&mut a, 4);
    let list_b = a.alloc(NodeArena::node(0x37, 0, child_b.0, 0, 0, 0));
    let list_a = a.alloc(NodeArena::node(0x37, 0, child_a.0, list_b.0, 0, 0));
    let mut expected = GL0.to_vec();
    expected.extend_from_slice(&GL4);
    assert_eq!(emit(&a, list_a), expected.as_slice());
}

#[test]
fn case_0x3a_traverse_then_0x1ca() {
    let mut a = NodeArena::new();
    let child = global_long_load(&mut a, 0);
    let n = w5_node(&mut a, 0x3a, child);
    assert_eq!(emit(&a, n), gl0_then(0x1ca).as_slice());
}

#[test]
fn case_0x3b_traverse_then_0x1cb() {
    let mut a = NodeArena::new();
    let child = global_long_load(&mut a, 0);
    let n = w5_node(&mut a, 0x3b, child);
    assert_eq!(emit(&a, n), gl0_then(0x1cb).as_slice());
}

#[test]
fn case_0x4d_emit_child_then_0x23d() {
    let mut a = NodeArena::new();
    let child = global_long_load(&mut a, 0);
    let n = w5_node(&mut a, 0x4d, child);
    assert_eq!(emit(&a, n), gl0_then(0x23d).as_slice());
}

#[test]
fn case_0x4e_emits_only_0x23e() {
    // No child emit — just the tail opcode 0x23e.
    let mut a = NodeArena::new();
    let n = a.alloc(NodeArena::node(0x4e, 0, 0, 0, 0, 0));
    assert_eq!(emit(&a, n), v2(0x23e).as_slice());
}

#[test]
fn case_0x4f_traverse_then_0xfb() {
    let mut a = NodeArena::new();
    let child = global_long_load(&mut a, 0);
    let n = w5_node(&mut a, 0x4f, child);
    assert_eq!(emit(&a, n), gl0_then(0xfb).as_slice());
}

#[test]
fn case_0x50_traverse_then_0xfa() {
    let mut a = NodeArena::new();
    let child = global_long_load(&mut a, 0);
    let n = w5_node(&mut a, 0x50, child);
    assert_eq!(emit(&a, n), gl0_then(0xfa).as_slice());
}

#[test]
fn case_0x53_byte5_selects_0x1c0_or_0x1bf() {
    // byte5 bit 0x40 clear → 0x1c0; set → 0x1bf. byte5 = (word[1] >> 8) & 0xff.
    let mut a = NodeArena::new();
    let child = global_long_load(&mut a, 0);
    let clear = w5_node(&mut a, 0x53, child);
    assert_eq!(emit(&a, clear), gl0_then(0x1c0).as_slice());

    let mut b = NodeArena::new();
    let child_b = global_long_load(&mut b, 0);
    let mut node = NodeArena::node(0x53, 0, 0, child_b.0, 0, 0);
    node.w[1] = 0x4000; // byte5 = 0x40 set
    let set = b.alloc(node);
    assert_eq!(emit(&b, set), gl0_then(0x1bf).as_slice());
}

#[test]
fn case_0x59_traverse_then_0x1c1() {
    let mut a = NodeArena::new();
    let child = global_long_load(&mut a, 0);
    let n = w5_node(&mut a, 0x59, child);
    assert_eq!(emit(&a, n), gl0_then(0x1c1).as_slice());
}

#[test]
fn case_0x5c_emits_only_0x162() {
    let mut a = NodeArena::new();
    let n = a.alloc(NodeArena::node(0x5c, 0, 0, 0, 0, 0));
    assert_eq!(emit(&a, n), v2(0x162).as_slice());
}

#[test]
fn case_0x5e_emit_child_then_0x40a() {
    let mut a = NodeArena::new();
    let child = global_long_load(&mut a, 0);
    let n = w5_node(&mut a, 0x5e, child);
    assert_eq!(emit(&a, n), gl0_then(0x40a).as_slice());
}

#[test]
fn case_0x5f_emit_child_then_0x40b() {
    let mut a = NodeArena::new();
    let child = global_long_load(&mut a, 0);
    let n = w5_node(&mut a, 0x5f, child);
    assert_eq!(emit(&a, n), gl0_then(0x40b).as_slice());
}

#[test]
fn case_0x73_emits_0x266_with_child_word4_low() {
    // case 0x73: emit opcode 0x266 with low16 of child's word[4].
    // The child is never emitted — only its word[4] is read.
    let mut a = NodeArena::new();
    let child = a.alloc(NodeArena::node(0, 0, 0x9999_1234, 0, 0, 0)); // word[4] low = 0x1234
    let n = a.alloc(NodeArena::node(0x73, 0, child.0, 0, 0, 0));
    let mut expected = v2(0x266);
    expected.extend_from_slice(&[0x34, 0x12]); // operand 0x1234 LE
    assert_eq!(emit(&a, n), expected.as_slice());
}

// ── case 0xf sub-dispatch (op-class in word[1] bits 8..10) ───────────────────

#[test]
fn case_0xf_class0_flag_clear_emits_child_ctx1() {
    // op-class 0, flag 0x8000 clear → emit child (word[4]) with context 1.
    let mut a = NodeArena::new();
    let child = global_long_load(&mut a, 0);
    let n = a.alloc(NodeArena::node(0xf, 0, child.0, 0, 0, 0)); // word[1]=0
    assert_eq!(emit(&a, n), GL0.as_slice());
}

#[test]
fn case_0xf_class2_emits_child_ctx3() {
    // op-class 2 → emit child (word[4]) with context 3, return.
    let mut a = NodeArena::new();
    let child = global_long_load(&mut a, 0);
    let mut node = NodeArena::node(0xf, 0, child.0, 0, 0, 0);
    node.w[1] = 2 << 8; // op-class 2
    let n = a.alloc(node);
    assert_eq!(emit(&a, n), GL0.as_slice());
}

#[test]
fn case_0xf_class3_sized_name_ref_emits_child_then_0x2c6() {
    // op-class 3 → emit child (context 3), then sized opcode 0x2c6 + type size.
    let mut a = NodeArena::new();
    let child = global_long_load(&mut a, 0);
    let type_desc = a.alloc(NodeArena::node(4, 0, 0x18, 0, 0, 0)); // kind 4, size 0x18
    let mut node = NodeArena::node(0xf, 0, child.0, 0, type_desc.0, 0); // word[6]=desc
    node.w[1] = 3 << 8; // op-class 3
    let n = a.alloc(node);
    let mut expected = GL0.to_vec();
    expected.extend(v2(0x2c6));
    expected.extend_from_slice(&[0x18, 0x00]);
    assert_eq!(emit(&a, n), expected.as_slice());
}

#[test]
fn case_0xf_class0_flag_set_plain_child_emits_ctx3() {
    // op-class 0, flag 0x8000 set, child opcode not in {0x60,0x69,0x5e} →
    // emit child with context 3, return (no tail opcode).
    let mut a = NodeArena::new();
    let child = global_long_load(&mut a, 0); // opcode 0x77
    let mut node = NodeArena::node(0xf, 0, child.0, 0, 0, 0);
    node.w[1] = 0x8000;
    let n = a.alloc(node);
    assert_eq!(emit(&a, n), GL0.as_slice());
}

#[test]
fn case_0xf_class0_flag_set_special_child_emits_ctx1_then_0x18c() {
    // op-class 0, flag 0x8000 set, child opcode == 0x5e → take the
    // emit-child(ctx1) + tail-0x18c branch. The 0x5e child itself emits its
    // own word[5] grandchild then opcode 0x40a (case 0x5e).
    let mut a = NodeArena::new();
    let grandchild = global_long_load(&mut a, 0);
    let child = w5_node(&mut a, 0x5e, grandchild); // opcode 0x5e
    let mut node = NodeArena::node(0xf, 0, child.0, 0, 0, 0);
    node.w[1] = 0x8000;
    let n = a.alloc(node);
    // child(0x5e) → GL0 ++ v2(0x40a); then outer tail v2(0x18c).
    let mut expected = GL0.to_vec();
    expected.extend(v2(0x40a));
    expected.extend(v2(0x18c));
    assert_eq!(emit(&a, n), expected.as_slice());
}

// ── case 0x12 (member deref) portable paths ──────────────────────────────────

#[test]
fn case_0x12_object_typeexpr_emits_child_ctx2() {
    // node hi == 0xf, child hi == 0x17 → emit child with context 2, return.
    let mut a = NodeArena::new();
    // child: global-load node carrying type tag 0x17 in word[0] high half.
    let child = a.alloc(NodeArena::node(0x77, 0x17, 0x0008, 2, 0, 0));
    let n = a.alloc(NodeArena::node(0x12, 0xf, child.0, 0, 0, 0));
    assert_eq!(emit(&a, n), GL0.as_slice());
}

#[test]
fn case_0x12_non_0x3f_child_emits_nothing() {
    // node hi != 0xf/0x17, child low16 != 0x3f → return with no bytes.
    let mut a = NodeArena::new();
    let child = global_long_load(&mut a, 0); // opcode 0x77 != 0x3f
    let n = a.alloc(NodeArena::node(0x12, 2, child.0, 0, 0, 0)); // node hi = 2
    assert_eq!(emit(&a, n), &[]);
}

// ── case 0x14 (object cast) portable path ────────────────────────────────────

#[test]
fn case_0x14_object_emits_child_ctx3_then_0x3f9() {
    // node hi == 0xf → emit child (word[4]) with context 3, tail 0x3f9.
    let mut a = NodeArena::new();
    let child = global_long_load(&mut a, 0);
    let n = a.alloc(NodeArena::node(0x14, 0xf, child.0, 0, 0, 0));
    assert_eq!(emit(&a, n), gl0_then(0x3f9).as_slice());
}

// ── cases 0x48..0x4b (traverse + opcode, 0x20000 result is self-contained) ───

#[test]
fn cases_0x48_to_0x4b_traverse_then_fixed_opcode() {
    // For a node whose result type is 0x20000, the case is fully self-contained:
    // traverse(word[5]) then emit the opcode, return. Opcode map:
    //   0x48→0x158, 0x49→0x15a, 0x4a→0x159, 0x4b→0x15b.
    for (op, value) in [(0x48u16, 0x158usize), (0x49, 0x15a), (0x4a, 0x159), (0x4b, 0x15b)] {
        let mut a = NodeArena::new();
        let child = global_long_load(&mut a, 0);
        // type_tag 2 → word[0] high = 2 → node hi = 0x20000.
        let n = a.alloc(NodeArena::node(op, 2, 0, child.0, 0, 0));
        assert_eq!(emit(&a, n), gl0_then(value).as_slice(), "op {op:#x}");
    }
}

// ── case 0x56 (0x4000 set: traverse + flag-selected opcode) ──────────────────

#[test]
fn case_0x56_set_selects_0x1bb_or_0x1bd() {
    // flag 0x4000 set: value = (flag 0x8000 ? 2 : 0) + 0x1bb, then traverse.
    let mut a = NodeArena::new();
    let child = global_long_load(&mut a, 0);
    let mut node = NodeArena::node(0x56, 0, 0, child.0, 0, 0);
    node.w[1] = 0x4000;
    let n = a.alloc(node);
    assert_eq!(emit(&a, n), gl0_then(0x1bb).as_slice());

    let mut b = NodeArena::new();
    let child_b = global_long_load(&mut b, 0);
    let mut node_b = NodeArena::node(0x56, 0, 0, child_b.0, 0, 0);
    node_b.w[1] = 0xc000; // 0x4000 | 0x8000
    let set = b.alloc(node_b);
    assert_eq!(emit(&b, set), gl0_then(0x1bd).as_slice());
}

// ── case 0x57 (non-0x4000: traverse + opcode + validate) ─────────────────────

#[test]
fn case_0x57_clear_selects_0x3fd_or_0x3fb() {
    // flag 0x4000 clear: value = (flag 0x8000 ? 0x3fb : 0x3fd). type_tag 0 makes
    // the trailing per-type validation a no-op.
    let mut a = NodeArena::new();
    let child = global_long_load(&mut a, 0);
    let n = w5_node(&mut a, 0x57, child); // word[1]=0
    assert_eq!(emit(&a, n), gl0_then(0x3fd).as_slice());

    let mut b = NodeArena::new();
    let child_b = global_long_load(&mut b, 0);
    let mut node_b = NodeArena::node(0x57, 0, 0, child_b.0, 0, 0);
    node_b.w[1] = 0x8000;
    let set = b.alloc(node_b);
    assert_eq!(emit(&b, set), gl0_then(0x3fb).as_slice());
}

// ── Unary negate (case 0xb) ──────────────────────────────────────────────────
//
// case 0xb emits the operand (context 2) then a type-selected opcode:
//   off = RT_TYPE_OFFSET[operand_tag];
//   if off==10 -> 0xf6 ; else { if off==9 -> 1 ; opcode = off + 0xf2 }
// The opcode is selected from the OPERAND's type tag, not the node's.

/// Build a negate node (op 0xb) over a typed operand load.
fn negate(a: &mut NodeArena, operand: NodeRef) -> NodeRef {
    a.alloc(NodeArena::node(0xb, 0, operand.0, 0, 0, 0))
}

#[test]
fn negate_long_emits_operand_then_0xf4() {
    // Long operand (tag 8): RT_TYPE_OFFSET[8]=2 -> 2+0xf2 = 0xf4.
    let mut a = NodeArena::new();
    let v = var_load_typed(&mut a, 8, 2, -8); // [0x6c,0xf8,0xff]
    let n = negate(&mut a, v);
    let mut expected = vec![0x6c, 0xf8, 0xff];
    expected.extend(v2(0xf4));
    assert_eq!(emit(&a, n), expected.as_slice());
}

#[test]
fn negate_integer_emits_operand_then_0xf3() {
    // Integer operand (tag 6): RT_TYPE_OFFSET[6]=1 -> 1+0xf2 = 0xf3.
    let mut a = NodeArena::new();
    let v = var_load_typed(&mut a, 6, 1, -4); // [0x6b,0xfc,0xff]
    let n = negate(&mut a, v);
    let mut expected = vec![0x6b, 0xfc, 0xff];
    expected.extend(v2(0xf3));
    assert_eq!(emit(&a, n), expected.as_slice());
}

#[test]
fn negate_single_emits_operand_then_0xf5() {
    // Single operand (tag 10): RT_TYPE_OFFSET[10]=3 -> 3+0xf2 = 0xf5.
    let mut a = NodeArena::new();
    let v = var_load_typed(&mut a, 10, 3, -4); // [0x6e,0xfc,0xff]
    let n = negate(&mut a, v);
    let mut expected = vec![0x6e, 0xfc, 0xff];
    expected.extend(v2(0xf5));
    assert_eq!(emit(&a, n), expected.as_slice());
}

#[test]
fn negate_double_emits_operand_then_0xf6() {
    // Double operand (tag 11): RT_TYPE_OFFSET[11]=4 -> 4+0xf2 = 0xf6.
    let mut a = NodeArena::new();
    let v = var_load_typed(&mut a, 11, 4, 0xff74u16 as i16); // [0x6f,0x74,0xff]
    let n = negate(&mut a, v);
    let mut expected = vec![0x6f, 0x74, 0xff];
    expected.extend(v2(0xf6));
    assert_eq!(emit(&a, n), expected.as_slice());
}

#[test]
fn negate_currency_emits_operand_then_0xf6() {
    // Currency operand (tag 12): RT_TYPE_OFFSET[12]=10 -> special 0xf6.
    let mut a = NodeArena::new();
    let v = var_load_typed(&mut a, 12, 6, -8); // [0x6d,0xf8,0xff]
    let n = negate(&mut a, v);
    let mut expected = vec![0x6d, 0xf8, 0xff];
    expected.extend(v2(0xf6));
    assert_eq!(emit(&a, n), expected.as_slice());
}

// ── Unary Not (case 0x10) ────────────────────────────────────────────────────
//
// case 0x10 emits the operand (context 1) then opcode 0x135.

#[test]
fn not_long_emits_operand_then_0x135() {
    let mut a = NodeArena::new();
    let v = var_load_typed(&mut a, 8, 2, -8); // Long [0x6c,0xf8,0xff]
    let n = a.alloc(NodeArena::node(0x10, 0, v.0, 0, 0, 0));
    let mut expected = vec![0x6c, 0xf8, 0xff];
    expected.extend(v2(0x135));
    assert_eq!(emit(&a, n), expected.as_slice());
}

// ── Power operator (case 0x1a) ───────────────────────────────────────────────
//
// case 0x1a emits both operands (context 2) then opcode 0xcf for the numeric
// (non-object) form, with no extra validation for a Double result.

#[test]
fn pow_double_emits_operands_then_0xcf() {
    let mut a = NodeArena::new();
    let lhs = var_load_typed(&mut a, 11, 4, 0xff74u16 as i16); // Double [0x6f,0x74,0xff]
    let rhs = var_load_typed(&mut a, 11, 4, 0xff6cu16 as i16); // Double [0x6f,0x6c,0xff]
    let n = a.alloc(NodeArena::node(0x1a, 11, lhs.0, rhs.0, 0, 0));
    let mut expected = vec![0x6f, 0x74, 0xff, 0x6f, 0x6c, 0xff];
    expected.extend(v2(0xcf));
    assert_eq!(emit(&a, n), expected.as_slice());
}

// ── Call-opcode computation kernel (RT_CALL_TYPECODE + inline map) ───────────
//
// The call emitter computes its base opcode as map(type_code(kind,ref,mask)).
// These pin the kernel against the extracted table; the full call emitter (case
// 0x61) builds on this once the symbol/type-pool model supplies the inputs.

#[test]
fn call_type_code_value_and_reference_paths() {
    use crate::emit::call_type_code;
    // value path: index (callee_type != 1) + mask*2.
    assert_eq!(call_type_code(1, false, false), 0x340); // idx 0
    assert_eq!(call_type_code(2, false, false), 0x310); // idx 1
    assert_eq!(call_type_code(1, false, true), 0x350); // idx 2
    assert_eq!(call_type_code(2, false, true), 0x300); // idx 3
    // reference path: index (callee_type != 1) + 4.
    assert_eq!(call_type_code(1, true, false), 0x330); // idx 4
    assert_eq!(call_type_code(2, true, true), 0x320); // idx 5
}

#[test]
fn map_call_type_code_covers_every_entry() {
    use crate::emit::map_call_type_code;
    assert_eq!(map_call_type_code(0x300), 0x16a);
    assert_eq!(map_call_type_code(0x310), 0x169);
    assert_eq!(map_call_type_code(0x320), 0x34f);
    assert_eq!(map_call_type_code(0x340), 0x16b);
    assert_eq!(map_call_type_code(0x350), 0x16c);
    // 0x330 (and anything unmapped) → the default/invalid slot.
    assert_eq!(map_call_type_code(0x330), 0x446);
    assert_eq!(map_call_type_code(0x000), 0x446);
}

// ── Call site: common (emit_mode 0) by-reference path + finalize ─────────────
//
// A by-reference method/Sub call (convention kind 4, ByRef) whose callee node
// carries the 0x20000 type region: emit the callee with context 6, the call
// opcode (base 0x320 from RT_CALL_TYPECODE[5], + 0xc for the 0x20000 region =
// 0x32c), then the finalize step's trailing member-id word. The selected
// dispatch record's 0x08 bit drives a per-type validation that emits nothing
// for a type-tag-2 operation, so the call site ends at the member word.

#[test]
fn emit_call_byref_common_path_emits_callee_opcode_then_member_word() {
    let mut a = NodeArena::new();
    let _null = a.alloc(NodeArena::node(0, 0, 0, 0, 0, 0)); // null sentinel at idx 0
    let callee = global_long_load(&mut a, 0); // emits GL0 (context-independent)
    let desc = CallDescriptor {
        kind: 4,
        byref: 1,
        flags: 0,
        node_word0: 0x0002_0000, // region 0x20000, type tag 2
        callee,
        arg_list: _null, // index 0 → no argument list
        member_id: 0x0007,
        size: 0,
    };
    let mut e = Emitter::new(&a);
    e.emit_call(&desc, 0);

    let mut expected = GL0.to_vec();
    expected.extend(v2(0x32c)); // call opcode 0x320 + 0xc
    expected.extend_from_slice(&[0x07, 0x00]); // finalize trailing member-id word
    assert_eq!(e.into_bytes(), expected);
}

// The finalize step's type-node path (call node type region 0x140000) builds a
// synthetic dispatch-type node and re-enters the emitter — it needs the type-
// pool allocator, so it must remain explicitly gated, never silently skipped.
#[test]
#[should_panic(expected = "region 0x140000")]
fn emit_call_finalize_type_node_region_is_gated() {
    let mut a = NodeArena::new();
    let _null = a.alloc(NodeArena::node(0, 0, 0, 0, 0, 0));
    let callee = global_long_load(&mut a, 0);
    let desc = CallDescriptor {
        kind: 4,
        byref: 1,
        flags: 0,
        node_word0: 0x0014_0000, // region 0x140000 → gated type-node path
        callee,
        arg_list: _null,
        member_id: 1,
        size: 0,
    };
    Emitter::new(&a).emit_call(&desc, 0);
}

// A 0x61 call node routed through emit_expr assembles the CallDescriptor from
// its words (word[2]=kind, word[3]=byref, word[5]=arg list, word[6]=callee,
// word[7]=member id) and must produce the same bytes as the equivalent direct
// emit_call (kind 4, ByRef, region 0x20000, member 7).
#[test]
fn case_0x61_call_node_assembles_descriptor_and_dispatches() {
    let mut a = NodeArena::new();
    let _null = a.alloc(NodeArena::node(0, 0, 0, 0, 0, 0));
    let callee = global_long_load(&mut a, 0); // emits GL0
    // node(opcode, type_tag, w4, w5=arglist, w6=callee, w7=member)
    let mut raw = NodeArena::node(0x61, 2, 0, 0, callee.0, 7); // type tag 2 → region 0x20000
    raw.w[2] = 4; // convention kind
    raw.w[3] = 1; // by-reference
    let n = a.alloc(raw);

    let mut expected = GL0.to_vec();
    expected.extend(v2(0x32c)); // call opcode 0x320 + 0xc
    expected.extend_from_slice(&[0x07, 0x00]); // finalize trailing member-id word
    assert_eq!(emit(&a, n), expected.as_slice());
}

// ── Type coercion / conversion cases ─────────────────────────────────────────

#[test]
fn case_0x32_coercion_emits_opcode_pool_index_and_size() {
    // No child list (size 0 + 4 = 4); flag byte clear → target 0x40d; descriptor
    // word[4] interns to pool index 0; emit 0x40d + 0x0000 + size 0x0004.
    let mut a = NodeArena::new();
    let _null = a.alloc(NodeArena::node(0, 0, 0, 0, 0, 0));
    let desc = a.alloc(NodeArena::node(0, 0, 0x55, 0, 0, 0)); // descriptor word[4]=0x55
    let n = a.alloc(NodeArena::node(0x32, 0, desc.0, 0, 0, 0)); // word[4]=desc, word[5]=0
    let mut expected = v2(0x40d);
    expected.extend_from_slice(&[0x00, 0x00, 0x04, 0x00]);
    assert_eq!(emit(&a, n), expected.as_slice());
}

#[test]
fn case_0x34_coercion_target_0x40f() {
    let mut a = NodeArena::new();
    let _null = a.alloc(NodeArena::node(0, 0, 0, 0, 0, 0));
    let desc = a.alloc(NodeArena::node(0, 0, 0x55, 0, 0, 0));
    let n = a.alloc(NodeArena::node(0x34, 0, desc.0, 0, 0, 0));
    let mut expected = v2(0x40f);
    expected.extend_from_slice(&[0x00, 0x00, 0x04, 0x00]);
    assert_eq!(emit(&a, n), expected.as_slice());
}

#[test]
fn case_0x4c_conversion_literal_type_node() {
    // inner.word[5]=0 (no list); inner.word[4]=p; p opcode 1 → emit 0x35d + p.word[4].
    let mut a = NodeArena::new();
    let _null = a.alloc(NodeArena::node(0, 0, 0, 0, 0, 0));
    let p = a.alloc(NodeArena::node(1, 0, 0x42, 0, 0, 0)); // opcode 1, word[4]=0x42
    let inner = a.alloc(NodeArena::node(0, 0, p.0, 0, 0, 0)); // word[4]=p, word[5]=0
    let n = a.alloc(NodeArena::node(0x4c, 0, 0, inner.0, 0, 0)); // word[5]=inner
    let mut expected = v2(0x35d);
    expected.extend_from_slice(&[0x42, 0x00]);
    assert_eq!(emit(&a, n), expected.as_slice());
}

#[test]
fn case_0x51_operator_class0_traverse_then_0x175() {
    let mut a = NodeArena::new();
    let _null = a.alloc(NodeArena::node(0, 0, 0, 0, 0, 0));
    let child = global_long_load(&mut a, 0);
    let n = a.alloc(NodeArena::node(0x51, 0, 0, child.0, 0, 0)); // op-class 0 (word[1]=0)
    let mut expected = GL0.to_vec();
    expected.extend(v2(0x175));
    assert_eq!(emit(&a, n), expected.as_slice());
}

#[test]
fn case_0x52_operator_class0_traverse_then_0x177() {
    let mut a = NodeArena::new();
    let _null = a.alloc(NodeArena::node(0, 0, 0, 0, 0, 0));
    let child = global_long_load(&mut a, 0);
    let n = a.alloc(NodeArena::node(0x52, 0, 0, child.0, 0, 0));
    let mut expected = GL0.to_vec();
    expected.extend(v2(0x177));
    assert_eq!(emit(&a, n), expected.as_slice());
}

#[test]
fn case_0x51_operator_class1_value_emits_0x176() {
    // op-class 1 (word[1] bits 8..10 = 1): value emit, 0x51 → 0x176.
    let mut a = NodeArena::new();
    let _null = a.alloc(NodeArena::node(0, 0, 0, 0, 0, 0));
    let child = global_long_load(&mut a, 0);
    let mut node = NodeArena::node(0x51, 0, 0, child.0, 0, 0);
    node.w[1] = 1 << 8;
    let n = a.alloc(node);
    let mut expected = GL0.to_vec();
    expected.extend(v2(0x176));
    assert_eq!(emit(&a, n), expected.as_slice());
}

#[test]
fn case_0x52_operator_class4_value_emits_0x414() {
    // op-class 4: value emit, 0x52 → 0x414.
    let mut a = NodeArena::new();
    let _null = a.alloc(NodeArena::node(0, 0, 0, 0, 0, 0));
    let child = global_long_load(&mut a, 0);
    let mut node = NodeArena::node(0x52, 0, 0, child.0, 0, 0);
    node.w[1] = 4 << 8;
    let n = a.alloc(node);
    let mut expected = GL0.to_vec();
    expected.extend(v2(0x414));
    assert_eq!(emit(&a, n), expected.as_slice());
}

#[test]
fn case_0x51_operator_class5_value_emits_0x413() {
    // op-class 5: value emit, 0x51 → 0x413.
    let mut a = NodeArena::new();
    let _null = a.alloc(NodeArena::node(0, 0, 0, 0, 0, 0));
    let child = global_long_load(&mut a, 0);
    let mut node = NodeArena::node(0x51, 0, 0, child.0, 0, 0);
    node.w[1] = 5 << 8;
    let n = a.alloc(node);
    let mut expected = GL0.to_vec();
    expected.extend(v2(0x413));
    assert_eq!(emit(&a, n), expected.as_slice());
}

#[test]
fn case_0x52_operator_class2_typed_emits_0x418_with_pooled_type() {
    // op-class 2: typed emit. word[5]=inner; traverse inner.word[5]; type
    // value pooled from inner.word[4].word[4]. 0x52 → 0x418.
    let mut a = NodeArena::new();
    let _null = a.alloc(NodeArena::node(0, 0, 0, 0, 0, 0));
    let child = global_long_load(&mut a, 0);
    let p = a.alloc(NodeArena::node(0, 0, 0x99, 0, 0, 0)); // word[4]=type value 0x99
    let inner = a.alloc(NodeArena::node(0, 0, p.0, child.0, 0, 0)); // w4=p, w5=child
    let mut node = NodeArena::node(0x52, 0, 0, inner.0, 0, 0); // word[5]=inner
    node.w[1] = 2 << 8;
    let n = a.alloc(node);
    let mut expected = GL0.to_vec();
    expected.extend(v2(0x418));
    expected.extend_from_slice(&[0x00, 0x00]); // pool index 0
    assert_eq!(emit(&a, n), expected.as_slice());
}

#[test]
fn case_0x51_operator_class3_typed_emits_0x417() {
    // op-class 3: typed emit. 0x51 → 0x417.
    let mut a = NodeArena::new();
    let _null = a.alloc(NodeArena::node(0, 0, 0, 0, 0, 0));
    let child = global_long_load(&mut a, 0);
    let p = a.alloc(NodeArena::node(0, 0, 0x99, 0, 0, 0));
    let inner = a.alloc(NodeArena::node(0, 0, p.0, child.0, 0, 0));
    let mut node = NodeArena::node(0x51, 0, 0, inner.0, 0, 0);
    node.w[1] = 3 << 8;
    let n = a.alloc(node);
    let mut expected = GL0.to_vec();
    expected.extend(v2(0x417));
    expected.extend_from_slice(&[0x00, 0x00]);
    assert_eq!(emit(&a, n), expected.as_slice());
}

#[test]
fn case_0x42_nondispatch_emits_0x42a_with_pooled_type() {
    // Non-dispatch path: type tag != 2, second child not 0x60.
    // word[5]=w5; w5.word[4]=a (a.word[4]=type value); w5.word[5]=second;
    // second.word[4]=typed child (a no-op). Emits emit_typed_node(child) [none]
    // then opcode 0x42a + pooled type value (index 0), then validate (type 8 →
    // no extra byte).
    let mut a = NodeArena::new();
    let _null = a.alloc(NodeArena::node(0, 0, 0, 0, 0, 0));
    let typed_child = a.alloc(NodeArena::node(0x1b, 0, 0, 0, 0, 0)); // no-op
    let second = a.alloc(NodeArena::node(0, 0, typed_child.0, 0, 0, 0)); // word[4]=child
    let avalue = a.alloc(NodeArena::node(0, 0, 0x55, 0, 0, 0)); // word[4]=type value 0x55
    let w5 = a.alloc(NodeArena::node(0, 0, avalue.0, second.0, 0, 0)); // w4=a, w5=second
    let n = a.alloc(NodeArena::node(0x42, 8, 0, w5.0, 0, 0)); // type tag 8, word[5]=w5
    let mut expected = v2(0x42a);
    expected.extend_from_slice(&[0x00, 0x00]); // pool index 0
    assert_eq!(emit(&a, n), expected.as_slice());
}

#[test]
fn case_0x43_nondispatch_same_as_0x42_common_path() {
    // 0x43 shares the common non-dispatch path with 0x42 (the form bit only
    // matters on the gated dispatch-binding path).
    let mut a = NodeArena::new();
    let _null = a.alloc(NodeArena::node(0, 0, 0, 0, 0, 0));
    let typed_child = a.alloc(NodeArena::node(0x1b, 0, 0, 0, 0, 0));
    let second = a.alloc(NodeArena::node(0, 0, typed_child.0, 0, 0, 0));
    let avalue = a.alloc(NodeArena::node(0, 0, 0x55, 0, 0, 0));
    let w5 = a.alloc(NodeArena::node(0, 0, avalue.0, second.0, 0, 0));
    let n = a.alloc(NodeArena::node(0x43, 8, 0, w5.0, 0, 0));
    let mut expected = v2(0x42a);
    expected.extend_from_slice(&[0x00, 0x00]);
    assert_eq!(emit(&a, n), expected.as_slice());
}

// ── Operand-dispatch + pooled member-reference cases ─────────────────────────

#[test]
fn case_0x66_emits_pooled_member_opcode_0x2f4() {
    let mut a = NodeArena::new();
    let _null = a.alloc(NodeArena::node(0, 0, 0, 0, 0, 0));
    let n = a.alloc(NodeArena::node(0x66, 0, 0, 0x99, 0, 0)); // word[5]=0x99 type value
    let mut expected = v2(0x2f4);
    expected.extend_from_slice(&[0x00, 0x00]); // pool index 0
    assert_eq!(emit(&a, n), expected.as_slice());
}

#[test]
fn case_0x67_emits_child_then_pooled_0x2f5() {
    let mut a = NodeArena::new();
    let _null = a.alloc(NodeArena::node(0, 0, 0, 0, 0, 0));
    let child = global_long_load(&mut a, 0);
    let n = a.alloc(NodeArena::node(0x67, 0, child.0, 0x99, 0, 0)); // word[4]=child, word[5]=type
    let mut expected = GL0.to_vec();
    expected.extend(v2(0x2f5));
    expected.extend_from_slice(&[0x00, 0x00]);
    assert_eq!(emit(&a, n), expected.as_slice());
}

#[test]
fn case_0x68_member_ref_160000_child_emits_0x2f2_then_pooled_word() {
    // node region 0x160000, child region 0x160000, context != 6 → opcode 0x2f2
    // then the pooled type value from word[5].
    let mut a = NodeArena::new();
    let _null = a.alloc(NodeArena::node(0, 0, 0, 0, 0, 0));
    let child = a.alloc(NodeArena::node(0x1b, 0x16, 0, 0, 0, 0)); // region 0x160000, emits nothing
    let n = a.alloc(NodeArena::node(0x68, 0x16, child.0, 0x99, 0, 0)); // region 0x160000, word[5]=type
    let mut expected = v2(0x2f2);
    expected.extend_from_slice(&[0x00, 0x00]); // pool index 0
    assert_eq!(emit(&a, n), expected.as_slice());
}

#[test]
fn case_0x63_emits_child_then_0x38d_for_non_object() {
    let mut a = NodeArena::new();
    let _null = a.alloc(NodeArena::node(0, 0, 0, 0, 0, 0));
    let child = global_long_load(&mut a, 0); // region 0 (not 0xf0000) → 0x38d
    let n = a.alloc(NodeArena::node(0x63, 0, child.0, 0x99, 0, 0));
    let mut expected = GL0.to_vec();
    expected.extend(v2(0x38d));
    expected.extend_from_slice(&[0x00, 0x00]);
    assert_eq!(emit(&a, n), expected.as_slice());
}

#[test]
fn case_0x3f_emits_member_size_opcode_0x20f() {
    let mut a = NodeArena::new();
    let _null = a.alloc(NodeArena::node(0, 0, 0, 0, 0, 0));
    let type_desc = a.alloc(NodeArena::node(4, 0, 0x10, 0, 0, 0)); // kind 4, size 0x10
    let child = a.alloc(NodeArena::node(0, 0, 0, type_desc.0, 0, 0)); // word[5]=type_desc
    let n = a.alloc(NodeArena::node(0x3f, 0, child.0, 0, 0, 0)); // word[4]=child, word[5]=0
    let mut expected = v2(0x20f);
    expected.extend_from_slice(&[0x10, 0x00]);
    assert_eq!(emit(&a, n), expected.as_slice());
}

#[test]
fn case_0x3e_no_member_emits_0x1c3() {
    let mut a = NodeArena::new();
    let _null = a.alloc(NodeArena::node(0, 0, 0, 0, 0, 0));
    let inner = a.alloc(NodeArena::node(0, 0, 0, 0, 0, 0)); // word[4]=0 (no member), word[5]=0
    let n = a.alloc(NodeArena::node(0x3e, 0, 0, inner.0, 0, 0));
    assert_eq!(emit(&a, n), v2(0x1c3).as_slice());
}

#[test]
fn case_0x44_conversion_0x3f5_no_dispatch_in_0x20000_region() {
    let mut a = NodeArena::new();
    let _null = a.alloc(NodeArena::node(0, 0, 0, 0, 0, 0));
    let p = a.alloc(NodeArena::node(1, 0, 0x07, 0, 0, 0)); // opcode 1, word[4]=7
    let inner = a.alloc(NodeArena::node(0, 0, p.0, 0, 0, 0)); // word[4]=p, word[5]=0
    let n = a.alloc(NodeArena::node(0x44, 2, 0, inner.0, 0, 0)); // type tag 2 → region 0x20000
    let mut expected = v2(0x3f5);
    expected.extend_from_slice(&[0x07, 0x00]);
    assert_eq!(emit(&a, n), expected.as_slice());
}

#[test]
fn case_0x48_traverse_then_0x158_no_dispatch_in_0x20000_region() {
    let mut a = NodeArena::new();
    let _null = a.alloc(NodeArena::node(0, 0, 0, 0, 0, 0));
    let child = global_long_load(&mut a, 0);
    let n = a.alloc(NodeArena::node(0x48, 2, 0, child.0, 0, 0)); // region 0x20000
    let mut expected = GL0.to_vec();
    expected.extend(v2(0x158));
    assert_eq!(emit(&a, n), expected.as_slice());
}

// ── Instruction-emitter cases (EbEmitInstruction2) ───────────────────────────

#[test]
fn case_0x6e_emits_target_opcode_and_pooled_word() {
    // flags 0 → opcode 0x398; target child (region 0) → context 6 (GL0); byte5
    // bit 0x80 clear → pooled type word (index 0). 0x6e always forces the member
    // word (fHasArg = 1), so word[8] (0) is emitted even though it is zero.
    let mut a = NodeArena::new();
    let _null = a.alloc(NodeArena::node(0, 0, 0, 0, 0, 0));
    let child = global_long_load(&mut a, 0);
    let n = a.alloc(NodeArena::node(0x6e, 0, child.0, 0, 0, 0x99)); // w4=child, w7=type value
    let mut expected = GL0.to_vec();
    expected.extend(v2(0x398));
    expected.extend_from_slice(&[0x00, 0x00]); // pooled type word
    expected.extend_from_slice(&[0x00, 0x00]); // member word (forced by fHasArg)
    assert_eq!(emit(&a, n), expected.as_slice());
}

#[test]
fn case_0x6c_walks_arg_then_emits_target_opcode_member_word() {
    // 0x6c walks word[5] (a non-list arg that emits nothing), then the target
    // (GL0), opcode 0x39b (0x443-0xa8 for a non-object target), and the pooled
    // word. With flags 0 the member word is NOT forced (fHasArg = (0x2000 set) =
    // false) and word[8] is zero, so no member word is emitted.
    let mut a = NodeArena::new();
    let _null = a.alloc(NodeArena::node(0, 0, 0, 0, 0, 0));
    let arg = a.alloc(NodeArena::node(0, 0, 0, 0, 0, 0)); // opcode 0 → emits nothing
    let child = global_long_load(&mut a, 0);
    let n = a.alloc(NodeArena::node(0x6c, 0, child.0, arg.0, 0, 0x99)); // w4=child, w5=arg, w7=type
    let mut expected = GL0.to_vec();
    expected.extend(v2(0x39b));
    expected.extend_from_slice(&[0x00, 0x00]); // pooled word only
    assert_eq!(emit(&a, n), expected.as_slice());
}

#[test]
fn case_0x6c_member_word_forced_when_0x2000_set() {
    // With flags bit 0x2000 set, fHasArg is true, so the (zero) member word is
    // emitted. Opcode: 0x8000 clear, 0x2000 set, non-object child → 0x39c.
    let mut a = NodeArena::new();
    let _null = a.alloc(NodeArena::node(0, 0, 0, 0, 0, 0));
    let arg = a.alloc(NodeArena::node(0, 0, 0, 0, 0, 0));
    let child = global_long_load(&mut a, 0);
    let mut node = NodeArena::node(0x6c, 0, child.0, arg.0, 0, 0x99);
    node.w[1] = 0x2000;
    let n = a.alloc(node);
    let mut expected = GL0.to_vec();
    expected.extend(v2(0x39c));
    expected.extend_from_slice(&[0x00, 0x00]); // pooled word
    expected.extend_from_slice(&[0x00, 0x00]); // member word (forced by 0x2000)
    assert_eq!(emit(&a, n), expected.as_slice());
}

#[test]
fn case_0x6a_class0_emits_callee_then_0x3ae() {
    // No argument list; emit the callee (GL0, context 1); op-class 0 → opcode
    // 0x3ae with the word[6] type size (0 here, word[6]=0), then the word[8]
    // trailing word (0). Validation for node type 0 emits nothing.
    let mut a = NodeArena::new();
    let _null = a.alloc(NodeArena::node(0, 0, 0, 0, 0, 0));
    let child = global_long_load(&mut a, 0);
    let n = a.alloc(NodeArena::node(0x6a, 0, child.0, 0, 0, 0)); // w4=child, w5=0, w6=0
    let mut expected = GL0.to_vec();
    expected.extend(v2(0x3ae));
    expected.extend_from_slice(&[0x00, 0x00]); // size 0 (word[6]==0)
    expected.extend_from_slice(&[0x00, 0x00]); // word[8] trailing word
    assert_eq!(emit(&a, n), expected.as_slice());
}

#[test]
fn case_0x6a_class1_emits_0x3b0_with_word8() {
    // op-class 1 → emit callee then sized opcode 0x3b0 with the word[8] operand
    // and finish (no trailing validation).
    let mut a = NodeArena::new();
    let _null = a.alloc(NodeArena::node(0, 0, 0, 0, 0, 0));
    let child = global_long_load(&mut a, 0);
    let mut node = NodeArena::node(0x6a, 0, child.0, 0, 0, 0);
    node.w[1] = 1 << 8; // op-class 1
    node.w[8] = 0x0044; // word[8] operand
    let n = a.alloc(node);
    let mut expected = GL0.to_vec();
    expected.extend(v2(0x3b0));
    expected.extend_from_slice(&[0x44, 0x00]);
    assert_eq!(emit(&a, n), expected.as_slice());
}

#[test]
fn case_0x6a_class0_flag80_emits_0x3f2_and_two_sizes() {
    // op-class 0 with flag byte bit 0x80 set → opcode 0x3f2 with the word[6]
    // size, then the word[7] size, then the word[8] trailing word.
    let mut a = NodeArena::new();
    let _null = a.alloc(NodeArena::node(0, 0, 0, 0, 0, 0));
    let child = global_long_load(&mut a, 0);
    let d6 = a.alloc(NodeArena::node(4, 0, 0x12, 0, 0, 0)); // size 0x12
    let d7 = a.alloc(NodeArena::node(4, 0, 0x34, 0, 0, 0)); // size 0x34
    let mut node = NodeArena::node(0x6a, 0, child.0, 0, d6.0, 0); // word[6]=d6
    node.w[1] = 0x8000; // byte5 bit 0x80
    node.w[7] = d7.0;
    node.w[8] = 0x0055;
    let n = a.alloc(node);
    let mut expected = GL0.to_vec();
    expected.extend(v2(0x3f2));
    expected.extend_from_slice(&[0x12, 0x00]); // word[6] size
    expected.extend_from_slice(&[0x34, 0x00]); // word[7] size
    expected.extend_from_slice(&[0x55, 0x00]); // word[8] trailing word
    assert_eq!(emit(&a, n), expected.as_slice());
}

// ── Type-descriptor cases (0x05, 0x11, 0x38) ─────────────────────────────────

#[test]
fn case_0x05_object_child_emits_guard() {
    // Child opcode >= 0x12 with flag byte 0x40 set → object guard 0x202; the
    // 0xf-type validation emits nothing at context 0.
    let mut a = NodeArena::new();
    let _null = a.alloc(NodeArena::node(0, 0, 0, 0, 0, 0));
    let child = a.alloc(NodeArena::node(0x20, 0, 0, 0, 0, 0)); // opcode 0x20 >= 0x12
    let mut raw = NodeArena::node(5, 0, child.0, 0, 0, 0);
    raw.w[1] = 0x4000; // byte5 bit 0x40
    let n = a.alloc(raw);
    assert_eq!(emit(&a, n), v2(0x202).as_slice());
}

#[test]
fn case_0x11_emits_wrapped_typed_node() {
    // word[4] is a plain (non-object) node → emitted with context 5 (GL0).
    let mut a = NodeArena::new();
    let _null = a.alloc(NodeArena::node(0, 0, 0, 0, 0, 0));
    let child = global_long_load(&mut a, 0);
    let n = a.alloc(NodeArena::node(0x11, 0, child.0, 0, 0, 0));
    assert_eq!(emit(&a, n), GL0.as_slice());
}

#[test]
fn case_0x38_emits_member_size_0x20d() {
    // inner.word[5] = dummy (emits nothing); inner.word[4] -> mid -> piv; piv's
    // type descriptor has size 0x10 → opcode 0x20d + 0x0010; tag 0 → no validation.
    let mut a = NodeArena::new();
    let _null = a.alloc(NodeArena::node(0, 0, 0, 0, 0, 0));
    let dummy = a.alloc(NodeArena::node(0, 0, 0, 0, 0, 0)); // opcode 0 → emits nothing
    let type_desc = a.alloc(NodeArena::node(4, 0, 0x10, 0, 0, 0)); // kind 4, size 0x10
    let piv = a.alloc(NodeArena::node(0, 0, 0, type_desc.0, 0, 0)); // word[5]=type_desc, tag 0
    let mid = a.alloc(NodeArena::node(0, 0, piv.0, 0, 0, 0)); // word[4]=piv
    let inner = a.alloc(NodeArena::node(0, 0, mid.0, dummy.0, 0, 0)); // word[4]=mid, word[5]=dummy
    let n = a.alloc(NodeArena::node(0x38, 0, 0, inner.0, 0, 0)); // word[5]=inner
    let mut expected = v2(0x20d);
    expected.extend_from_slice(&[0x10, 0x00]);
    assert_eq!(emit(&a, n), expected.as_slice());
}

#[test]
fn case_0x12_member_deref_0x160000_0x67_emits_pooled_2f6() {
    // node region 0x160000 with a 0x67 child: emit the grandchild (context 5),
    // then the pooled member opcode 0x2f6.
    let mut a = NodeArena::new();
    let _null = a.alloc(NodeArena::node(0, 0, 0, 0, 0, 0));
    let grandchild = global_long_load(&mut a, 0);
    let child = a.alloc(NodeArena::node(0x67, 0, grandchild.0, 0x99, 0, 0)); // w4=gc, w5=type
    let n = a.alloc(NodeArena::node(0x12, 0x16, child.0, 0, 0, 0)); // region 0x160000, w4=child
    let mut expected = GL0.to_vec();
    expected.extend(v2(0x2f6));
    expected.extend_from_slice(&[0x00, 0x00]);
    assert_eq!(emit(&a, n), expected.as_slice());
}

#[test]
fn case_0xc_expression_code_0xf_child_sized_opcode() {
    // 0xf-type child (emits nothing) → sized opcode 0xeb + size 8; node tag 0 so
    // the validation emits nothing.
    let mut a = NodeArena::new();
    let _null = a.alloc(NodeArena::node(0, 0, 0, 0, 0, 0));
    let child = a.alloc(NodeArena::node(0, 0xf, 0, 0, 0, 0)); // opcode 0, tag 0xf
    let type_desc = a.alloc(NodeArena::node(4, 0, 8, 0, 0, 0)); // kind 4, size 8
    let n = a.alloc(NodeArena::node(0xc, 0, child.0, 0, type_desc.0, 0)); // w4=child, w6=desc
    let mut expected = v2(0xeb);
    expected.extend_from_slice(&[0x08, 0x00]);
    assert_eq!(emit(&a, n), expected.as_slice());
}

#[test]
fn case_0xc_expression_code_0x10_child_emits_0xec() {
    // 0x10-type child (emits nothing) → value opcode 0xec (load); node tag 0 so
    // the validation emits nothing.
    let mut a = NodeArena::new();
    let _null = a.alloc(NodeArena::node(0, 0, 0, 0, 0, 0));
    let child = a.alloc(NodeArena::node(0, 0x10, 0, 0, 0, 0)); // opcode 0, tag 0x10
    let n = a.alloc(NodeArena::node(0xc, 0, child.0, 0, 0, 0)); // w4=child
    assert_eq!(emit(&a, n), v2(0xec).as_slice());
}

#[test]
fn case_0xd_expression_code_0x10_child_emits_0xee() {
    // store variant (0xd) → opcode 0xee.
    let mut a = NodeArena::new();
    let _null = a.alloc(NodeArena::node(0, 0, 0, 0, 0, 0));
    let child = a.alloc(NodeArena::node(0, 0x10, 0, 0, 0, 0));
    let n = a.alloc(NodeArena::node(0xd, 0, child.0, 0, 0, 0));
    assert_eq!(emit(&a, n), v2(0xee).as_slice());
}

#[test]
fn case_0x2d_typed_assign_0x10_child_emits_0x2c7() {
    // byte5 clear, non-object child of type 0x10: emit word[5] (here nothing)
    // then sized opcode 0x2c7 with the child's type size.
    let mut a = NodeArena::new();
    let _null = a.alloc(NodeArena::node(0, 0, 0, 0, 0, 0));
    let type_desc = a.alloc(NodeArena::node(4, 0, 0x10, 0, 0, 0)); // kind 4, size 0x10
    let child = a.alloc(NodeArena::node(0, 0x10, 0, type_desc.0, 0, 0)); // tag 0x10, w5=desc
    let nothing = a.alloc(NodeArena::node(0x1b, 0, 0, 0, 0, 0)); // word[5] emits nothing
    let n = a.alloc(NodeArena::node(0x2d, 0, child.0, nothing.0, 0, 0));
    let mut expected = v2(0x2c7);
    expected.extend_from_slice(&[0x10, 0x00]);
    assert_eq!(emit(&a, n), expected.as_slice());
}

#[test]
fn case_0x5a_complex_binop_branch2() {
    // flags 0x4000 set, 0x2000/0x8000 clear → traverse (nothing), emit operand
    // (GL0), opcode 0x390, then the two type words 0x22 and 0x11.
    let mut a = NodeArena::new();
    let _null = a.alloc(NodeArena::node(0, 0, 0, 0, 0, 0));
    let dummy = a.alloc(NodeArena::node(0, 0, 0, 0, 0, 0)); // traverse target → nothing
    let gl = global_long_load(&mut a, 0); // emit operand → GL0
    let n5w4 = a.alloc(NodeArena::node(0, 0, 0x11, 0, 0, 0)); // uvar2 source
    let b1w4 = a.alloc(NodeArena::node(0, 0, 0x22, 0, 0, 0)); // uvar3 source
    let b1 = a.alloc(NodeArena::node(0, 0, b1w4.0, dummy.0, 0, 0));
    let b0 = a.alloc(NodeArena::node(0, 0, gl.0, b1.0, 0, 0));
    let n5 = a.alloc(NodeArena::node(0, 0, n5w4.0, b0.0, 0, 0));
    let mut raw = NodeArena::node(0x5a, 0, 0, n5.0, 0, 0);
    raw.w[1] = 0x4000;
    let n = a.alloc(raw);

    let mut expected = GL0.to_vec();
    expected.extend(v2(0x390));
    expected.extend_from_slice(&[0x22, 0x00, 0x11, 0x00]);
    assert_eq!(emit(&a, n), expected.as_slice());
}

#[test]
fn case_0x6e_with_member_emits_arglist_first() {
    // member id 1 → EbFindActualNode emits the arg (GL0, context 3) first, then
    // the target (GL4, context 6), opcode 0x398, pooled word, and member word 1.
    let mut a = NodeArena::new();
    let _null = a.alloc(NodeArena::node(0, 0, 0, 0, 0, 0));
    let target = global_long_load(&mut a, 4); // GL4
    let arg = global_long_load(&mut a, 0); // GL0
    let mut raw = NodeArena::node(0x6e, 0, target.0, arg.0, 0, 0x99); // w4=target, w5=arg, w7=type
    raw.w[8] = 1; // member id 1 (find-actual depth + trailing member word)
    let n = a.alloc(raw);
    let mut expected = GL0.to_vec(); // arg emitted first
    expected.extend_from_slice(&GL4);
    expected.extend(v2(0x398));
    expected.extend_from_slice(&[0x00, 0x00, 0x01, 0x00]); // pooled word + member word
    assert_eq!(emit(&a, n), expected.as_slice());
}

#[test]
fn case_0x41_arglist_single_arg_emits_size_word() {
    // flag 0x8000 set, op-class 0 → opcode 0x3b2-0x119=0x299; count stays 1 →
    // bare opcode; one argument with a size-4 descriptor → trailing word 0x0004.
    let mut a = NodeArena::new();
    let _null = a.alloc(NodeArena::node(0, 0, 0, 0, 0, 0));
    let type_desc = a.alloc(NodeArena::node(4, 0, 4, 0, 0, 0)); // kind 4, size 4
    let arg = a.alloc(NodeArena::node(0, 0, 0, type_desc.0, 0, 0)); // word[5]=desc
    let mut raw = NodeArena::node(0x41, 0, 0, arg.0, 0, 0); // word[5]=arg
    raw.w[1] = 0x8000;
    let n = a.alloc(raw);
    let mut expected = v2(0x299);
    expected.extend_from_slice(&[0x04, 0x00]);
    assert_eq!(emit(&a, n), expected.as_slice());
}

// ── case 0x58 (byte5 0x40 clear: traverse + 0x3ff) ───────────────────────────

#[test]
fn case_0x58_byte5_clear_traverse_then_0x3ff() {
    let mut a = NodeArena::new();
    let child = global_long_load(&mut a, 0);
    let n = w5_node(&mut a, 0x58, child); // byte5 = 0
    assert_eq!(emit(&a, n), gl0_then(0x3ff).as_slice());
}

// ── Type-descriptor size model (sized UDT / object cases) ────────────────────
//
// A type descriptor is an arena record whose word[0] is the descriptor kind
// (4 = fixed-size) and whose word[4] low half carries the resolved byte size.
// `emit_get_type_size3` reads that size back (or the 0xffff_ffff sentinel for a
// null / non-fixed descriptor). The sized cases emit `opcode + size` (LE u16).

/// Allocate a fixed-size type descriptor carrying `size` bytes.
fn type_desc(a: &mut NodeArena, size: u16) -> NodeRef {
    a.alloc(NodeArena::node(4, 0, size as u32, 0, 0, 0))
}

/// Expected: opcode `n` (as emit_value2 would encode it) followed by `operand`
/// as a little-endian u16.
fn opc2(n: usize, operand: u16) -> Vec<u8> {
    let mut v = v2(n);
    v.extend_from_slice(&operand.to_le_bytes());
    v
}

#[test]
fn case_0x39_emits_0x20e_with_resolved_size() {
    // inner = word[5]; descriptor = inner.word[5]; size 0x10 → opcode 0x20e + 0x0010.
    let mut a = NodeArena::new();
    let desc = type_desc(&mut a, 0x10);
    let inner = a.alloc(NodeArena::node(0, 0, 0, desc.0, 0, 0)); // word[5] = desc
    let n = a.alloc(NodeArena::node(0x39, 0, 0, inner.0, 0, 0)); // word[5] = inner
    assert_eq!(emit(&a, n), opc2(0x20e, 0x10).as_slice());
}

#[test]
fn case_0x39_null_descriptor_emits_0xffff_sentinel() {
    // inner.word[5] == 0 → no descriptor → size sentinel 0xffff_ffff → operand 0xffff.
    let mut a = NodeArena::new();
    let inner = a.alloc(NodeArena::node(0, 0, 0, 0, 0, 0));
    let n = a.alloc(NodeArena::node(0x39, 0, 0, inner.0, 0, 0));
    assert_eq!(emit(&a, n), opc2(0x20e, 0xffff).as_slice());
}

#[test]
fn case_0x39_non_fixed_descriptor_emits_sentinel() {
    // descriptor kind != 4 → sentinel.
    let mut a = NodeArena::new();
    let desc = a.alloc(NodeArena::node(7, 0, 0x10, 0, 0, 0)); // kind 7, not fixed-size
    let inner = a.alloc(NodeArena::node(0, 0, 0, desc.0, 0, 0));
    let n = a.alloc(NodeArena::node(0x39, 0, 0, inner.0, 0, 0));
    assert_eq!(emit(&a, n), opc2(0x20e, 0xffff).as_slice());
}

#[test]
fn case_0x24_udt_emits_0xef_with_size() {
    // type_tag 0xf (UDT): emit both operands (ctx 1) then opcode 0xef + size.
    // word[6] = descriptor (size 8). validate(0xf, 0x17, ctx=0) emits nothing.
    let mut a = NodeArena::new();
    let lhs = int_lit(&mut a, 6, 5); // [0xf4, 0x05]
    let rhs = int_lit(&mut a, 6, 3); // [0xf4, 0x03]
    let desc = type_desc(&mut a, 8);
    let n = a.alloc(NodeArena::node(0x24, 0xf, lhs.0, rhs.0, desc.0, 0));
    let mut expected = vec![0xf4, 0x05, 0xf4, 0x03];
    expected.extend(opc2(0xef, 8));
    assert_eq!(emit(&a, n), expected.as_slice());
}

#[test]
fn case_0x1a_object_pow_emits_0xce_with_size() {
    // type_tag 0xf → node hi 0xf0000: emit operands (ctx 2) then opcode 0xce + size.
    let mut a = NodeArena::new();
    let lhs = int_lit(&mut a, 6, 5);
    let rhs = int_lit(&mut a, 6, 3);
    let desc = type_desc(&mut a, 0x20);
    let n = a.alloc(NodeArena::node(0x1a, 0xf, lhs.0, rhs.0, desc.0, 0));
    let mut expected = vec![0xf4, 0x05, 0xf4, 0x03];
    expected.extend(opc2(0xce, 0x20));
    assert_eq!(emit(&a, n), expected.as_slice());
}

// ── case 0xe op-kinds 1..7 (object/UDT assignment forms) ─────────────────────
//
// Built with a non-object LHS (so the object-LHS bail is not taken), a
// no-emit RHS (opcode 0, emits nothing), op-class in word[1] bits 8..10, and a
// size descriptor in word[6]. type_tag 8 (Long) makes the trailing validation a
// no-op, so output is exactly the case's own opcode (+ size where sized).

/// Build a case-0xe assignment node: op-class `class`, size descriptor `size`.
fn assign_op_node(a: &mut NodeArena, class: u32, size: u16) -> NodeRef {
    let rhs = a.alloc(NodeArena::node(0, 0, 0, 0, 0, 0)); // no-emit RHS
    let desc = type_desc(a, size);
    let mut node = NodeArena::node(0xe, 8, rhs.0, 0, desc.0, 0); // LHS kind 8 (Long)
    node.w[1] = class << 8; // op-class
    a.alloc(node)
}

#[test]
fn case_0xe_opclass1_udt_copy_emits_0x2fe_with_size() {
    let mut a = NodeArena::new();
    let n = assign_op_node(&mut a, 1, 0x10);
    assert_eq!(emit(&a, n), opc2(0x2fe, 0x10).as_slice());
}

#[test]
fn case_0xe_opclass2_set_emits_0x2fd() {
    let mut a = NodeArena::new();
    let n = assign_op_node(&mut a, 2, 0);
    assert_eq!(emit(&a, n), v2(0x2fd).as_slice());
}

#[test]
fn case_0xe_opclass3_emits_0x2f9_with_size() {
    let mut a = NodeArena::new();
    let n = assign_op_node(&mut a, 3, 0x18);
    assert_eq!(emit(&a, n), opc2(0x2f9, 0x18).as_slice());
}

#[test]
fn case_0xe_opclass4_emits_0x2fa_with_size() {
    let mut a = NodeArena::new();
    let n = assign_op_node(&mut a, 4, 0x18);
    assert_eq!(emit(&a, n), opc2(0x2fa, 0x18).as_slice());
}

#[test]
fn case_0xe_opclass5_emits_0x2fc_with_size() {
    let mut a = NodeArena::new();
    let n = assign_op_node(&mut a, 5, 0x18);
    assert_eq!(emit(&a, n), opc2(0x2fc, 0x18).as_slice());
}

#[test]
fn case_0xe_opclass7_me_assign_emits_0x41b() {
    let mut a = NodeArena::new();
    let n = assign_op_node(&mut a, 7, 0);
    assert_eq!(emit(&a, n), v2(0x41b).as_slice());
}

// ── case 0x57 / 0x58 sized (0x4000 set / byte5 0x40 set) ──────────────────────

#[test]
fn case_0x57_set_emits_flag_selected_sized_opcode() {
    // flag 0x4000 set, 0x8000 clear: opcode = ((!flags & 0x8000)|0xff0000)>>0xe = 0x3fe.
    let mut a = NodeArena::new();
    let child = global_long_load(&mut a, 0);
    let desc = type_desc(&mut a, 0x10);
    let mut node = NodeArena::node(0x57, 8, 0, child.0, desc.0, 0); // type_tag 8 → validate no-op
    node.w[1] = 0x4000;
    let n = a.alloc(node);
    let mut expected = GL0.to_vec();
    expected.extend(opc2(0x3fe, 0x10));
    assert_eq!(emit(&a, n), expected.as_slice());
}

#[test]
fn case_0x57_set_8000_emits_0x3fc_sized() {
    // flag 0x4000 and 0x8000 set: opcode = ((0) | 0xff0000) >> 0xe = 0x3fc.
    let mut a = NodeArena::new();
    let child = global_long_load(&mut a, 0);
    let desc = type_desc(&mut a, 0x10);
    let mut node = NodeArena::node(0x57, 8, 0, child.0, desc.0, 0);
    node.w[1] = 0xc000;
    let n = a.alloc(node);
    let mut expected = GL0.to_vec();
    expected.extend(opc2(0x3fc, 0x10));
    assert_eq!(emit(&a, n), expected.as_slice());
}

#[test]
fn case_0x58_set_emits_0x400_sized() {
    // byte5 bit 0x40 set (word[1] = 0x4000): opcode 0x400 + size.
    let mut a = NodeArena::new();
    let child = global_long_load(&mut a, 0);
    let desc = type_desc(&mut a, 0x10);
    let mut node = NodeArena::node(0x58, 8, 0, child.0, desc.0, 0);
    node.w[1] = 0x4000;
    let n = a.alloc(node);
    let mut expected = GL0.to_vec();
    expected.extend(opc2(0x400, 0x10));
    assert_eq!(emit(&a, n), expected.as_slice());
}

// ── Member reference (0x60) end-to-end through the resolver + value emitter ──

#[test]
fn member_reference_0x60_resolves_and_emits() {
    use crate::emit::SymbolContext;

    let mut a = NodeArena::new();
    // Size descriptor (word[0]==4, size in word[4]) for init_expr_descriptor.
    let d = a.alloc(NodeArena::node(4, 0, 4, 0, 0, 0));
    // 0x60 member-reference node: type tag 8, no member sub-expr (w4=0), size
    // descriptor in w5.
    let mut n = NodeArena::node(0x60, 8, 0, 0, 0, 0);
    n.w[5] = d.0;
    let node = a.alloc(n);

    // Records heap: a member record at 0x10 classifying to resolver category 4
    // (inline Long operand at +0xc, record byte+1 low3 = 3).
    let mut heap = vec![0u8; 0x40];
    heap[0x10] = 0x40; // +0 bit 6 → inline operand at +0xc
    heap[0x11] = 0x03; // +1 low3 = 3
    heap[0x10 + 0xc] = 8; // inline Long operand opcode

    let sym = SymbolContext {
        heap,
        member_off: 0x10,
        ctx_flag_c: 0,
        binding: Some((4, 0)), // binder-resolved (kind 4, byref 0)
    };

    let mut e = Emitter::new(&a).with_symbol_context(sym);
    e.emit_expr(node, 1);
    let bytes = e.into_bytes();

    // The resolver yields a kind-1 (local, by-ref) descriptor with operand 4;
    // the value emitter emits its load opcode + operand. Locked from the ported
    // EbEmitExpression2 path.
    assert_eq!(bytes, MEMBER_REF_0X60_BYTES);
}

const MEMBER_REF_0X60_BYTES: &[u8] = &[0x6c, 0x04, 0x00];

#[test]
fn member_reference_0x60_without_context_is_gated() {
    let mut a = NodeArena::new();
    let d = a.alloc(NodeArena::node(4, 0, 4, 0, 0, 0));
    let mut n = NodeArena::node(0x60, 8, 0, 0, 0, 0);
    n.w[5] = d.0;
    let node = a.alloc(n);
    let r = std::panic::catch_unwind(std::panic::AssertUnwindSafe(|| {
        let mut e = Emitter::new(&a);
        e.emit_expr(node, 1);
    }));
    assert!(r.is_err());
}

// ── Assignment statement (0x2c) common scalar path ──────────────────────────

#[test]
fn assignment_0x2c_emits_rhs_then_resolved_store() {
    use crate::emit::SymbolContext;

    let mut a = NodeArena::new();
    // RHS: a Long local load at frame offset -8.
    let s = a.alloc(NodeArena::node(0, 0, (0xfff8u32) << 16, 0, 0, 0));
    let rhs = a.alloc(NodeArena::node(0x74, 8, s.0, 2 /* CTX_LONG */, 0, 0));
    // LHS: a 0x60 member reference (type tag 8), size descriptor in w5.
    let d = a.alloc(NodeArena::node(4, 0, 4, 0, 0, 0));
    let mut lhs_n = NodeArena::node(0x60, 8, 0, 0, 0, 0);
    lhs_n.w[5] = d.0;
    let lhs = a.alloc(lhs_n);
    // 0x2c assignment node: w4 = LHS, w5 = RHS, region 0 (not 0x20000).
    let mut asn = NodeArena::node(0x2c, 0, 0, 0, 0, 0);
    asn.w[4] = lhs.0;
    asn.w[5] = rhs.0;
    let node = a.alloc(asn);

    let mut heap = vec![0u8; 0x40];
    heap[0x10] = 0x40;
    heap[0x11] = 0x03;
    heap[0x10 + 0xc] = 8;
    let sym = SymbolContext {
        heap,
        member_off: 0x10,
        ctx_flag_c: 0,
        binding: Some((4, 0)),
    };

    let mut e = Emitter::new(&a).with_symbol_context(sym);
    e.emit_expr(node, 0);
    let bytes = e.into_bytes();
    assert_eq!(bytes, ASSIGN_0X2C_BYTES);
}

// RHS Long load (6c f8 ff) then the resolved store of the LHS member, through
// the ported EbEmitAssignmentStmt → resolve_reference2 → value-emitter chain.
const ASSIGN_0X2C_BYTES: &[u8] = &[0x6c, 0xf8, 0xff, 0x71, 0x04, 0x00, 0x6c, 0x04, 0x00];
