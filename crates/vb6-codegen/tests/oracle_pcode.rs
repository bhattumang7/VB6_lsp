//! Byte-exact oracle confirmation for the emitter.
//!
//! Every expected byte vector here was captured from a real VB6-compiled p-code
//! exe by `re_lab/pcode_lab/oracle_survey.py` (and `cmp_survey.py`): compile a
//! one-line `Sub Main`, locate Sub Main's p-code block, and read the exact
//! bytes. These tests assert the Rust emitter reproduces that ground truth.
//!
//! Each statement `r = <expr>` becomes: emit the bound expression
//! (`Emitter::emit_expr`) followed by the typed store of `r`
//! (`Emitter::emit_var_store`). Frame offsets match the oracle's layout for the
//! corresponding `Dim` order.

use vb6_codegen::{Emitter, NodeArena, NodeRef};

/// VB6 internal type tags used in `word[0]` high half.
const T_INTEGER: u16 = 6;
const T_LONG: u16 = 8;
const T_DOUBLE: u16 = 11;

/// Runtime load type-contexts (`word[5]` of a load node).
const CTX_INTEGER: u16 = 1;
const CTX_LONG: u16 = 2;
const CTX_DOUBLE: u16 = 4;

/// Bound symbol node carrying a frame offset in its `type_info` field.
fn sym(a: &mut NodeArena, offset: i16) -> NodeRef {
    a.alloc(NodeArena::node(0, 0, (offset as u16 as u32) << 16, 0, 0, 0))
}

/// Typed local-load node (opcode 0x74): VB6 type tag in `word[0]` high half,
/// load context in `word[5]`.
fn load(a: &mut NodeArena, vb_type: u16, ctx: u16, offset: i16) -> NodeRef {
    let s = sym(a, offset);
    a.alloc(NodeArena::node(0x74, vb_type, s.0, ctx as u32, 0, 0))
}

/// Binary-op node with its own result type tag (used by the arithmetic path).
fn binop(a: &mut NodeArena, op: u16, type_tag: u16, lhs: NodeRef, rhs: NodeRef) -> NodeRef {
    a.alloc(NodeArena::node(op, type_tag, lhs.0, rhs.0, 0, 0))
}

/// Comparison node (own type tag 0; the opcode comes from the LHS operand type).
fn cmp(a: &mut NodeArena, op: u16, lhs: NodeRef, rhs: NodeRef) -> NodeRef {
    a.alloc(NodeArena::node(op, 0, lhs.0, rhs.0, 0, 0))
}

/// Emit `r = <root>`: the expression, then a typed store of `r`.
fn stmt(a: &NodeArena, root: NodeRef, store_ctx: u16, r_offset: i16) -> Vec<u8> {
    let mut e = Emitter::new(a);
    e.emit_expr(root, 0);
    e.emit_var_store(store_ctx as usize, r_offset);
    e.into_bytes()
}

// ── Long arithmetic: Dim a As Long, b As Long, r As Long ─────────────────────
// Frame: a=-136 (0xff78), b=-140 (0xff74), r=-144 (0xff70).

/// Build `a <op> b` with Long operands and a Long result, return `r = ...` bytes.
fn long_arith(op: u16) -> Vec<u8> {
    let mut a = NodeArena::new();
    let lhs = load(&mut a, T_LONG, CTX_LONG, 0xff78u16 as i16);
    let rhs = load(&mut a, T_LONG, CTX_LONG, 0xff74u16 as i16);
    let n = binop(&mut a, op, T_LONG, lhs, rhs);
    stmt(&a, n, CTX_LONG, 0xff70u16 as i16)
}

#[test]
fn long_add() {
    // r = a + b   (op 0x16)
    assert_eq!(long_arith(0x16), &[0x6c, 0x78, 0xff, 0x6c, 0x74, 0xff, 0xaa, 0x71, 0x70, 0xff]);
}

#[test]
fn long_sub() {
    // r = a - b   (op 0x17)
    assert_eq!(long_arith(0x17), &[0x6c, 0x78, 0xff, 0x6c, 0x74, 0xff, 0xae, 0x71, 0x70, 0xff]);
}

#[test]
fn long_mul() {
    // r = a * b   (op 0x18)
    assert_eq!(long_arith(0x18), &[0x6c, 0x78, 0xff, 0x6c, 0x74, 0xff, 0xb2, 0x71, 0x70, 0xff]);
}

#[test]
fn long_and() {
    // r = a And b   (op 0x23)
    assert_eq!(long_arith(0x23), &[0x6c, 0x78, 0xff, 0x6c, 0x74, 0xff, 0xc4, 0x71, 0x70, 0xff]);
}

#[test]
fn long_or() {
    // r = a Or b   (op 0x21)
    assert_eq!(long_arith(0x21), &[0x6c, 0x78, 0xff, 0x6c, 0x74, 0xff, 0xc5, 0x71, 0x70, 0xff]);
}

#[test]
fn long_xor() {
    // r = a Xor b   (op 0x22) — extended form 0xfb 0x13
    assert_eq!(long_arith(0x22), &[0x6c, 0x78, 0xff, 0x6c, 0x74, 0xff, 0xfb, 0x13, 0x71, 0x70, 0xff]);
}

// ── Long comparisons into Integer: Dim a As Long, b As Long, r As Integer ────
// Frame: a=-136 (0xff78), b=-140 (0xff74), r=-142 (0xff72). Store Integer 0x70.

/// Build `(a <op> b)` with Long operands, return `r = ...` (Integer store) bytes.
fn long_cmp(op: u16) -> Vec<u8> {
    let mut a = NodeArena::new();
    let lhs = load(&mut a, T_LONG, CTX_LONG, 0xff78u16 as i16);
    let rhs = load(&mut a, T_LONG, CTX_LONG, 0xff74u16 as i16);
    let n = cmp(&mut a, op, lhs, rhs);
    stmt(&a, n, CTX_INTEGER, 0xff72u16 as i16)
}

#[test]
fn long_eq() {
    // r = (a = b)   (op 0x26)
    assert_eq!(long_cmp(0x26), &[0x6c, 0x78, 0xff, 0x6c, 0x74, 0xff, 0xc7, 0x70, 0x72, 0xff]);
}

#[test]
fn long_ne() {
    // r = (a <> b)   (op 0x27)
    assert_eq!(long_cmp(0x27), &[0x6c, 0x78, 0xff, 0x6c, 0x74, 0xff, 0xcc, 0x70, 0x72, 0xff]);
}

#[test]
fn long_le() {
    // r = (a <= b)   (op 0x28)
    assert_eq!(long_cmp(0x28), &[0x6c, 0x78, 0xff, 0x6c, 0x74, 0xff, 0xd6, 0x70, 0x72, 0xff]);
}

#[test]
fn long_ge() {
    // r = (a >= b)   (op 0x29)
    assert_eq!(long_cmp(0x29), &[0x6c, 0x78, 0xff, 0x6c, 0x74, 0xff, 0xe0, 0x70, 0x72, 0xff]);
}

#[test]
fn long_lt() {
    // r = (a < b)   (op 0x2a)
    assert_eq!(long_cmp(0x2a), &[0x6c, 0x78, 0xff, 0x6c, 0x74, 0xff, 0xd1, 0x70, 0x72, 0xff]);
}

#[test]
fn long_gt() {
    // r = (a > b)   (op 0x2b)
    assert_eq!(long_cmp(0x2b), &[0x6c, 0x78, 0xff, 0x6c, 0x74, 0xff, 0xdb, 0x70, 0x72, 0xff]);
}

// ── Integer arithmetic: Dim a As Integer, b As Integer, r As Integer ─────────
// Frame: a=-134 (0xff7a), b=-136 (0xff78), r=-138 (0xff76).

#[test]
fn integer_add() {
    // r = a + b   (op 0x16, Integer result)
    let mut a = NodeArena::new();
    let lhs = load(&mut a, T_INTEGER, CTX_INTEGER, 0xff7au16 as i16);
    let rhs = load(&mut a, T_INTEGER, CTX_INTEGER, 0xff78u16 as i16);
    let n = binop(&mut a, 0x16, T_INTEGER, lhs, rhs);
    assert_eq!(
        stmt(&a, n, CTX_INTEGER, 0xff76u16 as i16),
        &[0x6b, 0x7a, 0xff, 0x6b, 0x78, 0xff, 0xa9, 0x70, 0x76, 0xff]
    );
}

// ── Double arithmetic: Dim a As Double, b As Double, r As Double ─────────────
// Frame: a=-140 (0xff74), b=-148 (0xff6c), r=-156 (0xff64).

/// Build `a <op> b` with Double operands and a Double result.
fn double_arith(op: u16) -> Vec<u8> {
    let mut a = NodeArena::new();
    let lhs = load(&mut a, T_DOUBLE, CTX_DOUBLE, 0xff74u16 as i16);
    let rhs = load(&mut a, T_DOUBLE, CTX_DOUBLE, 0xff6cu16 as i16);
    let n = binop(&mut a, op, T_DOUBLE, lhs, rhs);
    stmt(&a, n, CTX_DOUBLE, 0xff64u16 as i16)
}

#[test]
fn double_add() {
    // r = a + b   (op 0x16, Double result)
    assert_eq!(double_arith(0x16), &[0x6f, 0x74, 0xff, 0x6f, 0x6c, 0xff, 0xab, 0x74, 0x64, 0xff]);
}

#[test]
fn double_div() {
    // r = a / b   (op 0x19, floating divide). Oracle: 0xb6.
    let mut a = NodeArena::new();
    let lhs = load(&mut a, T_DOUBLE, CTX_DOUBLE, 0xff74u16 as i16);
    let rhs = load(&mut a, T_DOUBLE, CTX_DOUBLE, 0xff6cu16 as i16);
    let n = binop(&mut a, 0x19, T_DOUBLE, lhs, rhs);
    assert_eq!(
        stmt(&a, n, CTX_DOUBLE, 0xff64u16 as i16),
        &[0x6f, 0x74, 0xff, 0x6f, 0x6c, 0xff, 0xb6, 0x74, 0x64, 0xff]
    );
}
