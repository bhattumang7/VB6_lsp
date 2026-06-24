//! End-to-end pipeline tests: VB6 source text → parse → sema → lower → P-code bytes.
//!
//! Every expected byte vector here was captured from a real VB6-compiled p-code
//! exe by `re_lab/pcode_lab/e2e_survey.py`.  These tests assert that the full
//! Rust pipeline (ScannerContext → Parser → bind → lower_proc) reproduces the
//! exact byte stream the VB6 compiler emits.
//!
//! The oracle bytes for the first group duplicate the vectors already in
//! `oracle_pcode.rs` (which verify the emitter in isolation); here the same
//! bytes verify the **complete pipeline** from source text.

use vb6_codegen::lower_proc;
use vb6_sema::frontend::ast::ExprArena;
use vb6_sema::frontend::parser::Parser;
use vb6_sema::frontend::scanner::ScannerContext;
use vb6_sema::sema::bind;

// ── Pipeline helper ───────────────────────────────────────────────────────────

/// Parse, bind, and lower `src` (a complete VB6 module source), returning the
/// P-code bytes for the first proc (proc index 0).
///
/// `module_desc` is the compiled module-object descriptor word (`0x0008` for
/// the primary module in a single-module project — oracle-confirmed).
fn compile(src: &str, module_desc: u16) -> Vec<u8> {
    let mut ctx = ScannerContext::new(1, 1, 0x0409);
    ctx.intern_keywords();
    let mut arena = ExprArena::new();
    let mut parser = Parser::new(&mut ctx, src.as_bytes());
    let top = parser.parse_module(&mut arena);
    let spans = std::mem::take(&mut parser.node_spans);
    let vis = std::mem::take(&mut parser.decl_public);
    drop(parser);
    let module = bind(&ctx, &arena, &top, &spans, &vis);
    lower_proc(&module, 0, &arena, module_desc)
        .unwrap_or_else(|e| panic!("lower_proc failed: {e:?}"))
}

// ── Long arithmetic (Dim a As Long, b As Long, r As Long) ────────────────────
//
// Oracle: compile with real VB6 in p-code mode, extract Sub Main body bytes
// (stripping the trailing 0x14 End-Sub marker).  These vectors match the ones
// in oracle_pcode.rs, confirming the pipeline and the emitter agree.

#[test]
fn e2e_long_add() {
    // r = a + b  →  load a, load b, Long-Add, store r
    assert_eq!(
        compile(
            "Attribute VB_Name = \"Module1\"\r\n\
             Sub Main()\r\n\
             Dim a As Long, b As Long, r As Long\r\n\
             r = a + b\r\n\
             End Sub\r\n",
            0x0008,
        ),
        &[0x6c, 0x78, 0xff, 0x6c, 0x74, 0xff, 0xaa, 0x71, 0x70, 0xff]
    );
}

#[test]
fn e2e_long_sub() {
    assert_eq!(
        compile(
            "Attribute VB_Name = \"Module1\"\r\n\
             Sub Main()\r\n\
             Dim a As Long, b As Long, r As Long\r\n\
             r = a - b\r\n\
             End Sub\r\n",
            0x0008,
        ),
        &[0x6c, 0x78, 0xff, 0x6c, 0x74, 0xff, 0xae, 0x71, 0x70, 0xff]
    );
}

#[test]
fn e2e_long_mul() {
    assert_eq!(
        compile(
            "Attribute VB_Name = \"Module1\"\r\n\
             Sub Main()\r\n\
             Dim a As Long, b As Long, r As Long\r\n\
             r = a * b\r\n\
             End Sub\r\n",
            0x0008,
        ),
        &[0x6c, 0x78, 0xff, 0x6c, 0x74, 0xff, 0xb2, 0x71, 0x70, 0xff]
    );
}

// ── Long comparisons (r As Integer) ──────────────────────────────────────────
//
// Frame: a=-136 (0xff78), b=-140 (0xff74), r=-142 (0xff72).
// Integer is 2 bytes, so r immediately follows b's 4-byte slot.

#[test]
fn e2e_long_eq_into_integer() {
    assert_eq!(
        compile(
            "Attribute VB_Name = \"Module1\"\r\n\
             Sub Main()\r\n\
             Dim a As Long, b As Long, r As Integer\r\n\
             r = (a = b)\r\n\
             End Sub\r\n",
            0x0008,
        ),
        &[0x6c, 0x78, 0xff, 0x6c, 0x74, 0xff, 0xc7, 0x70, 0x72, 0xff]
    );
}

#[test]
fn e2e_long_ne_into_integer() {
    assert_eq!(
        compile(
            "Attribute VB_Name = \"Module1\"\r\n\
             Sub Main()\r\n\
             Dim a As Long, b As Long, r As Integer\r\n\
             r = (a <> b)\r\n\
             End Sub\r\n",
            0x0008,
        ),
        &[0x6c, 0x78, 0xff, 0x6c, 0x74, 0xff, 0xcc, 0x70, 0x72, 0xff]
    );
}

#[test]
fn e2e_long_lt_into_integer() {
    assert_eq!(
        compile(
            "Attribute VB_Name = \"Module1\"\r\n\
             Sub Main()\r\n\
             Dim a As Long, b As Long, r As Integer\r\n\
             r = (a < b)\r\n\
             End Sub\r\n",
            0x0008,
        ),
        &[0x6c, 0x78, 0xff, 0x6c, 0x74, 0xff, 0xd1, 0x70, 0x72, 0xff]
    );
}

#[test]
fn e2e_long_le_into_integer() {
    assert_eq!(
        compile(
            "Attribute VB_Name = \"Module1\"\r\n\
             Sub Main()\r\n\
             Dim a As Long, b As Long, r As Integer\r\n\
             r = (a <= b)\r\n\
             End Sub\r\n",
            0x0008,
        ),
        &[0x6c, 0x78, 0xff, 0x6c, 0x74, 0xff, 0xd6, 0x70, 0x72, 0xff]
    );
}

#[test]
fn e2e_long_gt_into_integer() {
    assert_eq!(
        compile(
            "Attribute VB_Name = \"Module1\"\r\n\
             Sub Main()\r\n\
             Dim a As Long, b As Long, r As Integer\r\n\
             r = (a > b)\r\n\
             End Sub\r\n",
            0x0008,
        ),
        &[0x6c, 0x78, 0xff, 0x6c, 0x74, 0xff, 0xdb, 0x70, 0x72, 0xff]
    );
}

#[test]
fn e2e_long_ge_into_integer() {
    assert_eq!(
        compile(
            "Attribute VB_Name = \"Module1\"\r\n\
             Sub Main()\r\n\
             Dim a As Long, b As Long, r As Integer\r\n\
             r = (a >= b)\r\n\
             End Sub\r\n",
            0x0008,
        ),
        &[0x6c, 0x78, 0xff, 0x6c, 0x74, 0xff, 0xe0, 0x70, 0x72, 0xff]
    );
}

// ── Integer arithmetic ────────────────────────────────────────────────────────
//
// Frame: a=-134 (0xff7a), b=-136 (0xff78), r=-138 (0xff76).
// All Integer (2 bytes each); no 4-byte alignment applied.

#[test]
fn e2e_integer_add() {
    assert_eq!(
        compile(
            "Attribute VB_Name = \"Module1\"\r\n\
             Sub Main()\r\n\
             Dim a As Integer, b As Integer, r As Integer\r\n\
             r = a + b\r\n\
             End Sub\r\n",
            0x0008,
        ),
        &[0x6b, 0x7a, 0xff, 0x6b, 0x78, 0xff, 0xa9, 0x70, 0x76, 0xff]
    );
}

// ── Double arithmetic ─────────────────────────────────────────────────────────
//
// Frame: a=-140 (0xff74), b=-148 (0xff6c), r=-156 (0xff64).
// Double = 8 bytes each.

#[test]
fn e2e_double_add() {
    assert_eq!(
        compile(
            "Attribute VB_Name = \"Module1\"\r\n\
             Sub Main()\r\n\
             Dim a As Double, b As Double, r As Double\r\n\
             r = a + b\r\n\
             End Sub\r\n",
            0x0008,
        ),
        &[0x6f, 0x74, 0xff, 0x6f, 0x6c, 0xff, 0xab, 0x74, 0x64, 0xff]
    );
}

// ── Control flow — If/While/Do (Long locals a=-136/0xff78, r=-140/0xff74) ────
//
// Oracle bytes captured from real VB6 p-code exe via cf_survey.py.
// Condition `a > 0`: load a (0x6c 0x78 0xff), push Long 0 (0xf5 0x00..0x00),
// Long-Gt (0xdb).
// Store r=1: push Long 1 (0xf5 0x01 0x00 0x00 0x00), store Long r (0x71 0x74 0xff).

#[test]
fn e2e_if_no_else() {
    // If a > 0 Then r = 1 End If  — Long locals a=-136, r=-140
    // BranchFalse (0x1c) to absolute offset 20 (end of if block)
    assert_eq!(
        compile(
            "Attribute VB_Name = \"Module1\"\r\n\
             Sub Main()\r\n\
             Dim a As Long, r As Long\r\n\
             If a > 0 Then\r\n\
             r = 1\r\n\
             End If\r\n\
             End Sub\r\n",
            0x0008,
        ),
        &[
            0x6c, 0x78, 0xff,                   // load Long a
            0xf5, 0x00, 0x00, 0x00, 0x00,       // push Long 0
            0xdb,                                // Long >
            0x1c, 0x14, 0x00,                   // BranchFalse → offset 20
            0xf5, 0x01, 0x00, 0x00, 0x00,       // push Long 1
            0x71, 0x74, 0xff,                    // store Long r
        ]
    );
}

#[test]
fn e2e_if_else() {
    // If a > 0 Then r = 1 Else r = 2 End If
    // BranchFalse to else start (offset 23), Jump to end (offset 31)
    assert_eq!(
        compile(
            "Attribute VB_Name = \"Module1\"\r\n\
             Sub Main()\r\n\
             Dim a As Long, r As Long\r\n\
             If a > 0 Then\r\n\
             r = 1\r\n\
             Else\r\n\
             r = 2\r\n\
             End If\r\n\
             End Sub\r\n",
            0x0008,
        ),
        &[
            0x6c, 0x78, 0xff,                   // load Long a
            0xf5, 0x00, 0x00, 0x00, 0x00,       // push Long 0
            0xdb,                                // Long >
            0x1c, 0x17, 0x00,                   // BranchFalse → offset 23 (else start)
            0xf5, 0x01, 0x00, 0x00, 0x00,       // push Long 1
            0x71, 0x74, 0xff,                    // store Long r
            0x1e, 0x1f, 0x00,                   // Jump → offset 31 (end)
            0xf5, 0x02, 0x00, 0x00, 0x00,       // push Long 2
            0x71, 0x74, 0xff,                    // store Long r
        ]
    );
}

#[test]
fn e2e_while_loop() {
    // While a > 0: a = a - 1: Wend  — one Long local a=-136
    // BranchFalse to offset 27 (past end), Jump to offset 0 (loop start)
    assert_eq!(
        compile(
            "Attribute VB_Name = \"Module1\"\r\n\
             Sub Main()\r\n\
             Dim a As Long\r\n\
             While a > 0\r\n\
             a = a - 1\r\n\
             Wend\r\n\
             End Sub\r\n",
            0x0008,
        ),
        &[
            0x6c, 0x78, 0xff,                   // load Long a
            0xf5, 0x00, 0x00, 0x00, 0x00,       // push Long 0
            0xdb,                                // Long >
            0x1c, 0x1b, 0x00,                   // BranchFalse → offset 27
            0x6c, 0x78, 0xff,                    // load Long a
            0xf5, 0x01, 0x00, 0x00, 0x00,       // push Long 1
            0xae,                                // Long sub
            0x71, 0x78, 0xff,                    // store Long a
            0x1e, 0x00, 0x00,                    // Jump → offset 0 (loop start)
        ]
    );
}

#[test]
fn e2e_do_while_loop() {
    // Do While a > 0: a = a - 1: Loop  — identical byte layout to While/Wend
    assert_eq!(
        compile(
            "Attribute VB_Name = \"Module1\"\r\n\
             Sub Main()\r\n\
             Dim a As Long\r\n\
             Do While a > 0\r\n\
             a = a - 1\r\n\
             Loop\r\n\
             End Sub\r\n",
            0x0008,
        ),
        &[
            0x6c, 0x78, 0xff,
            0xf5, 0x00, 0x00, 0x00, 0x00,
            0xdb,
            0x1c, 0x1b, 0x00,
            0x6c, 0x78, 0xff,
            0xf5, 0x01, 0x00, 0x00, 0x00,
            0xae,
            0x71, 0x78, 0xff,
            0x1e, 0x00, 0x00,
        ]
    );
}

#[test]
fn e2e_do_loop_while() {
    // Do: a = a - 1: Loop While a > 0  — body first, BranchTrue back
    assert_eq!(
        compile(
            "Attribute VB_Name = \"Module1\"\r\n\
             Sub Main()\r\n\
             Dim a As Long\r\n\
             Do\r\n\
             a = a - 1\r\n\
             Loop While a > 0\r\n\
             End Sub\r\n",
            0x0008,
        ),
        &[
            0x6c, 0x78, 0xff,                   // load Long a
            0xf5, 0x01, 0x00, 0x00, 0x00,       // push Long 1
            0xae,                                // Long sub
            0x71, 0x78, 0xff,                    // store Long a
            0x6c, 0x78, 0xff,                    // load Long a
            0xf5, 0x00, 0x00, 0x00, 0x00,       // push Long 0
            0xdb,                                // Long >
            0x1d, 0x00, 0x00,                    // BranchTrue → offset 0 (loop start)
        ]
    );
}

#[test]
fn e2e_nested_if() {
    // If a > 0 Then: If b > 0 Then r = 1 End If: End If
    // Nested BranchFalse — oracle: inner If jumps to 0x20 (32), outer also jumps to 0x20
    assert_eq!(
        compile(
            "Attribute VB_Name = \"Module1\"\r\n\
             Sub Main()\r\n\
             Dim a As Long, b As Long, r As Long\r\n\
             If a > 0 Then\r\n\
             If b > 0 Then\r\n\
             r = 1\r\n\
             End If\r\n\
             End If\r\n\
             End Sub\r\n",
            0x0008,
        ),
        &[
            0x6c, 0x78, 0xff,                   // load Long a   (a=-136)
            0xf5, 0x00, 0x00, 0x00, 0x00,       // push Long 0
            0xdb,                                // Long >
            0x1c, 0x20, 0x00,                   // BranchFalse → 32 (end)
            0x6c, 0x74, 0xff,                    // load Long b   (b=-140)
            0xf5, 0x00, 0x00, 0x00, 0x00,       // push Long 0
            0xdb,                                // Long >
            0x1c, 0x20, 0x00,                   // BranchFalse → 32 (end)
            0xf5, 0x01, 0x00, 0x00, 0x00,       // push Long 1
            0x71, 0x70, 0xff,                    // store Long r  (r=-144)
        ]
    );
}

// ── For loop (Long, no Step) ──────────────────────────────────────────────────
//
// For i = 1 To 10: r = r + i: Next i
// Two Long locals: i=-136 (0xff78), r=-140 (0xff74)
// Two hidden Long slots: hidden_0=-144 (0xff70), hidden_1=-148 (0xff6c)
//
// Byte layout:
//   push Long 1              [f5 01 00 00 00]
//   LdAddr i                 [04 78 ff]
//   push Long 10             [f5 0a 00 00 00]
//   ForInit no-step          [fe 64 6c ff 25 00]   frame_hidden=0xff6c, exit=37
//   load r                   [6c 74 ff]
//   load i                   [6c 78 ff]
//   Long add                 [aa]
//   store r                  [71 74 ff]
//   LdAddr i                 [04 78 ff]
//   ForNext no-step          [66 6c ff 13 00]      frame_hidden=0xff6c, back=19

#[test]
fn e2e_for_loop_long_no_step() {
    assert_eq!(
        compile(
            "Attribute VB_Name = \"Module1\"\r\n\
             Sub Main()\r\n\
             Dim i As Long, r As Long\r\n\
             For i = 1 To 10\r\n\
             r = r + i\r\n\
             Next i\r\n\
             End Sub\r\n",
            0x0008,
        ),
        &[
            0xf5, 0x01, 0x00, 0x00, 0x00,           // push Long 1 (start)
            0x04, 0x78, 0xff,                         // LdAddr i
            0xf5, 0x0a, 0x00, 0x00, 0x00,             // push Long 10 (end)
            0xfe, 0x64, 0x6c, 0xff, 0x25, 0x00,       // ForInit no-step: hidden=0xff6c, exit=37
            0x6c, 0x74, 0xff,                         // load Long r
            0x6c, 0x78, 0xff,                         // load Long i
            0xaa,                                     // Long add
            0x71, 0x74, 0xff,                         // store Long r
            0x04, 0x78, 0xff,                         // LdAddr i
            0x66, 0x6c, 0xff, 0x13, 0x00,             // ForNext no-step: hidden=0xff6c, back=19
        ]
    );
}

// ── For loop (Long, with Step) ────────────────────────────────────────────────
//
// For i = 1 To 10 Step 2: r = r + i: Next i
// Same frame layout as above.
//
// Byte layout:
//   push Long 1              [f5 01 00 00 00]
//   LdAddr i                 [04 78 ff]
//   push Long 10             [f5 0a 00 00 00]
//   push Long 2              [f5 02 00 00 00]
//   ForInit with-step        [fe 6c 6c ff 2a 00]   frame_hidden=0xff6c, exit=42
//   load r                   [6c 74 ff]
//   load i                   [6c 78 ff]
//   Long add                 [aa]
//   store r                  [71 74 ff]
//   LdAddr i                 [04 78 ff]
//   ForNext with-step        [67 6c ff 18 00]      frame_hidden=0xff6c, back=24

#[test]
fn e2e_for_loop_long_with_step() {
    assert_eq!(
        compile(
            "Attribute VB_Name = \"Module1\"\r\n\
             Sub Main()\r\n\
             Dim i As Long, r As Long\r\n\
             For i = 1 To 10 Step 2\r\n\
             r = r + i\r\n\
             Next i\r\n\
             End Sub\r\n",
            0x0008,
        ),
        &[
            0xf5, 0x01, 0x00, 0x00, 0x00,           // push Long 1 (start)
            0x04, 0x78, 0xff,                         // LdAddr i
            0xf5, 0x0a, 0x00, 0x00, 0x00,             // push Long 10 (end)
            0xf5, 0x02, 0x00, 0x00, 0x00,             // push Long 2 (step)
            0xfe, 0x6c, 0x6c, 0xff, 0x2a, 0x00,       // ForInit with-step: hidden=0xff6c, exit=42
            0x6c, 0x74, 0xff,                         // load Long r
            0x6c, 0x78, 0xff,                         // load Long i
            0xaa,                                     // Long add
            0x71, 0x74, 0xff,                         // store Long r
            0x04, 0x78, 0xff,                         // LdAddr i
            0x67, 0x6c, 0xff, 0x18, 0x00,             // ForNext with-step: hidden=0xff6c, back=24
        ]
    );
}

// ── Two sequential assignments ────────────────────────────────────────────────
//
// `b = a` then `a = b` with two Long locals:
// frame a=-136 (0xff78), b=-140 (0xff74).
// Bytes: load a, store b, load b, store a.

#[test]
fn e2e_two_sequential_long_assigns() {
    assert_eq!(
        compile(
            "Attribute VB_Name = \"Module1\"\r\n\
             Sub Main()\r\n\
             Dim a As Long, b As Long\r\n\
             b = a\r\n\
             a = b\r\n\
             End Sub\r\n",
            0x0008,
        ),
        &[
            0x6c, 0x78, 0xff, 0x71, 0x74, 0xff,
            0x6c, 0x74, 0xff, 0x71, 0x78, 0xff,
        ]
    );
}

// ── Regression: bugs caught by the live oracle comparison (compare_oracle.py) ──
// Long Xor: bitwise ops carry the operand-promoted type (Long), not Boolean, so
// the back-end keys the opcode on RT_TYPE_OFFSET[Long] (fb 13, not fb 12).
#[test]
fn e2e_long_xor() {
    assert_eq!(
        compile(
            "Attribute VB_Name = \"Module1\"\r\n\
             Sub Main()\r\n\
             Dim a As Long, b As Long, r As Long\r\n\
             r = a Xor b\r\n\
             End Sub\r\n",
            0x0008,
        ),
        &[0x6c, 0x78, 0xff, 0x6c, 0x74, 0xff, 0xfb, 0x13, 0x71, 0x70, 0xff]
    );
}

// The four operators the front-end operator table (DAT_0faa5e10) assigns by
// precedence, all routed to the generic operation emitter: `\`=0x1e, Mod=0x1d,
// Eqv=0x20, Imp=0x1f. Long operands; only the operator byte differs.
#[test]
fn e2e_long_idiv() {
    assert_eq!(
        compile(
            "Attribute VB_Name = \"Module1\"\r\n\
             Sub Main()\r\n\
             Dim a As Long, b As Long, r As Long\r\n\
             r = a \\ b\r\n\
             End Sub\r\n",
            0x0008,
        ),
        &[0x6c, 0x78, 0xff, 0x6c, 0x74, 0xff, 0xc0, 0x71, 0x70, 0xff]
    );
}

#[test]
fn e2e_long_mod() {
    assert_eq!(
        compile(
            "Attribute VB_Name = \"Module1\"\r\n\
             Sub Main()\r\n\
             Dim a As Long, b As Long, r As Long\r\n\
             r = a Mod b\r\n\
             End Sub\r\n",
            0x0008,
        ),
        &[0x6c, 0x78, 0xff, 0x6c, 0x74, 0xff, 0xc2, 0x71, 0x70, 0xff]
    );
}

#[test]
fn e2e_long_eqv() {
    assert_eq!(
        compile(
            "Attribute VB_Name = \"Module1\"\r\n\
             Sub Main()\r\n\
             Dim a As Long, b As Long, r As Long\r\n\
             r = a Eqv b\r\n\
             End Sub\r\n",
            0x0008,
        ),
        &[0x6c, 0x78, 0xff, 0x6c, 0x74, 0xff, 0xfb, 0x0b, 0x71, 0x70, 0xff]
    );
}

#[test]
fn e2e_long_imp() {
    assert_eq!(
        compile(
            "Attribute VB_Name = \"Module1\"\r\n\
             Sub Main()\r\n\
             Dim a As Long, b As Long, r As Long\r\n\
             r = a Imp b\r\n\
             End Sub\r\n",
            0x0008,
        ),
        &[0x6c, 0x78, 0xff, 0x6c, 0x74, 0xff, 0xfb, 0x03, 0x71, 0x70, 0xff]
    );
}

// Mixed-type operand coercion: a narrower binary-op operand is widened to the
// operation type via the conversion opcode
// assign_store_base(target) + assign_source_adjust(src). Integer→Long emits 0xe7
// (0x11c+1); the wider operand and the add follow unchanged.
#[test]
fn e2e_mixed_int_long_add() {
    assert_eq!(
        compile(
            "Attribute VB_Name = \"Module1\"\r\n\
             Sub Main()\r\n\
             Dim a As Integer, b As Long, r As Long\r\n\
             r = a + b\r\n\
             End Sub\r\n",
            0x0008,
        ),
        &[0x6b, 0x7a, 0xff, 0xe7, 0x6c, 0x74, 0xff, 0xaa, 0x71, 0x70, 0xff]
    );
}

// Byte: 2-byte escape-paged load (fc e0) / store (fc f0) via the value-emitter
// index path; add via the generic emitter (tag 5 -> RT_OPCODE_BYTE[0x8e]=fb escape).
#[test]
fn e2e_byte_add() {
    assert_eq!(
        compile(
            "Attribute VB_Name = \"Module1\"\r\n\
             Sub Main()\r\n\
             Dim a As Byte, b As Byte, r As Byte\r\n\
             r = a + b\r\n\
             End Sub\r\n",
            0x0008,
        ),
        &[0xfc, 0xe0, 0x7a, 0xff, 0xfc, 0xe0, 0x78, 0xff, 0xfb, 0x8e, 0xfc, 0xf0,
          0x76, 0xff]
    );
}

// Byte widened to Long: the Byte operand loads (fc e0) then coerces to Long (e7).
#[test]
fn e2e_mixed_byte_long_add() {
    assert_eq!(
        compile(
            "Attribute VB_Name = \"Module1\"\r\n\
             Sub Main()\r\n\
             Dim a As Byte, b As Long, r As Long\r\n\
             r = a + b\r\n\
             End Sub\r\n",
            0x0008,
        ),
        &[0xfc, 0xe0, 0x7a, 0xff, 0xe7, 0x6c, 0x74, 0xff, 0xaa, 0x71, 0x70, 0xff]
    );
}

// Const folding: a Const local has no frame slot and is folded to its literal at
// each use site — `r = K` emits the literal 42 (f5 2a..) + store, not a load.
#[test]
fn e2e_const_fold() {
    assert_eq!(
        compile(
            "Attribute VB_Name = \"Module1\"\r\n\
             Sub Main()\r\n\
             Const K As Long = 42\r\n\
             Dim r As Long\r\n\
             r = K\r\n\
             End Sub\r\n",
            0x0008,
        ),
        &[0xf5, 0x2a, 0x00, 0x00, 0x00, 0x71, 0x78, 0xff]
    );
}

// Select Case: subject evaluated once into a hidden temp (store 0x71 to -0x8c),
// each Case loads the temp, compares `=` (0xc7), BranchFalse past its body to the
// next case; the matched body jumps (0x1e) to the end. Case Else falls through.
#[test]
fn e2e_select_case() {
    assert_eq!(
        compile(
            "Attribute VB_Name = \"Module1\"\r\n\
             Sub Main()\r\n\
             Dim a As Long\r\n\
             Select Case a\r\n\
             Case 1\r\n\
             a = 2\r\n\
             Case Else\r\n\
             a = 3\r\n\
             End Select\r\n\
             End Sub\r\n",
            0x0008,
        ),
        &[0x6c, 0x78, 0xff, 0x71, 0x74, 0xff, 0x6c, 0x74, 0xff, 0xf5, 0x01, 0x00,
          0x00, 0x00, 0xc7, 0x1c, 0x1d, 0x00, 0xf5, 0x02, 0x00, 0x00, 0x00, 0x71,
          0x78, 0xff, 0x1e, 0x25, 0x00, 0xf5, 0x03, 0x00, 0x00, 0x00, 0x71, 0x78,
          0xff]
    );
}

// Date is Double-backed: load 0x6f / store 0x74 (Double load/store opcodes).
#[test]
fn e2e_date_copy() {
    assert_eq!(
        compile(
            "Attribute VB_Name = \"Module1\"\r\n\
             Sub Main()\r\n\
             Dim a As Date, r As Date\r\n\
             r = a\r\n\
             End Sub\r\n",
            0x0008,
        ),
        &[0x6f, 0x74, 0xff, 0x74, 0x6c, 0xff]
    );
}

// GoTo: unconditional jump (0x1e) to the label's byte offset; the label emits
// nothing. Here `GoTo L` jumps to offset 3 (past the jump), where `a = 1` sits.
#[test]
fn e2e_goto_label() {
    assert_eq!(
        compile(
            "Attribute VB_Name = \"Module1\"\r\n\
             Sub Main()\r\n\
             Dim a As Long\r\n\
             GoTo L\r\n\
             L:\r\n\
             a = 1\r\n\
             End Sub\r\n",
            0x0008,
        ),
        &[0x1e, 0x03, 0x00, 0xf5, 0x01, 0x00, 0x00, 0x00, 0x71, 0x78, 0xff]
    );
}

// Exit For: jump (0x1e) to the loop-end offset (the same target the ForInit exit
// slot is patched to).
#[test]
fn e2e_exit_for() {
    assert_eq!(
        compile(
            "Attribute VB_Name = \"Module1\"\r\n\
             Sub Main()\r\n\
             Dim i As Long\r\n\
             For i = 1 To 9\r\n\
             Exit For\r\n\
             Next i\r\n\
             End Sub\r\n",
            0x0008,
        ),
        &[0xf5, 0x01, 0x00, 0x00, 0x00, 0x04, 0x78, 0xff, 0xf5, 0x09, 0x00, 0x00,
          0x00, 0xfe, 0x64, 0x70, 0xff, 0x1e, 0x00, 0x1e, 0x1e, 0x00, 0x04, 0x78,
          0xff, 0x66, 0x70, 0xff, 0x13, 0x00]
    );
}

// Do Until: VB6 compiles `Until cond` as `While Not cond` — the comparison
// negates (> becomes <=, opcode db->d6) and the exit branch is BranchFalse (1c),
// identical structure to Do While.
#[test]
fn e2e_do_until_negates() {
    assert_eq!(
        compile(
            "Attribute VB_Name = \"Module1\"\r\n\
             Sub Main()\r\n\
             Dim a As Long\r\n\
             Do Until a > 9\r\n\
             a = a + 1\r\n\
             Loop\r\n\
             End Sub\r\n",
            0x0008,
        ),
        &[0x6c, 0x78, 0xff, 0xf5, 0x09, 0x00, 0x00, 0x00, 0xd6, 0x1c, 0x1b, 0x00,
          0x6c, 0x78, 0xff, 0xf5, 0x01, 0x00, 0x00, 0x00, 0xaa, 0x71, 0x78, 0xff,
          0x1e, 0x00, 0x00]
    );
}

// Currency arithmetic: node tag 0x0d (grounded from the kind->VARTYPE table),
// add opcode 0xac. Load/store use the Currency frame class (6): 6d/72.
#[test]
fn e2e_currency_add() {
    assert_eq!(
        compile(
            "Attribute VB_Name = \"Module1\"\r\n\
             Sub Main()\r\n\
             Dim a As Currency, b As Currency, r As Currency\r\n\
             r = a + b\r\n\
             End Sub\r\n",
            0x0008,
        ),
        &[0x6d, 0x74, 0xff, 0x6d, 0x6c, 0xff, 0xac, 0x72, 0x64, 0xff]
    );
}

// Boolean is operated on as Integer (tag 6, Integer-class load/store): And = 0xc4.
#[test]
fn e2e_boolean_and() {
    assert_eq!(
        compile(
            "Attribute VB_Name = \"Module1\"\r\n\
             Sub Main()\r\n\
             Dim a As Boolean, b As Boolean, r As Boolean\r\n\
             r = a And b\r\n\
             End Sub\r\n",
            0x0008,
        ),
        &[0x6b, 0x7a, 0xff, 0x6b, 0x78, 0xff, 0xc4, 0x70, 0x76, 0xff]
    );
}

// Long widened to Currency (the wider operand): conversion opcode 0xf0, then the
// Currency add 0xac.
#[test]
fn e2e_mixed_long_currency_add() {
    assert_eq!(
        compile(
            "Attribute VB_Name = \"Module1\"\r\n\
             Sub Main()\r\n\
             Dim a As Long, b As Currency, r As Currency\r\n\
             r = a + b\r\n\
             End Sub\r\n",
            0x0008,
        ),
        &[0x6c, 0x78, 0xff, 0xf0, 0x6d, 0x70, 0xff, 0xac, 0x72, 0x68, 0xff]
    );
}

// Single→Double widening emits NO conversion opcode: a floating-point operand
// widened to a wider float is consumed directly by the operation (only
// integer-typed operands carry an explicit widening conversion). Load Single,
// load Double, add Double — no 0xed.
#[test]
fn e2e_mixed_single_double_no_coerce() {
    assert_eq!(
        compile(
            "Attribute VB_Name = \"Module1\"\r\n\
             Sub Main()\r\n\
             Dim a As Single, b As Double, r As Double\r\n\
             r = a + b\r\n\
             End Sub\r\n",
            0x0008,
        ),
        &[0x6e, 0x78, 0xff, 0x6f, 0x70, 0xff, 0xab, 0x74, 0x68, 0xff]
    );
}

// Long→Double widening emits 0xec (0x12c+2).
#[test]
fn e2e_mixed_long_double_add() {
    assert_eq!(
        compile(
            "Attribute VB_Name = \"Module1\"\r\n\
             Sub Main()\r\n\
             Dim a As Long, b As Double, r As Double\r\n\
             r = a + b\r\n\
             End Sub\r\n",
            0x0008,
        ),
        &[0x6c, 0x78, 0xff, 0xec, 0x6f, 0x70, 0xff, 0xab, 0x74, 0x68, 0xff]
    );
}

// Unary minus: emitted through the generic operation emitter as the single-operand
// op 7 (base 0x00c6); arithmetic dispatch → RT_TYPE_OFFSET[Long] selects 0xb8.
#[test]
fn e2e_long_negate() {
    assert_eq!(
        compile(
            "Attribute VB_Name = \"Module1\"\r\n\
             Sub Main()\r\n\
             Dim a As Long, r As Long\r\n\
             r = -a\r\n\
             End Sub\r\n",
            0x0008,
        ),
        &[0x6c, 0x78, 0xff, 0xb8, 0x71, 0x74, 0xff]
    );
}

// Unary Not: the single-operand op 6 (base 0x00be); arithmetic dispatch →
// RT_TYPE_OFFSET[Long] selects 0xc3.
#[test]
fn e2e_long_not() {
    assert_eq!(
        compile(
            "Attribute VB_Name = \"Module1\"\r\n\
             Sub Main()\r\n\
             Dim a As Long, r As Long\r\n\
             r = Not a\r\n\
             End Sub\r\n",
            0x0008,
        ),
        &[0x6c, 0x78, 0xff, 0xc3, 0x71, 0x74, 0xff]
    );
}

// Double division: the `/` operator's bound opcode is 0x19 (the arithmetic-block
// gap), previously unmapped → UnsupportedNode.
#[test]
fn e2e_double_div() {
    assert_eq!(
        compile(
            "Attribute VB_Name = \"Module1\"\r\n\
             Sub Main()\r\n\
             Dim a As Double, b As Double, r As Double\r\n\
             r = a / b\r\n\
             End Sub\r\n",
            0x0008,
        ),
        &[0x6f, 0x74, 0xff, 0x6f, 0x6c, 0xff, 0xb6, 0x74, 0x64, 0xff]
    );
}
