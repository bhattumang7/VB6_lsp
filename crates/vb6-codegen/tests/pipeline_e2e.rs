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

use vb6_codegen::{lower_module, lower_proc};
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

/// Like `conv`, but returns the lowering `Result` so a test can assert that an
/// unsupported construct is gated (an error) rather than mis-emitted.
fn try_compile(decl: &str, stmt: &str) -> Result<Vec<u8>, vb6_codegen::LowerError> {
    let src =
        format!("Attribute VB_Name = \"Module1\"\r\nSub Main()\r\n{decl}\r\n{stmt}\r\nEnd Sub\r\n");
    let mut ctx = ScannerContext::new(1, 1, 0x0409);
    ctx.intern_keywords();
    let mut arena = ExprArena::new();
    let mut parser = Parser::new(&mut ctx, src.as_bytes());
    let top = parser.parse_module(&mut arena);
    let spans = std::mem::take(&mut parser.node_spans);
    let vis = std::mem::take(&mut parser.decl_public);
    drop(parser);
    let module = bind(&ctx, &arena, &top, &spans, &vis);
    lower_proc(&module, 0, &arena, 0x0008)
}

/// Like `compile`, but lowers every procedure of the module sharing one
/// module-global string pool, returning each procedure's byte stream.
fn compile_module(src: &str, module_desc: u16) -> Vec<Vec<u8>> {
    let mut ctx = ScannerContext::new(1, 1, 0x0409);
    ctx.intern_keywords();
    let mut arena = ExprArena::new();
    let mut parser = Parser::new(&mut ctx, src.as_bytes());
    let top = parser.parse_module(&mut arena);
    let spans = std::mem::take(&mut parser.node_spans);
    let vis = std::mem::take(&mut parser.decl_public);
    drop(parser);
    let module = bind(&ctx, &arena, &top, &spans, &vis);
    lower_module(&module, &arena, module_desc)
        .unwrap_or_else(|e| panic!("lower_module failed: {e:?}"))
}

// ── Type-conversion intrinsics (CInt/CLng/CSng/CDbl/CCur/CStr) ───────────────
//
// Each converts its argument to the target type with the explicit-conversion
// opcode family (distinct from implicit assignment coercion). Confirmed against
// the VB6-compiled exe.

fn conv(decl: &str, stmt: &str) -> Vec<u8> {
    compile(
        &format!("Attribute VB_Name = \"Module1\"\r\nSub Main()\r\n{decl}\r\n{stmt}\r\nEnd Sub\r\n"),
        0x0008,
    )
}

#[test]
fn e2e_cint_from_double() {
    assert_eq!(
        conv("Dim a As Double, r As Integer", "r = CInt(a)"),
        &[0x6f, 0x74, 0xff, 0xe5, 0x70, 0x72, 0xff]
    );
}

#[test]
fn e2e_cint_from_integer_noop() {
    // CInt of an Integer is a no-op (no conversion opcode).
    assert_eq!(
        conv("Dim a As Integer, r As Integer", "r = CInt(a)"),
        &[0x6b, 0x7a, 0xff, 0x70, 0x78, 0xff]
    );
}

#[test]
fn e2e_clng_from_single() {
    assert_eq!(
        conv("Dim a As Single, r As Long", "r = CLng(a)"),
        &[0x6e, 0x78, 0xff, 0xe8, 0x71, 0x74, 0xff]
    );
}

#[test]
fn e2e_csng_from_long() {
    // The →Single explicit conversion uses the dedicated 0xfc 0x3e form.
    assert_eq!(
        conv("Dim a As Long, r As Single", "r = CSng(a)"),
        &[0x6c, 0x78, 0xff, 0xfc, 0x3e, 0x73, 0x74, 0xff]
    );
}

#[test]
fn e2e_cdbl_from_single() {
    assert_eq!(
        conv("Dim a As Single, r As Double", "r = CDbl(a)"),
        &[0x6e, 0x78, 0xff, 0xfc, 0x3a, 0x74, 0x70, 0xff]
    );
}

#[test]
fn e2e_ccur_from_long() {
    assert_eq!(
        conv("Dim a As Long, r As Currency", "r = CCur(a)"),
        &[0x6c, 0x78, 0xff, 0xf0, 0x72, 0x70, 0xff]
    );
}

#[test]
fn e2e_cstr_from_long_move_store() {
    // CStr yields a fresh string temp → move-store (0x31), not copy (0x43).
    assert_eq!(
        conv("Dim a As Long, s As String", "s = CStr(a)"),
        &[0x6c, 0x78, 0xff, 0xfb, 0xfe, 0x31, 0x74, 0xff]
    );
}

// ── String-result runtime intrinsics (Chr/Space) ─────────────────────────────

#[test]
fn e2e_space_string_result() {
    // s = Space(n): load n, LdAddr result temp, runtime call (0x0a), load result
    // (0x60), move to s (0x31), free temp (0x35).
    assert_eq!(
        conv("Dim n As Long, s As String", "s = Space(n)"),
        &[0x6c, 0x78, 0xff, 0x04, 0x64, 0xff, 0x0a, 0x00, 0x00, 0x08, 0x00,
          0x04, 0x64, 0xff, 0x60, 0x31, 0x74, 0xff, 0x35, 0x64, 0xff]
    );
}

#[test]
fn e2e_chr_from_integer() {
    // Chr's Integer argument is widened to Long (0xe7) before the runtime call.
    assert_eq!(
        conv("Dim n As Integer, s As String", "s = Chr(n)"),
        &[0x6b, 0x7a, 0xff, 0xe7, 0x04, 0x64, 0xff, 0x0a, 0x00, 0x00, 0x08, 0x00,
          0x04, 0x64, 0xff, 0x60, 0x31, 0x74, 0xff, 0x35, 0x64, 0xff]
    );
}

// ── String-input string-result intrinsics (UCase/LCase/Trim/LTrim/RTrim) ─────
//
// The String argument is copied into an input temp (0x4d) then the runtime call
// writes a second result temp, which is moved to the target and freed. The two
// 16-byte string temps sit at consecutive frame slots (input temp1 above result
// temp2). The caller-side byte stream is identical across the family.

#[test]
fn e2e_ucase_string_input() {
    // s = UCase(t): LdAddr t, copy to input temp (0x4d), 08 40, LdAddr result temp,
    // runtime call (0x0a, argbytes 8), load result (0x60), move to s, free temp.
    assert_eq!(
        conv("Dim t As String, s As String", "s = UCase(t)"),
        &[0x04, 0x78, 0xff, 0x4d, 0x64, 0xff, 0x08, 0x40, 0x04, 0x54, 0xff,
          0x0a, 0x00, 0x00, 0x08, 0x00, 0x04, 0x54, 0xff, 0x60, 0x31, 0x74, 0xff,
          0x35, 0x54, 0xff]
    );
}

#[test]
fn e2e_string_input_family_identical_caller_bytes() {
    // LCase/Trim/LTrim/RTrim emit the same caller-side stream as UCase (the
    // runtime function differs only in the gated proc-descriptor binding).
    let ucase = conv("Dim t As String, s As String", "s = UCase(t)");
    for f in ["LCase", "Trim", "LTrim", "RTrim"] {
        assert_eq!(
            conv("Dim t As String, s As String", &format!("s = {f}(t)")),
            ucase,
            "{f} caller bytes diverged from UCase"
        );
    }
}

// ── Multi-argument string-result intrinsics (Left/Right/Mid/String) ──────────
//
// Each argument is pushed right-to-left as either a by-value load (numeric) or a
// boxed 16-byte temp tagged with its VARTYPE (`04 <var> 4d <temp> <vt> 40`; 0x08
// for a String, 0x03 for a Long). A final result temp is passed by address; the
// runtime call's arg-byte count sums each push (boxed → 4) plus the result (4).

#[test]
fn e2e_left_string_and_length() {
    // Left(t, n): length n pushed by value (0x6c), string t boxed (0x4d … 08 40),
    // result temp, runtime call with arg-bytes 0x0c (n 4 + t-box 4 + result 4).
    assert_eq!(
        conv("Dim t As String, n As Long, s As String", "s = Left(t, n)"),
        &[0x6c, 0x74, 0xff, 0x04, 0x78, 0xff, 0x4d, 0x60, 0xff, 0x08, 0x40,
          0x04, 0x50, 0xff, 0x0a, 0x00, 0x00, 0x0c, 0x00, 0x04, 0x50, 0xff,
          0x60, 0x31, 0x70, 0xff, 0x35, 0x50, 0xff]
    );
}

#[test]
fn e2e_right_matches_left_caller_bytes() {
    // Right shares Left's caller-side stream (signature-identical).
    assert_eq!(
        conv("Dim t As String, n As Long, s As String", "s = Right(t, n)"),
        conv("Dim t As String, n As Long, s As String", "s = Left(t, n)")
    );
}

#[test]
fn e2e_mid_three_args() {
    // Mid(t, n, m): the optional length m is a Variant boxed as VT_I4 (0x03 0x40);
    // the start n is a by-value Long; t is boxed VT_BSTR. Arg-bytes 0x10.
    assert_eq!(
        conv("Dim t As String, n As Long, m As Long, s As String", "s = Mid(t, n, m)"),
        &[0x04, 0x70, 0xff, 0x4d, 0x4c, 0xff, 0x03, 0x40, 0x6c, 0x74, 0xff,
          0x04, 0x78, 0xff, 0x4d, 0x5c, 0xff, 0x08, 0x40, 0x04, 0x3c, 0xff,
          0x0a, 0x00, 0x00, 0x10, 0x00, 0x04, 0x3c, 0xff, 0x60, 0x31, 0x6c, 0xff,
          0x35, 0x3c, 0xff]
    );
}

#[test]
fn e2e_mid_two_args_omitted_optional() {
    // Mid with the optional length omitted: the missing parameter is a hidden
    // Missing variant (pushed by address, 0x27), and the cleanup frees both that
    // variant temp and the result temp with the combined free (0x36 <count*2> …).
    // The Missing literal reserves a value-buffer slot, so its temp (0x40) sits one
    // 16-byte slot below the boxed string temp (0x60).
    assert_eq!(
        conv("Dim t As String, n As Long, s As String", "s = Mid(t, n)"),
        &[0x27, 0x40, 0xff, 0x6c, 0x74, 0xff, 0x04, 0x78, 0xff, 0x4d, 0x60, 0xff,
          0x08, 0x40, 0x04, 0x30, 0xff, 0x0a, 0x00, 0x00, 0x10, 0x00, 0x04, 0x30,
          0xff, 0x60, 0x31, 0x70, 0xff, 0x36, 0x04, 0x00, 0x40, 0xff, 0x30, 0xff]
    );
}

#[test]
fn e2e_mid_two_args_literal_start() {
    // Same omitted-optional form with a literal start position (pushed as a Long
    // literal, 0xf5).
    assert_eq!(
        conv("Dim t As String, s As String", "s = Mid(t, 2)"),
        &[0x27, 0x44, 0xff, 0xf5, 0x02, 0x00, 0x00, 0x00, 0x04, 0x78, 0xff, 0x4d,
          0x64, 0xff, 0x08, 0x40, 0x04, 0x34, 0xff, 0x0a, 0x00, 0x00, 0x10, 0x00,
          0x04, 0x34, 0xff, 0x60, 0x31, 0x74, 0xff, 0x36, 0x04, 0x00, 0x44, 0xff,
          0x34, 0xff]
    );
}

// ── InStr (dedicated opcode 0xfe 0xfd) ───────────────────────────────────────
//
// Four operands push in order — start (Long), string1, string2, compare-mode
// (Long) — then the InStr opcode, leaving a Long on the stack. An omitted leading
// start defaults to literal 1; the compare-mode is literal 0 (Option Compare
// Binary). The result composes in expressions (the opcode takes no ref operand).

#[test]
fn e2e_instr_two_args() {
    // InStr(a, b): start defaults to 1 (0xf5 imm), a, b, compare-mode 0, then fe fd.
    assert_eq!(
        conv("Dim a As String, b As String, r As Long", "r = InStr(a, b)"),
        &[0xf5, 0x01, 0x00, 0x00, 0x00, 0x6c, 0x78, 0xff, 0x6c, 0x74, 0xff,
          0xf5, 0x00, 0x00, 0x00, 0x00, 0xfe, 0xfd, 0x71, 0x70, 0xff]
    );
}

#[test]
fn e2e_instr_three_args() {
    // InStr(n, a, b): explicit start n loaded by value; compare-mode still 0.
    assert_eq!(
        conv("Dim a As String, b As String, n As Long, r As Long", "r = InStr(n, a, b)"),
        &[0x6c, 0x70, 0xff, 0x6c, 0x78, 0xff, 0x6c, 0x74, 0xff,
          0xf5, 0x00, 0x00, 0x00, 0x00, 0xfe, 0xfd, 0x71, 0x6c, 0xff]
    );
}

#[test]
fn e2e_instr_literal_operand() {
    // A string-literal search operand loads via the pool reference (0x1b).
    assert_eq!(
        conv("Dim a As String, r As Long", "r = InStr(a, \"x\")"),
        &[0xf5, 0x01, 0x00, 0x00, 0x00, 0x6c, 0x78, 0xff, 0x1b, 0x00, 0x00,
          0xf5, 0x00, 0x00, 0x00, 0x00, 0xfe, 0xfd, 0x71, 0x74, 0xff]
    );
}

#[test]
fn e2e_instr_in_expression() {
    // InStr leaves its Long result on the stack, so it composes with a following
    // operator (`+ 1` → push 1, add 0xaa) before the store.
    assert_eq!(
        conv("Dim a As String, b As String, r As Long", "r = InStr(a, b) + 1"),
        &[0xf5, 0x01, 0x00, 0x00, 0x00, 0x6c, 0x78, 0xff, 0x6c, 0x74, 0xff,
          0xf5, 0x00, 0x00, 0x00, 0x00, 0xfe, 0xfd, 0xf5, 0x01, 0x00, 0x00, 0x00,
          0xaa, 0x71, 0x70, 0xff]
    );
}

#[test]
fn e2e_instr_four_args_gated() {
    // An explicit compare-mode argument (4-arg form) needs Option Compare handling
    // → gated, not mis-emitted.
    assert!(try_compile(
        "Dim a As String, b As String, r As Long",
        "r = InStr(1, a, b, 1)"
    )
    .is_err());
}

#[test]
fn e2e_string_number_and_char() {
    // String(n, t): number n by value (Long), character t boxed (VT_BSTR).
    assert_eq!(
        conv("Dim n As Long, t As String, s As String", "s = String(n, t)"),
        &[0x04, 0x74, 0xff, 0x4d, 0x60, 0xff, 0x08, 0x40, 0x6c, 0x78, 0xff,
          0x04, 0x50, 0xff, 0x0a, 0x00, 0x00, 0x0c, 0x00, 0x04, 0x50, 0xff,
          0x60, 0x31, 0x70, 0xff, 0x35, 0x50, 0xff]
    );
}

// ── Number→String boxed-Variant intrinsics (Str/Hex/Oct) ─────────────────────

#[test]
fn e2e_str_boxes_number_as_variant() {
    // Str(n): the Long argument is boxed as a Variant (VT_I4 = 0x03), result temp,
    // runtime call arg-bytes 0x08.
    assert_eq!(
        conv("Dim n As Long, s As String", "s = Str(n)"),
        &[0x04, 0x78, 0xff, 0x4d, 0x64, 0xff, 0x03, 0x40, 0x04, 0x54, 0xff,
          0x0a, 0x00, 0x00, 0x08, 0x00, 0x04, 0x54, 0xff, 0x60, 0x31, 0x74, 0xff,
          0x35, 0x54, 0xff]
    );
}

#[test]
fn e2e_hex_oct_match_str_caller_bytes() {
    let str_bytes = conv("Dim n As Long, s As String", "s = Str(n)");
    for f in ["Hex", "Oct"] {
        assert_eq!(
            conv("Dim n As Long, s As String", &format!("s = {f}(n)")),
            str_bytes,
            "{f} caller bytes diverged from Str"
        );
    }
}

// ── Numeric-result runtime-library intrinsics (Asc/Sqr/Val) ──────────────────

#[test]
fn e2e_asc_runtime_call() {
    // Asc → Integer: size-load the String arg, runtime call opcode 0x0b.
    assert_eq!(
        conv("Dim s As String, r As Integer", "r = Asc(s)"),
        &[0x6c, 0x78, 0xff, 0x0b, 0x00, 0x00, 0x04, 0x00, 0x70, 0x76, 0xff]
    );
}

#[test]
fn e2e_sqr_runtime_call() {
    // Sqr → Double: size-8 load the Double arg, runtime call opcode 0x0a.
    assert_eq!(
        conv("Dim d As Double, r As Double", "r = Sqr(d)"),
        &[0x6d, 0x74, 0xff, 0x0a, 0x00, 0x00, 0x08, 0x00, 0x74, 0x6c, 0xff]
    );
}

#[test]
fn e2e_val_runtime_call() {
    assert_eq!(
        conv("Dim s As String, r As Double", "r = Val(s)"),
        &[0x6c, 0x78, 0xff, 0x0a, 0x00, 0x00, 0x04, 0x00, 0x74, 0x70, 0xff]
    );
}

// ── Dedicated-opcode unary intrinsics (Len/Abs/Sgn/Int/Fix) ──────────────────

#[test]
fn e2e_len_string() {
    assert_eq!(conv("Dim s As String, r As Long", "r = Len(s)"), &[0x6c, 0x78, 0xff, 0x4a, 0x71, 0x74, 0xff]);
}

#[test]
fn e2e_abs_by_type() {
    // Abs returns its argument type; opcode is type-indexed (0xbb/0xbc/0xbd).
    assert_eq!(conv("Dim a As Integer, r As Integer", "r = Abs(a)"), &[0x6b, 0x7a, 0xff, 0xbb, 0x70, 0x78, 0xff]);
    assert_eq!(conv("Dim a As Long, r As Long", "r = Abs(a)"), &[0x6c, 0x78, 0xff, 0xbc, 0x71, 0x74, 0xff]);
    assert_eq!(conv("Dim a As Double, r As Double", "r = Abs(a)"), &[0x6f, 0x74, 0xff, 0xbd, 0x74, 0x6c, 0xff]);
}

#[test]
fn e2e_sgn_returns_integer() {
    // Sgn → Integer; opcode indexed by argument type (0xfb 0xf4 for Long).
    assert_eq!(conv("Dim a As Long, r As Integer", "r = Sgn(a)"), &[0x6c, 0x78, 0xff, 0xfb, 0xf4, 0x70, 0x76, 0xff]);
}

#[test]
fn e2e_int_fix_float_only() {
    // Int/Fix are no-ops on integral args; floats use 0xfb 0xe7 / 0xfb 0xdf.
    assert_eq!(conv("Dim a As Long, r As Long", "r = Int(a)"), &[0x6c, 0x78, 0xff, 0x71, 0x74, 0xff]);
    assert_eq!(conv("Dim a As Double, r As Double", "r = Int(a)"), &[0x6f, 0x74, 0xff, 0xfb, 0xe7, 0x74, 0x6c, 0xff]);
    assert_eq!(conv("Dim a As Double, r As Double", "r = Fix(a)"), &[0x6f, 0x74, 0xff, 0xfb, 0xdf, 0x74, 0x6c, 0xff]);
}

// ── Intra-module Sub/Function calls ──────────────────────────────────────────
//
// The caller (Sub Main = proc 0) pushes each argument (ByVal value / ByRef
// address) then the call opcode: Sub call 0x0a, Function call 0x5e, operand =
// <call-site index><total arg bytes>. Confirmed byte-for-byte against the exe.

/// Build a module whose Sub Main contains `main_body`, plus the given callee
/// procedures, and return Sub Main's (proc 0) byte stream.
fn caller_bytes(main_body: &str, callees: &str) -> Vec<u8> {
    let src = format!(
        "Attribute VB_Name = \"Module1\"\r\nSub Main()\r\n{main_body}End Sub\r\n{callees}"
    );
    compile_module(&src, 0x0008).remove(0)
}

#[test]
fn e2e_call_sub_no_args() {
    assert_eq!(
        caller_bytes("    Call Foo\r\n", "Sub Foo()\r\n    Dim z As Long\r\n    z = 1\r\nEnd Sub\r\n"),
        &[0x0a, 0x00, 0x00, 0x00, 0x00]
    );
}

#[test]
fn e2e_call_sub_byval_literal() {
    assert_eq!(
        caller_bytes(
            "    Call Foo(5)\r\n",
            "Sub Foo(ByVal x As Long)\r\n    Dim z As Long\r\n    z = x\r\nEnd Sub\r\n"
        ),
        &[0xf5, 0x05, 0x00, 0x00, 0x00, 0x0a, 0x00, 0x00, 0x04, 0x00]
    );
}

#[test]
fn e2e_call_sub_byref_var() {
    // `Dim v; v = 3; Call Foo(v)` with Foo(x As Long) ByRef → LdAddr v before call.
    assert_eq!(
        caller_bytes(
            "    Dim v As Long\r\n    v = 3\r\n    Call Foo(v)\r\n",
            "Sub Foo(x As Long)\r\n    Dim z As Long\r\n    z = x\r\nEnd Sub\r\n"
        ),
        &[0xf5, 0x03, 0x00, 0x00, 0x00, 0x71, 0x78, 0xff, 0x04, 0x78, 0xff, 0x0a, 0x00, 0x00, 0x04, 0x00]
    );
}

#[test]
fn e2e_call_sub_byval_var() {
    assert_eq!(
        caller_bytes(
            "    Dim v As Long\r\n    v = 3\r\n    Call Foo(v)\r\n",
            "Sub Foo(ByVal x As Long)\r\n    Dim z As Long\r\n    z = x\r\nEnd Sub\r\n"
        ),
        &[0xf5, 0x03, 0x00, 0x00, 0x00, 0x71, 0x78, 0xff, 0x6c, 0x78, 0xff, 0x0a, 0x00, 0x00, 0x04, 0x00]
    );
}

#[test]
fn e2e_call_two_sites_indexed() {
    // Two call sites get sequential call-site indices 0 then 1.
    assert_eq!(
        caller_bytes(
            "    Call Foo\r\n    Call Bar\r\n",
            "Sub Foo()\r\n    Dim z As Long\r\n    z = 1\r\nEnd Sub\r\n\
             Sub Bar()\r\n    Dim z As Long\r\n    z = 2\r\nEnd Sub\r\n"
        ),
        &[0x0a, 0x00, 0x00, 0x00, 0x00, 0x0a, 0x01, 0x00, 0x00, 0x00]
    );
}

#[test]
fn e2e_call_function_in_expression() {
    // `r = F() + 1`: the Function call (0x5e) leaves its result on the stack, then
    // the `+ 1` consumes it.
    assert_eq!(
        caller_bytes("    Dim r As Long\r\n    r = F() + 1\r\n", "Function F() As Long\r\n    F = 7\r\nEnd Function\r\n"),
        &[0x5e, 0x00, 0x00, 0x00, 0x00, 0xf5, 0x01, 0x00, 0x00, 0x00, 0xaa, 0x71, 0x78, 0xff]
    );
}

#[test]
fn e2e_call_function_as_argument() {
    // `Call Foo(F())`: F() (reference index 0) is evaluated as Foo's argument; the
    // outer Sub call takes reference index 1.
    assert_eq!(
        caller_bytes(
            "    Call Foo(F())\r\n",
            "Function F() As Long\r\n    F = 7\r\nEnd Function\r\n\
             Sub Foo(ByVal x As Long)\r\n    Dim z As Long\r\n    z = x\r\nEnd Sub\r\n"
        ),
        &[0x5e, 0x00, 0x00, 0x00, 0x00, 0x0a, 0x01, 0x00, 0x04, 0x00]
    );
}

#[test]
fn e2e_call_two_args_reverse_order() {
    // Arguments are pushed right-to-left: `Foo(5, 6)` pushes 6 then 5; arg-bytes 8.
    assert_eq!(
        caller_bytes(
            "    Call Foo(5, 6)\r\n",
            "Sub Foo(ByVal a As Long, ByVal b As Long)\r\n    Dim z As Long\r\n    z = a\r\nEnd Sub\r\n"
        ),
        &[0xf5, 0x06, 0x00, 0x00, 0x00, 0xf5, 0x05, 0x00, 0x00, 0x00, 0x0a, 0x00, 0x00, 0x08, 0x00]
    );
}

#[test]
fn e2e_call_arg_coercion() {
    // An Integer argument to a Long parameter is widened (0xe7) before the call.
    assert_eq!(
        caller_bytes(
            "    Dim i As Integer\r\n    Call Foo(i)\r\n",
            "Sub Foo(ByVal x As Long)\r\n    Dim z As Long\r\n    z = x\r\nEnd Sub\r\n"
        ),
        &[0x6b, 0x7a, 0xff, 0xe7, 0x0a, 0x00, 0x00, 0x04, 0x00]
    );
}

#[test]
fn e2e_call_bare_implicit_no_args() {
    // A bare `Foo` statement (no `Call`, no parens) is an implicit Sub call.
    assert_eq!(
        caller_bytes("    Foo\r\n", "Sub Foo()\r\n    Dim z As Long\r\n    z = 1\r\nEnd Sub\r\n"),
        &[0x0a, 0x00, 0x00, 0x00, 0x00]
    );
}

#[test]
fn e2e_call_byval_double_sized_load() {
    // A same-typed ByVal Double argument is pushed with the size-8 value load
    // (0x6d), not the Double type load (0x6f); arg-bytes 8.
    assert_eq!(
        caller_bytes(
            "    Dim a As Double\r\n    Call Foo(a)\r\n",
            "Sub Foo(ByVal d As Double)\r\n    Dim z As Double\r\n    z = d\r\nEnd Sub\r\n"
        ),
        &[0x6d, 0x74, 0xff, 0x0a, 0x00, 0x00, 0x08, 0x00]
    );
}

#[test]
fn e2e_call_byval_integer_argbytes_padded() {
    // A ByVal Integer argument loads with 0x6b but its arg-bytes round up to 4.
    assert_eq!(
        caller_bytes(
            "    Dim a As Integer\r\n    Call Foo(a)\r\n",
            "Sub Foo(ByVal x As Integer)\r\n    Dim z As Integer\r\n    z = x\r\nEnd Sub\r\n"
        ),
        &[0x6b, 0x7a, 0xff, 0x0a, 0x00, 0x00, 0x04, 0x00]
    );
}

#[test]
fn e2e_call_byval_string_arg() {
    assert_eq!(
        caller_bytes(
            "    Dim a As String\r\n    Call Foo(a)\r\n",
            "Sub Foo(ByVal x As String)\r\n    Dim z As String\r\n    z = x\r\nEnd Sub\r\n"
        ),
        &[0x6c, 0x78, 0xff, 0x0a, 0x00, 0x00, 0x04, 0x00]
    );
}

#[test]
fn e2e_call_byref_integer_arg() {
    // ByRef pushes the address (4-byte pointer) regardless of element type.
    assert_eq!(
        caller_bytes(
            "    Dim a As Integer\r\n    Call Foo(a)\r\n",
            "Sub Foo(x As Integer)\r\n    Dim z As Integer\r\n    z = x\r\nEnd Sub\r\n"
        ),
        &[0x04, 0x7a, 0xff, 0x0a, 0x00, 0x00, 0x04, 0x00]
    );
}

#[test]
fn e2e_call_byref_double_arg() {
    assert_eq!(
        caller_bytes(
            "    Dim a As Double\r\n    Call Foo(a)\r\n",
            "Sub Foo(x As Double)\r\n    Dim z As Double\r\n    z = x\r\nEnd Sub\r\n"
        ),
        &[0x04, 0x74, 0xff, 0x0a, 0x00, 0x00, 0x04, 0x00]
    );
}

#[test]
fn e2e_call_discarded_function_result() {
    // `Call F(5)` discards the result → statement-form 0x0a (not 0x5e).
    assert_eq!(
        caller_bytes(
            "    Call F(5)\r\n",
            "Function F(ByVal x As Long) As Long\r\n    F = x\r\nEnd Function\r\n"
        ),
        &[0xf5, 0x05, 0x00, 0x00, 0x00, 0x0a, 0x00, 0x00, 0x04, 0x00]
    );
}

#[test]
fn e2e_call_string_arg_reference_index() {
    // A string-literal argument consumes a per-proc reference slot, so the call's
    // own reference index is 1 (the string took slot 0): `1b 00` then `0a 01 …`.
    assert_eq!(
        caller_bytes(
            "    Call Foo(\"x\")\r\n",
            "Sub Foo(ByVal s As String)\r\n    Dim z As String\r\n    z = s\r\nEnd Sub\r\n"
        ),
        &[0x1b, 0x00, 0x00, 0x0a, 0x01, 0x00, 0x04, 0x00]
    );
}

#[test]
fn e2e_call_function_result_no_args() {
    // `r = F()` → Function call 0x5e (result on stack) then store to r.
    assert_eq!(
        caller_bytes("    Dim r As Long\r\n    r = F()\r\n", "Function F() As Long\r\n    F = 7\r\nEnd Function\r\n"),
        &[0x5e, 0x00, 0x00, 0x00, 0x00, 0x71, 0x78, 0xff]
    );
}

#[test]
fn e2e_call_function_result_byval_arg() {
    assert_eq!(
        caller_bytes(
            "    Dim r As Long\r\n    r = F(5)\r\n",
            "Function F(ByVal x As Long) As Long\r\n    F = x\r\nEnd Function\r\n"
        ),
        &[0xf5, 0x05, 0x00, 0x00, 0x00, 0x5e, 0x00, 0x00, 0x04, 0x00, 0x71, 0x78, 0xff]
    );
}

// ── Multi-procedure modules: string pool is module-global ───────────────────
//
// String-literal indices are assigned across the whole module in procedure
// declaration order (deduped by value), not reset per procedure. Confirmed
// byte-for-byte against the VB6-compiled module exe.

#[test]
fn e2e_module_global_string_pool() {
    // Two procedures, each with a distinct string literal: proc 0's "aaa" gets
    // pool index 0 (`1b 00`), proc 1's "bbb" gets index 1 (`1b 01`).
    let procs = compile_module(
        "Attribute VB_Name = \"Module1\"\r\n\
         Sub Main()\r\n\
         Dim s As String\r\n\
         s = \"aaa\"\r\n\
         End Sub\r\n\
         Sub Two()\r\n\
         Dim t As String\r\n\
         t = \"bbb\"\r\n\
         End Sub\r\n",
        0x0008,
    );
    assert_eq!(procs.len(), 2);
    assert_eq!(procs[0], &[0x1b, 0x00, 0x00, 0x43, 0x78, 0xff]);
    assert_eq!(procs[1], &[0x1b, 0x01, 0x00, 0x43, 0x78, 0xff]);
}

// ── Multi-procedure: parameter types, Property Let, cross-proc globals ───────

/// Build a two-procedure module (Sub Main + the given second proc lines) and
/// return the lowered byte stream of the second procedure.
fn second_proc(second: &str) -> Vec<u8> {
    let src = format!(
        "Attribute VB_Name = \"Module1\"\r\n\
         Sub Main()\r\n    Dim a As Long\r\n    a = 1\r\nEnd Sub\r\n{second}"
    );
    compile_module(&src, 0x0008).remove(1)
}

#[test]
fn e2e_mp_param_string() {
    // ByVal String param load + copy-store.
    assert_eq!(
        second_proc("Sub Foo(ByVal s As String)\r\n    Dim t As String\r\n    t = s\r\nEnd Sub\r\n"),
        &[0x6c, 0x0c, 0x00, 0x43, 0x78, 0xff]
    );
}

#[test]
fn e2e_mp_param_double() {
    assert_eq!(
        second_proc("Sub Foo(ByVal d As Double)\r\n    Dim e As Double\r\n    e = d\r\nEnd Sub\r\n"),
        &[0x6f, 0x0c, 0x00, 0x74, 0x74, 0xff]
    );
}

#[test]
fn e2e_mp_param_byref_integer() {
    assert_eq!(
        second_proc("Sub Foo(x As Integer)\r\n    Dim y As Integer\r\n    y = x\r\nEnd Sub\r\n"),
        &[0x7f, 0x0c, 0x00, 0x70, 0x7a, 0xff]
    );
}

#[test]
fn e2e_mp_property_let() {
    assert_eq!(
        second_proc("Property Let P(ByVal v As Long)\r\n    Dim y As Long\r\n    y = v\r\nEnd Property\r\n"),
        &[0x6c, 0x0c, 0x00, 0x71, 0x78, 0xff]
    );
}

#[test]
fn e2e_mp_global_from_two_procs() {
    // A module-level global written from two procedures: same global store
    // (0x99 <module_desc> <field_offset>) in each.
    let procs = compile_module(
        "Attribute VB_Name = \"Module1\"\r\n\
         Dim g As Long\r\n\
         Sub Main()\r\n    g = 1\r\nEnd Sub\r\n\
         Sub Foo()\r\n    g = 2\r\nEnd Sub\r\n",
        0x0008,
    );
    assert_eq!(procs[0], &[0xf5, 0x01, 0x00, 0x00, 0x00, 0x99, 0x08, 0x00, 0x00, 0x00]);
    assert_eq!(procs[1], &[0xf5, 0x02, 0x00, 0x00, 0x00, 0x99, 0x08, 0x00, 0x00, 0x00]);
}

// ── Function return values (the proc name is an implicit slot-0 local) ───────

#[test]
fn e2e_function_return_long() {
    // `F = 5` stores into the return slot (frame slot 0 = 0xff78), like a normal
    // Long local store.
    let procs = compile_module(
        "Attribute VB_Name = \"Module1\"\r\n\
         Sub Main()\r\n\
         Dim a As Long\r\n\
         a = 1\r\n\
         End Sub\r\n\
         Function F() As Long\r\n\
         F = 5\r\n\
         End Function\r\n",
        0x0008,
    );
    assert_eq!(procs[1], &[0xf5, 0x05, 0x00, 0x00, 0x00, 0x71, 0x78, 0xff]);
}

#[test]
fn e2e_function_return_with_local() {
    // The return value takes slot 0 (0xff78); the user local `y` takes slot 1
    // (0xff74). `y = 7; F = y`.
    let procs = compile_module(
        "Attribute VB_Name = \"Module1\"\r\n\
         Sub Main()\r\n\
         Dim a As Long\r\n\
         a = 1\r\n\
         End Sub\r\n\
         Function F() As Long\r\n\
         Dim y As Long\r\n\
         y = 7\r\n\
         F = y\r\n\
         End Function\r\n",
        0x0008,
    );
    assert_eq!(
        procs[1],
        &[0xf5, 0x07, 0x00, 0x00, 0x00, 0x71, 0x74, 0xff, 0x6c, 0x74, 0xff, 0x71, 0x78, 0xff]
    );
}

#[test]
fn e2e_function_return_from_param() {
    // `Function F(ByVal x As Long) As Long: F = x` — load the ByVal param, store
    // into the return slot.
    let procs = compile_module(
        "Attribute VB_Name = \"Module1\"\r\n\
         Sub Main()\r\n\
         Dim a As Long\r\n\
         a = 1\r\n\
         End Sub\r\n\
         Function F(ByVal x As Long) As Long\r\n\
         F = x\r\n\
         End Function\r\n",
        0x0008,
    );
    assert_eq!(procs[1], &[0x6c, 0x0c, 0x00, 0x71, 0x78, 0xff]);
}

#[test]
fn e2e_module_global_static_block() {
    // The per-procedure static block is module-global: proc 0's Static `a` (Long)
    // takes block offset 0, proc 1's Static `b` takes offset 4 (after a's 4 bytes).
    let procs = compile_module(
        "Attribute VB_Name = \"Module1\"\r\n\
         Sub Main()\r\n\
         Static a As Long\r\n\
         a = 1\r\n\
         End Sub\r\n\
         Sub Foo()\r\n\
         Static b As Long\r\n\
         b = 2\r\n\
         End Sub\r\n",
        0x0008,
    );
    assert_eq!(procs.len(), 2);
    assert_eq!(procs[0], &[0xf5, 0x01, 0x00, 0x00, 0x00, 0x5f, 0x08, 0x00, 0x04, 0x00, 0x8f, 0x00, 0x00]);
    assert_eq!(procs[1], &[0xf5, 0x02, 0x00, 0x00, 0x00, 0x5f, 0x08, 0x00, 0x04, 0x00, 0x8f, 0x04, 0x00]);
}

#[test]
fn e2e_module_global_string_pool_dedup() {
    // A repeated string value is interned once across the module: proc 1's "aaa"
    // reuses index 0; its new "ccc" gets index 1.
    let procs = compile_module(
        "Attribute VB_Name = \"Module1\"\r\n\
         Sub Main()\r\n\
         Dim s As String\r\n\
         s = \"aaa\"\r\n\
         End Sub\r\n\
         Sub Two()\r\n\
         Dim t As String, u As String\r\n\
         t = \"aaa\"\r\n\
         u = \"ccc\"\r\n\
         End Sub\r\n",
        0x0008,
    );
    assert_eq!(procs.len(), 2);
    // proc 1 first stores "aaa" (reused index 0) then "ccc" (new index 1).
    assert_eq!(&procs[1][0..3], &[0x1b, 0x00, 0x00]);
    assert!(procs[1].windows(3).any(|w| w == [0x1b, 0x01, 0x00]));
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

// String literal: interns into the module string pool (index 0 here) and pushes
// it via 0x1b + 2-byte pool index, then copy-stores to s (0x43).
#[test]
fn e2e_string_literal() {
    assert_eq!(
        compile(
            "Attribute VB_Name = \"Module1\"\r\n\
             Sub Main()\r\n\
             Dim s As String\r\n\
             s = \"hi\"\r\n\
             End Sub\r\n",
            0x0008,
        ),
        &[0x1b, 0x00, 0x00, 0x43, 0x78, 0xff]
    );
}

// Fixed-length string source copy: `As String * 8` is a 16-byte inline Unicode
// buffer (2 bytes/char); `s = a` reads it length-aware (LdAddr a + 0x33<8>) and
// moves the BSTR temp into s (0x31).
#[test]
fn e2e_fixed_string_copy() {
    assert_eq!(
        compile(
            "Attribute VB_Name = \"Module1\"\r\n\
             Sub Main()\r\n\
             Dim a As String * 8, s As String\r\n\
             s = a\r\n\
             End Sub\r\n",
            0x0008,
        ),
        &[0x04, 0x6c, 0xff, 0x33, 0x08, 0x00, 0x31, 0x68, 0xff]
    );
}

// String copy: a String var loads its BSTR pointer (0x6c) and stores via the
// refcounted assign opcode (0x43).
#[test]
fn e2e_string_copy() {
    assert_eq!(
        compile(
            "Attribute VB_Name = \"Module1\"\r\n\
             Sub Main()\r\n\
             Dim a As String, s As String\r\n\
             s = a\r\n\
             End Sub\r\n",
            0x0008,
        ),
        &[0x6c, 0x78, 0xff, 0x43, 0x74, 0xff]
    );
}

// String comparison: both operands load as BSTR pointers (0x6c); the `=` compare
// emits the string-compare opcode (fb 30) and stores the Integer result.
#[test]
fn e2e_string_compare() {
    assert_eq!(
        compile(
            "Attribute VB_Name = \"Module1\"\r\n\
             Sub Main()\r\n\
             Dim a As String, b As String, r As Integer\r\n\
             r = (a = b)\r\n\
             End Sub\r\n",
            0x0008,
        ),
        &[0x6c, 0x78, 0xff, 0x6c, 0x74, 0xff, 0xfb, 0x30, 0x70, 0x72, 0xff]
    );
}

// ── String relational operators (Option Compare Binary) ─────────────────────
//
// Both operands load as 4-byte BSTR pointers (0x6c / string literal 0x1b); the
// comparison emits a dedicated two-byte string-compare opcode (0xfb <op>) whose
// second byte is selected by the operator, leaving an Integer (Boolean) result.

#[test]
fn e2e_string_compare_all_operators() {
    let decl = "Dim a As String, b As String, r As Integer";
    // load a (0x78), load b (0x74), fb <op>, store to r (0x72).
    for (stmt, op) in [
        ("r = (a = b)", 0x30u8),
        ("r = (a <> b)", 0x3d),
        ("r = (a < b)", 0x64),
        ("r = (a > b)", 0x71),
        ("r = (a <= b)", 0x4a),
        ("r = (a >= b)", 0x57),
    ] {
        assert_eq!(
            conv(decl, stmt),
            &[0x6c, 0x78, 0xff, 0x6c, 0x74, 0xff, 0xfb, op, 0x70, 0x72, 0xff],
            "string compare {stmt}"
        );
    }
}

#[test]
fn e2e_string_compare_literal_operand() {
    // A string-literal operand loads via 0x1b <pool-ref>; the compare is unchanged.
    assert_eq!(
        conv("Dim a As String, r As Integer", "r = (a = \"x\")"),
        &[0x6c, 0x78, 0xff, 0x1b, 0x00, 0x00, 0xfb, 0x30, 0x70, 0x76, 0xff]
    );
    assert_eq!(
        conv("Dim a As String, r As Integer", "r = (\"x\" = a)"),
        &[0x1b, 0x00, 0x00, 0x6c, 0x78, 0xff, 0xfb, 0x30, 0x70, 0x76, 0xff]
    );
}

#[test]
fn e2e_string_compare_in_condition() {
    // As an `If` condition the same compare opcode feeds the branch-false (0x1c).
    assert_eq!(
        conv("Dim a As String, b As String", "If a < b Then\r\na = b\r\nEnd If"),
        &[0x6c, 0x78, 0xff, 0x6c, 0x74, 0xff, 0xfb, 0x64, 0x1c, 0x11, 0x00,
          0x6c, 0x74, 0xff, 0x43, 0x78, 0xff]
    );
}

// String concat (`&`): node 0x24 with String tag emits the concat opcode (0x2a);
// the fresh temp result is moved into s via 0x31 (not the copy store 0x43).
#[test]
fn e2e_string_concat() {
    assert_eq!(
        compile(
            "Attribute VB_Name = \"Module1\"\r\n\
             Sub Main()\r\n\
             Dim a As String, b As String, s As String\r\n\
             s = a & b\r\n\
             End Sub\r\n",
            0x0008,
        ),
        &[0x6c, 0x78, 0xff, 0x6c, 0x74, 0xff, 0x2a, 0x31, 0x70, 0xff]
    );
}

// Select Case `1 To 5` (range): push lo, push hi, range-test (fb 86), branch-false.
#[test]
fn e2e_select_case_range() {
    assert_eq!(
        compile(
            "Attribute VB_Name = \"Module1\"\r\nSub Main()\r\n\
             Dim a As Long\r\n Select Case a\r\n Case 1 To 5\r\n a = 2\r\n End Select\r\nEnd Sub\r\n",
            0x0008,
        ),
        &[0x6c, 0x78, 0xff, 0x71, 0x74, 0xff, 0x6c, 0x74, 0xff, 0xf5, 0x01, 0x00,
          0x00, 0x00, 0xf5, 0x05, 0x00, 0x00, 0x00, 0xfb, 0x86, 0x1c, 0x20, 0x00,
          0xf5, 0x02, 0x00, 0x00, 0x00, 0x71, 0x78, 0xff]
    );
}

// Select Case `Is > 5`: push value, gt-compare (db), branch-false.
#[test]
fn e2e_select_case_is() {
    assert_eq!(
        compile(
            "Attribute VB_Name = \"Module1\"\r\nSub Main()\r\n\
             Dim a As Long\r\n Select Case a\r\n Case Is > 5\r\n a = 2\r\n End Select\r\nEnd Sub\r\n",
            0x0008,
        ),
        &[0x6c, 0x78, 0xff, 0x71, 0x74, 0xff, 0x6c, 0x74, 0xff, 0xf5, 0x05, 0x00,
          0x00, 0x00, 0xdb, 0x1c, 0x1a, 0x00, 0xf5, 0x02, 0x00, 0x00, 0x00, 0x71,
          0x78, 0xff]
    );
}

// Select Case `1, 2, 3` (multi-value OR): each value branches TRUE to the body;
// the last branches FALSE to the next case.
#[test]
fn e2e_select_case_multi() {
    assert_eq!(
        compile(
            "Attribute VB_Name = \"Module1\"\r\nSub Main()\r\n\
             Dim a As Long\r\n Select Case a\r\n Case 1, 2, 3\r\n a = 2\r\n End Select\r\nEnd Sub\r\n",
            0x0008,
        ),
        &[0x6c, 0x78, 0xff, 0x71, 0x74, 0xff, 0x6c, 0x74, 0xff, 0xf5, 0x01, 0x00,
          0x00, 0x00, 0xc7, 0x1d, 0x2a, 0x00, 0x6c, 0x74, 0xff, 0xf5, 0x02, 0x00,
          0x00, 0x00, 0xc7, 0x1d, 0x2a, 0x00, 0x6c, 0x74, 0xff, 0xf5, 0x03, 0x00,
          0x00, 0x00, 0xc7, 0x1c, 0x32, 0x00, 0xf5, 0x02, 0x00, 0x00, 0x00, 0x71,
          0x78, 0xff]
    );
}

// Concat chain `a & b & c`: the intermediate (a & b) is materialized to a hidden
// temp (store-keep 0x23), concatenated with c, the result moved into s (0x31), and
// the temp freed (0x2f).
#[test]
fn e2e_string_concat_chain() {
    assert_eq!(
        compile(
            "Attribute VB_Name = \"Module1\"\r\nSub Main()\r\n\
             Dim a As String, b As String, c As String, s As String\r\n s = a & b & c\r\nEnd Sub\r\n",
            0x0008,
        ),
        &[0x6c, 0x78, 0xff, 0x6c, 0x74, 0xff, 0x2a, 0x23, 0x68, 0xff, 0x6c, 0x70,
          0xff, 0x2a, 0x31, 0x6c, 0xff, 0x2f, 0x68, 0xff]
    );
}

// On Error GoTo label: opcode 0x4b + the backpatched label offset.
#[test]
fn e2e_on_error_goto() {
    assert_eq!(
        compile(
            "Attribute VB_Name = \"Module1\"\r\nSub Main()\r\n\
             On Error GoTo L\r\n Dim a As Long\r\n L:\r\n a = 1\r\nEnd Sub\r\n",
            0x0008,
        ),
        &[0x4b, 0x03, 0x00, 0xf5, 0x01, 0x00, 0x00, 0x00, 0x71, 0x78, 0xff]
    );
}

// LSet target = value: load value, load target, LSet opcode 0x47.
#[test]
fn e2e_lset() {
    assert_eq!(
        compile(
            "Attribute VB_Name = \"Module1\"\r\nSub Main()\r\n\
             Dim a As String, s As String\r\n LSet s = a\r\nEnd Sub\r\n",
            0x0008,
        ),
        &[0x6c, 0x78, 0xff, 0x6c, 0x74, 0xff, 0x47, 0x00, 0x00]
    );
}

// Mid(s, start, len) = value: LdAddr s, push start, push len, push value, Mid 0x4f.
#[test]
fn e2e_mid_statement() {
    assert_eq!(
        compile(
            "Attribute VB_Name = \"Module1\"\r\nSub Main()\r\n\
             Dim s As String\r\n Mid(s, 1, 1) = \"x\"\r\nEnd Sub\r\n",
            0x0008,
        ),
        &[0x04, 0x78, 0xff, 0xf5, 0x01, 0x00, 0x00, 0x00, 0xf5, 0x01, 0x00, 0x00,
          0x00, 0x1b, 0x00, 0x00, 0x4f, 0x00, 0x00]
    );
}

// Negative integer literal folds to a single push of the negated value.
#[test]
fn e2e_negative_literal() {
    assert_eq!(
        compile(
            "Attribute VB_Name = \"Module1\"\r\nSub Main()\r\n\
             Dim r As Long\r\n r = -5\r\nEnd Sub\r\n",
            0x0008,
        ),
        &[0xf5, 0xfb, 0xff, 0xff, 0xff, 0x71, 0x78, 0xff]
    );
}

// Like (string pattern compare): operands load as BSTR, compare fb 7e.
#[test]
fn e2e_string_like() {
    assert_eq!(
        compile(
            "Attribute VB_Name = \"Module1\"\r\nSub Main()\r\n\
             Dim a As String, b As String, r As Integer\r\n r = (a Like b)\r\nEnd Sub\r\n",
            0x0008,
        ),
        &[0x6c, 0x78, 0xff, 0x6c, 0x74, 0xff, 0xfb, 0x7e, 0x70, 0x72, 0xff]
    );
}

// Date literal: push 0xfa + 8-byte OLE serial (#1/1/2000# = 36526.0), Double store.
#[test]
fn e2e_date_literal() {
    assert_eq!(
        compile(
            "Attribute VB_Name = \"Module1\"\r\nSub Main()\r\n\
             Dim d As Date\r\n d = #1/1/2000#\r\nEnd Sub\r\n",
            0x0008,
        ),
        &[0xfa, 0x00, 0x00, 0x00, 0x00, 0xc0, 0xd5, 0xe1, 0x40, 0x74, 0x74, 0xff]
    );
}

// Exit Sub emits the procedure-return opcode 0x14.
#[test]
fn e2e_exit_sub() {
    assert_eq!(
        compile(
            "Attribute VB_Name = \"Module1\"\r\nSub Main()\r\n Exit Sub\r\nEnd Sub\r\n",
            0x0008,
        ),
        &[0x14]
    );
}

// Integer array element store: value pushed as Integer (f4 05), element-store 0xa2.
#[test]
fn e2e_integer_array_store() {
    assert_eq!(
        compile(
            "Attribute VB_Name = \"Module1\"\r\nSub Main()\r\n\
             Dim a(10) As Integer\r\n a(0) = 5\r\nEnd Sub\r\n",
            0x0008,
        ),
        &[0xf4, 0x05, 0xf5, 0x00, 0x00, 0x00, 0x00, 0x04, 0x64, 0xff, 0xa2]
    );
}

// Multi-dimensional array element store `a(1,1) = 5`: push value, push both Long
// indices, LdAddr the 2-D descriptor, indexed store (0xa7 <dims=2> 0x8f). The
// descriptor is 36 bytes (20 + 8 per dimension).
#[test]
fn e2e_multidim_array_store() {
    assert_eq!(
        compile(
            "Attribute VB_Name = \"Module1\"\r\nSub Main()\r\n\
             Dim a(3, 3) As Long\r\n a(1, 1) = 5\r\nEnd Sub\r\n",
            0x0008,
        ),
        &[0xf5, 0x05, 0x00, 0x00, 0x00, 0xf5, 0x01, 0x00, 0x00, 0x00, 0xf5, 0x01,
          0x00, 0x00, 0x00, 0x04, 0x5c, 0xff, 0xa7, 0x02, 0x00, 0x8f, 0x00, 0x00]
    );
}

// ReDim of a dynamic array: push the lower (0) and upper (10) bound, LdAddr the
// array pointer, then the ReDim opcode (fe 8e) + dim-count, element VARTYPE (Long
// = 3), element size (4), and flags (0x80).
#[test]
fn e2e_redim() {
    assert_eq!(
        compile(
            "Attribute VB_Name = \"Module1\"\r\nSub Main()\r\n\
             Dim a() As Long\r\n ReDim a(10)\r\nEnd Sub\r\n",
            0x0008,
        ),
        &[0xf5, 0x00, 0x00, 0x00, 0x00, 0xf5, 0x0a, 0x00, 0x00, 0x00, 0x04, 0x78,
          0xff, 0xfe, 0x8e, 0x01, 0x00, 0x03, 0x00, 0x04, 0x00, 0x80, 0x00]
    );
}

// Array element store `a(0) = 5`: push value, push Long index, LdAddr the array
// descriptor (0x04), element-store (0xa3). a(10) As Long is a 28-byte descriptor.
#[test]
fn e2e_array_store() {
    assert_eq!(
        compile(
            "Attribute VB_Name = \"Module1\"\r\n\
             Sub Main()\r\n\
             Dim a(10) As Long\r\n\
             a(0) = 5\r\n\
             End Sub\r\n",
            0x0008,
        ),
        &[0xf5, 0x05, 0x00, 0x00, 0x00, 0xf5, 0x00, 0x00, 0x00, 0x00, 0x04, 0x64,
          0xff, 0xa3]
    );
}

// Array element load `r = a(0)`: push Long index, LdAddr, element-load (0x9e),
// store to r.
#[test]
fn e2e_array_load() {
    assert_eq!(
        compile(
            "Attribute VB_Name = \"Module1\"\r\n\
             Sub Main()\r\n\
             Dim a(10) As Long, r As Long\r\n\
             r = a(0)\r\n\
             End Sub\r\n",
            0x0008,
        ),
        &[0xf5, 0x00, 0x00, 0x00, 0x00, 0x04, 0x64, 0xff, 0x9e, 0x71, 0x5c, 0xff]
    );
}

// Variant assignment from an integer literal: init the hidden 16-byte Variant
// temp from the inline literal (0x28), then variant-store (fc f6) into v.
#[test]
fn e2e_variant_from_literal() {
    assert_eq!(
        compile(
            "Attribute VB_Name = \"Module1\"\r\n\
             Sub Main()\r\n\
             Dim v As Variant\r\n\
             v = 5\r\n\
             End Sub\r\n",
            0x0008,
        ),
        &[0x28, 0x5c, 0xff, 0x05, 0x00, 0xfc, 0xf6, 0x6c, 0xff]
    );
}

// Variant assignment from a Long: load it, convert Long->Variant into the temp
// (fd 69), then variant-store (fc f6).
#[test]
fn e2e_variant_from_long() {
    assert_eq!(
        compile(
            "Attribute VB_Name = \"Module1\"\r\n\
             Sub Main()\r\n\
             Dim a As Long, v As Variant\r\n\
             v = a\r\n\
             End Sub\r\n",
            0x0008,
        ),
        &[0x6c, 0x78, 0xff, 0xfd, 0x69, 0x58, 0xff, 0xfc, 0xf6, 0x68, 0xff]
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

// ── Matrix-completion constructs (Phase 3) ───────────────────────────────────
//
// Exact-byte vectors for the COM-free single-procedure constructs closed in the
// comprehensive coverage pass: every literal form, operator/type coercion,
// statement, and array/string/Variant data access. Each vector is confirmed
// byte-for-byte against the real VB6 p-code compiler.

#[test]
fn e2e_a8_single_literal() {
    assert_eq!(
        compile(
            "Attribute VB_Name = \"Module1\"\r\nSub Main()\r\nDim r As Single\r\nr = 1.5\r\nEnd Sub\r\n",
            0x0008,
        ),
        &[0xfa, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0xf8, 0x3f, 0x73, 0x78, 0xff]
    );
}

#[test]
fn e2e_a9_double_literal() {
    assert_eq!(
        compile(
            "Attribute VB_Name = \"Module1\"\r\nSub Main()\r\nDim r As Double\r\nr = 1.5\r\nEnd Sub\r\n",
            0x0008,
        ),
        &[0xfa, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0xf8, 0x3f, 0x74, 0x74, 0xff]
    );
}

#[test]
fn e2e_a10_double_hash_literal() {
    assert_eq!(
        compile(
            "Attribute VB_Name = \"Module1\"\r\nSub Main()\r\nDim r As Double\r\nr = 1.5#\r\nEnd Sub\r\n",
            0x0008,
        ),
        &[0xfa, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0xf8, 0x3f, 0x74, 0x74, 0xff]
    );
}

#[test]
fn e2e_a11_double_scientific_literal() {
    assert_eq!(
        compile(
            "Attribute VB_Name = \"Module1\"\r\nSub Main()\r\nDim r As Double\r\nr = 1.5E10\r\nEnd Sub\r\n",
            0x0008,
        ),
        &[0xfa, 0x00, 0x00, 0x00, 0xb0, 0x8e, 0xf0, 0x0b, 0x42, 0x74, 0x74, 0xff]
    );
}

#[test]
fn e2e_a12_currency_literal() {
    assert_eq!(
        compile(
            "Attribute VB_Name = \"Module1\"\r\nSub Main()\r\nDim r As Currency\r\nr = 1.5@\r\nEnd Sub\r\n",
            0x0008,
        ),
        &[0xf6, 0x98, 0x3a, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x72, 0x74, 0xff]
    );
}

#[test]
fn e2e_a14_empty_string_literal() {
    assert_eq!(
        compile(
            "Attribute VB_Name = \"Module1\"\r\nSub Main()\r\nDim s As String\r\ns = \"\"\r\nEnd Sub\r\n",
            0x0008,
        ),
        &[0x1b, 0x00, 0x00, 0x43, 0x78, 0xff]
    );
}

#[test]
fn e2e_a16_date_time_literal() {
    assert_eq!(
        compile(
            "Attribute VB_Name = \"Module1\"\r\nSub Main()\r\nDim d As Date\r\nd = #1/1/2000 3:00:00 PM#\r\nEnd Sub\r\n",
            0x0008,
        ),
        &[0xfa, 0x00, 0x00, 0x00, 0x00, 0xd4, 0xd5, 0xe1, 0x40, 0x74, 0x74, 0xff]
    );
}

#[test]
fn e2e_a18_bool_false_literal() {
    assert_eq!(
        compile(
            "Attribute VB_Name = \"Module1\"\r\nSub Main()\r\nDim b As Boolean\r\nb = False\r\nEnd Sub\r\n",
            0x0008,
        ),
        &[0xf4, 0x00, 0x70, 0x7a, 0xff]
    );
}

#[test]
fn e2e_a19_variant_empty() {
    assert_eq!(
        compile(
            "Attribute VB_Name = \"Module1\"\r\nSub Main()\r\nDim v As Variant\r\nv = Empty\r\nEnd Sub\r\n",
            0x0008,
        ),
        &[0xfc, 0x67, 0x5c, 0xff, 0xfc, 0xf6, 0x6c, 0xff]
    );
}

#[test]
fn e2e_a20_variant_null() {
    assert_eq!(
        compile(
            "Attribute VB_Name = \"Module1\"\r\nSub Main()\r\nDim v As Variant\r\nv = Null\r\nEnd Sub\r\n",
            0x0008,
        ),
        &[0xfc, 0x64, 0x5c, 0xff, 0xfc, 0xf6, 0x6c, 0xff]
    );
}

#[test]
fn e2e_b4_single_divide() {
    assert_eq!(
        compile(
            "Attribute VB_Name = \"Module1\"\r\nSub Main()\r\nDim a As Single, b As Single, r As Single\r\nr = a / b\r\nEnd Sub\r\n",
            0x0008,
        ),
        &[0x6e, 0x78, 0xff, 0x6e, 0x74, 0xff, 0xb6, 0x73, 0x70, 0xff]
    );
}

#[test]
fn e2e_b4_currency_divide() {
    assert_eq!(
        compile(
            "Attribute VB_Name = \"Module1\"\r\nSub Main()\r\nDim a As Currency, b As Currency, r As Currency\r\nr = a / b\r\nEnd Sub\r\n",
            0x0008,
        ),
        &[0x6d, 0x74, 0xff, 0xee, 0x6d, 0x6c, 0xff, 0xee, 0xb6, 0xf1, 0x72, 0x64, 0xff]
    );
}

#[test]
fn e2e_b5_integer_idiv() {
    assert_eq!(
        compile(
            "Attribute VB_Name = \"Module1\"\r\nSub Main()\r\nDim a As Integer, b As Integer, r As Integer\r\nr = a \\ b\r\nEnd Sub\r\n",
            0x0008,
        ),
        &[0x6b, 0x7a, 0xff, 0x6b, 0x78, 0xff, 0xbf, 0x70, 0x76, 0xff]
    );
}

#[test]
fn e2e_b6_integer_mod() {
    assert_eq!(
        compile(
            "Attribute VB_Name = \"Module1\"\r\nSub Main()\r\nDim a As Integer, b As Integer, r As Integer\r\nr = a Mod b\r\nEnd Sub\r\n",
            0x0008,
        ),
        &[0x6b, 0x7a, 0xff, 0x6b, 0x78, 0xff, 0xc1, 0x70, 0x76, 0xff]
    );
}

#[test]
fn e2e_b11_integer_eqv() {
    assert_eq!(
        compile(
            "Attribute VB_Name = \"Module1\"\r\nSub Main()\r\nDim a As Integer, b As Integer, r As Integer\r\nr = a Eqv b\r\nEnd Sub\r\n",
            0x0008,
        ),
        &[0x6b, 0x7a, 0xff, 0x6b, 0x78, 0xff, 0xfb, 0x0a, 0x70, 0x76, 0xff]
    );
}

#[test]
fn e2e_b12_integer_imp() {
    assert_eq!(
        compile(
            "Attribute VB_Name = \"Module1\"\r\nSub Main()\r\nDim a As Integer, b As Integer, r As Integer\r\nr = a Imp b\r\nEnd Sub\r\n",
            0x0008,
        ),
        &[0x6b, 0x7a, 0xff, 0x6b, 0x78, 0xff, 0xfb, 0x02, 0x70, 0x76, 0xff]
    );
}

#[test]
fn e2e_b1_date_plus_numeric() {
    assert_eq!(
        compile(
            "Attribute VB_Name = \"Module1\"\r\nSub Main()\r\nDim d As Date, r As Date\r\nr = d + 1\r\nEnd Sub\r\n",
            0x0008,
        ),
        &[0x6f, 0x74, 0xff, 0xf4, 0x01, 0xeb, 0xab, 0xf2, 0x74, 0x6c, 0xff]
    );
}

#[test]
fn e2e_c1_negate_integer() {
    assert_eq!(
        compile(
            "Attribute VB_Name = \"Module1\"\r\nSub Main()\r\nDim a As Integer, r As Integer\r\nr = -a\r\nEnd Sub\r\n",
            0x0008,
        ),
        &[0x6b, 0x7a, 0xff, 0xb7, 0x70, 0x78, 0xff]
    );
}

#[test]
fn e2e_c1_negate_currency() {
    assert_eq!(
        compile(
            "Attribute VB_Name = \"Module1\"\r\nSub Main()\r\nDim a As Currency, r As Currency\r\nr = -a\r\nEnd Sub\r\n",
            0x0008,
        ),
        &[0x6d, 0x74, 0xff, 0xba, 0x72, 0x6c, 0xff]
    );
}

#[test]
fn e2e_c3_not_integer() {
    assert_eq!(
        compile(
            "Attribute VB_Name = \"Module1\"\r\nSub Main()\r\nDim a As Integer, r As Integer\r\nr = Not a\r\nEnd Sub\r\n",
            0x0008,
        ),
        &[0x6b, 0x7a, 0xff, 0xc3, 0x70, 0x78, 0xff]
    );
}

#[test]
fn e2e_c3_not_boolean() {
    assert_eq!(
        compile(
            "Attribute VB_Name = \"Module1\"\r\nSub Main()\r\nDim a As Boolean, r As Boolean\r\nr = Not a\r\nEnd Sub\r\n",
            0x0008,
        ),
        &[0x6b, 0x7a, 0xff, 0xc3, 0x70, 0x78, 0xff]
    );
}

#[test]
fn e2e_d2_integer_from_long() {
    assert_eq!(
        compile(
            "Attribute VB_Name = \"Module1\"\r\nSub Main()\r\nDim a As Long, r As Integer\r\nr = a\r\nEnd Sub\r\n",
            0x0008,
        ),
        &[0x6c, 0x78, 0xff, 0xe4, 0x70, 0x76, 0xff]
    );
}

#[test]
fn e2e_d4_double_from_single() {
    assert_eq!(
        compile(
            "Attribute VB_Name = \"Module1\"\r\nSub Main()\r\nDim a As Single, r As Double\r\nr = a\r\nEnd Sub\r\n",
            0x0008,
        ),
        &[0x6e, 0x78, 0xff, 0x74, 0x70, 0xff]
    );
}

#[test]
fn e2e_d5_double_from_currency() {
    assert_eq!(
        compile(
            "Attribute VB_Name = \"Module1\"\r\nSub Main()\r\nDim a As Currency, r As Double\r\nr = a\r\nEnd Sub\r\n",
            0x0008,
        ),
        &[0x6d, 0x74, 0xff, 0xee, 0x74, 0x6c, 0xff]
    );
}

#[test]
fn e2e_d6_single_from_double() {
    assert_eq!(
        compile(
            "Attribute VB_Name = \"Module1\"\r\nSub Main()\r\nDim a As Double, r As Single\r\nr = a\r\nEnd Sub\r\n",
            0x0008,
        ),
        &[0x6f, 0x74, 0xff, 0x73, 0x70, 0xff]
    );
}

#[test]
fn e2e_d7_single_from_currency() {
    assert_eq!(
        compile(
            "Attribute VB_Name = \"Module1\"\r\nSub Main()\r\nDim a As Currency, r As Single\r\nr = a\r\nEnd Sub\r\n",
            0x0008,
        ),
        &[0x6d, 0x74, 0xff, 0xee, 0x73, 0x70, 0xff]
    );
}

#[test]
fn e2e_d8_currency_from_long() {
    assert_eq!(
        compile(
            "Attribute VB_Name = \"Module1\"\r\nSub Main()\r\nDim a As Long, r As Currency\r\nr = a\r\nEnd Sub\r\n",
            0x0008,
        ),
        &[0x6c, 0x78, 0xff, 0xf0, 0x72, 0x70, 0xff]
    );
}

#[test]
fn e2e_d9_date_from_double() {
    assert_eq!(
        compile(
            "Attribute VB_Name = \"Module1\"\r\nSub Main()\r\nDim a As Double, r As Date\r\nr = a\r\nEnd Sub\r\n",
            0x0008,
        ),
        &[0x6f, 0x74, 0xff, 0xf2, 0x74, 0x6c, 0xff]
    );
}

#[test]
fn e2e_d12_string_from_numeric() {
    assert_eq!(
        compile(
            "Attribute VB_Name = \"Module1\"\r\nSub Main()\r\nDim n As Long, s As String\r\ns = n\r\nEnd Sub\r\n",
            0x0008,
        ),
        &[0x6c, 0x78, 0xff, 0xfb, 0xfe, 0x31, 0x74, 0xff]
    );
}

#[test]
fn e2e_d13_numeric_from_string() {
    assert_eq!(
        compile(
            "Attribute VB_Name = \"Module1\"\r\nSub Main()\r\nDim s As String, r As Long\r\nr = s\r\nEnd Sub\r\n",
            0x0008,
        ),
        &[0x6c, 0x78, 0xff, 0x50, 0x71, 0x74, 0xff]
    );
}

#[test]
fn e2e_d14_byte_from_integer() {
    assert_eq!(
        compile(
            "Attribute VB_Name = \"Module1\"\r\nSub Main()\r\nDim a As Integer, r As Byte\r\nr = a\r\nEnd Sub\r\n",
            0x0008,
        ),
        &[0x6b, 0x7a, 0xff, 0xfc, 0x0d, 0xfc, 0xf0, 0x78, 0xff]
    );
}

#[test]
fn e2e_e4_const_string() {
    assert_eq!(
        compile(
            "Attribute VB_Name = \"Module1\"\r\nSub Main()\r\nConst K As String = \"x\"\r\nDim s As String\r\ns = K\r\nEnd Sub\r\n",
            0x0008,
        ),
        &[0x1b, 0x00, 0x00, 0x43, 0x78, 0xff]
    );
}

#[test]
fn e2e_e4_const_double() {
    assert_eq!(
        compile(
            "Attribute VB_Name = \"Module1\"\r\nSub Main()\r\nConst K As Double = 1.5\r\nDim r As Double\r\nr = K\r\nEnd Sub\r\n",
            0x0008,
        ),
        &[0xfa, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0xf8, 0x3f, 0x74, 0x74, 0xff]
    );
}

#[test]
fn e2e_f1_double_array_store() {
    assert_eq!(
        compile(
            "Attribute VB_Name = \"Module1\"\r\nSub Main()\r\nDim a(10) As Double, b As Double\r\na(0) = b\r\nEnd Sub\r\n",
            0x0008,
        ),
        &[0x6f, 0x58, 0xff, 0xf5, 0x00, 0x00, 0x00, 0x00, 0x04, 0x64, 0xff, 0xa6]
    );
}

#[test]
fn e2e_f1_string_array_store() {
    assert_eq!(
        compile(
            "Attribute VB_Name = \"Module1\"\r\nSub Main()\r\nDim a(10) As String, s As String\r\na(0) = s\r\nEnd Sub\r\n",
            0x0008,
        ),
        &[0x6c, 0x5c, 0xff, 0xf5, 0x00, 0x00, 0x00, 0x00, 0x04, 0x64, 0xff, 0x3b]
    );
}

#[test]
fn e2e_f1_single_array_store() {
    assert_eq!(
        compile(
            "Attribute VB_Name = \"Module1\"\r\nSub Main()\r\nDim a(10) As Single, b As Single\r\na(0) = b\r\nEnd Sub\r\n",
            0x0008,
        ),
        &[0x6e, 0x5c, 0xff, 0xf5, 0x00, 0x00, 0x00, 0x00, 0x04, 0x64, 0xff, 0xa5]
    );
}

#[test]
fn e2e_f1_currency_array_store() {
    assert_eq!(
        compile(
            "Attribute VB_Name = \"Module1\"\r\nSub Main()\r\nDim a(10) As Currency, b As Currency\r\na(0) = b\r\nEnd Sub\r\n",
            0x0008,
        ),
        &[0x6d, 0x58, 0xff, 0xf5, 0x00, 0x00, 0x00, 0x00, 0x04, 0x64, 0xff, 0xa4]
    );
}

#[test]
fn e2e_f1_byte_array_store() {
    assert_eq!(
        compile(
            "Attribute VB_Name = \"Module1\"\r\nSub Main()\r\nDim a(10) As Byte, b As Byte\r\na(0) = b\r\nEnd Sub\r\n",
            0x0008,
        ),
        &[0xfc, 0xe0, 0x5e, 0xff, 0xf5, 0x00, 0x00, 0x00, 0x00, 0x04, 0x64, 0xff, 0xfc, 0xa0]
    );
}

#[test]
fn e2e_f5_redim_preserve() {
    assert_eq!(
        compile(
            "Attribute VB_Name = \"Module1\"\r\nSub Main()\r\nDim a() As Long\r\nReDim a(10)\r\nReDim Preserve a(20)\r\nEnd Sub\r\n",
            0x0008,
        ),
        &[0xf5, 0x00, 0x00, 0x00, 0x00, 0xf5, 0x0a, 0x00, 0x00, 0x00, 0x04, 0x78, 0xff, 0xfe, 0x8e, 0x01, 0x00, 0x03, 0x00, 0x04, 0x00, 0x80, 0x00, 0xf5, 0x00, 0x00, 0x00, 0x00, 0xf5, 0x14, 0x00, 0x00, 0x00, 0x04, 0x78, 0xff, 0xfe, 0x8f, 0x01, 0x00, 0x03, 0x00, 0x04, 0x00, 0x80, 0x00]
    );
}

#[test]
fn e2e_f6_erase_dynamic() {
    assert_eq!(
        compile(
            "Attribute VB_Name = \"Module1\"\r\nSub Main()\r\nDim a() As Long\r\nReDim a(10)\r\nErase a\r\nEnd Sub\r\n",
            0x0008,
        ),
        &[0xf5, 0x00, 0x00, 0x00, 0x00, 0xf5, 0x0a, 0x00, 0x00, 0x00, 0x04, 0x78, 0xff, 0xfe, 0x8e, 0x01, 0x00, 0x03, 0x00, 0x04, 0x00, 0x80, 0x00, 0x04, 0x78, 0xff, 0x5a]
    );
}

#[test]
fn e2e_g6_midb_statement() {
    assert_eq!(
        compile(
            "Attribute VB_Name = \"Module1\"\r\nSub Main()\r\nDim s As String\r\nMidB(s, 1, 2) = \"x\"\r\nEnd Sub\r\n",
            0x0008,
        ),
        &[0x04, 0x78, 0xff, 0xf5, 0x01, 0x00, 0x00, 0x00, 0xf5, 0x02, 0x00, 0x00, 0x00, 0x1b, 0x00, 0x00, 0xfc, 0xbe, 0x00, 0x00]
    );
}

#[test]
fn e2e_g6_mid_dollar_statement() {
    assert_eq!(
        compile(
            "Attribute VB_Name = \"Module1\"\r\nSub Main()\r\nDim s As String\r\nMid$(s, 1, 1) = \"x\"\r\nEnd Sub\r\n",
            0x0008,
        ),
        &[0x04, 0x78, 0xff, 0xf5, 0x01, 0x00, 0x00, 0x00, 0xf5, 0x01, 0x00, 0x00, 0x00, 0x1b, 0x00, 0x00, 0x4f, 0x00, 0x00]
    );
}

#[test]
fn e2e_g7_rset() {
    assert_eq!(
        compile(
            "Attribute VB_Name = \"Module1\"\r\nSub Main()\r\nDim a As String, s As String\r\nRSet s = a\r\nEnd Sub\r\n",
            0x0008,
        ),
        &[0x6c, 0x78, 0xff, 0x6c, 0x74, 0xff, 0xfe, 0x1e, 0x00, 0x00]
    );
}

#[test]
fn e2e_h3_do_infinite_exit() {
    assert_eq!(
        compile(
            "Attribute VB_Name = \"Module1\"\r\nSub Main()\r\nDim a As Long\r\nDo\r\na = a + 1\r\nExit Do\r\nLoop\r\nEnd Sub\r\n",
            0x0008,
        ),
        &[0x6c, 0x78, 0xff, 0xf5, 0x01, 0x00, 0x00, 0x00, 0xaa, 0x71, 0x78, 0xff, 0x1e, 0x12, 0x00, 0x1e, 0x00, 0x00]
    );
}

#[test]
fn e2e_h7_for_negative_step() {
    assert_eq!(
        compile(
            "Attribute VB_Name = \"Module1\"\r\nSub Main()\r\nDim i As Long\r\nFor i = 10 To 1 Step -1\r\nNext i\r\nEnd Sub\r\n",
            0x0008,
        ),
        &[0xf5, 0x0a, 0x00, 0x00, 0x00, 0x04, 0x78, 0xff, 0xf5, 0x01, 0x00, 0x00, 0x00, 0xf5, 0xff, 0xff, 0xff, 0xff, 0xfe, 0x6c, 0x70, 0xff, 0x20, 0x00, 0x04, 0x78, 0xff, 0x67, 0x70, 0xff, 0x18, 0x00]
    );
}

#[test]
fn e2e_h8_nested_for() {
    assert_eq!(
        compile(
            "Attribute VB_Name = \"Module1\"\r\nSub Main()\r\nDim i As Long, j As Long\r\nFor i = 1 To 3\r\nFor j = 1 To 3\r\nNext j\r\nNext i\r\nEnd Sub\r\n",
            0x0008,
        ),
        &[0xf5, 0x01, 0x00, 0x00, 0x00, 0x04, 0x78, 0xff, 0xf5, 0x03, 0x00, 0x00, 0x00, 0xfe, 0x64, 0x6c, 0xff, 0x36, 0x00, 0xf5, 0x01, 0x00, 0x00, 0x00, 0x04, 0x74, 0xff, 0xf5, 0x03, 0x00, 0x00, 0x00, 0xfe, 0x64, 0x64, 0xff, 0x2e, 0x00, 0x04, 0x74, 0xff, 0x66, 0x64, 0xff, 0x26, 0x00, 0x04, 0x78, 0xff, 0x66, 0x6c, 0xff, 0x13, 0x00]
    );
}

#[test]
fn e2e_i2_gosub_return() {
    assert_eq!(
        compile(
            "Attribute VB_Name = \"Module1\"\r\nSub Main()\r\nGoSub L\r\nExit Sub\r\nL:\r\nReturn\r\nEnd Sub\r\n",
            0x0008,
        ),
        &[0xfd, 0x0a, 0x05, 0x00, 0x14, 0xfc, 0xc9]
    );
}

#[test]
fn e2e_i3_on_expr_goto() {
    assert_eq!(
        compile(
            "Attribute VB_Name = \"Module1\"\r\nSub Main()\r\nDim a As Long\r\nOn a GoTo L1, L2\r\nL1:\r\na = 1\r\nL2:\r\na = 2\r\nEnd Sub\r\n",
            0x0008,
        ),
        &[0x6c, 0x78, 0xff, 0xe4, 0xfe, 0x96, 0x04, 0x00, 0x0c, 0x00, 0x14, 0x00, 0xf5, 0x01, 0x00, 0x00, 0x00, 0x71, 0x78, 0xff, 0xf5, 0x02, 0x00, 0x00, 0x00, 0x71, 0x78, 0xff]
    );
}

#[test]
fn e2e_i4_on_expr_gosub() {
    assert_eq!(
        compile(
            "Attribute VB_Name = \"Module1\"\r\nSub Main()\r\nDim a As Long\r\nOn a GoSub L1, L2\r\nExit Sub\r\nL1:\r\nReturn\r\nL2:\r\nReturn\r\nEnd Sub\r\n",
            0x0008,
        ),
        &[0x6c, 0x78, 0xff, 0xe4, 0xfe, 0x95, 0x04, 0x00, 0x0d, 0x00, 0x0f, 0x00, 0x14, 0xfc, 0xc9, 0xfc, 0xc9]
    );
}

#[test]
fn e2e_i6_on_error_resume_next() {
    assert_eq!(
        compile(
            "Attribute VB_Name = \"Module1\"\r\nSub Main()\r\nOn Error Resume Next\r\nDim a As Long\r\na = 1\r\nEnd Sub\r\n",
            0x0008,
        ),
        &[0x00, 0x02, 0x00, 0x05, 0x4b, 0xff, 0xff, 0x00, 0x0a, 0xf5, 0x01, 0x00, 0x00, 0x00, 0x71, 0x78, 0xff, 0x00, 0x00]
    );
}

#[test]
fn e2e_i7_on_error_goto_zero() {
    assert_eq!(
        compile(
            "Attribute VB_Name = \"Module1\"\r\nSub Main()\r\nOn Error GoTo 0\r\nDim a As Long\r\na = 1\r\nEnd Sub\r\n",
            0x0008,
        ),
        &[0x4b, 0xfe, 0xff, 0xf5, 0x01, 0x00, 0x00, 0x00, 0x71, 0x78, 0xff]
    );
}

#[test]
fn e2e_i8_resume_next() {
    assert_eq!(
        compile(
            "Attribute VB_Name = \"Module1\"\r\nSub Main()\r\nOn Error GoTo L\r\nDim a As Long\r\na = 1\r\nL:\r\nResume Next\r\nEnd Sub\r\n",
            0x0008,
        ),
        &[0x00, 0x02, 0x00, 0x05, 0x4b, 0x11, 0x00, 0x00, 0x0a, 0xf5, 0x01, 0x00, 0x00, 0x00, 0x71, 0x78, 0xff, 0x00, 0x06, 0xfd, 0x0c, 0xff, 0xff, 0x00, 0x00]
    );
}

#[test]
fn e2e_i8_resume() {
    assert_eq!(
        compile(
            "Attribute VB_Name = \"Module1\"\r\nSub Main()\r\nOn Error GoTo L\r\nDim a As Long\r\na = 1\r\nL:\r\nResume\r\nEnd Sub\r\n",
            0x0008,
        ),
        &[0x00, 0x02, 0x00, 0x05, 0x4b, 0x11, 0x00, 0x00, 0x0a, 0xf5, 0x01, 0x00, 0x00, 0x00, 0x71, 0x78, 0xff, 0x00, 0x06, 0xfd, 0x0c, 0xfe, 0xff, 0x00, 0x00]
    );
}

#[test]
fn e2e_i10_stop() {
    assert_eq!(
        compile(
            "Attribute VB_Name = \"Module1\"\r\nSub Main()\r\nStop\r\nEnd Sub\r\n",
            0x0008,
        ),
        &[0xfc, 0xc2]
    );
}

#[test]
fn e2e_i11_end() {
    assert_eq!(
        compile(
            "Attribute VB_Name = \"Module1\"\r\nSub Main()\r\nEnd\r\nEnd Sub\r\n",
            0x0008,
        ),
        &[0xfc, 0xc8]
    );
}

#[test]
fn e2e_i12_error_statement() {
    assert_eq!(
        compile(
            "Attribute VB_Name = \"Module1\"\r\nSub Main()\r\nError 5\r\nEnd Sub\r\n",
            0x0008,
        ),
        &[0xf5, 0x05, 0x00, 0x00, 0x00, 0x45]
    );
}

#[test]
fn e2e_i13_numeric_line_label() {
    assert_eq!(
        compile(
            "Attribute VB_Name = \"Module1\"\r\nSub Main()\r\nDim a As Long\r\nGoTo 100\r\n100 a = 1\r\nEnd Sub\r\n",
            0x0008,
        ),
        &[0x00, 0x02, 0x00, 0x05, 0x1e, 0x07, 0x00, 0x00, 0x0a, 0xf5, 0x01, 0x00, 0x00, 0x00, 0x71, 0x78, 0xff, 0x00, 0x00]
    );
}

// ── Variant source paths, concat conversions, and fixed-array Erase ──────────

#[test]
fn e2e_d11_long_from_variant() {
    assert_eq!(
        compile(
            "Attribute VB_Name = \"Module1\"\r\nSub Main()\r\nDim v As Variant, r As Long\r\nr = v\r\nEnd Sub\r\n",
            0x0008,
        ),
        &[0x04, 0x6c, 0xff, 0xfc, 0x22, 0x71, 0x68, 0xff]
    );
}

#[test]
fn e2e_f1_variant_array_store() {
    assert_eq!(
        compile(
            "Attribute VB_Name = \"Module1\"\r\nSub Main()\r\nDim a(10) As Variant, v As Variant\r\na(0) = v\r\nEnd Sub\r\n",
            0x0008,
        ),
        &[0x04, 0x50, 0xff, 0xf5, 0x00, 0x00, 0x00, 0x00, 0x04, 0x64, 0xff, 0xfc, 0xb0]
    );
}

#[test]
fn e2e_b15_numeric_string_concat() {
    assert_eq!(
        compile(
            "Attribute VB_Name = \"Module1\"\r\nSub Main()\r\nDim n As Long, s As String\r\ns = n & \"x\"\r\nEnd Sub\r\n",
            0x0008,
        ),
        &[0x6c, 0x78, 0xff, 0xfb, 0xfe, 0x23, 0x70, 0xff, 0x1b, 0x00, 0x00, 0x2a, 0x31, 0x74, 0xff, 0x2f, 0x70, 0xff]
    );
}

#[test]
fn e2e_g8_fixed_string_concat() {
    assert_eq!(
        compile(
            "Attribute VB_Name = \"Module1\"\r\nSub Main()\r\nDim a As String * 4, b As String * 4, s As String\r\ns = a & b\r\nEnd Sub\r\n",
            0x0008,
        ),
        &[0x04, 0x74, 0xff, 0x33, 0x04, 0x00, 0x23, 0x64, 0xff, 0x04, 0x6c, 0xff, 0x33, 0x04, 0x00, 0x23, 0x60, 0xff, 0x2a, 0x31, 0x68, 0xff, 0x32, 0x04, 0x00, 0x64, 0xff, 0x60, 0xff]
    );
}

#[test]
fn e2e_concat_var_then_numeric() {
    assert_eq!(
        compile(
            "Attribute VB_Name = \"Module1\"\r\nSub Main()\r\nDim a As String, n As Long, s As String\r\ns = a & n\r\nEnd Sub\r\n",
            0x0008,
        ),
        &[0x6c, 0x78, 0xff, 0x6c, 0x74, 0xff, 0xfb, 0xfe, 0x23, 0x6c, 0xff, 0x2a, 0x31, 0x70, 0xff, 0x2f, 0x6c, 0xff]
    );
}

#[test]
fn e2e_f6_erase_fixed() {
    assert_eq!(
        compile(
            "Attribute VB_Name = \"Module1\"\r\nSub Main()\r\nDim a(10) As Long\r\nErase a\r\nEnd Sub\r\n",
            0x0008,
        ),
        &[0x04, 0x64, 0xff, 0x59, 0x5c, 0xff, 0x5a]
    );
}

#[test]
fn e2e_f6_erase_fixed_2d() {
    assert_eq!(
        compile(
            "Attribute VB_Name = \"Module1\"\r\nSub Main()\r\nDim a(3, 3) As Long\r\nErase a\r\nEnd Sub\r\n",
            0x0008,
        ),
        &[0x04, 0x5c, 0xff, 0x59, 0x54, 0xff, 0x5a]
    );
}

// ── Static locals (per-procedure static block, 0x5f-addressed) ───────────────

#[test]
fn e2e_e3_static_long_store() {
    assert_eq!(
        compile(
            "Attribute VB_Name = \"Module1\"\r\nSub Main()\r\nStatic x As Long\r\nx = 1\r\nEnd Sub\r\n",
            0x0008,
        ),
        &[0xf5, 0x01, 0x00, 0x00, 0x00, 0x5f, 0x08, 0x00, 0x04, 0x00, 0x8f, 0x00, 0x00]
    );
}

#[test]
fn e2e_static_integer_store() {
    assert_eq!(
        compile(
            "Attribute VB_Name = \"Module1\"\r\nSub Main()\r\nStatic x As Integer\r\nx = 1\r\nEnd Sub\r\n",
            0x0008,
        ),
        &[0xf4, 0x01, 0x5f, 0x08, 0x00, 0x04, 0x00, 0x8e, 0x00, 0x00]
    );
}

#[test]
fn e2e_static_double_store() {
    assert_eq!(
        compile(
            "Attribute VB_Name = \"Module1\"\r\nSub Main()\r\nDim b As Double\r\nStatic x As Double\r\nx = b\r\nEnd Sub\r\n",
            0x0008,
        ),
        &[0x6f, 0x74, 0xff, 0x5f, 0x08, 0x00, 0x04, 0x00, 0x92, 0x00, 0x00]
    );
}

#[test]
fn e2e_static_byte_store() {
    assert_eq!(
        compile(
            "Attribute VB_Name = \"Module1\"\r\nSub Main()\r\nDim b As Byte\r\nStatic x As Byte\r\nx = b\r\nEnd Sub\r\n",
            0x0008,
        ),
        &[0xfc, 0xe0, 0x7a, 0xff, 0x5f, 0x08, 0x00, 0x04, 0x00, 0xfd, 0x80, 0x00, 0x00]
    );
}

#[test]
fn e2e_static_string_store() {
    assert_eq!(
        compile(
            "Attribute VB_Name = \"Module1\"\r\nSub Main()\r\nStatic x As String\r\nx = \"a\"\r\nEnd Sub\r\n",
            0x0008,
        ),
        &[0x1b, 0x00, 0x00, 0x5f, 0x08, 0x00, 0x04, 0x00, 0xfd, 0x91, 0x00, 0x00]
    );
}

#[test]
fn e2e_static_long_load() {
    assert_eq!(
        compile(
            "Attribute VB_Name = \"Module1\"\r\nSub Main()\r\nStatic x As Long\r\nDim r As Long\r\nr = x\r\nEnd Sub\r\n",
            0x0008,
        ),
        &[0x5f, 0x08, 0x00, 0x04, 0x00, 0x8a, 0x00, 0x00, 0x71, 0x78, 0xff]
    );
}

#[test]
fn e2e_static_two_longs() {
    assert_eq!(
        compile(
            "Attribute VB_Name = \"Module1\"\r\nSub Main()\r\nStatic x As Long\r\nStatic y As Long\r\nx = 1\r\ny = 2\r\nEnd Sub\r\n",
            0x0008,
        ),
        &[0xf5, 0x01, 0x00, 0x00, 0x00, 0x5f, 0x08, 0x00, 0x04, 0x00, 0x8f, 0x00, 0x00, 0xf5, 0x02, 0x00, 0x00, 0x00, 0x5f, 0x08, 0x00, 0x04, 0x00, 0x8f, 0x04, 0x00]
    );
}

// ── Multi-procedure: parameters, Function/Property returns, arrays ───────────

#[test]
fn e2e_mp_byval_long_param() {
    let procs = compile_module(
            "Attribute VB_Name = \"Module1\"\r\nSub Main()\r\n    Dim a As Long\r\n    a = 1\r\nEnd Sub\r\nSub Foo(ByVal x As Long)\r\n    Dim y As Long\r\n    y = x\r\nEnd Sub\r\n",
        0x0008);
    assert_eq!(procs[1], &[0x6c, 0x0c, 0x00, 0x71, 0x78, 0xff]);
}

#[test]
fn e2e_mp_byref_long_param() {
    let procs = compile_module(
            "Attribute VB_Name = \"Module1\"\r\nSub Main()\r\n    Dim a As Long\r\n    a = 1\r\nEnd Sub\r\nSub Foo(x As Long)\r\n    Dim y As Long\r\n    y = x\r\nEnd Sub\r\n",
        0x0008);
    assert_eq!(procs[1], &[0x80, 0x0c, 0x00, 0x71, 0x78, 0xff]);
}

#[test]
fn e2e_mp_optional_long_param() {
    let procs = compile_module(
            "Attribute VB_Name = \"Module1\"\r\nSub Main()\r\n    Dim a As Long\r\n    a = 1\r\nEnd Sub\r\nSub Foo(Optional ByVal x As Long)\r\n    Dim y As Long\r\n    y = x\r\nEnd Sub\r\n",
        0x0008);
    assert_eq!(procs[1], &[0x6c, 0x0c, 0x00, 0x71, 0x78, 0xff]);
}

#[test]
fn e2e_mp_function_string_return() {
    let procs = compile_module(
            "Attribute VB_Name = \"Module1\"\r\nSub Main()\r\n    Dim a As Long\r\n    a = 1\r\nEnd Sub\r\nFunction F() As String\r\n    F = \"x\"\r\nEnd Function\r\n",
        0x0008);
    assert_eq!(procs[1], &[0x1b, 0x00, 0x00, 0x43, 0x78, 0xff]);
}

#[test]
fn e2e_mp_function_double_return() {
    let procs = compile_module(
            "Attribute VB_Name = \"Module1\"\r\nSub Main()\r\n    Dim a As Long\r\n    a = 1\r\nEnd Sub\r\nFunction F() As Double\r\n    Dim d As Double\r\n    F = d\r\nEnd Function\r\n",
        0x0008);
    assert_eq!(procs[1], &[0x6f, 0x6c, 0xff, 0x74, 0x74, 0xff]);
}

#[test]
fn e2e_mp_function_coerced_return() {
    let procs = compile_module(
            "Attribute VB_Name = \"Module1\"\r\nSub Main()\r\n    Dim a As Long\r\n    a = 1\r\nEnd Sub\r\nFunction F() As Long\r\n    Dim i As Integer\r\n    F = i\r\nEnd Function\r\n",
        0x0008);
    assert_eq!(procs[1], &[0x6b, 0x76, 0xff, 0xe7, 0x71, 0x78, 0xff]);
}

#[test]
fn e2e_mp_property_get_long() {
    let procs = compile_module(
            "Attribute VB_Name = \"Module1\"\r\nSub Main()\r\n    Dim a As Long\r\n    a = 1\r\nEnd Sub\r\nProperty Get P() As Long\r\n    P = 5\r\nEnd Property\r\n",
        0x0008);
    assert_eq!(procs[1], &[0xf5, 0x05, 0x00, 0x00, 0x00, 0x71, 0x78, 0xff]);
}

#[test]
fn e2e_mp_array_store_second_proc() {
    let procs = compile_module(
            "Attribute VB_Name = \"Module1\"\r\nSub Main()\r\n    Dim a(10) As Long\r\n    a(0) = 1\r\nEnd Sub\r\nSub Foo()\r\n    Dim b(10) As Long\r\n    b(0) = 2\r\nEnd Sub\r\n",
        0x0008);
    assert_eq!(procs[1], &[0xf5, 0x02, 0x00, 0x00, 0x00, 0xf5, 0x00, 0x00, 0x00, 0x00, 0x04, 0x64, 0xff, 0xa3]);
}
