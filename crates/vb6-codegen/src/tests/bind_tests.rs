use super::*;

#[test]
fn slots_allocate_in_sequence_of_four() {
    let mut t = SlotTable::new();
    let ids: Vec<u16> = (0..6).map(|_| t.allocate_slot()).collect();
    assert_eq!(ids, vec![0x00, 0x04, 0x08, 0x0c, 0x10, 0x14]);
}

#[test]
fn freed_slot_is_reused_before_fresh_growth() {
    let mut t = SlotTable::new();
    let a = t.allocate_slot(); // 0
    let b = t.allocate_slot(); // 4
    let _c = t.allocate_slot(); // 8
    assert_eq!((a, b), (0x00, 0x04));
    t.free_slot(b); // return slot 4
    assert_eq!(t.allocate_slot(), 0x04); // reused (LIFO free list)
    assert_eq!(t.allocate_slot(), 0x0c); // then the next fresh slot
}

#[test]
fn assign_var_slot_sets_assigned_bit_and_vt() {
    let mut t = SlotTable::new();
    let s = t.allocate_slot();
    // VT 3 (Long): assigned bit 0x40 set, VT in bits 10-13 -> 3<<10 = 0xc00.
    t.assign_var_slot(s, 0xdead_beef, 3);
    let tf = t.desc(s).type_flags();
    assert_eq!(tf & 0x40, 0x40, "assigned bit");
    assert_eq!((tf >> 10) & 0xf, 3, "VT code");
    assert_eq!(t.frame_value(s), 0xdead_beef);
}

#[test]
fn assign_var_slot_masks_preserve_only_kept_bits() {
    // Starting from all-ones type flags, the two masks must reduce to exactly
    // (0x3d8f & 0xc3ff) | 0x40 | (vt<<10). For vt=0 that is 0x018f | 0x40.
    let mut t = SlotTable::new();
    let s = t.allocate_slot();
    t.descs[0].set_type_flags(0xffff);
    t.assign_var_slot(s, 0, 0);
    let expected = (((0xffffu16 & 0x3d8f) | 0x40) & 0xc3ff) | 0;
    assert_eq!(t.desc(s).type_flags(), expected);
}

// ── ProcFrame tests (probe-verified) ─────────────────────────────────────

#[test]
fn proc_frame_first_integer_at_minus_134() {
    // typeCtx 1 (Integer, 2 bytes, no alignment). P0 - 2 = -132 - 2 = -134.
    // Probe: `Dim a As Integer` → a at 0xff7a = -134. ✓
    let mut f = ProcFrame::new();
    let v = f.declare_local("a", 1).unwrap();
    assert_eq!(v.type_ctx, 1);
    assert_eq!(v.frame_offset, -134);
}

#[test]
fn proc_frame_first_long_at_minus_136() {
    // typeCtx 2 (Long, 4 bytes). P0=-132 is 4-aligned; cursor -132-4 = -136.
    // Probe: `Dim a As Long` → a at 0xff78 = -136. ✓
    let mut f = ProcFrame::new();
    let v = f.declare_local("a", 2).unwrap();
    assert_eq!(v.frame_offset, -136);
}

#[test]
fn proc_frame_first_single_at_minus_136() {
    // typeCtx 3 (Single, 4 bytes). Same alignment as Long.
    // Probe: `Dim a As Single` → a at 0xff78 = -136. ✓
    let mut f = ProcFrame::new();
    let v = f.declare_local("a", 3).unwrap();
    assert_eq!(v.frame_offset, -136);
}

#[test]
fn proc_frame_first_double_at_minus_140() {
    // typeCtx 4 (Double, 8 bytes). P0=-132 is 4-aligned; cursor -132-8=-140.
    // Probe: `Dim a As Double` → a at 0xff74 = -140. ✓
    let mut f = ProcFrame::new();
    let v = f.declare_local("a", 4).unwrap();
    assert_eq!(v.frame_offset, -140);
}

#[test]
fn proc_frame_first_currency_at_minus_140() {
    // typeCtx 6 (Currency, 8 bytes). Same as Double.
    // Probe: `Dim a As Currency` → a at 0xff74 = -140. ✓
    let mut f = ProcFrame::new();
    let v = f.declare_local("a", 6).unwrap();
    assert_eq!(v.frame_offset, -140);
}

#[test]
fn proc_frame_four_doubles_match_probe() {
    // Probe: `Dim a, b, c, r As Double` (all Doubles) in the 4-Double Sub:
    // a=-140, b=-148, c=-156, r=-164.
    let mut f = ProcFrame::new();
    let a = f.declare_local("a", 4).unwrap();
    let b = f.declare_local("b", 4).unwrap();
    let c = f.declare_local("c", 4).unwrap();
    let r = f.declare_local("r", 4).unwrap();
    assert_eq!(a.frame_offset, -140, "a");
    assert_eq!(b.frame_offset, -148, "b");
    assert_eq!(c.frame_offset, -156, "c");
    assert_eq!(r.frame_offset, -164, "r");
}

#[test]
fn proc_frame_integer_then_double_matches_probe() {
    // Probe: `Dim a As Integer, b As Double`:
    // a Integer at -134; cursor=-134, align to -136, b Double: -136-8=-144.
    // Probe confirmed: Double b at 0xff70 = -144. ✓
    let mut f = ProcFrame::new();
    let a = f.declare_local("a", 1).unwrap(); // Integer
    let b = f.declare_local("b", 4).unwrap(); // Double
    assert_eq!(a.frame_offset, -134);
    assert_eq!(b.frame_offset, -144);
}

#[test]
fn proc_frame_long_then_double_matches_probe() {
    // `Dim a As Long, b As Double`: a Long at -136; cursor already 4-aligned;
    // b Double: -136-8=-144. Probe: 0xff70 = -144. ✓
    let mut f = ProcFrame::new();
    let a = f.declare_local("a", 2).unwrap(); // Long
    let b = f.declare_local("b", 4).unwrap(); // Double
    assert_eq!(a.frame_offset, -136);
    assert_eq!(b.frame_offset, -144);
}

#[test]
fn proc_frame_string_then_integer_matches_probe() {
    // `Dim a As String, b As Integer`: String (4 bytes) at -136; Integer b -138.
    // Probe confirmed: Integer b at 0xff76 = -138. ✓
    let mut f = ProcFrame::new();
    let a = f.declare_local("a", 5).unwrap(); // String
    let b = f.declare_local("b", 1).unwrap(); // Integer
    assert_eq!(a.frame_offset, -136);
    assert_eq!(b.frame_offset, -138);
}

#[test]
fn proc_frame_string_then_long_matches_probe() {
    // `Dim a As String, b As Long`: String at -136; Long b: cursor=-136
    // (already 4-aligned), b = -136-4 = -140. Probe: 0xff74 = -140. ✓
    let mut f = ProcFrame::new();
    let a = f.declare_local("a", 5).unwrap(); // String
    let b = f.declare_local("b", 2).unwrap(); // Long
    assert_eq!(a.frame_offset, -136);
    assert_eq!(b.frame_offset, -140);
}

#[test]
fn proc_frame_byte_same_size_as_integer() {
    // Byte uses typeCtx 1 (same as Integer, 2-byte frame slot).
    // Probe: `Dim a As Byte, b As Integer` → Integer b at 0xff78 = -136
    // (a Byte at -134, b Integer at -134-2=-136). ✓
    let mut f = ProcFrame::new();
    let a = f.declare_local("a", 1).unwrap(); // Byte → typeCtx 1
    let b = f.declare_local("b", 1).unwrap(); // Integer → typeCtx 1
    assert_eq!(a.frame_offset, -134);
    assert_eq!(b.frame_offset, -136);
}

#[test]
fn proc_frame_resolve_returns_declared_var() {
    let mut f = ProcFrame::new();
    f.declare_local("x", 4).unwrap();
    let v = f.resolve("x").expect("x should resolve");
    assert_eq!(v.type_ctx, 4);
    assert_eq!(v.frame_offset, -140);
    assert!(f.resolve("y").is_none());
}

#[test]
fn proc_frame_redeclare_returns_error() {
    let mut f = ProcFrame::new();
    f.declare_local("x", 4).unwrap();
    assert_eq!(f.declare_local("x", 2), Err(DeclError::AlreadyDeclared));
}

#[test]
fn proc_frame_locals_frame_bytes_grows_with_allocations() {
    let mut f = ProcFrame::new();
    assert_eq!(f.locals_frame_bytes(), 0);
    f.declare_local("a", 4).unwrap(); // Double: 8 bytes + 0 align bytes
    assert_eq!(f.locals_frame_bytes(), 8);
    f.declare_local("b", 1).unwrap(); // Integer: 2 bytes (no align from -140)
    assert_eq!(f.locals_frame_bytes(), 10);
}

// ── make_load_node / bind→emit integration ───────────────────────────────

#[test]
fn make_load_node_produces_double_load_at_correct_offset() {
    // Probe: `Dim a As Double` → a at frame offset -140 (0xff74).
    // make_load_node should produce a node that the emitter turns into
    // [0x6f, 0x74, 0xff].
    use crate::emit::Emitter;
    let mut f = ProcFrame::new();
    f.declare_local("a", 4).unwrap(); // Double at -140
    let mut arena = NodeArena::new();
    let load = f.make_load_node(&mut arena, "a").expect("a declared");
    let mut emitter = Emitter::new(&arena);
    emitter.emit_expr(load, 0);
    assert_eq!(emitter.into_bytes(), &[0x6f, 0x74, 0xff]);
}

#[test]
fn make_load_node_unknown_name_returns_none() {
    let f = ProcFrame::new();
    let mut arena = NodeArena::new();
    assert!(f.make_load_node(&mut arena, "x").is_none());
}

#[test]
fn make_load_node_long_at_correct_offset() {
    // `Dim n As Long` → Long at -136 (0xff78). Emitter: [0x6c, 0x78, 0xff].
    use crate::emit::Emitter;
    let mut f = ProcFrame::new();
    f.declare_local("n", 2).unwrap(); // Long at -136
    let mut arena = NodeArena::new();
    let load = f.make_load_node(&mut arena, "n").expect("n declared");
    let mut emitter = Emitter::new(&arena);
    emitter.emit_expr(load, 0);
    assert_eq!(emitter.into_bytes(), &[0x6c, 0x78, 0xff]);
}

#[test]
fn make_load_node_integer_at_correct_offset() {
    // `Dim a As Integer, b As Integer` → a at -134, b at -136.
    use crate::emit::Emitter;
    let mut f = ProcFrame::new();
    f.declare_local("a", 1).unwrap(); // Integer at -134
    f.declare_local("b", 1).unwrap(); // Integer at -136
    let mut arena = NodeArena::new();
    let la = f.make_load_node(&mut arena, "a").unwrap();
    let lb = f.make_load_node(&mut arena, "b").unwrap();
    // a: [0x6b, 0x7a, 0xff]; b: [0x6b, 0x78, 0xff]
    let mut emitter = Emitter::new(&arena);
    emitter.emit_expr(la, 0);
    emitter.emit_expr(lb, 0);
    assert_eq!(emitter.into_bytes(), &[0x6b, 0x7a, 0xff, 0x6b, 0x78, 0xff]);
}

#[test]
fn full_pipeline_add_two_longs() {
    // Full bind→emit pipeline for `a + b` where both are Long.
    // Dim a As Long: frame -136 (0xff78); Dim b As Long: frame -140 (0xff74).
    // ADD Long (op 0x16, type_tag 8) → n_opc=144, RT_OPCODE_BYTE[144]=0xaa.
    // Expected bytes:
    //   load a:  [0x6c, 0x78, 0xff]
    //   load b:  [0x6c, 0x74, 0xff]
    //   ADD Long: [0xaa]
    // Total: [0x6c, 0x78, 0xff, 0x6c, 0x74, 0xff, 0xaa]
    use crate::emit::Emitter;
    let mut f = ProcFrame::new();
    f.declare_local("a", 2).unwrap(); // Long at -136
    f.declare_local("b", 2).unwrap(); // Long at -140
    let mut arena = NodeArena::new();
    let la = f.make_load_node(&mut arena, "a").unwrap();
    let lb = f.make_load_node(&mut arena, "b").unwrap();
    // Binary ADD node: op=0x16 (22), type_tag=8 (Long result)
    let add_node = arena.alloc(NodeArena::node(0x16, 8, la.0, lb.0, 0, 0));
    let mut emitter = Emitter::new(&arena);
    emitter.emit_expr(add_node, 0);
    assert_eq!(
        emitter.into_bytes(),
        &[0x6c, 0x78, 0xff, 0x6c, 0x74, 0xff, 0xaa]
    );
}

#[test]
fn full_pipeline_assign_sum_of_two_longs() {
    // `r = a + b` where a, b, r are all Long.
    // Dim a As Long: -136 (0xff78)
    // Dim b As Long: -140 (0xff74)
    // Dim r As Long: -144 (0xff70)
    // Sequence: load a, load b, ADD Long, store r.
    // ADD Long n_opc=144, RT_OPCODE_BYTE[144]=0xaa → [0xaa]
    // Long load opcode: 0x6c; Long store opcode: 0x71 (RT_STORE_BY_CTX[2])
    // Expected: [0x6c,0x78,0xff, 0x6c,0x74,0xff, 0xaa, 0x71,0x70,0xff]
    use crate::emit::Emitter;
    let mut f = ProcFrame::new();
    f.declare_local("a", 2).unwrap(); // Long at -136
    f.declare_local("b", 2).unwrap(); // Long at -140
    f.declare_local("r", 2).unwrap(); // Long at -144
    let rv = f.resolve("r").unwrap();
    let mut arena = NodeArena::new();
    let la = f.make_load_node(&mut arena, "a").unwrap();
    let lb = f.make_load_node(&mut arena, "b").unwrap();
    let add_node = arena.alloc(NodeArena::node(0x16, 8, la.0, lb.0, 0, 0));
    let mut emitter = Emitter::new(&arena);
    emitter.emit_expr(add_node, 0);
    emitter.emit_var_store(rv.type_ctx, rv.frame_offset);
    assert_eq!(
        emitter.into_bytes(),
        &[0x6c, 0x78, 0xff, 0x6c, 0x74, 0xff, 0xaa, 0x71, 0x70, 0xff]
    );
}

#[test]
fn full_pipeline_and_two_integers() {
    // `a And b` where a, b are Integer.
    // Dim a As Integer: -134 (0xff7a); Dim b As Integer: -136 (0xff78).
    // AND Integer (op 0x23=35, type_tag=6): base=0x0021=33, offset=1, n_opc=34,
    // RT_OPCODE_BYTE[34]=0xc4.
    // load a: [0x6b, 0x7a, 0xff]; load b: [0x6b, 0x78, 0xff]; AND: [0xc4]
    use crate::emit::Emitter;
    let mut f = ProcFrame::new();
    f.declare_local("a", 1).unwrap(); // Integer at -134
    f.declare_local("b", 1).unwrap(); // Integer at -136
    let mut arena = NodeArena::new();
    let la = f.make_load_node(&mut arena, "a").unwrap();
    let lb = f.make_load_node(&mut arena, "b").unwrap();
    let and_node = arena.alloc(NodeArena::node(0x23, 6, la.0, lb.0, 0, 0)); // AND, Integer
    let mut emitter = Emitter::new(&arena);
    emitter.emit_expr(and_node, 0);
    assert_eq!(
        emitter.into_bytes(),
        &[0x6b, 0x7a, 0xff, 0x6b, 0x78, 0xff, 0xc4]
    );
}

// ── ParamFrame tests (oracle-verified) ───────────────────────────────────

#[test]
fn param_frame_first_long_at_plus_12() {
    // Oracle: `Sub Foo(ByVal p As Long)` → load p at frame offset +12 (0x000c). ✓
    let mut f = ParamFrame::new();
    let p = f.declare_param("p", 2, false).unwrap();
    assert_eq!(p.frame_offset, 12);
    assert_eq!(p.type_ctx, 2);
    assert!(!p.byref);
}

#[test]
fn param_frame_two_longs_at_12_and_16() {
    // Oracle: `(ByVal p As Long, ByVal q As Long)` → p at +12, q at +16. ✓
    let mut f = ParamFrame::new();
    let p = f.declare_param("p", 2, false).unwrap();
    let q = f.declare_param("q", 2, false).unwrap();
    assert_eq!(p.frame_offset, 12);
    assert_eq!(q.frame_offset, 16);
}

#[test]
fn param_frame_integer_padded_to_dword() {
    // Integer (2 bytes) occupies a full DWORD slot (step=4) in the param area.
    // Oracle: `(ByVal p As Integer, ByVal q As Long)` → q at +16 (+4 from +12). ✓
    let mut f = ParamFrame::new();
    let p = f.declare_param("p", 1, false).unwrap(); // Integer
    let q = f.declare_param("q", 2, false).unwrap(); // Long
    assert_eq!(p.frame_offset, 12);
    assert_eq!(q.frame_offset, 16);
}

#[test]
fn param_frame_double_step_is_eight() {
    // Double (8 bytes) occupies two DWORDs; next param is at +12+8=+20.
    let mut f = ParamFrame::new();
    let p = f.declare_param("p", 4, false).unwrap(); // Double
    let q = f.declare_param("q", 2, false).unwrap(); // Long
    assert_eq!(p.frame_offset, 12);
    assert_eq!(q.frame_offset, 20);
}

#[test]
fn param_frame_byref_flag_stored() {
    // ByRef params have the same frame offsets as ByVal — the byref flag
    // only selects a different opcode at emit time.
    let mut f = ParamFrame::new();
    let p = f.declare_param("p", 2, true).unwrap(); // ByRef Long
    assert_eq!(p.frame_offset, 12);
    assert!(p.byref);
}

#[test]
fn param_frame_redeclare_returns_error() {
    let mut f = ParamFrame::new();
    f.declare_param("p", 2, false).unwrap();
    assert_eq!(f.declare_param("p", 2, false), Err(DeclError::AlreadyDeclared));
}

#[test]
fn param_frame_resolve() {
    let mut f = ParamFrame::new();
    f.declare_param("p", 2, false).unwrap();
    let v = f.resolve("p").expect("p declared");
    assert_eq!(v.frame_offset, 12);
    assert!(f.resolve("q").is_none());
}

// ── ParamFrame::make_load_node ────────────────────────────────────────────────

#[test]
fn param_make_load_node_byval_long_emits_load_opcode() {
    // ByVal Long at +12 → opcode 0x74 node → emit_typed_load → [0x6c, 0x0c, 0x00].
    // Oracle: `Sub Foo(ByVal p As Long) ... r = p`. ✓
    use crate::emit::Emitter;
    let mut f = ParamFrame::new();
    f.declare_param("p", 2, false).unwrap(); // ByVal Long at +12
    let mut arena = NodeArena::new();
    let load = f.make_load_node(&mut arena, "p").expect("p declared");
    let mut e = Emitter::new(&arena);
    e.emit_expr(load, 0);
    assert_eq!(e.into_bytes(), &[0x6c, 0x0c, 0x00]);
}

#[test]
fn param_make_load_node_byref_long_emits_byref_opcode() {
    // ByRef Long at +12 → opcode 0x75 node → emit_byref_load → [0x80, 0x0c, 0x00].
    // 0x80 = RT_LOAD_BY_CTX[2] (0x6c) + 0x14. Oracle: `Sub Foo(ByRef p As Long)`. ✓
    use crate::emit::Emitter;
    let mut f = ParamFrame::new();
    f.declare_param("p", 2, true).unwrap(); // ByRef Long at +12
    let mut arena = NodeArena::new();
    let load = f.make_load_node(&mut arena, "p").expect("p declared");
    let mut e = Emitter::new(&arena);
    e.emit_expr(load, 0);
    assert_eq!(e.into_bytes(), &[0x80, 0x0c, 0x00]);
}

#[test]
fn param_make_load_node_returns_none_for_undeclared() {
    let f = ParamFrame::new();
    let mut arena = NodeArena::new();
    assert!(f.make_load_node(&mut arena, "x").is_none());
}

// ── GlobalFrame ───────────────────────────────────────────────────────────────

#[test]
fn global_frame_first_long_at_offset_zero() {
    // `Public a As Long` → a at field_offset 0.
    let mut f = GlobalFrame::new(0x0008);
    let a = f.declare_global("a", 2).unwrap();
    assert_eq!(a.type_ctx, 2);
    assert_eq!(a.module_desc, 0x0008);
    assert_eq!(a.field_offset, 0);
}

#[test]
fn global_frame_two_longs_at_0_and_4() {
    // `Public a As Long : Public b As Long` → a at 0, b at 4.
    // Oracle: second Long variable at field_offset 4 (probe-confirmed). ✓
    let mut f = GlobalFrame::new(0x0008);
    let a = f.declare_global("a", 2).unwrap();
    let b = f.declare_global("b", 2).unwrap();
    assert_eq!(a.field_offset, 0);
    assert_eq!(b.field_offset, 4);
}

#[test]
fn global_frame_long_then_double() {
    // `Public a As Long : Public b As Double` → a at 0 (4 bytes), b at 4 (8 bytes).
    let mut f = GlobalFrame::new(0x0008);
    let a = f.declare_global("a", 2).unwrap(); // Long: 4 bytes
    let b = f.declare_global("b", 4).unwrap(); // Double: 8 bytes
    assert_eq!(a.field_offset, 0);
    assert_eq!(b.field_offset, 4);
}

#[test]
fn global_frame_redeclare_returns_error() {
    let mut f = GlobalFrame::new(0x0008);
    f.declare_global("g", 2).unwrap();
    assert_eq!(f.declare_global("g", 2), Err(DeclError::AlreadyDeclared));
}

#[test]
fn global_frame_resolve() {
    let mut f = GlobalFrame::new(0x0008);
    f.declare_global("g", 4).unwrap(); // Double at 0
    let v = f.resolve("g").expect("g declared");
    assert_eq!(v.type_ctx, 4);
    assert_eq!(v.field_offset, 0);
    assert!(f.resolve("x").is_none());
}

#[test]
fn global_make_load_node_long_emits_global_opcode() {
    // `Public g As Long` → make_load_node → emit_global_load → [0x94, 0x08, 0x00, 0x00, 0x00].
    // 0x94 = RT_LOAD_BY_CTX[Long=2] (0x6c) + 0x28. Oracle-confirmed. ✓
    use crate::emit::Emitter;
    let mut f = GlobalFrame::new(0x0008);
    f.declare_global("g", 2).unwrap(); // Long at field_offset 0
    let mut arena = NodeArena::new();
    let load = f.make_load_node(&mut arena, "g").expect("g declared");
    let mut e = Emitter::new(&arena);
    e.emit_expr(load, 0);
    assert_eq!(e.into_bytes(), &[0x94, 0x08, 0x00, 0x00, 0x00]);
}

#[test]
fn global_make_load_node_second_long_at_field_4() {
    // `Public a As Long : Public b As Long` → b at field_offset 4.
    // Load b → [0x94, 0x08, 0x00, 0x04, 0x00]. Oracle-confirmed. ✓
    use crate::emit::Emitter;
    let mut f = GlobalFrame::new(0x0008);
    f.declare_global("a", 2).unwrap();
    f.declare_global("b", 2).unwrap();
    let mut arena = NodeArena::new();
    let load = f.make_load_node(&mut arena, "b").expect("b declared");
    let mut e = Emitter::new(&arena);
    e.emit_expr(load, 0);
    assert_eq!(e.into_bytes(), &[0x94, 0x08, 0x00, 0x04, 0x00]);
}

#[test]
fn global_make_load_node_returns_none_for_undeclared() {
    let f = GlobalFrame::new(0x0008);
    let mut arena = NodeArena::new();
    assert!(f.make_load_node(&mut arena, "x").is_none());
}

#[test]
fn global_make_load_node_double_emits_global_opcode() {
    // `Public g As Double` → [0x97, 0x08, 0x00, 0x00, 0x00].
    // 0x97 = RT_LOAD_BY_CTX[Double=4] (0x6f) + 0x28. Oracle-confirmed. ✓
    use crate::emit::Emitter;
    let mut f = GlobalFrame::new(0x0008);
    f.declare_global("g", 4).unwrap();
    let mut arena = NodeArena::new();
    let load = f.make_load_node(&mut arena, "g").expect("g declared");
    let mut e = Emitter::new(&arena);
    e.emit_expr(load, 0);
    assert_eq!(e.into_bytes(), &[0x97, 0x08, 0x00, 0x00, 0x00]);
}
