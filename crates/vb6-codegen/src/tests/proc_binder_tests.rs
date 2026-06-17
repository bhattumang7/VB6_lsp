use super::*;
use crate::emit::Emitter;
use crate::node::NodeArena;

fn emit_expr(arena: &NodeArena, root: NodeRef) -> Vec<u8> {
    let mut e = Emitter::new(arena);
    e.emit_expr(root, 0);
    e.into_bytes()
}

#[test]
fn declare_and_bind_single_long() {
    // `Dim a As Long` (typeCtx 2) → frame offset -136.
    let mut binder = ProcBinder::new();
    let v = binder.declare_local("a", 2).unwrap();
    assert_eq!(v.type_ctx, 2);
    assert_eq!(v.frame_offset, -136);
}

#[test]
fn bind_local_load_produces_correct_bytes() {
    // `Dim a As Long` → bind_local_load → emit → [0x6c, 0x78, 0xff]
    let mut binder = ProcBinder::new();
    binder.declare_local("a", 2).unwrap();
    let mut arena = NodeArena::new();
    let load = binder.bind_local_load(&mut arena, "a").unwrap();
    assert_eq!(emit_expr(&arena, load), &[0x6c, 0x78, 0xff]);
}

#[test]
fn bind_local_load_returns_none_for_undeclared() {
    let binder = ProcBinder::new();
    let mut arena = NodeArena::new();
    assert!(binder.bind_local_load(&mut arena, "x").is_none());
}

#[test]
fn full_proc_binder_add_two_longs_and_store() {
    // Simulate compiling: `Dim a As Long : Dim b As Long : Dim r As Long`
    //                      `r = a + b`
    // Expected bytes: load a [0x6c,0x78,0xff] + load b [0x6c,0x74,0xff]
    //                 + ADD Long [0xaa] + store r [0x71,0x70,0xff]
    let mut binder = ProcBinder::new();
    binder.declare_local("a", 2).unwrap(); // Long at -136
    binder.declare_local("b", 2).unwrap(); // Long at -140
    binder.declare_local("r", 2).unwrap(); // Long at -144
    let rv = binder.resolve_local("r").unwrap();
    let mut arena = NodeArena::new();
    let la = binder.bind_local_load(&mut arena, "a").unwrap();
    let lb = binder.bind_local_load(&mut arena, "b").unwrap();
    let add = arena.alloc(NodeArena::node(0x16, 8, la.0, lb.0, 0, 0));
    let mut emitter = Emitter::new(&arena);
    emitter.emit_expr(add, 0);
    emitter.emit_var_store(rv.type_ctx, rv.frame_offset);
    assert_eq!(
        emitter.into_bytes(),
        &[0x6c, 0x78, 0xff, 0x6c, 0x74, 0xff, 0xaa, 0x71, 0x70, 0xff]
    );
}

#[test]
fn locals_frame_bytes_grows_with_declarations() {
    let mut binder = ProcBinder::new();
    assert_eq!(binder.locals_frame_bytes(), 0);
    binder.declare_local("a", 4).unwrap(); // Double: 8 bytes
    assert_eq!(binder.locals_frame_bytes(), 8);
    binder.declare_local("b", 2).unwrap(); // Long: 4 bytes
    assert_eq!(binder.locals_frame_bytes(), 12);
}

// ── Parameter resolution and binding ─────────────────────────────────────────

#[test]
fn declare_param_and_resolve() {
    let mut binder = ProcBinder::new();
    let p = binder.declare_param("p", 2, false).unwrap(); // ByVal Long at +12
    assert_eq!(p.frame_offset, 12);
    assert_eq!(p.type_ctx, 2);
    assert!(!p.byref);
    let v = binder.resolve_param("p").expect("p declared");
    assert_eq!(v.frame_offset, 12);
}

#[test]
fn bind_name_finds_byval_param_emits_load() {
    // bind_name for a ByVal Long param → load opcode 0x6c.
    // Oracle: ByVal Long at +12 → [0x6c, 0x0c, 0x00]. ✓
    let mut binder = ProcBinder::new();
    binder.declare_param("p", 2, false).unwrap();
    let mut arena = NodeArena::new();
    let load = binder.bind_name(&mut arena, "p");
    assert_eq!(emit_expr(&arena, load), &[0x6c, 0x0c, 0x00]);
}

#[test]
fn bind_name_finds_byref_param_emits_byref_load() {
    // bind_name for a ByRef Long param → load opcode 0x80.
    // 0x80 = RT_LOAD_BY_CTX[Long=2] (0x6c) + 0x14. Oracle-confirmed. ✓
    let mut binder = ProcBinder::new();
    binder.declare_param("p", 2, true).unwrap();
    let mut arena = NodeArena::new();
    let load = binder.bind_name(&mut arena, "p");
    assert_eq!(emit_expr(&arena, load), &[0x80, 0x0c, 0x00]);
}

#[test]
fn bind_name_prefers_local_over_param() {
    // When a local and a param share the same name, the local takes priority
    // (locals shadow params in the resolution order: locals → params → globals).
    // Local Long at -136 → [0x6c, 0x78, 0xff].
    let mut binder = ProcBinder::new();
    binder.declare_param("x", 2, false).unwrap(); // ByVal Long param at +12
    binder.declare_local("x", 2).unwrap();        // Long local at -136
    let mut arena = NodeArena::new();
    let load = binder.bind_name(&mut arena, "x");
    assert_eq!(emit_expr(&arena, load), &[0x6c, 0x78, 0xff]);
}

// ── Global resolution and binding ─────────────────────────────────────────────

#[test]
fn declare_global_and_resolve() {
    let mut binder = ProcBinder::new();
    let g = binder.declare_global("g", 2).unwrap(); // Long
    assert_eq!(g.type_ctx, 2);
    assert_eq!(g.field_offset, 0);
    assert_eq!(g.module_desc, 0x0008);
    let v = binder.resolve_global("g").expect("g declared");
    assert_eq!(v.field_offset, 0);
}

#[test]
fn bind_name_finds_global_long_emits_global_load() {
    // bind_name for a module-level Long global → [0x94, 0x08, 0x00, 0x00, 0x00].
    // 0x94 = RT_LOAD_BY_CTX[Long=2] (0x6c) + 0x28. Oracle-confirmed. ✓
    let mut binder = ProcBinder::new();
    binder.declare_global("g", 2).unwrap();
    let mut arena = NodeArena::new();
    let load = binder.bind_name(&mut arena, "g");
    assert_eq!(emit_expr(&arena, load), &[0x94, 0x08, 0x00, 0x00, 0x00]);
}

#[test]
fn bind_name_local_shadows_global() {
    // A local variable shadows a global of the same name.
    // Long local at -136 → [0x6c, 0x78, 0xff], not the global load sequence.
    let mut binder = ProcBinder::new();
    binder.declare_global("g", 2).unwrap();
    binder.declare_local("g", 2).unwrap();
    let mut arena = NodeArena::new();
    let load = binder.bind_name(&mut arena, "g");
    assert_eq!(emit_expr(&arena, load), &[0x6c, 0x78, 0xff]);
}

#[test]
fn bind_name_param_shadows_global() {
    // A ByVal param shadows a global of the same name.
    // ByVal Long at +12 → [0x6c, 0x0c, 0x00].
    let mut binder = ProcBinder::new();
    binder.declare_global("g", 2).unwrap();
    binder.declare_param("g", 2, false).unwrap();
    let mut arena = NodeArena::new();
    let load = binder.bind_name(&mut arena, "g");
    assert_eq!(emit_expr(&arena, load), &[0x6c, 0x0c, 0x00]);
}

#[test]
fn with_module_desc_uses_given_desc_word() {
    // ProcBinder::with_module_desc sets the module descriptor for globals.
    let mut binder = ProcBinder::with_module_desc(0x0010);
    binder.declare_global("g", 2).unwrap();
    let v = binder.resolve_global("g").expect("g declared");
    assert_eq!(v.module_desc, 0x0010);
}
