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
