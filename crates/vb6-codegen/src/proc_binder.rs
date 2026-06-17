//! Procedure-level binding driver: declaration, name resolution, and expression
//! tree construction for one procedure under compilation.
//!
//! ## What is ported
//!
//! The local-variable path — `Dim x As T` declaration, frame-offset allocation,
//! and building bound load/store nodes — is complete and tested.
//!
//! ## What requires the symbol table (not yet ported)
//!
//! Global name resolution (`EbBindName` @ 0fab7ad0, 1923 lines), member access
//! (`EbAdjustBoundExpr` @ 0fab8795, `EbResolveMemberAccess3`), object refs
//! (`EbResolveObjectRef`), and proc-level symbols (`EbGetProcEntry`) all require
//! the VBA6 module symbol table and compilation context structures (ECX+0xd8
//! proc entry, ECX+0x2c module table, etc.) to be ported before they can be
//! implemented.  Until that work is done, any attempt to resolve a name that
//! isn't a declared local will panic with `unimplemented!`.
//!
//! ## Name-binding kinds (word\[7\] of a bound name node)
//!
//! After EbBindName resolves a name, `word[7]` of the name node encodes what
//! was found:
//!
//! | word\[7\] | Kind |
//! |----------|------|
//! | 2 | Local variable (frame offset in word\[4\] high 16) |
//! | 3 | Sub/Function call |
//! | 5 | Array element |
//! | 9 | Resolved expression (EbResolveAndAdjustExpr) |
//! | 10 | Object/member reference |

use crate::bind::{DeclError, LocalVar, ProcFrame};
use crate::node::{NodeArena, NodeRef};

/// The binding kind stored in `word[7]` of a name node after EbBindName runs.
/// Only `Local` (2) is directly emittable via `make_load_node`; the others
/// require additional rewriting that depends on the symbol table.
#[derive(Clone, Copy, Debug, PartialEq, Eq)]
#[repr(u32)]
pub enum BindKind {
    Local = 2,
    FuncCall = 3,
    ArrayElem = 5,
    ResolvedExpr = 9,
    ObjectRef = 10,
}

/// Binding result for a single name in a procedure.
#[derive(Clone, Copy, Debug)]
pub struct NameBinding {
    pub kind: BindKind,
    pub var: LocalVar,
}

/// Procedure-level binder.  Owns the frame allocator for one procedure and
/// exposes declaration and resolution APIs.
///
/// To compile a statement like `r = a + b`:
/// 1. Declare the locals: `declare_local("a", 2)`, etc.
/// 2. Build load nodes: `bind_local_load(&mut arena, "a")`.
/// 3. Build expression nodes with `NodeArena::node`.
/// 4. Call `Emitter::emit_expr` on the root.
/// 5. Call `emit_var_store` for the assignment target.
#[derive(Debug)]
pub struct ProcBinder {
    frame: ProcFrame,
}

impl ProcBinder {
    pub fn new() -> Self {
        Self { frame: ProcFrame::new() }
    }

    /// Declare a local variable, allocating its frame slot.
    ///
    /// Returns the allocated `LocalVar` (frame offset + type context) on
    /// success, or `DeclError::AlreadyDeclared` if the name was already seen
    /// in this scope.
    pub fn declare_local(
        &mut self,
        name: &str,
        type_ctx: usize,
    ) -> Result<LocalVar, DeclError> {
        self.frame.declare_local(name, type_ctx)
    }

    /// Resolve a declared local and create a bound load node in `arena`.
    ///
    /// The returned `NodeRef` is a type-0x74 load node: its symbol child
    /// carries the frame offset; its `word[5]` carries the type context.
    /// Pass it directly to `Emitter::emit_expr` to produce the load bytes.
    ///
    /// Returns `None` when the name is not a declared local.  Full module-scope
    /// resolution (globals, proc calls) requires porting EbBindName.
    pub fn bind_local_load(
        &self,
        arena: &mut NodeArena,
        name: &str,
    ) -> Option<NodeRef> {
        self.frame.make_load_node(arena, name)
    }

    /// Resolve a declared local and return its `LocalVar` for use when building
    /// a store (the caller passes `var.type_ctx` and `var.frame_offset` to
    /// `Emitter::emit_var_store`).
    pub fn resolve_local(&self, name: &str) -> Option<LocalVar> {
        self.frame.resolve(name)
    }

    /// Total bytes the frame cursor has moved for locals.  This is the value
    /// that goes into the proc-level frame-size descriptor at `proc+0x74/0x76`.
    pub fn locals_frame_bytes(&self) -> u16 {
        self.frame.locals_frame_bytes()
    }

    /// Bind a name reference: tries locals first; panics with `unimplemented!`
    /// for anything outside the current local scope (globals, proc calls, etc.)
    /// until EbBindName is ported.
    pub fn bind_name(&self, arena: &mut NodeArena, name: &str) -> NodeRef {
        if let Some(load) = self.frame.make_load_node(arena, name) {
            return load;
        }
        unimplemented!(
            "ProcBinder::bind_name: '{}' is not a declared local; \
             global/proc resolution requires EbBindName @ 0fab7ad0",
            name
        );
    }
}

impl Default for ProcBinder {
    fn default() -> Self {
        Self::new()
    }
}

#[cfg(test)]
mod tests {
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
}
