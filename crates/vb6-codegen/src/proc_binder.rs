//! Procedure-level binding driver: declaration, name resolution, and expression
//! tree construction for one procedure under compilation.
//!
//! Covers locals, parameters, and module-level globals.  Member access, object
//! references, and proc-level symbols all require the module symbol table and
//! compilation-context structures and remain unimplemented.
//!
//! ## Name-binding kinds (word\[7\] of a bound name node)
//!
//! After a name is resolved, `word[7]` of the name node encodes what was found:
//!
//! | word\[7\] | Kind |
//! |----------|------|
//! | 2 | Local variable (frame offset in word\[4\] high 16) |
//! | 3 | Sub/Function call |
//! | 5 | Array element |
//! | 9 | Resolved expression |
//! | 10 | Object/member reference |

use crate::bind::{DeclError, GlobalFrame, GlobalVar, LocalVar, ParamFrame, ParamVar, ProcFrame};
use crate::node::{NodeArena, NodeRef};

/// The binding kind stored in `word[7]` of a name node after name resolution.
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

/// Procedure-level binder.  Owns the frame allocators for one procedure and
/// exposes declaration and resolution APIs for locals and parameters.
///
/// To compile a statement like `r = a + b`:
/// 1. Declare params: `declare_param("p", 2, false)`, etc.
/// 2. Declare locals: `declare_local("a", 2)`, etc.
/// 3. Build load nodes: `bind_local_load(&mut arena, "a")`.
/// 4. Build expression nodes with `NodeArena::node`.
/// 5. Call `Emitter::emit_expr` on the root.
/// 6. Call `emit_var_store` or `emit_byval_param_store` for the assignment
///    target.
#[derive(Debug)]
pub struct ProcBinder {
    frame: ProcFrame,
    params: ParamFrame,
    globals: GlobalFrame,
}

impl ProcBinder {
    pub fn new() -> Self {
        Self { frame: ProcFrame::new(), params: ParamFrame::new(), globals: GlobalFrame::default() }
    }

    /// Create a binder with a specific module descriptor for global allocation.
    /// `module_desc` must be the compiled module-object descriptor offset
    /// (oracle-confirmed 0x0008 for single-module programs).
    pub fn with_module_desc(module_desc: u16) -> Self {
        Self {
            frame: ProcFrame::new(),
            params: ParamFrame::new(),
            globals: GlobalFrame::new(module_desc),
        }
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

    /// Declare a parameter, allocating its frame slot (positive offset).
    ///
    /// Parameters must be declared in left-to-right order (i.e., the order they
    /// appear in the `Sub`/`Function` signature), before any locals.
    pub fn declare_param(
        &mut self,
        name: &str,
        type_ctx: usize,
        byref: bool,
    ) -> Result<ParamVar, DeclError> {
        self.params.declare_param(name, type_ctx, byref)
    }

    /// Resolve a declared local and create a bound load node in `arena`.
    ///
    /// The returned `NodeRef` is a type-0x74 load node: its symbol child
    /// carries the frame offset; its `word[5]` carries the type context.
    /// Pass it directly to `Emitter::emit_expr` to produce the load bytes.
    ///
    /// Returns `None` when the name is not a declared local.  Full module-scope
    /// resolution (globals, proc calls) requires the module symbol table.
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

    /// Resolve a declared parameter and return its `ParamVar`.
    pub fn resolve_param(&self, name: &str) -> Option<ParamVar> {
        self.params.resolve(name)
    }

    /// Declare a module-level global variable, allocating its field slot in the
    /// global data block.
    pub fn declare_global(
        &mut self,
        name: &str,
        type_ctx: usize,
    ) -> Result<GlobalVar, DeclError> {
        self.globals.declare_global(name, type_ctx)
    }

    /// Resolve a declared global and return its `GlobalVar`.
    pub fn resolve_global(&self, name: &str) -> Option<GlobalVar> {
        self.globals.resolve(name)
    }

    /// Total bytes the frame cursor has moved for locals.  This is the value
    /// that goes into the proc-level frame-size descriptor at `proc+0x74/0x76`.
    pub fn locals_frame_bytes(&self) -> u16 {
        self.frame.locals_frame_bytes()
    }

    /// Bind a name reference: tries locals, then parameters, then module globals.
    /// Panics with `unimplemented!` for member refs, object refs, proc-level
    /// symbols, and built-ins — those require the full module symbol table.
    pub fn bind_name(&self, arena: &mut NodeArena, name: &str) -> NodeRef {
        if let Some(load) = self.frame.make_load_node(arena, name) {
            return load;
        }
        if let Some(load) = self.params.make_load_node(arena, name) {
            return load;
        }
        if let Some(load) = self.globals.make_load_node(arena, name) {
            return load;
        }
        unimplemented!(
            "ProcBinder::bind_name: '{}' is not a declared local, parameter, or module global; \
             member refs, object refs, proc calls, and built-ins require the full module \
             symbol table",
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
#[path = "tests/proc_binder_tests.rs"]
mod tests;
