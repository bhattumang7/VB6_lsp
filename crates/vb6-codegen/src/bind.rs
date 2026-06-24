//! Name binding and the proc-compilation context.
//!
//! Two subsystems live here:
//!
//! **`SlotTable`** — the compile-time slot allocator.  Allocates and frees
//! typed descriptor records that the binder uses as scratch space before frame
//! offsets are computed.  Slot IDs are `descriptorIndex << 2`.
//!
//! **`ProcFrame`** — the runtime frame allocator.  Given a sequence of local
//! variable declarations (name + type context), it assigns each variable a
//! signed 16-bit frame offset (relative to the proc's virtual frame pointer)
//! using the alignment and sizing rules confirmed by empirical probes against
//! the real VB6 compiler.
//!
//! ## Frame layout (confirmed by probe)
//! Frame cursor starts at `PROC_FRAME_BASE = -132`.  For each local, if the
//! type's frame size is ≥ 4 the cursor is first rounded down to the nearest
//! multiple of 4, then decremented by the frame size; the result is the
//! variable's frame offset.
//!
//! | typeCtx | Type(s)              | Frame bytes |
//! |---------|----------------------|-------------|
//! | 0       | Object / untyped ptr | 4           |
//! | 1       | Integer, Boolean, Byte | 2         |
//! | 2       | Long                 | 4           |
//! | 3       | Single               | 4           |
//! | 4       | Double               | 8           |
//! | 5       | String (BSTR ptr)    | 4           |
//! | 6       | Currency             | 8           |
//!
//! Variant (16 bytes) and Date (8 bytes) share their type contexts with the
//! indirect-type path (typeCtx 0 / unconfirmed); they are handled via
//! `unimplemented!()` until the mapping is confirmed.

use std::collections::HashMap;

use crate::node::{NodeArena, NodeRef};

/// A 10-byte per-slot descriptor. Only the fields the codegen path reads are
/// named; the remainder is preserved as raw bytes so the record keeps its exact
/// 10-byte footprint.
#[derive(Clone, Copy, Debug, Default, PartialEq, Eq)]
pub struct SlotDesc {
    bytes: [u8; 10],
}

impl SlotDesc {
    /// Type-flag word at descriptor `+2`: bit 6 (0x40) = assigned; bits 10-13 =
    /// VT code; assignment preserves the other bits.
    pub fn type_flags(&self) -> u16 {
        u16::from_le_bytes([self.bytes[2], self.bytes[3]])
    }
    fn set_type_flags(&mut self, v: u16) {
        self.bytes[2..4].copy_from_slice(&v.to_le_bytes());
    }
    /// Free-list link at descriptor `+4` (next free slot ID, 0xffff = none).
    fn next(&self) -> u16 {
        u16::from_le_bytes([self.bytes[4], self.bytes[5]])
    }
    fn set_next(&mut self, v: u16) {
        self.bytes[4..6].copy_from_slice(&v.to_le_bytes());
    }
}

/// The variable frame + slot descriptors of a procedure under compilation.
///
/// Slot IDs are `descriptorIndex << 2`: allocation returns `(descOffset / 10) << 2`
/// and assignment indexes the frame at the slot's byte offset and the descriptor
/// at `(offset >> 2) * 10`.
#[derive(Debug, Default, Clone)]
pub struct SlotTable {
    descs: Vec<SlotDesc>,
    /// Compile-time frame: one cell per slot, holding the node value an
    /// assignment writes. Index = slotId >> 2.
    frame: Vec<u32>,
    /// Free-list head as a slot ID; 0xffff when empty.
    free_head: u16,
}

const FREE_NONE: u16 = 0xffff;

impl SlotTable {
    pub fn new() -> Self {
        Self {
            descs: Vec::new(),
            frame: Vec::new(),
            free_head: FREE_NONE,
        }
    }

    /// Allocate a slot. When the free list is empty, grow the descriptor array by
    /// a batch (4 records minimum, doubling thereafter — the batch size only
    /// affects when growth happens, not the slot-ID order) and thread the new
    /// records onto the free list in ascending order; then pop the head. Slot ID
    /// = index << 2.
    pub fn allocate_slot(&mut self) -> u16 {
        if self.free_head == FREE_NONE {
            let old = self.descs.len();
            let batch = old.max(4);
            for i in old..old + batch {
                let mut d = SlotDesc::default();
                let next = if i + 1 < old + batch {
                    ((i + 1) as u16) << 2
                } else {
                    FREE_NONE
                };
                d.set_next(next);
                self.descs.push(d);
                self.frame.push(0);
            }
            self.free_head = (old as u16) << 2;
        }
        let idx = (self.free_head >> 2) as usize;
        self.free_head = self.descs[idx].next();
        self.descs[idx].set_next(FREE_NONE); // mark allocated
        (idx as u16) << 2
    }

    /// Return a slot to the free list: push it onto the head so the most-recently
    /// freed slot is reused first.
    pub fn free_slot(&mut self, slot_id: u16) {
        let idx = (slot_id >> 2) as usize;
        self.descs[idx].set_next(self.free_head);
        self.free_head = slot_id;
    }

    /// Store `value` in the slot's frame cell, set the assigned bit (0x40) in the
    /// descriptor's type-flag word, and write the VT code into bits 10-13. The
    /// masks preserve exactly the bits VB6 keeps: `& 0x3d8f | 0x40`, then
    /// `& 0xc3ff | (vt&0xf)<<10`.
    pub fn assign_var_slot(&mut self, slot_id: u16, value: u32, vt: u16) {
        let idx = (slot_id >> 2) as usize;
        self.frame[idx] = value;
        let d = &mut self.descs[idx];
        let mut tf = d.type_flags();
        tf = (tf & 0x3d8f) | 0x40;
        tf = (tf & 0xc3ff) | ((vt & 0xf) << 10);
        d.set_type_flags(tf);
    }

    /// The frame value stored for a slot.
    pub fn frame_value(&self, slot_id: u16) -> u32 {
        self.frame[(slot_id >> 2) as usize]
    }

    /// The descriptor for a slot.
    pub fn desc(&self, slot_id: u16) -> SlotDesc {
        self.descs[(slot_id >> 2) as usize]
    }

    /// Number of descriptor records allocated so far (for tests/inspection).
    pub fn desc_count(&self) -> usize {
        self.descs.len()
    }
}

// ── ProcFrame ────────────────────────────────────────────────────────────────

/// Frame size in bytes for a given type context. The size determines how much
/// the frame cursor decrements per allocation (after 4-byte alignment for
/// sizes ≥ 4).
fn frame_size_of_ctx(type_ctx: usize) -> i16 {
    match type_ctx {
        0 => 4, // Object / untyped pointer
        1 => 2, // Integer, Boolean, Byte
        2 => 4, // Long
        3 => 4, // Single
        4 => 8, // Double
        5 => 4, // String (BSTR pointer)
        6 => 8, // Currency
        10 => 16, // Variant
        _ => unimplemented!(
            "frame_size_of_ctx: typeCtx {} not yet confirmed (Date/Variant \
             use the indirect-type path)",
            type_ctx
        ),
    }
}

/// Frame cursor before the first local is allocated, relative to the proc
/// virtual frame pointer. Confirmed by probing for all numeric types.
pub const PROC_FRAME_BASE: i16 = -132;

/// One local variable's binding: its type context and signed frame offset.
#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub struct LocalVar {
    /// Internal type context (0=Object, 1=Integer/Bool/Byte, 2=Long, 3=Single,
    /// 4=Double, 5=String, 6=Currency).
    pub type_ctx: usize,
    /// Signed frame offset from the proc virtual frame pointer.  Negative for
    /// stack locals.
    pub frame_offset: i16,
}

/// Reason a local declaration fails.
#[derive(Debug, PartialEq, Eq)]
pub enum DeclError {
    AlreadyDeclared,
}

/// Runtime frame allocator for one procedure under compilation.
///
/// Tracks the frame cursor and assigns frame offsets to declared locals.
/// The cursor starts at [`PROC_FRAME_BASE`] and decrements on each allocation.
/// Before allocating any type of frame size ≥ 4, the cursor is rounded down
/// (made more negative) to the nearest multiple of 4.
///
/// # Example
/// ```
/// use vb6_codegen::bind::{ProcFrame, PROC_FRAME_BASE};
/// let mut f = ProcFrame::new();
/// let a = f.declare_local("a", 4 /* Double */).unwrap(); // -140
/// let b = f.declare_local("b", 4).unwrap();              // -148
/// assert_eq!(a.frame_offset, -140);
/// assert_eq!(b.frame_offset, -148);
/// ```
#[derive(Debug)]
pub struct ProcFrame {
    cursor: i16,
    vars: HashMap<String, LocalVar>,
}

impl ProcFrame {
    pub fn new() -> Self {
        Self {
            cursor: PROC_FRAME_BASE,
            vars: HashMap::new(),
        }
    }

    /// Declare a local variable.  Allocates frame space and returns its
    /// `LocalVar`, or `Err(DeclError::AlreadyDeclared)` if the name is already
    /// in scope.
    pub fn declare_local(
        &mut self,
        name: &str,
        type_ctx: usize,
    ) -> Result<LocalVar, DeclError> {
        if self.vars.contains_key(name) {
            return Err(DeclError::AlreadyDeclared);
        }
        let var = self.alloc(type_ctx);
        self.vars.insert(name.to_string(), var);
        Ok(var)
    }

    /// Allocate a frame slot for an unnamed local, returning its `LocalVar`.
    /// Used when locals are identified by declaration index (the binder's
    /// `local_idx`) rather than by name.
    pub fn declare_anon(&mut self, type_ctx: usize) -> LocalVar {
        self.alloc(type_ctx)
    }

    /// Allocate a frame slot of an explicit byte size (4-byte aligned, like
    /// [`Self::alloc`]). Used for fixed-length strings, whose inline buffer is
    /// larger than the 4-byte String pointer slot. The returned `frame_offset`
    /// is the bottom of the slot (the `LdAddr` target).
    pub fn declare_anon_bytes(&mut self, size: i16) -> LocalVar {
        if size >= 4 {
            let rem = self.cursor.rem_euclid(4) as i16;
            if rem != 0 {
                self.cursor -= rem;
            }
        }
        self.cursor -= size;
        LocalVar { type_ctx: 5, frame_offset: self.cursor }
    }

    /// Move the frame cursor for one local of `type_ctx` (4-byte alignment for
    /// sizes ≥ 4, then decrement by the size) and return its `LocalVar`.
    fn alloc(&mut self, type_ctx: usize) -> LocalVar {
        let size = frame_size_of_ctx(type_ctx);
        if size >= 4 {
            // Align cursor down to the nearest multiple of 4.
            // rem_euclid gives a non-negative remainder in [0, 4).
            let rem = self.cursor.rem_euclid(4) as i16;
            if rem != 0 {
                self.cursor -= rem;
            }
        }
        self.cursor -= size;
        LocalVar {
            type_ctx,
            frame_offset: self.cursor,
        }
    }

    /// Resolve a declared local name to its `LocalVar`.
    pub fn resolve(&self, name: &str) -> Option<LocalVar> {
        self.vars.get(name).copied()
    }

    /// Total bytes used by locals so far (unsigned frame growth).
    pub fn locals_frame_bytes(&self) -> u16 {
        (PROC_FRAME_BASE - self.cursor) as u16
    }

    /// Allocate a bound symbol node followed by a typed variable-load node in
    /// `arena`, returning the load node's reference.
    ///
    /// The symbol node (opcode 0) stores the frame offset in `type_info()` =
    /// high 16 bits of `word[4]`.  The load node (opcode 0x74) stores the
    /// type context in `word[5]` and points to the symbol via `word[4]`.
    ///
    /// If `name` was not declared, returns `None`.
    pub fn make_load_node(&self, arena: &mut NodeArena, name: &str) -> Option<NodeRef> {
        let var = self.resolve(name)?;
        let sym = arena.alloc(NodeArena::node(
            0,
            0,
            (var.frame_offset as u16 as u32) << 16,
            0,
            0,
            0,
        ));
        let load = arena.alloc(NodeArena::node(
            0x74,
            0,
            sym.0,
            var.type_ctx as u32,
            0,
            0,
        ));
        Some(load)
    }
}

impl Default for ProcFrame {
    fn default() -> Self {
        Self::new()
    }
}

// ── ParamFrame ────────────────────────────────────────────────────────────────

/// Frame offset of the first parameter slot, relative to the proc virtual
/// frame pointer. Confirmed by oracle: `ByVal p As Long` → opcode loads at
/// frame offset +12.
pub const PROC_PARAM_BASE: i16 = 12;

/// Stack step in bytes for a parameter of a given type context.  Parameters
/// occupy at least one DWORD (4 bytes) of stack space; 8-byte types (Double,
/// Currency) occupy two DWORDs.  This mirrors the standard x86 calling
/// convention: the caller always pushes a whole DWORD per slot, rounded up to
/// the type's natural size if larger.
fn param_step(type_ctx: usize) -> i16 {
    let sz = frame_size_of_ctx(type_ctx);
    // DWORD-align upward: sizes ≤ 4 become 4, sizes > 4 are already multiples of 4.
    ((sz as i16 + 3) & !3).max(4)
}

/// One parameter's binding: its type context, frame offset, and whether it is
/// passed by reference.
#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub struct ParamVar {
    pub type_ctx: usize,
    /// Signed frame offset from the proc virtual frame pointer.  Positive for
    /// parameters (pushed by the caller above the frame pointer).
    pub frame_offset: i16,
    /// `true` if the parameter is ByRef (the slot contains a pointer to the
    /// actual value); `false` for ByVal (the slot contains the value directly).
    pub byref: bool,
}

/// Runtime frame allocator for one procedure's parameter list.
///
/// Assigns each declared parameter a signed positive frame offset, starting at
/// [`PROC_PARAM_BASE`] and incrementing by one or two DWORDs per slot.
///
/// # Example
/// ```
/// use vb6_codegen::bind::{ParamFrame, PROC_PARAM_BASE};
/// let mut f = ParamFrame::new();
/// let p = f.declare_param("p", 2 /* Long */, false).unwrap();  // +12
/// let q = f.declare_param("q", 2,            false).unwrap();  // +16
/// assert_eq!(p.frame_offset, 12);
/// assert_eq!(q.frame_offset, 16);
/// ```
#[derive(Debug)]
pub struct ParamFrame {
    cursor: i16,
    vars: HashMap<String, ParamVar>,
}

impl ParamFrame {
    pub fn new() -> Self {
        Self {
            cursor: PROC_PARAM_BASE,
            vars: HashMap::new(),
        }
    }

    /// Declare a named parameter.  Returns `Err(DeclError::AlreadyDeclared)` if a
    /// parameter with that name has already been declared.
    pub fn declare_param(
        &mut self,
        name: &str,
        type_ctx: usize,
        byref: bool,
    ) -> Result<ParamVar, DeclError> {
        if self.vars.contains_key(name) {
            return Err(DeclError::AlreadyDeclared);
        }
        let var = self.alloc_param(type_ctx, byref);
        self.vars.insert(name.to_string(), var);
        Ok(var)
    }

    /// Allocate a parameter slot by index (declaration order).  Used when
    /// parameters are identified by `param_idx` from `vb6_sema::NameResolution`.
    pub fn declare_anon_param(&mut self, type_ctx: usize, byref: bool) -> ParamVar {
        self.alloc_param(type_ctx, byref)
    }

    fn alloc_param(&mut self, type_ctx: usize, byref: bool) -> ParamVar {
        let offset = self.cursor;
        self.cursor += param_step(type_ctx);
        ParamVar { type_ctx, frame_offset: offset, byref }
    }

    /// Resolve a declared parameter name to its `ParamVar`.
    pub fn resolve(&self, name: &str) -> Option<ParamVar> {
        self.vars.get(name).copied()
    }

    /// Allocate a bound symbol node + a typed load node in `arena`.
    ///
    /// ByVal parameters use synthetic opcode `0x74` (same as locals — the frame
    /// offset is positive, but the opcode is identical). ByRef parameters use
    /// synthetic opcode `0x75`; `emit_expr` routes them to `emit_byref_load`.
    ///
    /// Returns `None` when `name` was not declared.
    pub fn make_load_node(&self, arena: &mut NodeArena, name: &str) -> Option<NodeRef> {
        let var = self.resolve(name)?;
        let sym = arena.alloc(NodeArena::node(
            0,
            0,
            (var.frame_offset as u16 as u32) << 16,
            0,
            0,
            0,
        ));
        // 0x74 = ByVal (local-load opcode), 0x75 = ByRef param load
        let opcode = if var.byref { 0x75 } else { 0x74 };
        let load = arena.alloc(NodeArena::node(
            opcode,
            0,
            sym.0,
            var.type_ctx as u32,
            0,
            0,
        ));
        Some(load)
    }
}

impl Default for ParamFrame {
    fn default() -> Self {
        Self::new()
    }
}

// ── GlobalFrame ───────────────────────────────────────────────────────────────

/// One module-level global variable's binding: its type context, the module
/// descriptor word, and the byte offset within the module's global data block.
#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub struct GlobalVar {
    /// Internal type context (same indices as [`LocalVar::type_ctx`]).
    pub type_ctx: usize,
    /// Module-object descriptor word assigned by the compiled form.  The real
    /// VB6 compiler uses `0x0008` for the first (and typically only) module in
    /// a single-module project; confirmed by oracle probes.
    pub module_desc: u16,
    /// Byte offset of this variable within the module's global data block.
    /// The first declared global starts at 0; each subsequent global advances
    /// by the variable's frame size (4 bytes for Integer/Long/Single/Object,
    /// 8 bytes for Double/Currency).
    pub field_offset: u16,
}

/// Module-level global variable frame allocator.
///
/// Assigns each declared global variable a `field_offset` within the module's
/// data block, starting at 0 and incrementing by the type's frame size.
///
/// # Example
/// ```
/// use vb6_codegen::bind::GlobalFrame;
/// let mut f = GlobalFrame::new(0x0008);
/// let a = f.declare_global("a", 2 /* Long */).unwrap();  // field_offset = 0
/// let b = f.declare_global("b", 2).unwrap();             // field_offset = 4
/// assert_eq!(a.field_offset, 0);
/// assert_eq!(b.field_offset, 4);
/// ```
#[derive(Debug)]
pub struct GlobalFrame {
    module_desc: u16,
    cursor: u16,
    vars: HashMap<String, GlobalVar>,
}

impl GlobalFrame {
    /// Create a new global frame for the module with the given descriptor word.
    pub fn new(module_desc: u16) -> Self {
        Self { module_desc, cursor: 0, vars: HashMap::new() }
    }

    /// Declare a named module-level global.  Returns `Err(DeclError::AlreadyDeclared)`
    /// if the name is already in scope.
    pub fn declare_global(
        &mut self,
        name: &str,
        type_ctx: usize,
    ) -> Result<GlobalVar, DeclError> {
        if self.vars.contains_key(name) {
            return Err(DeclError::AlreadyDeclared);
        }
        let var = self.alloc(type_ctx);
        self.vars.insert(name.to_string(), var);
        Ok(var)
    }

    /// Allocate a global slot by declaration index (anonymous).  Used when
    /// globals are identified by index rather than by name.
    pub fn declare_anon_global(&mut self, type_ctx: usize) -> GlobalVar {
        self.alloc(type_ctx)
    }

    fn alloc(&mut self, type_ctx: usize) -> GlobalVar {
        let offset = self.cursor;
        self.cursor += frame_size_of_ctx(type_ctx) as u16;
        GlobalVar { type_ctx, module_desc: self.module_desc, field_offset: offset }
    }

    /// Resolve a declared global name to its `GlobalVar`.
    pub fn resolve(&self, name: &str) -> Option<GlobalVar> {
        self.vars.get(name).copied()
    }

    /// Allocate a bound global-load node in `arena` for the named variable.
    ///
    /// Uses synthetic opcode `0x77`; `emit_expr` routes it to `emit_global_load`.
    /// The node carries `module_desc` in the low 16 bits of `word[4]` and
    /// `field_offset` in the high 16 bits; `word[5]` holds the type context.
    ///
    /// Returns `None` when `name` was not declared.
    pub fn make_load_node(&self, arena: &mut NodeArena, name: &str) -> Option<NodeRef> {
        let var = self.resolve(name)?;
        let packed = (var.module_desc as u32) | ((var.field_offset as u32) << 16);
        let load = arena.alloc(NodeArena::node(
            0x77,
            0,
            packed,
            var.type_ctx as u32,
            0,
            0,
        ));
        Some(load)
    }
}

impl Default for GlobalFrame {
    fn default() -> Self {
        Self::new(0x0008)
    }
}

// ─────────────────────────────────────────────────────────────────────────────

#[cfg(test)]
#[path = "tests/bind_tests.rs"]
mod tests;
