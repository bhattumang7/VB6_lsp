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
        let var = LocalVar {
            type_ctx,
            frame_offset: self.cursor,
        };
        self.vars.insert(name.to_string(), var);
        Ok(var)
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

// ─────────────────────────────────────────────────────────────────────────────

#[cfg(test)]
#[path = "tests/bind_tests.rs"]
mod tests;
