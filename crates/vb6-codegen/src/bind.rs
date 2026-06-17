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
mod tests {
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
}
