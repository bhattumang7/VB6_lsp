//! Name binding & the proc-compilation context.
//!
//! The binder is not a pure function: it reads and mutates a per-procedure
//! *compilation context* that the front-end builds up — an interned-identifier
//! table, a scope tree, declared-symbol records, the variable frame, and per-slot
//! descriptors. This module implements that context and the binding routines that
//! operate on it.
//!
//! Fidelity: every value that can affect the emitted P-code (slot-ID sequence,
//! descriptor flag bits, binding kinds) follows the VB6 rules exactly. The backing
//! storage is modelled with Rust collections. Final byte-exactness is confirmed
//! against real VB6 output; here the logic is checked behaviorally.

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
}
