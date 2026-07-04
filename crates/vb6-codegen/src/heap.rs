//! The module symbol heap — the free-list allocator VB6's declaration compiler
//! uses to lay out the compiled member records (the "bags") the resolver reads.
//!
//! Records are addressed by byte offset into a single heap buffer (the "base").
//! Free space is tracked as an offset-ordered singly linked list of blocks; each
//! block carries an 8-byte header at its start:
//!
//! ```text
//!   +0  u32   next-free offset   (0xffff_ffff = end of list)
//!   +4  u16   size               (usable bytes = total block size - 8)
//!   +6  u16   pad                (always 0)
//! ```
//!
//! The allocator must reproduce VB6's placement byte-for-byte so record offsets
//! match: the resolver dereferences record `+0xc` as an offset into this same
//! heap. This module ports the offset arithmetic exactly; the grow path (which
//! bottoms out in a COM buffer-manager call) and the deferred-free path are gated
//! until their reverse-engineered call-site objects are modelled.

/// End-of-list / null sentinel for a free-list offset.
pub const NIL: u32 = 0xffff_ffff;

/// The failure code `EbAllocateHeapSpace` returns when a request cannot be
/// satisfied (oversized, or growth disabled).
pub const EB_ALLOC_FAILED: i32 = -0x7ff8_fff2;

/// The module heap context — the object VB6 threads as `this` (`in_ECX`) through
/// the allocator family. Only the fields the ported routines touch are modelled.
///
/// * `mem` — the heap buffer; offset `0` is the "base" all block offsets are
///   relative to.
/// * `free_head` — context `+0xc`: the offset of the first free block.
/// * `flags` — context byte `+0x18`: bit `0` selects the in-place coalescing
///   mode (clear ⇒ the gated deferred-free path); bit `1` selects 8-byte
///   alignment (clear ⇒ 2-byte alignment).
#[derive(Clone, Debug, Default, PartialEq, Eq)]
pub struct HeapContext {
    pub mem: Vec<u8>,
    pub free_head: u32,
    pub flags: u8,
    /// Context `+8`: when `0`, an exhausted free list may grow the buffer (the
    /// gated COM path); when non-zero, growth is disabled and an over-capacity
    /// request fails.
    pub buffer_flag: i32,
}

/// Size of the backing buffer a fresh [`HeapContext`] starts with.
const INITIAL_BUFFER_SIZE: usize = 0x80;

impl HeapContext {
    /// Port of the module-heap initializer: allocate the context's starting
    /// buffer and format it as one free block spanning the whole thing.
    ///
    /// `coalesce_mode` sets flag bit `0` (in-place coalescing when set, the
    /// gated deferred-free path when clear); `align8` sets flag bit `1`
    /// (8-byte alignment when set, 2-byte alignment when clear).
    ///
    /// The initial buffer is `0x80` bytes, formatted as a single free block at
    /// offset `0` with usable size `0x78` (`0x80` minus the 8-byte header) and
    /// no successor (`NIL`). Growth is enabled (`buffer_flag = 0`).
    pub fn new(coalesce_mode: bool, align8: bool) -> Self {
        let mut mem = vec![0u8; INITIAL_BUFFER_SIZE];
        let usable = (INITIAL_BUFFER_SIZE - 8) as u16;
        mem[0..4].copy_from_slice(&NIL.to_le_bytes());
        mem[4..6].copy_from_slice(&usable.to_le_bytes());
        mem[6..8].copy_from_slice(&0u16.to_le_bytes());

        let flags = (coalesce_mode as u8) | ((align8 as u8) << 1);

        HeapContext { mem, free_head: 0, flags, buffer_flag: 0 }
    }

    /// Read the next-free offset stored in the block header at `off`.
    pub fn block_next(&self, off: u32) -> u32 {
        let o = off as usize;
        u32::from_le_bytes([self.mem[o], self.mem[o + 1], self.mem[o + 2], self.mem[o + 3]])
    }

    fn set_block_next(&mut self, off: u32, v: u32) {
        let o = off as usize;
        self.mem[o..o + 4].copy_from_slice(&v.to_le_bytes());
    }

    /// Read the usable-size field stored in the block header at `off`.
    pub fn block_size(&self, off: u32) -> u16 {
        let o = off as usize;
        u16::from_le_bytes([self.mem[o + 4], self.mem[o + 5]])
    }

    fn set_block_size(&mut self, off: u32, v: u16) {
        let o = off as usize;
        self.mem[o + 4..o + 6].copy_from_slice(&v.to_le_bytes());
    }

    fn set_block_pad(&mut self, off: u32, v: u16) {
        let o = off as usize;
        self.mem[o + 6..o + 8].copy_from_slice(&v.to_le_bytes());
    }

    /// Port of `EbAlignSize8`: round a requested byte size up to the heap's
    /// alignment, with a minimum of 8. Alignment is 8 bytes when context flag
    /// bit `1` is set, otherwise 2 bytes.
    pub fn align_size8(&self, cb: u32) -> u32 {
        if (self.flags >> 1) & 1 == 0 {
            if cb < 8 {
                8
            } else {
                (cb + 1) & 0xffff_fffe
            }
        } else if cb < 8 {
            8
        } else {
            (cb + 7) & 0xffff_fff8
        }
    }

    /// Port of `EbFindFreeBlock`: carve a block of (at least) `nsize` aligned
    /// bytes out of the free list, returning its offset, or [`NIL`] if no block
    /// fits.
    ///
    /// The list is searched first-fit by usable capacity (`block size + 8`, since
    /// a satisfied request reuses the block's own 8-byte header). A block with at
    /// least 8 spare bytes is split: the tail becomes a fresh free block at
    /// `found + aligned`, and the predecessor (or the list head) is relinked to
    /// it. A block with less than 8 spare bytes is consumed whole and unlinked.
    pub fn find_free_block(&mut self, nsize: u32) -> u32 {
        let aligned = self.align_size8(nsize);
        let aligned_lo = aligned & 0xffff;

        let mut prev = NIL;
        let mut cur = self.free_head;
        while cur != NIL {
            // Capacity that could satisfy the request: usable size + the header.
            if self.block_size(cur) as u32 + 8 >= aligned_lo {
                break;
            }
            prev = cur;
            cur = self.block_next(cur);
        }
        if cur == NIL {
            return NIL;
        }

        let found = cur;
        let stored = self.block_size(found) as u32;
        let leftover = stored.wrapping_sub(aligned).wrapping_add(8);
        let link_target = if leftover >= 8 {
            // Split: the remainder becomes a free block just past the allocation.
            let rem = found + aligned;
            self.set_block_pad(rem, 0);
            self.set_block_next(rem, self.block_next(found));
            self.set_block_size(rem, stored.wrapping_sub(aligned) as u16);
            rem
        } else {
            self.block_next(found)
        };
        if prev == NIL {
            self.free_head = link_target;
        } else {
            self.set_block_next(prev, link_target);
        }
        found
    }

    /// Port of `EbCoalesceMemory` (the in-place mode, context flag bit `0` set):
    /// return the `size`-byte region at offset `address` to the free list,
    /// merging it with an adjacent predecessor and/or successor block where the
    /// two are contiguous and lie in the same 64 KiB segment and the merged
    /// block would still fit in 64 KiB.
    ///
    /// The free list is kept ordered by ascending offset. The deferred-free mode
    /// (flag bit `0` clear, which hands the region to `EbPushStackEntry`) is
    /// gated.
    pub fn coalesce_memory(&mut self, address: u32, size: i32) {
        if self.flags & 1 == 0 {
            unimplemented!(
                "EbCoalesceMemory deferred-free path (context +0x18 bit 0 clear, \
                 EbPushStackEntry); Phase 6"
            );
        }

        let s16 = size as i16;
        let size_u = size as u32;
        let mut cur_off = address; // the region being inserted/merged
        let mut block = address; // the header currently treated as "new"
        let mut prev = NIL; // predecessor offset in the free list
        let mut merged = false;

        self.set_block_pad(address, 0);
        self.set_block_size(block, (s16.wrapping_sub(8)) as u16);

        let mut node = self.free_head;
        let mut last;
        'outer: while node != NIL {
            loop {
                last = node;
                if cur_off < last {
                    if prev == NIL {
                        self.free_head = cur_off;
                        self.set_block_next(block, last);
                    } else {
                        let prev_size = self.block_size(prev) as u32;
                        if (cur_off >> 16 == prev >> 16)
                            && prev.wrapping_add(8).wrapping_add(prev_size) == cur_off
                            && size_u.wrapping_add(8).wrapping_add(prev_size) < 0x1_0001
                        {
                            // Merge the new region onto its predecessor.
                            let merged_size = s16.wrapping_add(self.block_size(prev) as i16);
                            self.set_block_size(prev, merged_size as u16);
                            block = prev;
                            cur_off = prev;
                        } else {
                            self.set_block_next(prev, cur_off);
                            self.set_block_next(block, last);
                        }
                    }
                    // Merge the successor block onto the (possibly grown) new one.
                    let new_size = self.block_size(block) as u32;
                    let last_size = self.block_size(last) as u32;
                    if (cur_off >> 16 == last >> 16)
                        && cur_off.wrapping_add(8).wrapping_add(new_size) == last
                        && new_size.wrapping_add(0x10).wrapping_add(last_size) < 0x1_0001
                    {
                        let combined =
                            (last_size as i16).wrapping_add(8).wrapping_add(self.block_size(block) as i16);
                        self.set_block_size(block, combined as u16);
                        self.set_block_next(block, self.block_next(last));
                    }
                    merged = true;
                    break 'outer;
                }
                node = self.block_next(last);
                prev = last;
                if node == NIL {
                    break;
                }
            }
            // Walked off the end without finding an insertion point: try to merge
            // the region onto the final block.
            let last_size = self.block_size(last) as u32;
            if (last >> 16 == cur_off >> 16)
                && last.wrapping_add(8).wrapping_add(last_size) == cur_off
                && size_u.wrapping_add(8).wrapping_add(last_size) < 0x1_0001
            {
                merged = true;
                let merged_size = s16.wrapping_add(self.block_size(last) as i16);
                self.set_block_size(last, merged_size as u16);
            }
        }

        if !merged {
            self.set_block_next(block, NIL);
            self.set_block_size(block, (s16.wrapping_sub(8)) as u16);
            if prev == NIL {
                self.free_head = cur_off;
            } else {
                self.set_block_next(prev, cur_off);
            }
        }
    }

    /// Read a little-endian dword at an arbitrary heap offset.
    pub fn read_dword(&self, off: u32) -> u32 {
        let o = off as usize;
        u32::from_le_bytes([self.mem[o], self.mem[o + 1], self.mem[o + 2], self.mem[o + 3]])
    }

    /// Write a little-endian dword at an arbitrary heap offset.
    pub fn write_dword(&mut self, off: u32, v: u32) {
        let o = off as usize;
        self.mem[o..o + 4].copy_from_slice(&v.to_le_bytes());
    }

    /// Port of `EbLinkListNode3`: append the record at offset `node_value` to a
    /// singly-linked child list. The list's head pointer lives in the heap at
    /// `head_off` (e.g. a parent record's `+0x28` slot); each node's next pointer
    /// is at the node's `+0x14`; `tail` is the caller's running tail offset
    /// ([`NIL`] when the list is empty).
    pub fn link_list_node3(&mut self, node_value: u32, head_off: u32, tail: &mut u32) {
        if *tail == NIL {
            *tail = node_value;
            self.write_dword(head_off, node_value);
        } else {
            let old_tail = *tail;
            *tail = node_value;
            self.write_dword(old_tail + 0x14, node_value);
        }
    }

    /// Port of `EbAllocateHeapSpace`: allocate `size` aligned bytes, returning the
    /// record's offset into the heap, or [`EB_ALLOC_FAILED`].
    ///
    /// A request is served from the free list ([`find_free_block`]) when one
    /// fits. The grow path (an exhausted free list with growth enabled) is gated:
    /// it calls a COM buffer-manager (a global singleton's vtable) to realloc the
    /// backing buffer, then coalesces the new tail in — modelling that, and the
    /// initial heap state it depends on, is the remaining declaration-compiler
    /// front-half work.
    ///
    /// [`find_free_block`]: HeapContext::find_free_block
    pub fn allocate_heap_space(&mut self, size: u32) -> Result<u32, i32> {
        let aligned = self.align_size8(size);
        if aligned < 0x1_0000 {
            let off = self.find_free_block(aligned);
            if off != NIL {
                return Ok(off);
            }
            if self.buffer_flag == 0 {
                unimplemented!(
                    "EbAllocateHeapSpace grow path: COM buffer-manager realloc \
                     (global singleton vtable +0x10) + the heap-init seeding it \
                     depends on; Phase 6"
                );
            }
        }
        Err(EB_ALLOC_FAILED)
    }

    /// Zero a freshly-allocated record's first `len` bytes — the effect of the
    /// bag allocators' template copies (every bag template is all-zero).
    fn zero_record(&mut self, off: u32, len: usize) {
        let o = off as usize;
        self.mem[o..o + len].fill(0);
    }

    /// Port of `EbAllocateStructure2` (+ its `EbZeroMemory2` init): allocate a
    /// type-structure record sized for `n_elements` slots and zero its live
    /// bytes.
    ///
    /// The allocation is `n*8 + 0x10` bytes for `n != 0`, else `0x18`. The zero
    /// fill covers the leading 0x10 bytes plus 8 per element (`0x10 + 8*n`); for
    /// `n == 0` the trailing 8 bytes are left as the carved block's contents.
    pub fn allocate_structure2(&mut self, n_elements: u32) -> Result<u32, i32> {
        let size = if n_elements != 0 {
            n_elements * 8 + 0x10
        } else {
            0x18
        };
        let off = self.allocate_heap_space(size)?;
        self.zero_record(off, (0x10 + 8 * n_elements) as usize);
        Ok(off)
    }

    /// Port of `EbAllocateMethodBag`: a 0x40-byte member record with the low 3
    /// bits of byte `+0` set to the method tag (4).
    pub fn allocate_method_bag(&mut self) -> Result<u32, i32> {
        let off = self.allocate_heap_space(0x40)?;
        self.zero_record(off, 0x40);
        let o = off as usize;
        self.mem[o] = (self.mem[o] & 0xf8) ^ 4;
        Ok(off)
    }

    /// Port of `EbAllocateInterfaceBag`: a 0x40-byte all-zero member record.
    pub fn allocate_interface_bag(&mut self) -> Result<u32, i32> {
        let off = self.allocate_heap_space(0x40)?;
        self.zero_record(off, 0x40);
        Ok(off)
    }

    /// Port of `EbAllocatePropertyBag`: a 0x1c-byte all-zero member record.
    pub fn allocate_property_bag(&mut self) -> Result<u32, i32> {
        let off = self.allocate_heap_space(0x1c)?;
        self.zero_record(off, 0x1c);
        Ok(off)
    }

    /// Port of `EbAllocateParameterBag`: a 0x28-byte member record with the low 3
    /// bits of byte `+0` set to the parameter tag (2) and bytes `+0x12`/`+0x13`
    /// set to `0xff`.
    pub fn allocate_parameter_bag(&mut self) -> Result<u32, i32> {
        let off = self.allocate_heap_space(0x28)?;
        self.zero_record(off, 0x28);
        let o = off as usize;
        self.mem[o + 0x12] = 0xff;
        self.mem[o + 0x13] = 0xff;
        self.mem[o] = (self.mem[o] & 0xf8) ^ 2;
        Ok(off)
    }

    /// Port of `EbAllocateTypeDescriptor`: a 0x24-byte descriptor whose first
    /// 0x20 bytes are zeroed by the template copy; the trailing 4 bytes
    /// (`+0x20..0x24`) are left as the carved block's existing contents.
    pub fn allocate_type_descriptor(&mut self) -> Result<u32, i32> {
        let off = self.allocate_heap_space(0x24)?;
        self.zero_record(off, 0x20);
        Ok(off)
    }
}

#[cfg(test)]
#[path = "tests/heap_tests.rs"]
mod tests;
