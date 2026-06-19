//! Type-node construction helpers — the routines `EbBuildTypeNode2` uses to set
//! the flag bits on the operand type-node it builds (the mini-p-code the resolver
//! later classifies).
//!
//! A "type node" here is a byte-addressed operand node in the records/emit heap.
//! Only the bytes these routines touch are referenced: `+0` (opcode, low 6 bits
//! select the form), `+1`/`+5` (type-class flag nibbles), `+4` (the inline
//! follow-on opcode for a `0x1d`-form node), and `+2` (the `EbProcessType3`
//! attribute word).

/// Port of `EbToggleBitfield`: set the node's low-6-bit form to `mask` (via the
/// `(mask ^ b) & 0x3f ^ b` toggle), and — when the result is the `0x1d` form —
/// copy `word[0]` into `+4` and stamp the follow-on opcode (`+4 = (+4 & 0xe5) |
/// 0x25`).
pub fn toggle_bitfield(node: &mut [u8], mask: u8) {
    let b = ((mask ^ node[0]) & 0x3f) ^ node[0];
    node[0] = b;
    if b & 0x3f == 0x1d {
        // *(u16)(node+4) = *(u16)(node+0)
        node[4] = node[0];
        node[5] = node[1];
        node[4] = (node[4] & 0xe5) | 0x25;
    }
}

/// Whether a node carries the live `0x1d`-form follow-on opcode that the flag
/// mutators also update (`+0` low6 == `0x1d` and `+4` low6 != `0x25`).
fn has_secondary_flag_slot(node: &[u8]) -> bool {
    node[0] & 0x3f == 0x1d && node[4] & 0x3f != 0x25
}

/// Port of `EbSetTypeFlag4`: set or clear bit `0x08` of the node's type-class
/// nibble at `+1` (and at `+5` for a `0x1d`-form node).
pub fn set_type_flag4(node: &mut [u8], set: bool) {
    let bit = (set as u8) << 3;
    if has_secondary_flag_slot(node) {
        node[5] = bit | (node[5] & 0xf7);
    }
    node[1] = bit | (node[1] & 0xf7);
}

/// Port of `EbToggleTypeFlag`: toggle the low-3-bit type kind at `+1` (and at
/// `+5` for a `0x1d`-form node) toward `enable` (`(enable ^ b) & 7 ^ b`).
pub fn toggle_type_flag(node: &mut [u8], enable: u8) {
    if has_secondary_flag_slot(node) {
        node[5] = ((enable ^ node[5]) & 7) ^ node[5];
    }
    node[1] = ((enable ^ node[1]) & 7) ^ node[1];
}

/// Port of `EbProcessType3` (COM-free branches): set the attribute word at
/// `type_node + 2` according to the type operation.
///
/// The word's bits `0x0d00` are first cleared, then per op: `6` sets `0x80`
/// (kind nibble) and `0x800`; `7` sets `0x80` and `0x100`; `8` (with no slot
/// container, the `pParam4 == 0xffff` case) sets `0x40` and `0x400`. Any other
/// op leaves the cleared word.
///
/// The slot-container sub-paths of ops `8` (real container) and `0xc` resolve a
/// COM ITypeInfo slot (`EbEnsureSlotResolved2` + vtable calls) and are gated.
pub fn process_type3_simple(type_node: &mut [u8], n_type_op: i32) -> Result<(), i32> {
    // *(u16)(node+2) &= 0xf2ff
    let mut w = read_word2(type_node) & 0xf2ff;
    match n_type_op {
        8 => {
            // pParam4 == 0xffff path only (no slot container).
            w |= 0x40;
            w |= 0x400;
        }
        0xc => {
            unimplemented!(
                "EbProcessType3 op 0xc: COM ITypeInfo slot container \
                 (EbEnsureSlotResolved2 + vtable); Phase 6"
            );
        }
        6 => {
            w |= 0x80;
            w |= 0x800;
        }
        7 => {
            w |= 0x80;
            w |= 0x100;
        }
        _ => {
            // word already has 0x80 set in the source for the non-6/7 fallthrough
            // before returning 0; but ops other than 6/7/8/0xc return with only
            // the 0x80 set then 0 — match the source: set 0x80, write, return.
            w |= 0x80;
        }
    }
    write_word2(type_node, w);
    Ok(())
}

/// The inline type-node `EbBuildTypeNode2` produces for a scalar (base-type)
/// declaration with no out-of-line data (`pbData == 0`): a 4-byte node written
/// directly into the caller's output word, whose opcode (low 6 bits of byte 0)
/// is the type code. The variant-like codes `0x0a`/`0x16`/`0x19` remap to the
/// object form `3`. The caller's inline flag (`*piOutput`) is set to `1`.
///
/// Returns the inline node as a dword. The recursion (`0x1a`), struct/array
/// (`0x1b`/`0x1c`) and slot (`0x1d`) type forms take separate, COM-bearing
/// paths and are handled elsewhere; this is the scalar leaf only.
pub fn build_inline_type_node(type_word: u16) -> u32 {
    let code: u8 = if type_word == 0x16 || type_word == 0x19 || type_word == 0x0a {
        3
    } else {
        type_word as u8
    };
    // The node lives in a 4-byte output word; use a 6-byte scratch so the
    // (base-type-unreachable) 0x1d follow-on stamp can't index out of bounds,
    // then return the live 4 bytes.
    let mut node = [0u8; 6];
    toggle_bitfield(&mut node, code);
    u32::from_le_bytes([node[0], node[1], node[2], node[3]])
}

fn read_word2(node: &[u8]) -> u16 {
    u16::from_le_bytes([node[2], node[3]])
}

fn write_word2(node: &mut [u8], v: u16) {
    node[2] = (v & 0xff) as u8;
    node[3] = (v >> 8) as u8;
}

#[cfg(test)]
#[path = "tests/typenode_tests.rs"]
mod tests;
