//! Compiled symbol **member record** — the byte-laid record VB6 builds for each
//! module procedure during declaration, and that the call emitter reads back to
//! decide a call's convention.
//!
//! VB6 keeps these in the module's symbol heap as fixed 0x40-byte records,
//! addressed by byte offset (distinct from the 40-byte *expression* nodes in
//! [`crate::node`]). The declaration compiler writes them; the call-site emitter
//! reads `+0x3a` (convention kind), `+0x01`/`+0x3d` (by-reference mode) and
//! `+0x30` (the member-dispatch id emitted as the call's trailing word).
//!
//! Only the fields the call path uses are modelled here; the record is the full
//! 0x40 bytes so offsets match the source records exactly.

/// A compiled member record (0x40 bytes), byte-addressed like the source.
#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub struct MemberRecord {
    bytes: [u8; 0x40],
}

impl Default for MemberRecord {
    fn default() -> Self {
        Self { bytes: [0; 0x40] }
    }
}

impl MemberRecord {
    pub fn new() -> Self {
        Self::default()
    }

    pub fn byte(&self, off: usize) -> u8 {
        self.bytes[off]
    }

    pub fn set_byte(&mut self, off: usize, v: u8) {
        self.bytes[off] = v;
    }

    pub fn u16(&self, off: usize) -> u16 {
        u16::from_le_bytes([self.bytes[off], self.bytes[off + 1]])
    }

    pub fn set_u16(&mut self, off: usize, v: u16) {
        self.bytes[off..off + 2].copy_from_slice(&v.to_le_bytes());
    }

    /// The convention kind — `record[0x3a]`: bit `0x10` set ⇒ kind 8, else the
    /// low 3 bits. (Port of the `*param == 5` branch of `FUN_0fabac12`.)
    pub fn kind(&self) -> i32 {
        let b = self.byte(0x3a);
        if b & 0x10 == 0 {
            (b & 7) as i32
        } else {
            8
        }
    }

    /// The by-reference passing mode — `record[1] & 7` when `record[0x3d] & 6`
    /// is not `6`, otherwise 0. (Port of the `*param == 5` branch of
    /// `FUN_0fabac47`.)
    pub fn byref(&self) -> i32 {
        if self.byte(0x3d) & 6 != 6 {
            (self.byte(1) & 7) as i32
        } else {
            0
        }
    }

    /// The member-dispatch id — `record[0x30]` — emitted as the call's trailing
    /// word.
    pub fn member_id(&self) -> u16 {
        self.u16(0x30)
    }

    pub fn set_member_id(&mut self, id: u16) {
        self.set_u16(0x30, id);
    }

    /// Pack the type-size class into the high nibble of `record[1]` from a type
    /// code, preserving the low nibble. (Port of `FUN_0fac2536`: the nibble is
    /// `tc==1 ? 1 : tc==2 ? 2 : tc==4 ? 4 : tc==8 ? 8 : 0`.)
    pub fn pack_type_size_class(&mut self, type_code: i32) {
        let nibble = (((((type_code == 4) as u8 | ((type_code == 8) as u8) << 1) << 1
            | (type_code == 2) as u8)
            << 1)
            | (type_code == 1) as u8)
            << 4;
        self.bytes[1] = nibble | (self.bytes[1] & 0xf);
    }
}

/// The callee's compiled type-info descriptor, as the call emitter sees it
/// (`node[4]+0x14`). Its discriminator selects how the convention kind and
/// by-reference mode are read.
#[derive(Clone, Copy, Debug)]
pub enum CalleeTypeInfo<'a> {
    /// Discriminator 1: a type-library item. Kind comes from `FUN_0fbe1daa`,
    /// which is not yet ported.
    TypeLib,
    /// Discriminator 5: a resolved module member — read its record.
    ResolvedMember(&'a MemberRecord),
    /// Any other discriminator (4, …): a plain value callee.
    Default,
}

impl CalleeTypeInfo<'_> {
    /// Convention kind (port of `FUN_0fabac12`).
    pub fn kind(&self) -> i32 {
        match self {
            CalleeTypeInfo::TypeLib => {
                unimplemented!("type-library callee kind (FUN_0fbe1daa); Phase 6")
            }
            CalleeTypeInfo::ResolvedMember(m) => m.kind(),
            CalleeTypeInfo::Default => 4,
        }
    }

    /// By-reference mode (port of `FUN_0fabac47`).
    pub fn byref(&self) -> i32 {
        match self {
            CalleeTypeInfo::ResolvedMember(m) => m.byref(),
            _ => 0,
        }
    }
}

#[cfg(test)]
#[path = "tests/sym_record_tests.rs"]
mod tests;
