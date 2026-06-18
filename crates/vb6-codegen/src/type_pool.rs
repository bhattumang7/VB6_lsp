//! Type-intern pool — the per-compilation table that maps a distinct type value
//! to the 16-bit index emitted by type coercions/conversions, call nodes
//! (`node[9]`), and member references.
//!
//! In the runtime this is a hash table (`EbHashLookup2`) backing an array whose
//! current length is the next index (`context+0x28`). `EbRegisterTypeInfo2`
//! interns a type value: on first sight it stores the value at `array[count]`,
//! records `count` as the value's index in the hash entry, then increments
//! `count`; a repeat lookup returns the same index. The hash table only locates
//! the entry — it does not affect which index a value gets. The **observable**
//! behaviour is therefore exact: indices are assigned in first-seen order and
//! deduplicated by type value, which is what this models.
//!
//! `EbExtractTypeValue2` (`FUN_0fabd3fb`) is the thin wrapper that interns a type
//! value and returns the index truncated to 16 bits.

use std::collections::HashMap;

/// A type-intern pool: distinct type values mapped to first-seen 16-bit indices.
#[derive(Debug, Default, Clone)]
pub struct TypePool {
    index_of: HashMap<u32, u16>,
    next: u16,
}

impl TypePool {
    pub fn new() -> Self {
        Self::default()
    }

    /// Number of distinct type values interned so far (the next index to assign).
    pub fn len(&self) -> u16 {
        self.next
    }

    pub fn is_empty(&self) -> bool {
        self.next == 0
    }

    /// Intern a type value, returning its index. On first sight the value is
    /// assigned the current count and the count is incremented (`EbRegisterTypeInfo2`);
    /// a repeat returns the previously assigned index.
    pub fn intern(&mut self, type_value: u32) -> u16 {
        if let Some(&i) = self.index_of.get(&type_value) {
            return i;
        }
        let i = self.next;
        self.index_of.insert(type_value, i);
        self.next += 1;
        i
    }

    /// Port of `EbExtractTypeValue2`: intern the type value and return its index
    /// truncated to 16 bits (the value emitted as a coercion/conversion operand).
    pub fn extract_type_value2(&mut self, type_value: u32) -> u16 {
        self.intern(type_value) & 0xffff
    }
}

#[cfg(test)]
#[path = "tests/type_pool_tests.rs"]
mod tests;
