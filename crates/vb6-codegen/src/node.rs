//! Raw VB6 expression-node model — the 40-byte records the emitter walks.
//!
//! VB6 expression nodes are fixed 40-byte (ten 32-bit word) records. The code
//! generator reads them positionally:
//!
//! * `word[0]` low 16 bits = opcode / node-type; high 16 bits = type tag
//! * `word[1]` low 16 bits = flags
//! * `word[4]` = `lhs` — a child reference for operators, or inline literal
//!   payload for literals (the low 4 bytes of an 8-byte value)
//! * `word[5]` = `rhs` — child reference, or the high 4 bytes of a literal /
//!   a string length
//! * `word[6]` = extra (step, child ptr, string-data ptr, …)
//!
//! Whether `word[4]`/`word[5]` are child references or raw payload depends on the
//! opcode: for an operator they are node references; for a literal they are an
//! immediate. We keep the same raw representation so the emitter makes the same
//! per-opcode decision the format requires.
//!
//! Children are stored as [`NodeRef`] indices into a [`NodeArena`]; `NodeRef(0)`
//! is the null reference (an absent child).

/// A reference to a node in a [`NodeArena`]. `NodeRef(0)` is null (absent child).
#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub struct NodeRef(pub u32);

impl NodeRef {
    pub const NULL: NodeRef = NodeRef(0);
    pub fn is_null(self) -> bool {
        self.0 == 0
    }
}

/// A raw VB6 expression node: ten 32-bit words, the 40-byte record the code
/// generator walks.
#[derive(Clone, Copy, Debug, Default, PartialEq, Eq)]
pub struct RawNode {
    pub w: [u32; 10],
}

impl RawNode {
    /// Node type / opcode = low 16 bits of `word[0]`, read signed.
    pub fn opcode(&self) -> i32 {
        (self.w[0] as u16 as i16) as i32
    }

    /// Type tag = high 16 bits of `word[0]`, read signed.
    pub fn type_tag(&self) -> i32 {
        ((self.w[0] >> 16) as u16 as i16) as i32
    }

    /// Flags = low 16 bits of `word[1]`.
    pub fn flags(&self) -> u16 {
        self.w[1] as u16
    }

    /// `word[4]` as a child reference.
    pub fn lhs(&self) -> NodeRef {
        NodeRef(self.w[4])
    }

    /// `word[5]` as a child reference.
    pub fn rhs(&self) -> NodeRef {
        NodeRef(self.w[5])
    }

    /// Raw word access.
    pub fn word(&self, i: usize) -> u32 {
        self.w[i]
    }

    /// The 2-byte type-info operand at `node + 0x12` (high half of `word[4]`),
    /// emitted after a typed opcode.
    pub fn type_info(&self) -> u16 {
        (self.w[4] >> 16) as u16
    }

    /// The type-pool index at `node + 0x10` (low half of `word[4]`), used to
    /// follow a type indirection when resolving a typed opcode.
    pub fn type_pool_index(&self) -> u16 {
        self.w[4] as u16
    }

    /// Whether `node + 0x14` (low byte of `word[5]`) has bit 0 set — the
    /// "indirect type" flag checked when resolving a typed opcode.
    pub fn has_indirect_type(&self) -> bool {
        (self.w[5] & 1) != 0
    }

    /// The 8-byte literal payload stored at `word[4]`/`word[5]` (`node + 4`),
    /// little-endian, as read for an 8-byte literal.
    pub fn literal8(&self) -> [u8; 8] {
        let mut out = [0u8; 8];
        out[0..4].copy_from_slice(&self.w[4].to_le_bytes());
        out[4..8].copy_from_slice(&self.w[5].to_le_bytes());
        out
    }

    /// The 8-byte literal payload interpreted as an IEEE-754 double (`node + 4`),
    /// the source for the Single-literal conversion.
    pub fn literal_f64(&self) -> f64 {
        f64::from_le_bytes(self.literal8())
    }
}

/// An arena of raw nodes. Index 0 is a reserved null sentinel so that
/// [`NodeRef::NULL`] round-trips and child references use a null index.
#[derive(Debug, Default, Clone)]
pub struct NodeArena {
    nodes: Vec<RawNode>,
    /// Backing bytes for string/data literals. A node references its bytes by a
    /// byte offset stored in `word[6]` plus a logical length in `word[5]`.
    blobs: Vec<u8>,
}

impl NodeArena {
    pub fn new() -> Self {
        // nodes[0] is the null sentinel; never returned by alloc.
        Self {
            nodes: vec![RawNode::default()],
            blobs: Vec::new(),
        }
    }

    /// Append literal bytes to the blob store, returning their byte offset (the
    /// value a node carries in `word[6]`).
    pub fn alloc_blob(&mut self, bytes: &[u8]) -> u32 {
        let off = self.blobs.len() as u32;
        self.blobs.extend_from_slice(bytes);
        off
    }

    /// Read `len` bytes of literal data at byte `offset` in the blob store.
    pub fn blob(&self, offset: u32, len: usize) -> &[u8] {
        let o = offset as usize;
        &self.blobs[o..o + len]
    }

    /// Allocate a node, returning its non-null reference.
    pub fn alloc(&mut self, node: RawNode) -> NodeRef {
        let idx = self.nodes.len() as u32;
        self.nodes.push(node);
        NodeRef(idx)
    }

    /// Borrow a node by reference. Panics on the null reference.
    pub fn get(&self, r: NodeRef) -> &RawNode {
        assert!(!r.is_null(), "NodeArena::get on null NodeRef");
        &self.nodes[r.0 as usize]
    }

    // ── Test/lowering construction helpers ───────────────────────────────────

    /// Build a node from its opcode (node type), type tag, and the four payload
    /// words `word[4..8]`.
    pub fn node(opcode: u16, type_tag: u16, w4: u32, w5: u32, w6: u32, w7: u32) -> RawNode {
        let mut n = RawNode::default();
        n.w[0] = (opcode as u32) | ((type_tag as u32) << 16);
        n.w[4] = w4;
        n.w[5] = w5;
        n.w[6] = w6;
        n.w[7] = w7;
        n
    }
}
