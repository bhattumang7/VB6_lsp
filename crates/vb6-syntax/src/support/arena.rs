//! Typed bump arena for compiler-internal node trees.
//!
//! All expression-class AST nodes share the same 40-byte raw layout and are
//! allocated from a single bump arena hung off the compiler context. Here we
//! use a typed `Vec`-backed arena whose items are indexed by a strongly-typed
//! [`NodeId`] instead of a raw pointer.

/// Index into an [`Arena<T>`].
///
/// Replaces a raw `u32 *` pointer used to reference AST nodes.
#[derive(Clone, Copy, Debug, PartialEq, Eq, Hash)]
pub struct NodeId(pub u32);

/// A simple append-only arena that maps `NodeId → T`.
///
/// Allocation is O(amortized 1), access is O(1). No deallocation — the whole
/// arena is dropped when the compiler context owning it is dropped. This is a
/// bump-arena lifetime model.
pub struct Arena<T> {
    nodes: Vec<T>,
}

impl NodeId {
    /// Return the raw index as `u32` — convenience for storing in AST node fields.
    pub fn index(self) -> u32 {
        self.0
    }
}

impl<T> Arena<T> {
    pub fn new() -> Self {
        Self { nodes: Vec::new() }
    }

    /// Append `node` and return its [`NodeId`].
    ///
    /// Conceptually a bump-pointer advance: the cursor into the current arena
    /// page is advanced by the node size and the previous position is returned.
    pub fn alloc(&mut self, node: T) -> NodeId {
        let id = self.nodes.len() as u32;
        self.nodes.push(node);
        NodeId(id)
    }

    /// Borrow a node by id.
    pub fn get(&self, id: NodeId) -> &T {
        &self.nodes[id.0 as usize]
    }

    /// Mutably borrow a node by id.
    pub fn get_mut(&mut self, id: NodeId) -> &mut T {
        &mut self.nodes[id.0 as usize]
    }

    /// Number of allocated nodes.
    pub fn len(&self) -> usize {
        self.nodes.len()
    }

    pub fn is_empty(&self) -> bool {
        self.nodes.is_empty()
    }
}

impl<T> Default for Arena<T> {
    fn default() -> Self {
        Self::new()
    }
}
