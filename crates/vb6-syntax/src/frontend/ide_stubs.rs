//! IDE-layer stub functions: record/undo history and include-file support.
//!
//! These entry points fire during parsing to update the IDE's undo/redo state
//! machine and to handle include-file insertion.  They are orthogonal to the
//! core compilation engine and are **no-ops** in the engine crate.
//!
//! Full implementation deferred to the IDE/UI layer.

// TODO: implement IDE undo/redo recording when the IDE shell is added.

/// Records a "set" action into the IDE undo/redo history. No-op in the engine.
#[allow(unused)]
pub fn record_set(_kind: u32, _target: u32) {}

/// Ends an IDE record action. No-op in the engine.
#[allow(unused)]
pub fn record_end() {}

/// Records a user-initiated undo action. No-op in the engine.
#[allow(unused)]
pub fn record_user_action() {}

/// Erases a user-initiated undo action. No-op in the engine.
#[allow(unused)]
pub fn record_erase_user_action() {}

/// Begins an IDE record action with the given kind. No-op in the engine.
#[allow(unused)]
pub fn record_begin_action(_kind: u32) {}

/// Erases a record action. No-op in the engine.
#[allow(unused)]
pub fn record_erase_action(_kind: u32) {}

/// Records a source line into the IDE undo log. No-op in the engine.
#[allow(unused)]
pub fn record_line(_line_ptr: u32) {}

/// Inserts an include-file at the current parser position. No-op in the
/// engine; include-file support is an IDE feature deferred to the IDE/UI layer.
#[allow(unused)]
pub fn insert_file(_param1: u32, _param2: u32) {}
