//! Byte-offset ↔ line/column mapping for LSP positions.
//!
//! The engine works in byte offsets (the parser's [`Span`] is a byte range).
//! LSP positions are `(line, character)` where `character` counts UTF-16 code
//! units from the start of the line.
//!
//! VB6 source is Windows-1252 (single-byte): every byte decodes to exactly one
//! Unicode scalar in the Basic Multilingual Plane, which is exactly one UTF-16
//! code unit. So a byte offset within a line equals its UTF-16 character index,
//! and the whole mapping reduces to tracking line-start offsets. (If the engine
//! is ever fed UTF-8 source this assumption breaks and the column math must
//! switch to counting code units — see `column_assumption` test.)
//!
//! A line ends at a line feed (0x0A), at a carriage return (0x0D), or at a
//! carriage return and line feed together, which are one break. The break
//! characters stay part of the line's byte content, which matches how editors
//! count columns.

use crate::frontend::ast::Span;
use crate::frontend::line_reader::physical_line_starts;

/// A zero-based LSP-style position.
#[derive(Clone, Copy, Debug, Default, PartialEq, Eq)]
pub struct Position {
    /// Zero-based line number.
    pub line: u32,
    /// Zero-based UTF-16 code-unit offset from the start of the line.
    pub character: u32,
}

/// A zero-based half-open `(start, end)` position range.
#[derive(Clone, Copy, Debug, Default, PartialEq, Eq)]
pub struct Range {
    pub start: Position,
    pub end: Position,
}

/// Precomputed line-start table over a source buffer.
#[derive(Debug, Clone)]
pub struct LineIndex {
    /// Byte offset of the first character of each line. `line_starts[0] == 0`.
    line_starts: Vec<u32>,
    /// Total source length in bytes (clamps out-of-range offsets).
    len: u32,
}

impl LineIndex {
    /// Build a line index from source bytes.
    pub fn new(src: &[u8]) -> Self {
        LineIndex { line_starts: physical_line_starts(src), len: src.len() as u32 }
    }

    /// Convert a byte offset to a `(line, character)` position.
    ///
    /// Offsets past the end clamp to the end of the buffer.
    pub fn position(&self, offset: u32) -> Position {
        let offset = offset.min(self.len);
        // Greatest line whose start is <= offset.
        let line = match self.line_starts.binary_search(&offset) {
            Ok(exact) => exact,
            Err(next) => next - 1, // `next` is the first start > offset; back up one
        };
        Position {
            line: line as u32,
            character: offset - self.line_starts[line],
        }
    }

    /// Convert a `(line, character)` position back to a byte offset.
    ///
    /// Out-of-range lines/characters clamp to the buffer length.
    pub fn offset(&self, pos: Position) -> u32 {
        let Some(&start) = self.line_starts.get(pos.line as usize) else {
            return self.len;
        };
        (start + pos.character).min(self.len)
    }

    /// Convert a byte [`Span`] to an LSP [`Range`].
    pub fn range(&self, span: Span) -> Range {
        Range {
            start: self.position(span.start),
            end: self.position(span.start + span.len),
        }
    }

    /// Number of lines in the source.
    pub fn line_count(&self) -> usize {
        self.line_starts.len()
    }
}

#[cfg(test)]
#[path = "tests/line_index.rs"]
mod tests;
