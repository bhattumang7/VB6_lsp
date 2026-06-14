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
//! Lines are delimited by `\n` (0x0A); a preceding `\r` (CRLF) stays part of the
//! line's byte content, which matches how editors count columns.

use crate::frontend::ast::Span;

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
        let mut line_starts = Vec::with_capacity(src.len() / 32 + 1);
        line_starts.push(0u32);
        for (i, &b) in src.iter().enumerate() {
            if b == b'\n' {
                line_starts.push((i + 1) as u32);
            }
        }
        LineIndex { line_starts, len: src.len() as u32 }
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
mod tests {
    use super::*;

    #[test]
    fn single_line() {
        let idx = LineIndex::new(b"Dim x As Long");
        assert_eq!(idx.position(0), Position { line: 0, character: 0 });
        assert_eq!(idx.position(4), Position { line: 0, character: 4 });
    }

    #[test]
    fn multi_line_lf() {
        let src = b"Sub Foo()\n    Dim y\nEnd Sub\n";
        let idx = LineIndex::new(src);
        // "Dim y" — the 'D' is at byte offset 14 (after "Sub Foo()\n    ").
        let d = src.iter().position(|&b| b == b'D').unwrap() as u32;
        assert_eq!(idx.position(d), Position { line: 1, character: 4 });
        // Start of line 2 ("End Sub").
        let e = "Sub Foo()\n    Dim y\n".len() as u32;
        assert_eq!(idx.position(e), Position { line: 2, character: 0 });
    }

    #[test]
    fn crlf_keeps_cr_in_line() {
        let src = b"AB\r\nCD";
        let idx = LineIndex::new(src);
        // 'C' is line 1, char 0 (the \r\n both belong to line 0's terminator).
        assert_eq!(idx.position(4), Position { line: 1, character: 0 });
        // The \r is line 0, char 2.
        assert_eq!(idx.position(2), Position { line: 0, character: 2 });
    }

    #[test]
    fn round_trip() {
        let src = b"alpha\nbeta gamma\ndelta\n";
        let idx = LineIndex::new(src);
        for off in 0..=src.len() as u32 {
            let pos = idx.position(off);
            assert_eq!(idx.offset(pos), off.min(src.len() as u32), "offset {off}");
        }
    }

    #[test]
    fn column_assumption_cp1252() {
        // 0xE9 = 'é' in Windows-1252, a single byte = single UTF-16 unit.
        // Column math is pure byte distance, so the identifier after it lands
        // at the expected character index.
        let src = b"Dim caf\xE9 As Long";
        let idx = LineIndex::new(src);
        // "As" begins after "Dim café " = 9 bytes.
        let a = src.windows(2).position(|w| w == b"As").unwrap() as u32;
        assert_eq!(idx.position(a), Position { line: 0, character: a });
    }

    #[test]
    fn offset_past_end_clamps() {
        let idx = LineIndex::new(b"abc");
        assert_eq!(idx.position(99), Position { line: 0, character: 3 });
    }

    #[test]
    fn span_to_range() {
        let src = b"Sub Foo()\nEnd Sub\n";
        let idx = LineIndex::new(src);
        let span = Span { start: 4, len: 3 }; // "Foo"
        let r = idx.range(span);
        assert_eq!(r.start, Position { line: 0, character: 4 });
        assert_eq!(r.end, Position { line: 0, character: 7 });
    }
}
