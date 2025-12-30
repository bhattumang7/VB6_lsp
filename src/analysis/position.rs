//! Position and Range Types
//!
//! Precise position tracking for symbol table operations.
//! Converts between tree-sitter and LSP position formats.

use std::cmp::Ordering;

/// A precise position in source code (0-indexed line and column)
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash, Default)]
pub struct SourcePosition {
    pub line: u32,
    pub column: u32,
}

impl SourcePosition {
    pub fn new(line: u32, column: u32) -> Self {
        Self { line, column }
    }

    /// Create from tree-sitter Point
    pub fn from_ts_point(point: tree_sitter::Point) -> Self {
        Self {
            line: point.row as u32,
            column: point.column as u32,
        }
    }

    /// Convert to LSP Position
    pub fn to_lsp(&self) -> tower_lsp::lsp_types::Position {
        tower_lsp::lsp_types::Position {
            line: self.line,
            character: self.column,
        }
    }

    /// Create from LSP Position
    pub fn from_lsp(pos: tower_lsp::lsp_types::Position) -> Self {
        Self {
            line: pos.line,
            column: pos.character,
        }
    }
}

impl Ord for SourcePosition {
    fn cmp(&self, other: &Self) -> Ordering {
        match self.line.cmp(&other.line) {
            Ordering::Equal => self.column.cmp(&other.column),
            ord => ord,
        }
    }
}

impl PartialOrd for SourcePosition {
    fn partial_cmp(&self, other: &Self) -> Option<Ordering> {
        Some(self.cmp(other))
    }
}

/// A range in source code with start and end positions
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash, Default)]
pub struct SourceRange {
    pub start: SourcePosition,
    pub end: SourcePosition,
}

impl SourceRange {
    pub fn new(start: SourcePosition, end: SourcePosition) -> Self {
        Self { start, end }
    }

    /// Create from tree-sitter Node
    pub fn from_ts_node(node: &tree_sitter::Node) -> Self {
        Self {
            start: SourcePosition::from_ts_point(node.start_position()),
            end: SourcePosition::from_ts_point(node.end_position()),
        }
    }

    /// Check if this range contains a position
    pub fn contains(&self, pos: SourcePosition) -> bool {
        if pos.line < self.start.line || pos.line > self.end.line {
            return false;
        }
        if pos.line == self.start.line && pos.column < self.start.column {
            return false;
        }
        if pos.line == self.end.line && pos.column > self.end.column {
            return false;
        }
        true
    }

    /// Check if this range contains another range
    pub fn contains_range(&self, other: &SourceRange) -> bool {
        self.contains(other.start) && self.contains(other.end)
    }

    /// Check if this range overlaps with another
    pub fn overlaps(&self, other: &SourceRange) -> bool {
        self.contains(other.start) || self.contains(other.end) ||
        other.contains(self.start) || other.contains(self.end)
    }

    /// Convert to LSP Range
    pub fn to_lsp(&self) -> tower_lsp::lsp_types::Range {
        tower_lsp::lsp_types::Range {
            start: self.start.to_lsp(),
            end: self.end.to_lsp(),
        }
    }

    /// Create from LSP Range
    pub fn from_lsp(range: tower_lsp::lsp_types::Range) -> Self {
        Self {
            start: SourcePosition::from_lsp(range.start),
            end: SourcePosition::from_lsp(range.end),
        }
    }

    /// Calculate the "size" of this range (for finding innermost scope)
    pub fn size(&self) -> u64 {
        let line_diff = self.end.line.saturating_sub(self.start.line) as u64;
        let col_diff = if self.start.line == self.end.line {
            self.end.column.saturating_sub(self.start.column) as u64
        } else {
            self.end.column as u64
        };
        line_diff * 10000 + col_diff
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn test_position_ordering() {
        let p1 = SourcePosition::new(1, 5);
        let p2 = SourcePosition::new(1, 10);
        let p3 = SourcePosition::new(2, 0);

        assert!(p1 < p2);
        assert!(p2 < p3);
        assert!(p1 < p3);
    }

    #[test]
    fn test_range_contains() {
        let range = SourceRange::new(
            SourcePosition::new(1, 0),
            SourcePosition::new(5, 10),
        );

        assert!(range.contains(SourcePosition::new(1, 0)));
        assert!(range.contains(SourcePosition::new(3, 5)));
        assert!(range.contains(SourcePosition::new(5, 10)));
        assert!(!range.contains(SourcePosition::new(0, 0)));
        assert!(!range.contains(SourcePosition::new(5, 11)));
        assert!(!range.contains(SourcePosition::new(6, 0)));
    }

    #[test]
    fn test_range_contains_single_line() {
        let range = SourceRange::new(
            SourcePosition::new(5, 10),
            SourcePosition::new(5, 20),
        );

        assert!(range.contains(SourcePosition::new(5, 10)));
        assert!(range.contains(SourcePosition::new(5, 15)));
        assert!(range.contains(SourcePosition::new(5, 20)));
        assert!(!range.contains(SourcePosition::new(5, 9)));
        assert!(!range.contains(SourcePosition::new(5, 21)));
    }
}
