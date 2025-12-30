/**
 * Position and Range Types
 *
 * Precise position tracking for symbol table operations.
 * Converts between tree-sitter and MCP position formats.
 */

import type { Point, SyntaxNode } from "tree-sitter";

/**
 * A precise position in source code (0-indexed line and column)
 */
export interface SourcePosition {
  line: number;
  column: number;
}

export function createPosition(line: number, column: number): SourcePosition {
  return { line, column };
}

/**
 * Create from tree-sitter Point
 */
export function fromTsPoint(point: Point): SourcePosition {
  return {
    line: point.row,
    column: point.column,
  };
}

/**
 * Compare two positions
 * Returns: -1 if a < b, 0 if equal, 1 if a > b
 */
export function comparePositions(
  a: SourcePosition,
  b: SourcePosition
): number {
  if (a.line !== b.line) {
    return a.line < b.line ? -1 : 1;
  }
  if (a.column !== b.column) {
    return a.column < b.column ? -1 : 1;
  }
  return 0;
}

/**
 * A range in source code with start and end positions
 */
export interface SourceRange {
  start: SourcePosition;
  end: SourcePosition;
}

export function createRange(
  start: SourcePosition,
  end: SourcePosition
): SourceRange {
  return { start, end };
}

/**
 * Create from tree-sitter Node
 */
export function fromTsNode(node: SyntaxNode): SourceRange {
  return {
    start: fromTsPoint(node.startPosition),
    end: fromTsPoint(node.endPosition),
  };
}

/**
 * Check if this range contains a position
 */
export function rangeContains(
  range: SourceRange,
  pos: SourcePosition
): boolean {
  if (pos.line < range.start.line || pos.line > range.end.line) {
    return false;
  }
  if (pos.line === range.start.line && pos.column < range.start.column) {
    return false;
  }
  if (pos.line === range.end.line && pos.column > range.end.column) {
    return false;
  }
  return true;
}

/**
 * Check if this range contains another range
 */
export function rangeContainsRange(
  outer: SourceRange,
  inner: SourceRange
): boolean {
  return rangeContains(outer, inner.start) && rangeContains(outer, inner.end);
}

/**
 * Check if two ranges overlap
 */
export function rangesOverlap(a: SourceRange, b: SourceRange): boolean {
  return (
    rangeContains(a, b.start) ||
    rangeContains(a, b.end) ||
    rangeContains(b, a.start) ||
    rangeContains(b, a.end)
  );
}

/**
 * Calculate the "size" of a range (for finding innermost scope)
 */
export function rangeSize(range: SourceRange): number {
  const lineDiff = range.end.line - range.start.line;
  const colDiff =
    range.start.line === range.end.line
      ? range.end.column - range.start.column
      : range.end.column;
  return lineDiff * 10000 + colDiff;
}

/**
 * Default/empty range
 */
export function emptyRange(): SourceRange {
  return {
    start: { line: 0, column: 0 },
    end: { line: 0, column: 0 },
  };
}

/**
 * Module-level range (covers entire file)
 */
export function moduleRange(): SourceRange {
  return {
    start: { line: 0, column: 0 },
    end: { line: Number.MAX_SAFE_INTEGER - 1, column: 0 },
  };
}
