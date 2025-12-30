//! Document Management Utilities
//!
//! Helper functions for working with LSP documents.

use ropey::Rope;
use tower_lsp::lsp_types::Position;

/// Convert LSP Position to byte offset in a Rope
pub fn position_to_offset(rope: &Rope, position: Position) -> Option<usize> {
    let line = position.line as usize;
    if line >= rope.len_lines() {
        return None;
    }

    let line_start = rope.line_to_char(line);
    let offset = line_start + position.character as usize;

    if offset > rope.len_chars() {
        None
    } else {
        Some(offset)
    }
}

/// Convert byte offset to LSP Position in a Rope
pub fn offset_to_position(rope: &Rope, offset: usize) -> Position {
    let line = rope.char_to_line(offset);
    let line_start = rope.line_to_char(line);
    let character = offset - line_start;

    Position {
        line: line as u32,
        character: character as u32,
    }
}

/// Get the word at a given position
pub fn word_at_position(rope: &Rope, position: Position) -> Option<String> {
    let offset = position_to_offset(rope, position)?;
    let line = position.line as usize;

    if line >= rope.len_lines() {
        return None;
    }

    let line_text = rope.line(line).to_string();
    let char_offset = position.character as usize;

    // Find word boundaries
    let start = line_text[..char_offset]
        .rfind(|c: char| !c.is_alphanumeric() && c != '_')
        .map(|i| i + 1)
        .unwrap_or(0);

    let end = line_text[char_offset..]
        .find(|c: char| !c.is_alphanumeric() && c != '_')
        .map(|i| char_offset + i)
        .unwrap_or(line_text.len());

    if start < end {
        Some(line_text[start..end].to_string())
    } else {
        None
    }
}
