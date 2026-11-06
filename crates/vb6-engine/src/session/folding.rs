//! Folding range generation from VB6 source text.
//!
//! Scans the source line-by-line, tracking block-open/close keywords to build
//! start/end line pairs. Covers procedures (Sub/Function/Property), conditionals
//! (If/ElseIf/Else…End If), loops (For/Next, Do/Loop, While/Wend), and
//! structural blocks (With/End With, Select Case/End Select, Type/End Type,
//! Enum/End Enum).

use super::Session;

/// A foldable source region, expressed as inclusive 0-based line numbers.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct FoldRange {
    pub start_line: u32,
    pub end_line: u32,
}

impl Session {
    /// All foldable regions in `module`, in source order.
    pub fn folding_ranges(&self, module: usize) -> Vec<FoldRange> {
        let Some(m) = self.modules.get(module) else { return Vec::new() };
        folding_ranges_for_source(&m.source)
    }
}

pub(super) fn folding_ranges_for_source(source: &[u8]) -> Vec<FoldRange> {
    // Interpret as Latin-1 (VB6 is Windows-1252; ASCII subset is identical)
    let lines: Vec<String> = crate::frontend::line_reader::physical_lines(source)
        .map(|line| source[line.start..line.end].iter().map(|&b| b as char).collect())
        .collect();
    let mut ranges = Vec::new();
    // Stack of (start_line, BlockKind)
    let mut stack: Vec<(u32, Block)> = Vec::new();

    for (line_num, raw_line) in lines.iter().enumerate() {
        let line_num = line_num as u32;
        let trimmed = raw_line.trim();
        if trimmed.is_empty() || trimmed.starts_with('\'') {
            continue;
        }
        // Strip inline comment before keyword matching
        let no_comment = strip_comment(trimmed);
        let up = no_comment.to_ascii_uppercase();
        let up = up.trim();

        if let Some(blk) = block_open(up) {
            stack.push((line_num, blk));
        } else if let Some(close) = block_close(up) {
            // Pop the most-recent matching open block
            if let Some(pos) = stack.iter().rposition(|(_, b)| b.matches(close)) {
                let (start_line, _) = stack.remove(pos);
                if line_num > start_line {
                    ranges.push(FoldRange { start_line, end_line: line_num });
                }
            }
        }
    }
    ranges
}

/// Strip a trailing `'…` comment from a line (outside string literals).
fn strip_comment(line: &str) -> &str {
    let mut in_str = false;
    for (i, c) in line.char_indices() {
        match c {
            '"' => in_str = !in_str,
            '\'' if !in_str => return &line[..i],
            _ => {}
        }
    }
    line
}

#[derive(Clone, Copy, PartialEq, Eq, Debug)]
enum Block {
    Sub, Function, Property,
    If, For, Do, While, With, Select,
    Type, Enum,
}

#[derive(Clone, Copy, PartialEq, Eq)]
enum Close {
    Sub, Function, Property,
    If, For, Do, While, With, Select,
    Type, Enum,
}

impl Block {
    fn matches(self, c: Close) -> bool {
        matches!(
            (self, c),
            (Block::Sub,      Close::Sub)
            | (Block::Function, Close::Function)
            | (Block::Property, Close::Property)
            | (Block::If,       Close::If)
            | (Block::For,      Close::For)
            | (Block::Do,       Close::Do)
            | (Block::While,    Close::While)
            | (Block::With,     Close::With)
            | (Block::Select,   Close::Select)
            | (Block::Type,     Close::Type)
            | (Block::Enum,     Close::Enum)
        )
    }
}

/// Strip leading visibility/modifier keywords and return the rest.
fn strip_modifiers(up: &str) -> &str {
    let mut s = up;
    loop {
        let prev = s;
        for pfx in &["PRIVATE ", "PUBLIC ", "FRIEND ", "STATIC ", "DEFAULT "] {
            if let Some(rest) = s.strip_prefix(pfx) {
                s = rest.trim_start();
            }
        }
        if s == prev { break; }
    }
    s
}

fn block_open(up: &str) -> Option<Block> {
    let s = strip_modifiers(up);
    if s.starts_with("SUB ") || s == "SUB" { return Some(Block::Sub); }
    if s.starts_with("FUNCTION ") || s == "FUNCTION" { return Some(Block::Function); }
    if s.starts_with("PROPERTY ") { return Some(Block::Property); }
    if s.starts_with("TYPE ") { return Some(Block::Type); }
    if s.starts_with("ENUM ") { return Some(Block::Enum); }
    // Block If: line ends with THEN (nothing executable follows)
    if (s.starts_with("IF ") || s.starts_with("IF(")) && is_block_if(s) {
        return Some(Block::If);
    }
    if s.starts_with("FOR ") { return Some(Block::For); }
    if s == "DO" || s.starts_with("DO ") { return Some(Block::Do); }
    if s.starts_with("WHILE ") { return Some(Block::While); }
    if s.starts_with("WITH ") { return Some(Block::With); }
    if s.starts_with("SELECT ") { return Some(Block::Select); }
    None
}

fn block_close(up: &str) -> Option<Close> {
    if up == "END SUB" || up.starts_with("END SUB ") { return Some(Close::Sub); }
    if up == "END FUNCTION" || up.starts_with("END FUNCTION ") { return Some(Close::Function); }
    if up == "END PROPERTY" || up.starts_with("END PROPERTY ") { return Some(Close::Property); }
    if up == "END IF" || up.starts_with("END IF ") { return Some(Close::If); }
    if up == "END TYPE" || up.starts_with("END TYPE ") { return Some(Close::Type); }
    if up == "END ENUM" || up.starts_with("END ENUM ") { return Some(Close::Enum); }
    if up == "END WITH" || up.starts_with("END WITH ") { return Some(Close::With); }
    if up == "END SELECT" || up.starts_with("END SELECT ") { return Some(Close::Select); }
    // Loop closers (no "END" prefix)
    if up.starts_with("NEXT") && (up.len() == 4 || up.as_bytes()[4] == b' ') {
        return Some(Close::For);
    }
    if up.starts_with("LOOP") && (up.len() == 4 || up.as_bytes()[4] == b' ') {
        return Some(Close::Do);
    }
    if up == "WEND" || up.starts_with("WEND ") { return Some(Close::While); }
    None
}

/// A block `If` has nothing executable after `THEN` (comments are fine).
fn is_block_if(up: &str) -> bool {
    // Find the last THEN token
    let upper = up.to_ascii_uppercase();
    if let Some(pos) = upper.rfind("THEN") {
        let after = upper[pos + 4..].trim_start();
        // Nothing after, or only a comment marker
        return after.is_empty() || after.starts_with('\'');
    }
    false
}

#[cfg(test)]
#[path = "tests/folding.rs"]
mod tests;
