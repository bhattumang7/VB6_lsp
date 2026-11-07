//! Designer-file (`.frm`/`.ctl`/…) support for the session: form-control symbols.
//!
//! VB6 designer files are two documents in one: a `Begin…End` designer section
//! (parsed by [`crate::frm`]) followed by VB6 code. [`form_controls`] flattens the
//! control tree of the designer side into a symbol list for the LSP layer.

use crate::frm::parse_frm;
use crate::frm::parser::{BeginBlock, PropValue};
use crate::frontend::ast::Span;

/// A control declared in a designer file.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct FormControl {
    /// Instance name, e.g. `cmdOK`.
    pub name: String,
    /// Control type, e.g. `VB.CommandButton`.
    pub control_type: String,
    /// Nesting depth (0 = the form/usercontrol itself).
    pub depth: u32,
    /// Byte span of the control name on its `Begin` line, or [`Span::DUMMY`]
    /// when the line does not spell the name.
    pub span: Span,
    /// The control's `Index` property: `Some` for an element of a control array.
    pub index: Option<i32>,
    /// Byte span of the control type on the `Begin` line, or [`Span::DUMMY`].
    pub type_span: Span,
}

/// Parse a designer file's source and return its controls (the root form and
/// every nested control), depth-first.
///
/// Returns an empty list for `.cls`-style files with no designer block or on a
/// designer parse error (the code side is still handled by the normal binder).
pub fn form_controls(src: &str) -> Vec<FormControl> {
    controls_of(src, src.as_bytes())
}

/// Like [`form_controls`], from a file's raw (Windows-1252) bytes; spans are
/// byte offsets into `raw`.
pub fn form_controls_from_bytes(raw: &[u8]) -> Vec<FormControl> {
    let text: String = raw.iter().map(|&b| b as char).collect();
    controls_of(&text, raw)
}

/// `text` is what the designer parser reads; `raw` is the same file as the
/// bytes the returned spans index into (the two have the same lines).
fn controls_of(text: &str, raw: &[u8]) -> Vec<FormControl> {
    let Ok(frm) = parse_frm(text) else {
        return Vec::new();
    };
    let mut out = Vec::new();
    if let Some(root) = frm.root.as_ref() {
        let lines = line_starts(raw);
        walk(root, 0, raw, &lines, &mut out);
    }
    out
}

/// Byte offset of the start of each line (index 0 = line 1).
fn line_starts(raw: &[u8]) -> Vec<usize> {
    crate::frontend::line_reader::physical_line_starts(raw).into_iter().map(|s| s as usize).collect()
}

fn walk(block: &BeginBlock, depth: u32, raw: &[u8], lines: &[usize], out: &mut Vec<FormControl>) {
    out.push(FormControl {
        name: block.name.clone(),
        control_type: block.control_type.clone(),
        depth,
        span: name_span(raw, lines, block),
        index: index_of(block),
        type_span: type_span(raw, lines, block),
    });
    for child in &block.children {
        walk(child, depth + 1, raw, lines, out);
    }
}

/// The value of a control's `Index` property, when it has one.
fn index_of(block: &BeginBlock) -> Option<i32> {
    block.properties.iter().find_map(|p| match (&p.value, p.name.eq_ignore_ascii_case("Index")) {
        (PropValue::Scalar(v), true) => v.split_whitespace().next()?.parse().ok(),
        _ => None,
    })
}

/// Byte span of a control's type on its own `Begin <type> <name>` line.
fn type_span(raw: &[u8], lines: &[usize], block: &BeginBlock) -> Span {
    let Some(&start) = block.line.checked_sub(1).and_then(|i| lines.get(i)) else {
        return Span::DUMMY;
    };
    let end = lines.get(block.line).copied().unwrap_or(raw.len());
    let needle = block.control_type.as_bytes();
    match raw[start..end].windows(needle.len().max(1)).position(|w| w == needle) {
        Some(at) if !needle.is_empty() => {
            Span { start: (start + at) as u32, len: needle.len() as u32 }
        }
        _ => Span::DUMMY,
    }
}

/// Byte span of a control's name on its own `Begin <type> <name>` line.
fn name_span(raw: &[u8], lines: &[usize], block: &BeginBlock) -> Span {
    let Some(&start) = block.line.checked_sub(1).and_then(|i| lines.get(i)) else {
        return Span::DUMMY;
    };
    let end = lines.get(block.line).copied().unwrap_or(raw.len());
    let line = &raw[start..end];
    let find = |hay: &[u8], needle: &[u8]| hay.windows(needle.len()).position(|w| w == needle);
    // The name is the word after the type on the `Begin` line.
    let Some(type_at) = find(line, block.control_type.as_bytes()) else {
        return Span::DUMMY;
    };
    let after = type_at + block.control_type.len();
    let skipped = line[after..].iter().take_while(|b| b.is_ascii_whitespace()).count();
    let name_at = after + skipped;
    if line[name_at..].starts_with(block.name.as_bytes()) {
        Span { start: (start + name_at) as u32, len: block.name.len() as u32 }
    } else {
        Span::DUMMY
    }
}

#[cfg(test)]
#[path = "tests/forms.rs"]
mod tests;
