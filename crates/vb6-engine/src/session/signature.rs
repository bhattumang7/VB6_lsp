//! Signature help: find the proc call the cursor is inside and return parameter info.

use crate::frontend::ast::ProcKind;
use super::hover;
use super::Session;

/// Signature help result for one call site.
#[derive(Debug, Clone)]
pub struct SigHelp {
    /// Full rendered signature label, e.g. `Add(x As Long, y As Long) As Long`.
    pub label: String,
    /// For each parameter: `(start, end)` byte offsets within `label` that
    /// delimit that parameter's text. Used by editors for inline highlighting.
    pub params: Vec<(u32, u32)>,
    /// 0-based index of the parameter the cursor is currently inside.
    pub active_param: usize,
}

impl Session {
    /// Return signature help when the cursor at `offset` is inside a call argument list.
    ///
    /// Scans backwards through the source text to find the innermost unclosed `(`,
    /// reads the identifier before it, resolves it to a proc declaration, and
    /// returns the full signature with the active parameter index.
    pub fn signature_help(&self, module: usize, offset: u32) -> Option<SigHelp> {
        let m = self.modules.get(module)?;
        let (call_name, active_param) = find_call_context(&m.source, offset)?;

        // Search this module first, then public procs in other modules
        for p in &m.bound.procs {
            let pname = hover::name_at(&m.source, p.name_span);
            if pname.eq_ignore_ascii_case(&call_name) {
                return Some(build(&m.ctx, &m.source, p, active_param));
            }
        }
        for other in &self.modules {
            for p in &other.bound.procs {
                if !p.is_public { continue; }
                let pname = hover::name_at(&other.source, p.name_span);
                if pname.eq_ignore_ascii_case(&call_name) {
                    return Some(build(&other.ctx, &other.source, p, active_param));
                }
            }
        }
        None
    }
}

/// Scan backwards from `offset` through `source` to find the innermost
/// unclosed `(`. Returns the name of the identifier before `(` and the
/// 0-based comma count before the cursor (= active parameter index).
/// Returns `None` if the cursor is not inside a call argument list.
fn find_call_context(source: &[u8], offset: u32) -> Option<(String, usize)> {
    let end = (offset as usize).min(source.len());
    let text = &source[..end];
    let mut depth: i32 = 0;
    let mut commas: usize = 0;
    let mut i = end;

    while i > 0 {
        i -= 1;
        match text[i] {
            b')' => depth += 1,
            b'(' => {
                if depth == 0 {
                    // Found the matching open paren — read the ident before it
                    let name = ident_before(text, i)?;
                    return Some((name, commas));
                }
                depth -= 1;
            }
            b',' if depth == 0 => commas += 1,
            // VB6 is line-based; don't cross a line boundary
            b'\n' => return None,
            _ => {}
        }
    }
    None
}

/// Read the identifier that ends at byte position `paren_pos` in `text`
/// (scanning right-to-left, skipping leading whitespace).
fn ident_before(text: &[u8], paren_pos: usize) -> Option<String> {
    let mut i = paren_pos;
    // skip whitespace
    while i > 0 && matches!(text[i - 1], b' ' | b'\t') {
        i -= 1;
    }
    let end = i;
    // read identifier chars backwards
    while i > 0 && is_ident(text[i - 1]) {
        i -= 1;
    }
    if i == end { return None; }
    Some(text[i..end].iter().map(|&b| b as char).collect())
}

fn is_ident(b: u8) -> bool {
    b.is_ascii_alphanumeric() || b == b'_'
}

/// Build a `SigHelp` from a resolved proc declaration.
fn build(
    ctx: &crate::frontend::scanner::ScannerContext,
    src: &[u8],
    p: &crate::sema::symbol::BoundProc,
    active_param: usize,
) -> SigHelp {
    let name = hover::name_at(src, p.name_span);
    let mut label = format!("{}(", name);
    let mut params = Vec::new();

    for (i, param) in p.params.iter().enumerate() {
        if i > 0 { label.push_str(", "); }
        let start = label.len() as u32;
        let pname = hover::name_at(src, param.name_span);
        if param.flags.optional   { label.push_str("Optional "); }
        if param.flags.param_array { label.push_str("ParamArray "); }
        else if param.flags.by_val { label.push_str("ByVal "); }
        else if param.flags.by_ref { label.push_str("ByRef "); }
        label.push_str(&pname);
        if param.flags.is_array  { label.push_str("()"); }
        label.push_str(" As ");
        label.push_str(&hover::type_str(ctx, &param.vba_type));
        params.push((start, label.len() as u32));
    }

    label.push(')');
    if matches!(p.kind, ProcKind::Function | ProcKind::PropGet) {
        label.push_str(" As ");
        label.push_str(&hover::type_str(ctx, &p.ret_type));
    }

    let clamped = if p.params.is_empty() { 0 } else { active_param.min(p.params.len() - 1) };
    SigHelp { label, params, active_param: clamped }
}
