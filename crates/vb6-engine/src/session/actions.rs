//! Code actions (editor quick-fixes and a small refactor).
//!
//! These are editor conveniences; the code they generate is valid VB6 that the
//! engine itself accepts and respects VB6 semantics:
//!   * **Declare undefined variable** — fixes "Variable not defined" (only
//!     raised under `Option Explicit`); an undeclared variable defaults to
//!     `Variant`, and a type-suffixed use (`n%`) declares that type.
//!   * **Create missing Sub/Function** — fixes "Sub or Function not defined";
//!     statement-position calls become a `Sub`, value-position calls a
//!     `Function`, with one `Variant` parameter per argument.
//!   * **Toggle single-line / block If** — a purely lexical rewrite between
//!     `If c Then s` and the `If c Then` … `End If` block form.

use super::format::{classify, tokenize, Eff};
use super::{CodeAction, CodeActionKind, Position, Session, TextEdit};
use crate::frontend::ast::{ExprNode, NodeId};
use crate::frontend::token::{Kw, Span, TokenKind};
use crate::sema::binder::{ERR_SUB_OR_FUNCTION_NOT_DEFINED, ERR_VARIABLE_NOT_DEFINED};

impl Session {
    /// Code actions available for the byte range `[start, end]` in `module`.
    pub fn code_actions(&self, module: usize, start: u32, end: u32) -> Vec<CodeAction> {
        let mut out = Vec::new();
        if self.modules.get(module).is_none() {
            return out;
        }

        for d in self.diagnostics(module) {
            if !span_overlaps(d.span, start, end) {
                continue;
            }
            if d.code == ERR_VARIABLE_NOT_DEFINED as u32 {
                out.extend(self.declare_var_action(module, d.span));
            } else if d.code == ERR_SUB_OR_FUNCTION_NOT_DEFINED as u32 {
                out.extend(self.create_proc_action(module, d.span));
            }
        }

        out.extend(self.toggle_if_action(module, start));
        out
    }

    /// Quick-fix: insert a `Dim` for an undeclared variable, just above the
    /// statement that uses it (always valid — VB6 hoists declarations).
    fn declare_var_action(&self, module: usize, span: Span) -> Option<CodeAction> {
        let m = self.modules.get(module)?;
        let src = m.source.as_slice();
        let name = read_name_with_suffix(src, span.start as usize);
        if name.is_empty() {
            return None;
        }
        let line = m.line_index.position(span.start).line;
        let line_start = m.line_index.offset(Position { line, character: 0 });
        let indent = leading_ws(src, line_start as usize);
        let nl = newline_style(src);
        let has_suffix = name
            .bytes()
            .next_back()
            .map(is_suffix)
            .unwrap_or(false);
        let decl = if has_suffix {
            format!("{indent}Dim {name}{nl}")
        } else {
            format!("{indent}Dim {name} As Variant{nl}")
        };
        Some(CodeAction {
            title: format!("Declare '{name}' with Dim"),
            kind: CodeActionKind::QuickFix,
            edits: vec![TextEdit { module, span: Span { start: line_start, len: 0 }, new_text: decl }],
        })
    }

    /// Quick-fix: append an empty `Sub`/`Function` stub for an undefined call.
    fn create_proc_action(&self, module: usize, span: Span) -> Option<CodeAction> {
        let m = self.modules.get(module)?;
        let src = m.source.as_slice();
        let name = read_ident(src, span.start as usize);
        if name.is_empty() {
            return None;
        }
        let node = find_nameref_node(m, span)?;
        let (is_func, arity) = call_shape(m, node);
        let nl = newline_style(src);
        let params = (0..arity)
            .map(|i| format!("Param{} As Variant", i + 1))
            .collect::<Vec<_>>()
            .join(", ");
        let stub = if is_func {
            format!("{nl}Private Function {name}({params}) As Variant{nl}End Function{nl}")
        } else {
            format!("{nl}Private Sub {name}({params}){nl}End Sub{nl}")
        };
        let at = src.len() as u32;
        let kind_word = if is_func { "Function" } else { "Sub" };
        Some(CodeAction {
            title: format!("Create {kind_word} '{name}'"),
            kind: CodeActionKind::QuickFix,
            edits: vec![TextEdit { module, span: Span { start: at, len: 0 }, new_text: stub }],
        })
    }

    /// Refactor: convert a single-line `If` to a block, or collapse a trivial
    /// `If … End If` block back to a single line, when the cursor sits on it.
    fn toggle_if_action(&self, module: usize, off: u32) -> Option<CodeAction> {
        let m = self.modules.get(module)?;
        let src = m.source.as_slice();
        let li = &m.line_index;
        let toks = tokenize(src);
        let lines = logical_lines(li, &toks);

        let cur_line = li.position(off).line as usize;
        let idx = lines.iter().position(|l| l.first <= cur_line && cur_line <= l.last)?;
        let head = &lines[idx];

        // Must be led by `If` (block headers also start with `If`).
        if first_kw(&head.sig) != Some(Kw::If) {
            return None;
        }
        let header_is_block = matches!(head.sig.last(), Some((TokenKind::Kw(Kw::Then), _)));
        if header_is_block {
            self.collapse_if(module, &lines, idx)
        } else {
            self.expand_if(module, head)
        }
    }

    /// `If c Then s [Else t]` (one physical line) → block form.
    fn expand_if(&self, module: usize, head: &LogLine) -> Option<CodeAction> {
        let m = self.modules.get(module)?;
        let src = m.source.as_slice();
        // Keep this simple and safe: only single-physical-line single-line Ifs.
        if head.first != head.last {
            return None;
        }
        let (then_i, else_i) = then_else_indices(&head.sig)?;
        let if_tok = head.sig.first()?.1;
        let then_tok = head.sig[then_i].1;

        let (line_start, line_end) = line_bounds(src, &m.line_index, head.first);
        let cond = slice_cp1252(src, end_of(if_tok) as usize, then_tok.start as usize);
        let cond = cond.trim();
        if cond.is_empty() {
            return None;
        }

        let (then_end, else_part) = match else_i {
            Some(ei) => {
                let else_tok = head.sig[ei].1;
                let else_text =
                    slice_cp1252(src, end_of(else_tok) as usize, line_end).trim().to_string();
                (head.sig[ei].1.start as usize, Some(else_text))
            }
            None => (line_end, None),
        };
        let then_text = slice_cp1252(src, end_of(then_tok) as usize, then_end);
        let then_text = then_text.trim();
        if then_text.is_empty() {
            return None;
        }

        let indent = leading_ws(src, line_start);
        let nl = newline_style(src);
        let mut new_text = format!(
            "{indent}If {cond} Then{nl}{indent}{INDENT}{then_text}",
            INDENT = "    "
        );
        if let Some(else_text) = else_part {
            if !else_text.is_empty() {
                new_text.push_str(&format!("{nl}{indent}Else{nl}{indent}    {else_text}"));
            }
        }
        new_text.push_str(&format!("{nl}{indent}End If"));

        Some(CodeAction {
            title: "Convert to multi-line If".to_string(),
            kind: CodeActionKind::RefactorRewrite,
            edits: vec![TextEdit {
                module,
                span: Span { start: line_start as u32, len: (line_end - line_start) as u32 },
                new_text,
            }],
        })
    }

    /// `If c Then` … `End If` block → `If c Then s [Else t]`, only when the body
    /// is a single statement (and optionally a single `Else` statement), with no
    /// `ElseIf`, nested blocks, or comments.
    fn collapse_if(&self, module: usize, lines: &[LogLine], head_idx: usize) -> Option<CodeAction> {
        let m = self.modules.get(module)?;
        let src = m.source.as_slice();
        let head = &lines[head_idx];

        // cond from the header line.
        let then_tok = head.sig.iter().find(|(k, _)| matches!(k, TokenKind::Kw(Kw::Then)))?.1;
        let if_tok = head.sig.first()?.1;
        let cond = slice_cp1252(src, end_of(if_tok) as usize, then_tok.start as usize)
            .trim()
            .to_string();
        if cond.is_empty() {
            return None;
        }

        // Walk forward to the matching `End If`. Any nested block makes the
        // body non-trivial, so we bail rather than collapse.
        let body = collect_if_body(src, &m.line_index, lines, head_idx)?;
        let CollapsibleIf { then_stmts, else_stmts, end_idx } = body;
        if then_stmts.len() != 1 || else_stmts.len() > 1 {
            return None;
        }

        let (line_start, _) = line_bounds(src, &m.line_index, head.first);
        let (_, end_line_end) = line_bounds(src, &m.line_index, lines[end_idx].last);
        let indent = leading_ws(src, line_start);

        let mut new_text = format!("{indent}If {cond} Then {}", then_stmts[0]);
        if let Some(e) = else_stmts.first() {
            new_text.push_str(&format!(" Else {e}"));
        }

        Some(CodeAction {
            title: "Collapse to single-line If".to_string(),
            kind: CodeActionKind::RefactorRewrite,
            edits: vec![TextEdit {
                module,
                span: Span { start: line_start as u32, len: (end_line_end - line_start) as u32 },
                new_text,
            }],
        })
    }
}

/// The collected body of a collapsible `If … Then` … `End If` block.
struct CollapsibleIf {
    then_stmts: Vec<String>,
    else_stmts: Vec<String>,
    end_idx: usize,
}

/// Walk the lines after an `If` header to its matching `End If`, gathering the
/// `Then`/`Else` statement texts. Returns `None` if the block is not a
/// collapsible shape (comment, nested block, `ElseIf`, missing `End If`).
fn collect_if_body(
    src: &[u8],
    li: &super::LineIndex,
    lines: &[LogLine],
    head_idx: usize,
) -> Option<CollapsibleIf> {
    let mut then_stmts: Vec<String> = Vec::new();
    let mut else_stmts: Vec<String> = Vec::new();
    let mut in_else = false;
    for (i, l) in lines.iter().enumerate().skip(head_idx + 1) {
        if has_comment(l) {
            return None;
        }
        match classify(&l.sig) {
            Eff::Close => {
                return Some(CollapsibleIf { then_stmts, else_stmts, end_idx: i });
            }
            // A nested block opener/closer, `ElseIf`, or `Case` is not a
            // collapsible shape.
            Eff::OpenIf | Eff::OpenGeneric | Eff::OpenSelect | Eff::CloseSelect | Eff::Case => {
                return None
            }
            // Only a bare `Else` is collapsible; `ElseIf` is not.
            Eff::ElseLike if first_kw(&l.sig) == Some(Kw::Else) && !in_else => {
                in_else = true;
            }
            Eff::ElseLike => return None,
            Eff::Neutral => push_if_stmt(src, li, l, in_else, &mut then_stmts, &mut else_stmts),
        }
    }
    None
}

/// Append a neutral body line's trimmed text to the appropriate branch (skipping
/// blank lines).
fn push_if_stmt(
    src: &[u8],
    li: &super::LineIndex,
    l: &LogLine,
    in_else: bool,
    then_stmts: &mut Vec<String>,
    else_stmts: &mut Vec<String>,
) {
    let text = logical_text(src, li, l).trim().to_string();
    if text.is_empty() {
        return;
    }
    if in_else {
        else_stmts.push(text);
    } else {
        then_stmts.push(text);
    }
}

// ── Logical-line model ──────────────────────────────────────────────────────────

/// A logical line: a run of physical lines joined by `_` continuations, with the
/// tokens that fall on them and a comment/continuation-free `sig` view.
struct LogLine {
    first: usize,
    last: usize,
    toks: Vec<(TokenKind, Span)>,
    sig: Vec<(TokenKind, Span)>,
}

/// Split the token stream into logical lines, preserving order.
fn logical_lines(li: &super::LineIndex, toks: &[(TokenKind, Span)]) -> Vec<LogLine> {
    let line_count = li.line_count();
    let mut by_line: Vec<Vec<usize>> = vec![Vec::new(); line_count];
    for (i, (_, sp)) in toks.iter().enumerate() {
        let ln = li.position(sp.start).line as usize;
        if ln < line_count {
            by_line[ln].push(i);
        }
    }

    let mut out = Vec::new();
    let mut line = 0usize;
    while line < line_count {
        let mut last = line;
        while line_ends_with_cont(&by_line[last], toks) && last + 1 < line_count {
            last += 1;
        }
        let group_toks: Vec<(TokenKind, Span)> = (line..=last)
            .flat_map(|l| by_line[l].iter())
            .map(|&i| toks[i].clone())
            .collect();
        let sig: Vec<(TokenKind, Span)> = group_toks
            .iter()
            .filter(|(k, _)| !matches!(k, TokenKind::Kw(Kw::Apos) | TokenKind::Kw(Kw::LineCont)))
            .cloned()
            .collect();
        out.push(LogLine { first: line, last, toks: group_toks, sig });
        line = last + 1;
    }
    out
}

fn line_ends_with_cont(line_toks: &[usize], toks: &[(TokenKind, Span)]) -> bool {
    line_toks
        .last()
        .map(|&i| matches!(toks[i].0, TokenKind::Kw(Kw::LineCont)))
        .unwrap_or(false)
}

fn has_comment(l: &LogLine) -> bool {
    l.toks.iter().any(|(k, _)| matches!(k, TokenKind::Kw(Kw::Apos)))
}

fn first_kw(sig: &[(TokenKind, Span)]) -> Option<Kw> {
    sig.iter().find_map(|(k, _)| if let TokenKind::Kw(w) = k { Some(*w) } else { None })
}

/// Index of the `Then` and (optional) `Else` tokens at paren-depth 0, for a
/// single-line `If`. Returns `None` if there is no `Then`.
fn then_else_indices(sig: &[(TokenKind, Span)]) -> Option<(usize, Option<usize>)> {
    let mut depth = 0i32;
    let mut then_i = None;
    let mut else_i = None;
    for (i, (k, _)) in sig.iter().enumerate() {
        match k {
            TokenKind::Kw(Kw::LParen) => depth += 1,
            TokenKind::Kw(Kw::RParen) => depth -= 1,
            TokenKind::Kw(Kw::Then) if depth == 0 && then_i.is_none() => then_i = Some(i),
            TokenKind::Kw(Kw::Else) if depth == 0 && then_i.is_some() && else_i.is_none() => {
                else_i = Some(i)
            }
            _ => {}
        }
    }
    Some((then_i?, else_i))
}

// ── Source helpers ──────────────────────────────────────────────────────────────

fn end_of(span: Span) -> u32 {
    span.start + span.len
}

fn span_overlaps(span: Span, start: u32, end: u32) -> bool {
    span.start <= end && start <= span.start + span.len
}

fn is_ident_continue(b: u8) -> bool {
    matches!(b, b'A'..=b'Z' | b'a'..=b'z' | b'0'..=b'9' | b'_' | 0x80..=0xFF)
}

fn is_suffix(b: u8) -> bool {
    matches!(b, b'%' | b'&' | b'!' | b'#' | b'@' | b'$')
}

fn read_ident(src: &[u8], start: usize) -> String {
    let mut e = start;
    while e < src.len() && is_ident_continue(src[e]) {
        e += 1;
    }
    slice_cp1252(src, start, e)
}

fn read_name_with_suffix(src: &[u8], start: usize) -> String {
    let mut e = start;
    while e < src.len() && is_ident_continue(src[e]) {
        e += 1;
    }
    if e < src.len() && is_suffix(src[e]) {
        e += 1;
    }
    slice_cp1252(src, start, e)
}

fn leading_ws(src: &[u8], line_start: usize) -> String {
    let mut e = line_start;
    while e < src.len() && matches!(src[e], b' ' | b'\t') {
        e += 1;
    }
    slice_cp1252(src, line_start, e)
}

fn newline_style(src: &[u8]) -> &'static str {
    if src.windows(2).any(|w| w == b"\r\n") {
        "\r\n"
    } else {
        "\n"
    }
}

/// Decode a byte range as Windows-1252 — the encoding the engine's buffer is in
/// (the host encodes the editor's UTF-8 to CP-1252 before handing it over). Only
/// 0x80–0x9F differs from Latin-1; everything else is `b as char`. Undefined
/// CP-1252 slots (0x81/0x8D/0x8F/0x90/0x9D) pass through unchanged.
fn slice_cp1252(src: &[u8], a: usize, b: usize) -> String {
    src[a.min(src.len())..b.min(src.len())].iter().map(|&c| cp1252_char(c)).collect()
}

/// Map one Windows-1252 byte to its Unicode scalar.
fn cp1252_char(b: u8) -> char {
    match b {
        0x80 => '\u{20AC}', // €
        0x82 => '\u{201A}', // ‚
        0x83 => '\u{0192}', // ƒ
        0x84 => '\u{201E}', // „
        0x85 => '\u{2026}', // …
        0x86 => '\u{2020}', // †
        0x87 => '\u{2021}', // ‡
        0x88 => '\u{02C6}', // ˆ
        0x89 => '\u{2030}', // ‰
        0x8A => '\u{0160}', // Š
        0x8B => '\u{2039}', // ‹
        0x8C => '\u{0152}', // Œ
        0x8E => '\u{017D}', // Ž
        0x91 => '\u{2018}', // ‘
        0x92 => '\u{2019}', // ’
        0x93 => '\u{201C}', // “
        0x94 => '\u{201D}', // ”
        0x95 => '\u{2022}', // •
        0x96 => '\u{2013}', // –
        0x97 => '\u{2014}', // —
        0x98 => '\u{02DC}', // ˜
        0x99 => '\u{2122}', // ™
        0x9A => '\u{0161}', // š
        0x9B => '\u{203A}', // ›
        0x9C => '\u{0153}', // œ
        0x9E => '\u{017E}', // ž
        0x9F => '\u{0178}', // Ÿ
        _ => b as char,
    }
}

/// `(line_start, line_end)` byte offsets for a physical line, excluding the
/// line terminator.
fn line_bounds(src: &[u8], li: &super::LineIndex, line: usize) -> (usize, usize) {
    let start = li.offset(Position { line: line as u32, character: 0 }) as usize;
    let next = li.offset(Position { line: line as u32 + 1, character: 0 }) as usize;
    let mut end = next.max(start);
    while end > start && matches!(src[end - 1], b'\n' | b'\r') {
        end -= 1;
    }
    (start, end)
}

/// Trimmed text content of a logical line's first physical line (its statement
/// text), used when collapsing a block.
fn logical_text(src: &[u8], li: &super::LineIndex, l: &LogLine) -> String {
    let (s, e) = line_bounds(src, li, l.first);
    slice_cp1252(src, s, e)
}

// ── AST helpers (Sub vs Function, arity) ────────────────────────────────────────

fn find_nameref_node(m: &super::ModuleData, span: Span) -> Option<u32> {
    for i in 0..m.arena.len() as u32 {
        if let ExprNode::NameRef { .. } = m.arena.get(NodeId(i)) {
            let sp = m.spans.get(NodeId(i));
            if sp.start == span.start && sp.len == span.len {
                return Some(i);
            }
        }
    }
    None
}

fn arglist_len(arena: &crate::frontend::ast::ExprArena, args: NodeId) -> usize {
    match arena.get(args) {
        ExprNode::ArgList { args } => args.len(),
        _ => 0,
    }
}

fn block_contains(arena: &crate::frontend::ast::ExprArena, id: u32) -> bool {
    for i in 0..arena.len() as u32 {
        if let ExprNode::Block { stmts } = arena.get(NodeId(i)) {
            if stmts.iter().any(|n| n.0 == id) {
                return true;
            }
        }
    }
    false
}

#[cfg(test)]
mod tests {
    use crate::session::{CodeAction, Session, TextEdit};

    fn apply(src: &str, edits: &[TextEdit]) -> String {
        let mut bytes = src.as_bytes().to_vec();
        let mut es: Vec<&TextEdit> = edits.iter().collect();
        es.sort_by_key(|e| std::cmp::Reverse(e.span.start));
        for e in es {
            let s = e.span.start as usize;
            let end = s + e.span.len as usize;
            bytes.splice(s..end, e.new_text.bytes());
        }
        String::from_utf8(bytes).unwrap()
    }

    fn actions_at(src: &str, needle: &str) -> (Session, Vec<CodeAction>) {
        let s = Session::from_sources(vec![("m.bas".to_string(), src.as_bytes().to_vec())]);
        let off = src.find(needle).expect("needle in source") as u32;
        let acts = s.code_actions(0, off, off);
        (s, acts)
    }

    #[test]
    fn declare_undefined_variable() {
        let src = "Option Explicit\nSub Foo()\n    x = 1\nEnd Sub\n";
        let (_s, acts) = actions_at(src, "x =");
        let a = acts
            .iter()
            .find(|a| a.title.contains("Declare"))
            .expect("declare action");
        let out = apply(src, &a.edits);
        assert!(out.contains("    Dim x As Variant\n    x = 1"), "got:\n{out}");
    }

    #[test]
    fn create_missing_sub() {
        let src = "Sub Foo()\n    Bar\nEnd Sub\n";
        let (_s, acts) = actions_at(src, "Bar");
        let a = acts
            .iter()
            .find(|a| a.title.contains("Create Sub"))
            .expect("create sub action");
        let out = apply(src, &a.edits);
        assert!(out.contains("Private Sub Bar()\nEnd Sub"), "got:\n{out}");
    }

    #[test]
    fn expand_single_line_if() {
        let src = "If x > 1 Then y = 2\n";
        let (_s, acts) = actions_at(src, "If");
        let a = acts
            .iter()
            .find(|a| a.kind == crate::session::CodeActionKind::RefactorRewrite)
            .expect("toggle action");
        let out = apply(src, &a.edits);
        assert_eq!(out, "If x > 1 Then\n    y = 2\nEnd If\n");
    }

    #[test]
    fn expand_preserves_cp1252_string_literal() {
        // Source bytes are Windows-1252: 0x92 is a right single quote (’).
        let mut src = b"If x Then y = \"a".to_vec();
        src.push(0x92);
        src.extend_from_slice(b"b\"\n");
        let s = Session::from_sources(vec![("m.bas".to_string(), src.clone())]);
        let acts = s.code_actions(0, 0, 0);
        let a = acts
            .iter()
            .find(|a| a.kind == crate::session::CodeActionKind::RefactorRewrite)
            .expect("toggle action");
        // The relocated string text must carry the correct Unicode quote, not the
        // Latin-1 mis-decode (U+0092).
        let moved = &a.edits[0].new_text;
        assert!(moved.contains('\u{2019}'), "got: {moved:?}");
        assert!(!moved.contains('\u{0092}'), "got: {moved:?}");
    }

    #[test]
    fn collapse_block_if() {
        let src = "If x > 1 Then\n    y = 2\nEnd If\n";
        let (_s, acts) = actions_at(src, "If");
        let a = acts
            .iter()
            .find(|a| a.kind == crate::session::CodeActionKind::RefactorRewrite)
            .expect("toggle action");
        let out = apply(src, &a.edits);
        assert_eq!(out, "If x > 1 Then y = 2\n");
    }
}

/// Decide whether an undefined call should be created as a `Function` (used in
/// value position) or a `Sub` (called as a statement), and its argument count.
fn call_shape(m: &super::ModuleData, node: u32) -> (bool, usize) {
    let arena = &m.arena;
    let mut arity = 0;
    let mut call_node = None;
    for i in 0..arena.len() as u32 {
        match arena.get(NodeId(i)) {
            ExprNode::Call { func, args } if func.0 == node => {
                call_node = Some(i);
                arity = arglist_len(arena, *args);
            }
            ExprNode::CallStmt { callee, args } if callee.0 == node => {
                // Explicit `Call` keyword is always a statement → Sub.
                return (false, arglist_len(arena, *args));
            }
            _ => {}
        }
    }
    // Bare name as a statement, or a call expression used as a statement → Sub.
    let is_stmt = block_contains(arena, node)
        || call_node.map(|c| block_contains(arena, c)).unwrap_or(false);
    (!is_stmt, arity)
}
