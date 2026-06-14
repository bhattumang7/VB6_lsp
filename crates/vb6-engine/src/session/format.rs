//! Document formatting for a module.
//!
//! Three transformations, all expressed as minimal byte-span [`TextEdit`]s so
//! non-ASCII (Windows-1252) content is never re-encoded:
//!
//!   1. **Keyword case** — normalize each keyword token to the canonical
//!      spelling from the keyword table ([`Kw::name`]). This mirrors what the
//!      VB6 IDE does when it detokenizes a stored line for display.
//!   2. **Indentation** — reindent block bodies (a deliberate, editor-level
//!      addition; VB6 itself stores leading whitespace literally). Driven by
//!      the engine's own lexer so block keywords inside strings/comments and
//!      line-continued statements are never misread.
//!   3. **Trailing whitespace** — strip spaces/tabs at end of line.
//!
//! Edits are emitted in source order and never overlap (the indent edit covers
//! the leading-whitespace run, keyword edits cover token spans in the content,
//! and the trailing edit covers the run after the last content byte).

use std::collections::HashSet;

use super::{Position, Session, TextEdit};
use crate::frontend::ast::{ExprNode, NodeId};
use crate::frontend::scanner::{Scanner, ScannerContext};
use crate::frontend::token::{Kw, Span, TokenKind};

/// One indentation level, in spaces.
const INDENT: &str = "    ";

/// Tokenize an entire source buffer, dropping the synthetic `Eol`/`Eof`/`Error`
/// markers but keeping comment (`Apos`) and line-continuation (`LineCont`)
/// tokens, which formatting needs. Spans are absolute byte offsets.
pub(super) fn tokenize(src: &[u8]) -> Vec<(TokenKind, Span)> {
    let mut ctx = ScannerContext::new(1, 1, 0x0409);
    ctx.intern_keywords();
    let mut sc = Scanner::new(&mut ctx, src);
    let mut out = Vec::new();
    loop {
        let t = sc.next_token();
        match t.kind {
            TokenKind::Eof => break,
            TokenKind::Eol | TokenKind::Error => continue,
            _ => out.push((t.kind, t.span)),
        }
    }
    out
}

/// A block frame on the indentation stack. Only the distinctions that affect
/// indentation are tracked (`If` for `Else`/`ElseIf`, `Select`/`Case` for the
/// two-level `Select Case` layout); everything else is `Generic`.
#[derive(PartialEq, Eq, Clone, Copy)]
pub(super) enum Frame {
    Generic,
    If,
    Select,
    Case,
}

/// The indentation effect of one logical line, from its leading keyword(s).
pub(super) enum Eff {
    /// No block change; print at the current depth.
    Neutral,
    /// Opens a generic block (`Sub`, `For`, `Do`, `With`, `Type`, …).
    OpenGeneric,
    /// Opens an `If`/`#If` block (header ends in `Then`).
    OpenIf,
    /// Opens a `Select Case` block.
    OpenSelect,
    /// Closes a generic block (`End X`, `Loop`, `Next`, `Wend`, `#End`).
    Close,
    /// Closes a `Select Case` block (`End Select`).
    CloseSelect,
    /// A `Case`/`Case Else` arm.
    Case,
    /// `Else`/`ElseIf`/`#Else`/`#ElseIf`: dedented one level, no net change.
    ElseLike,
}

pub(super) fn is_modifier(kw: Kw) -> bool {
    matches!(kw, Kw::Public | Kw::Private | Kw::Friend | Kw::Global | Kw::Static)
}

/// Classify a logical line from its significant tokens (comments / line
/// continuations already removed).
pub(super) fn classify(sig: &[(TokenKind, Span)]) -> Eff {
    let Some(lead) = leading_keyword(sig) else {
        return Eff::Neutral;
    };
    let last_is_then = matches!(sig.last(), Some((TokenKind::Kw(Kw::Then), _)));

    match lead {
        Kw::Sub | Kw::Function | Kw::Property | Kw::For | Kw::Do | Kw::While | Kw::With
        | Kw::Type | Kw::Enum => Eff::OpenGeneric,
        Kw::Select => Eff::OpenSelect,
        Kw::If | Kw::CcIf if last_is_then => Eff::OpenIf,
        Kw::Else | Kw::ElseIf | Kw::CcElse | Kw::CcElseIf => Eff::ElseLike,
        Kw::Case => Eff::Case,
        Kw::Loop | Kw::Next | Kw::Wend | Kw::CcEnd | Kw::EndIf => Eff::Close,
        Kw::End => classify_end(sig),
        _ => Eff::Neutral,
    }
}

/// The leading keyword of a logical line, skipping visibility modifiers, or
/// `None` if the first significant token is not a keyword.
fn leading_keyword(sig: &[(TokenKind, Span)]) -> Option<Kw> {
    for (k, _) in sig {
        if let TokenKind::Kw(kw) = k {
            if is_modifier(*kw) {
                continue;
            }
            return Some(*kw);
        }
        break;
    }
    None
}

/// Effect of an `End …` line: `End Select` is special; other block enders close
/// one level; bare `End` (terminate) leaves indentation unchanged.
fn classify_end(sig: &[(TokenKind, Span)]) -> Eff {
    let after = sig
        .iter()
        .filter_map(|(k, _)| if let TokenKind::Kw(w) = k { Some(*w) } else { None })
        .skip_while(|&k| k != Kw::End)
        .nth(1);
    match after {
        Some(Kw::Select) => Eff::CloseSelect,
        Some(Kw::Sub | Kw::Function | Kw::Property | Kw::If | Kw::With | Kw::Type | Kw::Enum) => {
            Eff::Close
        }
        _ => Eff::Neutral,
    }
}

/// Apply a line's effect to the block stack, returning the print level for the
/// first physical line of the logical line.
pub(super) fn apply(stack: &mut Vec<Frame>, eff: Eff) -> usize {
    match eff {
        Eff::Neutral => stack.len(),
        Eff::OpenGeneric => {
            let l = stack.len();
            stack.push(Frame::Generic);
            l
        }
        Eff::OpenIf => {
            let l = stack.len();
            stack.push(Frame::If);
            l
        }
        Eff::OpenSelect => {
            let l = stack.len();
            stack.push(Frame::Select);
            l
        }
        Eff::ElseLike => stack.len().saturating_sub(1),
        Eff::Close => {
            stack.pop();
            stack.len()
        }
        Eff::Case => {
            if stack.last() == Some(&Frame::Case) {
                stack.pop();
            }
            let l = stack.len();
            stack.push(Frame::Case);
            l
        }
        Eff::CloseSelect => {
            if stack.last() == Some(&Frame::Case) {
                stack.pop();
            }
            if stack.last() == Some(&Frame::Select) {
                stack.pop();
            }
            stack.len()
        }
    }
}

impl Session {
    /// Reformat a module, returning minimal byte-span edits (or empty if the
    /// module is unknown or already formatted).
    pub fn format(&self, module: usize) -> Vec<TextEdit> {
        let Some(m) = self.modules.get(module) else {
            return Vec::new();
        };
        let src = m.source.as_slice();
        let li = &m.line_index;
        let line_count = li.line_count();
        let toks = tokenize(src);

        // Group token indices by physical line.
        let mut by_line: Vec<Vec<usize>> = vec![Vec::new(); line_count];
        for (i, (_, sp)) in toks.iter().enumerate() {
            let ln = li.position(sp.start).line as usize;
            if ln < line_count {
                by_line[ln].push(i);
            }
        }

        let mut edits = Vec::new();
        recase_keywords(&toks, src, module, &identifier_offsets(m), &mut edits);

        // Indentation + trailing trim, walking logical lines (physical lines
        // joined by a trailing `_` continuation).
        let mut stack: Vec<Frame> = Vec::new();
        let mut line = 0usize;
        while line < line_count {
            // Extend the logical line over continuation lines.
            let mut group = vec![line];
            while line_ends_with_cont(by_line[*group.last().unwrap()].as_slice(), &toks)
                && *group.last().unwrap() + 1 < line_count
            {
                let next = *group.last().unwrap() + 1;
                group.push(next);
            }

            let sig: Vec<(TokenKind, Span)> = group
                .iter()
                .flat_map(|&l| by_line[l].iter())
                .map(|&i| toks[i].clone())
                .filter(|(k, _)| {
                    !matches!(k, TokenKind::Kw(Kw::Apos) | TokenKind::Kw(Kw::LineCont))
                })
                .collect();
            let level = apply(&mut stack, classify(&sig));

            for (gi, &l) in group.iter().enumerate() {
                let lvl = if gi == 0 { level } else { level + 1 };
                let has_tokens = !by_line[l].is_empty();
                emit_line_layout(src, li, module, l, lvl, has_tokens, &mut edits);
            }

            line = *group.last().unwrap() + 1;
        }

        edits
    }
}

/// Whether the last token on a physical line is a line-continuation marker.
fn line_ends_with_cont(line_toks: &[usize], toks: &[(TokenKind, Span)]) -> bool {
    line_toks
        .last()
        .map(|&i| matches!(toks[i].0, TokenKind::Kw(Kw::LineCont)))
        .unwrap_or(false)
}

/// Emit the indentation and trailing-whitespace edits for one physical line.
fn emit_line_layout(
    src: &[u8],
    li: &super::LineIndex,
    module: usize,
    line: usize,
    level: usize,
    has_tokens: bool,
    edits: &mut Vec<TextEdit>,
) {
    let line_start = li.offset(Position { line: line as u32, character: 0 }) as usize;
    let next_start = li.offset(Position { line: line as u32 + 1, character: 0 }) as usize;
    let mut line_end = next_start.max(line_start);
    while line_end > line_start && matches!(src[line_end - 1], b'\n' | b'\r') {
        line_end -= 1;
    }

    // Content start: first non-space/tab byte.
    let mut cs = line_start;
    while cs < line_end && matches!(src[cs], b' ' | b'\t') {
        cs += 1;
    }

    // Desired indentation: blank/whitespace-only lines collapse to empty.
    let desired = if has_tokens && cs < line_end {
        INDENT.repeat(level)
    } else {
        String::new()
    };
    if &src[line_start..cs] != desired.as_bytes() {
        edits.push(TextEdit {
            module,
            span: Span { start: line_start as u32, len: (cs - line_start) as u32 },
            new_text: desired,
        });
    }

    // Trailing whitespace after the last content byte.
    if cs < line_end {
        let mut ce = line_end;
        while ce > cs && matches!(src[ce - 1], b' ' | b'\t') {
            ce -= 1;
        }
        if ce < line_end {
            edits.push(TextEdit {
                module,
                span: Span { start: ce as u32, len: (line_end - ce) as u32 },
                new_text: String::new(),
            });
        }
    }
}

#[cfg(test)]
mod tests {
    use crate::session::{Session, TextEdit};

    /// Apply span edits to a source string (right-to-left so offsets stay valid).
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

    fn formatted(src: &str) -> String {
        let s = Session::from_sources(vec![("m.bas".to_string(), src.as_bytes().to_vec())]);
        apply(src, &s.format(0))
    }

    #[test]
    fn keyword_case_and_indent() {
        let out = formatted("sub foo()\nx = 1\nend sub\n");
        assert_eq!(out, "Sub foo()\n    x = 1\nEnd Sub\n");
    }

    #[test]
    fn select_case_two_levels() {
        let src = "Select Case x\nCase 1\nFoo\nCase Else\nBar\nEnd Select\n";
        let out = formatted(src);
        assert_eq!(
            out,
            "Select Case x\n    Case 1\n        Foo\n    Case Else\n        Bar\nEnd Select\n"
        );
    }

    #[test]
    fn nested_if_inside_sub() {
        let src = "Sub S()\nIf a Then\nb = 1\nElse\nb = 2\nEnd If\nEnd Sub\n";
        let out = formatted(src);
        assert_eq!(
            out,
            "Sub S()\n    If a Then\n        b = 1\n    Else\n        b = 2\n    End If\nEnd Sub\n"
        );
    }

    #[test]
    fn trims_trailing_and_blank_lines() {
        let out = formatted("Sub Foo()   \n\n   \nEnd Sub\n");
        assert_eq!(out, "Sub Foo()\n\n\nEnd Sub\n");
    }

    #[test]
    fn keyword_in_string_is_left_alone() {
        // "end sub" lives in a string literal and is not a keyword token.
        let src = "x = \"end sub\"\n";
        assert_eq!(formatted(src), src);
    }

    #[test]
    fn member_named_like_keyword_not_recased() {
        // `.name` is a member access, not the `Name` keyword.
        let src = "x = obj.name\n";
        assert_eq!(formatted(src), src);
    }
}

/// Byte offsets of tokens the bound model treats as *identifiers* — every
/// `NameRef` use site and every declared name. Some VBA keywords (`B`, `Name`,
/// `Line`, `Error`, …) are also legal identifiers; the scanner still interns
/// them as keyword tokens, so without this set the formatter would "recase" a
/// variable named `b` to `B`. Skipping these offsets keeps casing changes to
/// tokens that are keywords *in context*.
fn identifier_offsets(m: &super::ModuleData) -> HashSet<u32> {
    let mut set = HashSet::new();
    collect_nameref_offsets(m, &mut set);
    collect_decl_name_offsets(m, &mut set);
    set
}

/// Add the start offset of every `NameRef` use site to `set`.
fn collect_nameref_offsets(m: &super::ModuleData, set: &mut HashSet<u32>) {
    for i in 0..m.arena.len() as u32 {
        if let ExprNode::NameRef { .. } = m.arena.get(NodeId(i)) {
            let sp = m.spans.get(NodeId(i));
            if sp.len > 0 {
                set.insert(sp.start);
            }
        }
    }
}

/// Add the start offset of every declared name (procs/params/locals, module
/// vars, types and members, enums and members) to `set`.
fn collect_decl_name_offsets(m: &super::ModuleData, set: &mut HashSet<u32>) {
    let mut add = |sp: Span| {
        if sp.len > 0 {
            set.insert(sp.start);
        }
    };
    for p in &m.bound.procs {
        add(p.name_span);
        for prm in &p.params {
            add(prm.name_span);
        }
        for loc in &p.locals {
            add(loc.name_span);
        }
    }
    for v in &m.bound.module_vars {
        add(v.name_span);
    }
    for t in &m.bound.type_decls {
        add(t.name_span);
        for mem in &t.members {
            add(mem.name_span);
        }
    }
    for e in &m.bound.enum_decls {
        add(e.name_span);
        for mem in &e.members {
            add(mem.name_span);
        }
    }
}

/// Emit a case-normalizing edit for every word keyword whose source casing
/// differs from the canonical table spelling. Member names after `.`/`!`
/// (`obj.Name`) and tokens the binder treats as identifiers are left alone even
/// when they collide with a keyword.
fn recase_keywords(
    toks: &[(TokenKind, Span)],
    src: &[u8],
    module: usize,
    ident_offsets: &HashSet<u32>,
    edits: &mut Vec<TextEdit>,
) {
    let mut prev: Option<&TokenKind> = None;
    for (k, sp) in toks {
        if let TokenKind::Kw(kw) = k {
            if let Some(edit) = recase_edit(*kw, *sp, src, module, prev, ident_offsets) {
                edits.push(edit);
            }
        }
        prev = Some(k);
    }
}

/// The case-normalizing edit for one keyword token, or `None` when it must be
/// left alone (member name, binder identifier, non-word keyword, or already
/// canonically cased).
fn recase_edit(
    kw: Kw,
    sp: Span,
    src: &[u8],
    module: usize,
    prev: Option<&TokenKind>,
    ident_offsets: &HashSet<u32>,
) -> Option<TextEdit> {
    let after_member = matches!(
        prev,
        Some(TokenKind::Kw(Kw::Dot)) | Some(TokenKind::Kw(Kw::DotStmt)) | Some(TokenKind::Kw(Kw::Bang))
    );
    if after_member || ident_offsets.contains(&sp.start) {
        return None;
    }
    let name = kw.name();
    // Only recase plain word keywords whose length matches the source token
    // (skips operators, `$`-suffixed and `VB_*` forms).
    if name.len() as u32 != sp.len || !name.bytes().all(|b| b.is_ascii_alphabetic()) {
        return None;
    }
    let cur = &src[sp.start as usize..(sp.start + sp.len) as usize];
    if cur != name.as_bytes() && cur.eq_ignore_ascii_case(name.as_bytes()) {
        Some(TextEdit { module, span: sp, new_text: name.to_string() })
    } else {
        None
    }
}
