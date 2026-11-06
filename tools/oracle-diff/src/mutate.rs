//! Token-level mutation of compiling modules. Each mutant differs from its
//! parent by one change, so the compiler's first error (if any) is caused by
//! that change.

use vb6_syntax::frontend::scanner::{Scanner, ScannerContext};
use vb6_syntax::frontend::token::{Kw, TokenKind};

/// One scanned token of the parent: its byte range and the kind of line break it is.
#[derive(Clone, Debug, PartialEq, Eq)]
pub struct Piece {
    pub start: usize,
    pub end: usize,
    pub is_eol: bool,
}

/// A single-change variant of a parent module.
#[derive(Clone, Debug)]
pub struct Mutant {
    pub description: String,
    /// 1-based line of the parent on which the change was made.
    pub line: u32,
    pub source: Vec<u8>,
}

/// Deterministic generator (xorshift64*), so a seed reproduces a run.
pub struct Rng(u64);

impl Rng {
    pub fn new(seed: u64) -> Rng {
        Rng(seed.wrapping_mul(0x9E3779B97F4A7C15) | 1)
    }

    pub fn next(&mut self) -> u64 {
        let mut x = self.0;
        x ^= x >> 12;
        x ^= x << 25;
        x ^= x >> 27;
        self.0 = x;
        x.wrapping_mul(0x2545F4914F6CDD1D)
    }

    pub fn below(&mut self, n: usize) -> usize {
        (self.next() % n as u64) as usize
    }
}

/// Words and punctuation substituted into or inserted before a token.
pub const VOCABULARY: &[&str] = &[
    "(", ")", ",", ".", "!", ":", "=", "+", "-", "*", "/", "\\", "^", "&", "<", ">", "<>", "<=", ">=",
    "And", "Or", "Not", "Mod", "Is", "Like", "Then", "Else", "End", "As", "To", "Step", "Next", "Dim",
    "Set", "Let", "Call", "New", "With", "ByVal", "ByRef", "Optional", "Const", "Static", "Public",
    "Private", "Integer", "String", "Variant", "Object", "Nothing", "True", "False", "0", "1", "&H10",
    "1.5", "\"s\"", "#1/1/2000#", "x", "Me", ";", "_", "$", "%", "#", "@",
];

/// Splits the code of `source` into tokens. Comments are left out.
pub fn pieces(source: &[u8]) -> Vec<Piece> {
    let mut ctx = ScannerContext::new(1, 1, 0x0409);
    ctx.intern_keywords();
    let mut scanner = Scanner::new(&mut ctx, source);
    let mut out = Vec::new();
    let mut in_comment = false;
    loop {
        let token = scanner.next_token();
        let start = token.span.start as usize;
        let end = start + token.span.len as usize;
        match token.kind {
            TokenKind::Eof => break,
            TokenKind::Eol => {
                in_comment = false;
                out.push(Piece { start, end: end.max(start), is_eol: true });
            }
            TokenKind::Kw(Kw::Apos) | TokenKind::Kw(Kw::Rem) => in_comment = true,
            _ if in_comment => {}
            _ => out.push(Piece { start, end, is_eol: false }),
        }
        if out.len() > source.len() + 8 {
            break; // a scanner that stops advancing must not hang the run
        }
    }
    out
}

fn line_of(source: &[u8], offset: usize) -> u32 {
    1 + source[..offset.min(source.len())].iter().filter(|b| **b == b'\n').count() as u32
}

/// Number of lines in the module header (class `VERSION` block and leading
/// `Attribute` lines). The compiler does not count them in the line numbers it
/// reports.
pub fn header_lines(source: &[u8]) -> u32 {
    let end = header_end(source);
    vb6_syntax::frontend::line_reader::physical_line_starts(&source[..end]).len() as u32 - 1
}

/// End of the module header (class `VERSION` block and leading `Attribute` lines):
/// tokens before it are not mutated.
fn header_end(source: &[u8]) -> usize {
    let mut end = 0;
    for line in vb6_syntax::frontend::line_reader::physical_lines(source).take(40) {
        let text = String::from_utf8_lossy(&source[line.start..line.end]);
        let t = text.trim_start().to_ascii_lowercase();
        if t.starts_with("attribute vb_") || t.starts_with("version ") || t.starts_with("begin")
            || t.trim_end() == "end" || t.starts_with("multiuse")
            || t.starts_with("persistable") || t.starts_with("databinding") || t.starts_with("datasource")
            || t.starts_with("mtstransaction")
        {
            end = line.next;
        } else if !t.trim().is_empty() {
            // The header is only the leading block: a code line ends it, so a later
            // `End` or `Attribute VB_...` line belongs to the code and is counted.
            break;
        }
    }
    end
}

/// Produces `count` mutants of `source` from `seed`.
pub fn mutants(source: &[u8], seed: u64, count: usize) -> Vec<Mutant> {
    let all = pieces(source);
    let floor = header_end(source);
    let eligible: Vec<usize> = (0..all.len()).filter(|i| all[*i].start >= floor).collect();
    let mut rng = Rng::new(seed);
    let mut out = Vec::new();
    if eligible.is_empty() {
        return out;
    }
    let mut attempts = 0;
    while out.len() < count && attempts < count * 20 {
        attempts += 1;
        let i = eligible[rng.below(eligible.len())];
        let piece = &all[i];
        let text = &source[piece.start..piece.end];
        let line = line_of(source, piece.start);
        let (description, replacement, from, to): (String, Vec<u8>, usize, usize) = match rng.below(5) {
            0 => (
                format!("delete `{}`", String::from_utf8_lossy(text).trim_end()),
                Vec::new(),
                piece.start,
                piece.end,
            ),
            1 if !piece.is_eol => {
                let mut dup = text.to_vec();
                dup.push(b' ');
                dup.extend_from_slice(text);
                (format!("duplicate `{}`", String::from_utf8_lossy(text)), dup, piece.start, piece.end)
            }
            2 => {
                let Some(next) = all.get(i + 1) else { continue };
                if piece.is_eol || next.is_eol {
                    continue;
                }
                let mut swapped = source[next.start..next.end].to_vec();
                swapped.extend_from_slice(&source[piece.end..next.start]);
                swapped.extend_from_slice(text);
                (
                    format!(
                        "swap `{}` and `{}`",
                        String::from_utf8_lossy(text),
                        String::from_utf8_lossy(&source[next.start..next.end])
                    ),
                    swapped,
                    piece.start,
                    next.end,
                )
            }
            3 if !piece.is_eol => {
                let word = VOCABULARY[rng.below(VOCABULARY.len())];
                (
                    format!("replace `{}` with `{word}`", String::from_utf8_lossy(text)),
                    word.as_bytes().to_vec(),
                    piece.start,
                    piece.end,
                )
            }
            _ => {
                let word = VOCABULARY[rng.below(VOCABULARY.len())];
                let mut ins = word.as_bytes().to_vec();
                ins.push(b' ');
                (format!("insert `{word}` before token {}", i), ins, piece.start, piece.start)
            }
        };
        let mut mutated = Vec::with_capacity(source.len() + replacement.len());
        mutated.extend_from_slice(&source[..from]);
        mutated.extend_from_slice(&replacement);
        mutated.extend_from_slice(&source[to..]);
        if mutated != source {
            out.push(Mutant { description, line, source: mutated });
        }
    }
    out
}
