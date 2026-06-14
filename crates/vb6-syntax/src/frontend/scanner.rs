//! VB6 scanner symbol table and keyword interning.
//!
//! Key functions:
//!   * `intern_keywords` — interns the 271-entry keyword/operator table into the
//!     scanner context, building the token→symbol and length tables and
//!     recording the "Left" string id.
//!   * `intern_string` — hash + insert/lookup a name in the symbol table.
//!
//! ## Implementation details
//!
//! The symbol table uses a Rust [`Vec`] for storing symbol records and a
//! case-insensitive [`HashMap`] for name lookups. Symbols are assigned IDs in
//! insertion order.
//!
//! The current implementation focuses on the keyword-interning path:
//! `intern_keywords` interns the standard set of keywords.

use std::collections::HashMap;

use super::keyword_table::KEYWORD_TABLE;
use super::token::{Kw, Lit, Span, Token, TokenKind, TypeSuffix};

/// `E_OUTOFMEMORY` — returned by `intern_keywords` when a symbol can't be
/// allocated (`0x8007000e`).
pub const E_OUTOFMEMORY: u32 = 0x8007_000e;

/// Index of a [`Symbol`] within a [`ScannerContext`]'s symbol arena.
pub type SymId = usize;

/// An interned name record — the symbol descriptor. Only the semantically
/// meaningful fields are modeled.
#[derive(Clone, Debug)]
pub struct Symbol {
    /// value/cookie.
    pub value: u32,
    /// dispid from the persistence path; 0 for keyword interning.
    pub dispid: i16,
    /// token id.
    pub token: i16,
    /// string id, assigned in insertion order as `counter << 1`.
    pub id: u16,
    /// flag bits. Bit 1 = name keeps a trailing `$`; bit 2 = this is the
    /// `$`-stripped form of another name; bit 4 = reserved word; bit 8 =
    /// persistence-path marker.
    pub flags: u8,
    /// default attribute word, always `0x3ff` at allocation.
    pub attr: u16,
    /// the interned name, original spelling.
    pub name: String,
}

/// The scanner context — contains the symbol table and keyword mapping.
pub struct ScannerContext {
    /// LCID used for name hashing.
    pub lcid: u32,
    /// Linkage to parent LCID if present.
    pub parent_lcid: Option<u32>,
    /// scanner options flags.
    pub options: u32,
    /// token-buffer capacity.
    pub token_buf_size: u32,
    /// mode flags.
    pub mode_flags: u8,
    /// keyword-init flag.
    pub do_keyword_init: u8,
    /// scanner state flags. Bit 1 (`0x2`) guards "keywords interned".
    pub flags6: u16,
    /// the string id of `Left` (token `0x6d`), captured during interning.
    pub field4: u16,
    /// set when a lookup replaced a flagged duplicate.
    pub dup_replaced: bool,
    /// count of interned names; the next id is `(counter << 1)`.
    pub counter: u32,

    /// Symbol arena (owns every record).
    symbols: Vec<Symbol>,
    /// Case-folded name → symbol index.
    by_key: HashMap<String, SymId>,
    /// Id-ordered name table (the slot array); only grown when
    /// `mode_flags != 0`. Index `k` holds the symbol with id `(k + 1) << 1`.
    name_index: Vec<SymId>,

    /// token → symbol map, indexed by token id.
    token_table: Vec<Option<SymId>>,
    /// per-keyword-entry name length, indexed by table position.
    keyword_len: Vec<u8>,
    /// keyword name pointers.
    keyword_names: Vec<&'static str>,
}

impl ScannerContext {
    /// Create a context with keyword pointers populated and tables cleared.
    pub fn new(mode_flags: u8, do_keyword_init: u8, lcid: u32) -> Self {
        Self::with_options(mode_flags, do_keyword_init, lcid, 0, 0x1000)
    }

    /// Full constructor: also records `options` and `token_buf_size`.
    pub fn with_options(
        mode_flags: u8,
        do_keyword_init: u8,
        lcid: u32,
        options: u32,
        token_buf_size: u32,
    ) -> Self {
        let n = KEYWORD_TABLE.len();
        Self {
            lcid,
            parent_lcid: None,
            options,
            token_buf_size,
            mode_flags,
            do_keyword_init,
            flags6: 0,
            field4: 0,
            dup_replaced: false,
            counter: 0,
            symbols: Vec::new(),
            by_key: HashMap::new(),
            name_index: Vec::new(),
            token_table: vec![None; n],
            keyword_len: vec![0u8; n],
            keyword_names: KEYWORD_TABLE.iter().map(|e| e.name).collect(),
        }
    }

    // --- read accessors -----------------------------------------------------

    /// Borrow a symbol by id.
    pub fn symbol(&self, id: SymId) -> &Symbol {
        &self.symbols[id]
    }

    /// Number of interned symbols.
    pub fn symbol_count(&self) -> usize {
        self.symbols.len()
    }

    /// Look up an interned symbol by name (case-insensitive).
    pub fn lookup(&self, name: &str) -> Option<&Symbol> {
        self.by_key
            .get(&name.to_ascii_lowercase())
            .map(|&id| &self.symbols[id])
    }

    /// The symbol id mapped to a token, if any.
    pub fn token_sym_id(&self, token: u16) -> Option<SymId> {
        self.token_table.get(token as usize).copied().flatten()
    }

    /// The symbol mapped to a token id, if any.
    pub fn token_symbol(&self, token: u16) -> Option<&Symbol> {
        self.token_sym_id(token).map(|id| &self.symbols[id])
    }

    /// The recorded length for a keyword-table entry.
    pub fn keyword_len(&self, index: usize) -> u8 {
        self.keyword_len[index]
    }

    // --- interning functions ------------------------------------------------

    /// Interns a string for keywords or non-persistence paths.
    fn intern_string(&mut self, name: &str, param_2: u32) -> Option<SymId> {
        self.dup_replaced = false;
        let sym = self.intern_core(name, param_2)?;
        // VB6's `param_2 & 2` (persistence flag) and `param_4 & 6 == 0`
        // (general-identifier lookup) paths are unused in an LSP context:
        // every caller passes param_2 = 4, and do_keyword_init is always
        // non-zero during keyword initialisation.  No implementation needed.
        Some(sym)
    }

    /// Look up `name` returning the existing symbol, or allocate and link a new one.
    fn intern_core(&mut self, name: &str, param_4: u32) -> Option<SymId> {
        let key = name.to_ascii_lowercase();

        // --- found path ---
        if let Some(&existing) = self.by_key.get(&key) {
            return Some(existing);
        }

        // --- not found: allocate ---
        if self.mode_flags != 0 && self.counter > 0x7ffe {
            return None;
        }

        let mut sym = Symbol {
            value: 0,
            dispid: 0,
            token: 0,
            id: 0,
            flags: 0,
            attr: 0x3ff,
            name: name.to_string(),
        };
        if param_4 & 2 != 0 {
            sym.flags |= 8;
        }

        if self.mode_flags != 0 {
            self.counter += 1;
        }
        sym.id = (self.counter as u16) << 1;

        let id = self.symbols.len();
        self.symbols.push(sym);
        self.by_key.insert(key, id);
        if self.mode_flags != 0 {
            self.name_index.push(id);
        }
        Some(id)
    }

    /// Interns the full set of keywords and operators.
    ///
    /// Walks the keyword/operator table and records: the token field on each
    /// symbol, the token→symbol table, the per-entry length, the `Left` string id,
    /// the reserved-word flag, and the `$`-suffix pair flags.
    pub fn intern_keywords(&mut self) -> u32 {
        // Already interned -> nothing to do.
        if self.flags6 & 2 != 0 {
            return 0;
        }
        // Adopt the parent LCID if linked.
        if let Some(parent_lcid) = self.parent_lcid {
            self.lcid = parent_lcid;
        }
        self.flags6 |= 2; // set the "interned" guard

        for (index, entry) in KEYWORD_TABLE.iter().enumerate() {
            let w0 = entry.w0;
            let w1 = entry.w1;
            let name = self.keyword_names[index];
            let token = (w0 & 0xffff) as u16;

            // In limited mode, intern only Left (0x6d) and Object (0x88).
            if self.mode_flags == 0 && token != 0x6d && token != 0x88 {
                continue;
            }

            if self.intern_keyword_entry(index, name, token, w1).is_none() {
                self.flags6 &= !2;
                return E_OUTOFMEMORY;
            }
        }
        0
    }

    /// Intern a single keyword-table entry: record length, token mapping,
    /// reserved/`Left` flags, and any `$`-stripped companion. Returns `None`
    /// on allocation failure.
    fn intern_keyword_entry(
        &mut self,
        index: usize,
        name: &'static str,
        token: u16,
        w1: u32,
    ) -> Option<()> {
        let len = name.len();
        self.keyword_len[index] = len as u8;

        let sym = self.intern_string(name, 4)?;

        self.symbols[sym].token = token as i16;
        self.token_table[token as usize] = Some(sym);

        if (w1 >> 4) & 1 == 0 {
            if token == 0x6d {
                // "Left": record its string id.
                self.field4 = self.symbols[sym].id;
            }
        } else {
            self.symbols[sym].flags |= 4; // reserved word
        }

        // trailing '$' -> also intern the '$'-stripped name and cross-flag
        if name.as_bytes()[len - 1] == b'$' {
            let stripped = &name[..len - 1];
            let s2 = self.intern_string(stripped, 4)?;
            self.symbols[sym].flags |= 1; // original (with '$'): flag 1
            self.symbols[s2].flags |= 2; // stripped form: flag 2
        }
        Some(())
    }

    // --- general identifier path --------------------------------------------

    /// Intern a user-defined identifier name, returning its arena index.
    pub fn intern_ident(&mut self, name: &str) -> Option<SymId> {
        let key = name.to_ascii_lowercase();
        if let Some(&existing) = self.by_key.get(&key) {
            return Some(existing);
        }
        if self.mode_flags != 0 && self.counter > 0x7ffe {
            return None;
        }
        if self.mode_flags != 0 {
            self.counter += 1;
        }
        let sym = Symbol {
            value: 0,
            dispid: 0,
            token: 0,        // token = 0 → identifier
            id: (self.counter as u16) << 1,
            flags: 0,
            attr: 0x3ff,
            name: name.to_string(),
        };
        let id = self.symbols.len();
        self.symbols.push(sym);
        self.by_key.insert(key, id);
        if self.mode_flags != 0 {
            self.name_index.push(id);
        }
        Some(id)
    }

    /// Return the canonical name for a symbol by arena index.
    /// Returns "" when `id` is out of range.
    pub fn sym_name(&self, id: usize) -> &str {
        self.symbols.get(id).map(|s| s.name.as_str()).unwrap_or("")
    }

    /// Set the "verbatim / bracketed" flag (bit 7) on a symbol.
    pub fn set_verbatim_flag(&mut self, id: SymId) {
        self.symbols[id].flags |= 0x80;
    }

    /// Returns the name of a keyword by its index.
    pub fn keyword_string(index: usize) -> &'static str {
        KEYWORD_TABLE[index].name
    }
}

// ── Char-class table ──────────────────────────────────────────────────────────

/// Char-class dispatch table.
///
/// Covers byte values 0x00–0xFF.
pub const CHAR_CLASSES: [u8; 256] = [
    // 0x00–0x0f  NUL + control → class 0 (end-of-source sentinel / skip)
    0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00,
    0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00,
    // 0x10–0x1f  more controls; 0x14 = class 1 (EOL marker)
    0x00, 0x00, 0x00, 0x00, 0x01, 0x02, 0x02, 0x03,
    0x03, 0x03, 0x03, 0x03, 0x03, 0x04, 0xb7, 0xb7,
    // 0x20–0x2f  space · ! " # $ % & ' ( ) * + , - . /
    0x05, 0x06, 0x07, 0x08, 0x09, 0x0a, 0x0b, 0x0c,
    0x0d, 0x0e, 0x0f, 0x10, 0x11, 0x12, 0x13, 0x14,
    // 0x30–0x3f  0–9 : ; < = > ?
    0x15, 0x16, 0x17, 0x18, 0x19, 0x1a, 0x1b, 0x1c,
    0x1d, 0x1e, 0x1f, 0x20, 0x21, 0x22, 0x23, 0x24,
    // 0x40–0x4f  @ A B C D E F G H I J K L M N O
    0x25, 0x26, 0x27, 0x28, 0x29, 0x2a, 0x2b, 0x2c,
    0xb7, 0x2d, 0x2e, 0x2f, 0x30, 0x31, 0x32, 0x33,
    // 0x50–0x5f  P Q R S T U V W X Y Z [ \ ] ^ _
    0x34, 0x35, 0x36, 0x37, 0x38, 0x39, 0x3a, 0x3a,
    0x3b, 0x3b, 0x3c, 0x3d, 0x3e, 0x3f, 0x40, 0x41,
    // 0x60–0x6f  ` a b c d e f g h i j k l m n o
    0xb7, 0x42, 0x43, 0x44, 0x45, 0x46, 0x47, 0x48,
    0x49, 0x4a, 0x4b, 0x4c, 0xb7, 0x4d, 0x4e, 0x4f,
    // 0x70–0x7f  p q r s t u v w x y z { | } ~ DEL
    0x50, 0x51, 0x52, 0x53, 0x54, 0x55, 0x56, 0x57,
    0x58, 0x59, 0x5a, 0x5a, 0x5a, 0xb7, 0xb7, 0xb7,
    // 0x80–0x8f  extended Latin (Windows-1252)
    0xb7, 0xb7, 0xb7, 0xb7, 0x5b, 0x5b, 0x5b, 0x5b,
    0x5b, 0x5b, 0x5c, 0xb7, 0xb7, 0x5d, 0x5d, 0xb7,
    // 0x90–0x9f
    0xb7, 0x5e, 0x5f, 0x60, 0x61, 0x62, 0x63, 0xb7,
    0x64, 0x65, 0x65, 0x66, 0x66, 0x67, 0x68, 0x69,
    // 0xa0–0xaf
    0x6a, 0x6a, 0x6a, 0x6b, 0xb7, 0x6c, 0x6d, 0x6e,
    0x6f, 0x70, 0x71, 0x72, 0x73, 0x74, 0x75, 0x76,
    // 0xb0–0xbf
    0x77, 0x78, 0x79, 0x7a, 0x7b, 0x7c, 0x7d, 0x7e,
    0x7f, 0x80, 0x81, 0x82, 0x83, 0x84, 0x84, 0x85,
    // 0xc0–0xcf
    0x86, 0x87, 0x88, 0x89, 0x89, 0x8a, 0x8b, 0x8c,
    0x8c, 0x8d, 0x8e, 0x8f, 0x90, 0x91, 0x92, 0x93,
    // 0xd0–0xdf
    0x94, 0x95, 0x96, 0x96, 0x96, 0x96, 0x96, 0x96,
    0x96, 0x96, 0x96, 0x96, 0x96, 0x96, 0x97, 0x98,
    // 0xe0–0xef
    0x99, 0x9a, 0x9a, 0x9b, 0x9c, 0x9d, 0x9e, 0x9f,
    0xa0, 0xa1, 0xa2, 0xa3, 0xa3, 0xb7, 0xb7, 0xa4,
    // 0xf0–0xff
    0xa5, 0xa6, 0xa7, 0xa8, 0xa9, 0xaa, 0xab, 0xac,
    0xad, 0xae, 0xb7, 0xaf, 0xb7, 0xb7, 0xb0, 0xb1,
];

/// Returns true if `b` may *continue* (not start) a VBA identifier.
#[inline]
fn is_ident_continue(b: u8) -> bool {
    matches!(b, b'A'..=b'Z' | b'a'..=b'z' | b'0'..=b'9' | b'_' | 0x80..=0xFF)
}

/// True if `b` can begin an identifier (or a bracketed/member name).
#[inline]
fn is_ident_start(b: u8) -> bool {
    matches!(b, b'A'..=b'Z' | b'a'..=b'z' | b'_' | b'[' | 0x80..=0xFF)
}

// ── Date literal parser ───────────────────────────────────────────────────────

/// Convert a Gregorian year/month/day to days since the proleptic Gregorian
/// epoch (March 1, 0000).
fn gregorian_to_days(y: i32, m: i32, d: i32) -> i32 {
    let y = if m <= 2 { y - 1 } else { y };
    let era = if y >= 0 { y } else { y - 399 } / 400;
    let yoe = y - era * 400;
    let m_adj = if m > 2 { m - 3 } else { m + 9 };
    let doy = (153 * m_adj + 2) / 5 + d - 1;
    let doe = yoe * 365 + yoe / 4 - yoe / 100 + doy;
    era * 146097 + doe
}

/// VBA date serial for a given Gregorian date.
///
/// Epoch = December 30, 1899 (serial 0).
fn vba_date_serial(y: i32, m: i32, d: i32) -> f64 {
    const VBA_EPOCH: i32 = 693899; // gregorian_to_days(1899, 12, 30)
    (gregorian_to_days(y, m, d) - VBA_EPOCH) as f64
}

/// Expand a two-digit year: 00–29 → 2000–2029, 30–99 → 1930–1999.
fn expand_two_digit_year(yy: i32) -> i32 {
    if yy < 30 { 2000 + yy } else { 1900 + yy }
}

/// Parse a time string `H:MM[:SS][ AM|PM]` as a day fraction (0.0–1.0).
fn parse_time_fraction(s: &str) -> f64 {
    let s = s.trim();
    let (time_str, ampm) = {
        let upper = s.to_ascii_uppercase();
        if upper.ends_with(" AM") {
            (&s[..s.len() - 3], 0i32)
        } else if upper.ends_with(" PM") {
            (&s[..s.len() - 3], 1i32)
        } else if upper.ends_with("AM") {
            (&s[..s.len() - 2], 0i32)
        } else if upper.ends_with("PM") {
            (&s[..s.len() - 2], 1i32)
        } else {
            (s, -1i32)
        }
    };
    let parts: Vec<&str> = time_str.trim().split(':').collect();
    if parts.is_empty() || parts.len() > 3 {
        return 0.0;
    }
    let h = parts[0].trim().parse::<f64>().unwrap_or(0.0);
    let m = parts.get(1).and_then(|p| p.trim().parse::<f64>().ok()).unwrap_or(0.0);
    let sec = parts.get(2).and_then(|p| p.trim().parse::<f64>().ok()).unwrap_or(0.0);
    let mut hours = h;
    if ampm == 1 && h < 12.0 { hours += 12.0; }
    if ampm == 0 && h == 12.0 { hours = 0.0; }
    (hours * 3600.0 + m * 60.0 + sec) / 86400.0
}

/// Parse a VBA date literal string and return the date serial.
fn scan_vba_date_literal(s: &str) -> f64 {
    let s = s.trim();

    let first_part = s.split_whitespace().next().unwrap_or(s);
    let has_date_sep = first_part.contains('/') || first_part.contains('-');
    if !has_date_sep && s.contains(':') {
        return parse_time_fraction(s);
    }

    let (date_part, time_part) = match s.find(' ') {
        Some(idx) => (&s[..idx], &s[idx + 1..]),
        None => (s, ""),
    };

    let time_frac = if time_part.is_empty() { 0.0 } else { parse_time_fraction(time_part) };

    for sep in [b'/', b'-'] {
        if let Some(serial) = parse_date_with_sep(date_part, sep) {
            return serial + time_frac;
        }
    }
    0.0
}

/// Parse `date_part` split on `sep` into a date serial, or `None` if it does
/// not form a valid 3-component date.
fn parse_date_with_sep(date_part: &str, sep: u8) -> Option<f64> {
    let parts: Vec<&str> = date_part.split(|c: char| c as u8 == sep).collect();
    if parts.len() != 3 {
        return None;
    }
    let (Ok(a), Ok(b), Ok(c)) = (
        parts[0].trim().parse::<i32>(),
        parts[1].trim().parse::<i32>(),
        parts[2].trim().parse::<i32>(),
    ) else {
        return None;
    };
    let (y, m, d) = if a >= 100 {
        (a, b, c)
    } else if c >= 100 {
        (c, a, b)
    } else {
        (expand_two_digit_year(c), a, b)
    };
    if m >= 1 && m <= 12 && d >= 1 && d <= 31 {
        Some(vba_date_serial(y, m, d))
    } else {
        None
    }
}

// ── Scanner ───────────────────────────────────────────────────────────────────

/// Per-statement tokenizer.
pub struct Scanner<'a> {
    ctx: &'a mut ScannerContext,
    /// Source bytes for one logical statement.
    src: &'a [u8],
    /// Current byte offset into `src`.
    pos: usize,
    /// Scratch buffer for identifier and literal accumulation.
    tok_buf: String,
}

impl<'a> Scanner<'a> {
    /// Return the canonical name for a symbol by id.
    pub fn sym_name(&self, id: u32) -> &str {
        self.ctx.sym_name(id as usize)
    }

    /// Create a scanner for `src`, borrowing `ctx` for symbol interning.
    pub fn new(ctx: &'a mut ScannerContext, src: &'a [u8]) -> Self {
        Self { ctx, src, pos: 0, tok_buf: String::with_capacity(64) }
    }

    /// Return the next token from the source, advancing the position.
    pub fn next_token(&mut self) -> Token {
        loop {
            let start = self.pos;
            let b = match self.advance() {
                None => return Token::eof(),
                Some(b) => b,
            };
            let cls = CHAR_CLASSES[b as usize];
            match cls {
                0x00 => return Token::eol(start as u32),

                0x01..=0x04 => { /* skip */ }

                0x05 => { /* skip */ }

                0x06 => return Token::op(Kw::Bang, self.span(start)),

                0x07 => return self.scan_string(start),

                0x08 => return self.scan_hash(start),

                0x09 | 0x0a | 0x25 => return Token::error(self.span(start)),

                0x0b => return self.scan_amp(start),

                0x0c => return self.consume_comment(start),

                0x0d => return Token::op(Kw::LParen, self.span(start)),

                0x0e => return Token::op(Kw::RParen, self.span(start)),

                0x0f => return Token::op(Kw::Star, self.span(start)),

                0x10 => return Token::op(Kw::Plus, self.span(start)),

                0x11 => return Token::op(Kw::Comma, self.span(start)),

                0x12 => return Token::op(Kw::Minus, self.span(start)),

                0x13 => return self.scan_dot(start),

                0x14 => return Token::op(Kw::Slash, self.span(start)),

                0x15..=0x1e => return self.scan_number(b, start),

                0x1f => return self.scan_colon(start),

                0x20 => return Token::op(Kw::Semi, self.span(start)),

                0x21 => return self.scan_lt(start),

                0x22 => return self.scan_eq(start),

                0x23 => return self.scan_gt(start),

                0x24 => return Token::op(Kw::Question, self.span(start)),

                0x3d => return self.scan_bracketed_ident(start),

                0x3e => return Token::op(Kw::Backslash, self.span(start)),

                0x3f => return Token::error(self.span(start)),

                0x40 => return Token::op(Kw::Caret, self.span(start)),

                0x41 => return self.scan_underscore(b, start),

                _ => {
                    let tok = self.scan_ident(b, start);
                    if matches!(tok.kind, TokenKind::Kw(Kw::Rem)) {
                        return self.consume_comment(start);
                    }
                    return tok;
                }
            }
        }
    }

    /// Consume the rest of the physical line as comment text.
    fn consume_comment(&mut self, start: usize) -> Token {
        while let Some(b) = self.peek() {
            let cls = CHAR_CLASSES[b as usize];
            if cls == 0x00 || cls == 0x01 {
                break;
            }
            self.pos += 1;
        }
        Token::op(Kw::Apos, self.span(start))
    }

    /// Scan a string literal.
    fn scan_string(&mut self, start: usize) -> Token {
        self.tok_buf.clear();
        loop {
            match self.advance() {
                None | Some(0x00) => return Token::error(self.span(start)),
                Some(b'"') => {
                    if self.peek() == Some(b'"') {
                        self.pos += 1;
                        self.tok_buf.push('"');
                    } else {
                        break;
                    }
                }
                Some(b) => self.tok_buf.push(b as char),
            }
        }
        let s = std::mem::take(&mut self.tok_buf).into_boxed_str();
        Token::lit(TokenKind::StrLit, Lit::Str(s), self.span(start))
    }

    /// Scan a numeric literal (integer, float, or type-suffixed).
    ///
    /// Exponent markers: `e`, `E`, `d`, `D`. `d`/`D` forces Double subtype.
    ///
    /// Type suffix mapping:
    /// * `%` → `IntLit`
    /// * `&` → `LongLit`
    /// * `!` → `SngLit`
    /// * `#` → `DblLit`
    /// * `@` → `CurLit`
    fn scan_number(&mut self, first: u8, start: usize) -> Token {
        self.tok_buf.clear();
        self.tok_buf.push(first as char);
        let mut has_decimal = first == b'.';
        let mut has_exp = false;

        loop {
            let Some(&b) = self.src.get(self.pos) else { break };
            match b {
                b'0'..=b'9' => {
                    self.tok_buf.push(b as char);
                    self.pos += 1;
                }
                b'.' if !has_decimal && !has_exp => {
                    has_decimal = true;
                    self.tok_buf.push('.');
                    self.pos += 1;
                }
                b'e' | b'E' | b'd' | b'D' if !has_exp => {
                    if self.try_consume_exponent() {
                        has_exp = true;
                    } else {
                        break;
                    }
                }
                _ => break,
            }
        }

        let is_float = has_decimal || has_exp;
        let span = self.span(start);
        self.finish_number_literal(is_float, span)
    }

    /// Try to consume an exponent marker plus optional sign at the current
    /// position. Returns true (and advances) only if a digit follows.
    fn try_consume_exponent(&mut self) -> bool {
        let exp_start = self.pos + 1;
        let sign_off = if self.src.get(exp_start)
            .map_or(false, |&s| s == b'+' || s == b'-') { 1 } else { 0 };
        if !self.src.get(exp_start + sign_off)
            .map_or(false, |&d| d.is_ascii_digit())
        {
            return false;
        }
        self.tok_buf.push('e');
        self.pos += 1;
        if sign_off > 0 {
            let s = self.src[self.pos];
            self.tok_buf.push(s as char);
            self.pos += 1;
        }
        true
    }

    /// Resolve the accumulated `tok_buf` into a numeric token, honoring an
    /// optional type suffix.
    fn finish_number_literal(&mut self, is_float: bool, span: Span) -> Token {
        match self.src.get(self.pos).copied() {
            Some(b'%') => {
                self.pos += 1;
                let v: i32 = self.tok_buf.parse().unwrap_or(0);
                Token::lit(TokenKind::IntLit, Lit::Int(v), span)
            }
            Some(b'&') if !is_float => {
                self.pos += 1;
                let v: i32 = self.tok_buf.parse().unwrap_or(0);
                Token::lit(TokenKind::LongLit, Lit::Long(v), span)
            }
            Some(b'!') => {
                self.pos += 1;
                let v: f32 = self.tok_buf.parse().unwrap_or(0.0);
                Token::lit(TokenKind::SngLit, Lit::Single(v), span)
            }
            Some(b'#') => {
                self.pos += 1;
                let v: f64 = self.tok_buf.parse().unwrap_or(0.0);
                Token::lit(TokenKind::DblLit, Lit::Double(v), span)
            }
            Some(b'@') => {
                self.pos += 1;
                let v: f64 = self.tok_buf.parse().unwrap_or(0.0);
                Token::lit(TokenKind::CurLit, Lit::Currency((v * 10_000.0) as i64), span)
            }
            _ if is_float => {
                let v: f64 = self.tok_buf.parse().unwrap_or(0.0);
                Token::lit(TokenKind::DblLit, Lit::Double(v), span)
            }
            _ => {
                let v: i64 = self.tok_buf.parse().unwrap_or(0);
                if v >= i16::MIN as i64 && v <= i16::MAX as i64 {
                    Token::lit(TokenKind::IntLit, Lit::Int(v as i32), span)
                } else {
                    Token::lit(TokenKind::LongLit, Lit::Long(v as i32), span)
                }
            }
        }
    }

    /// Handle `'#'`: conditional-compile directives or date literals.
    fn scan_hash(&mut self, start: usize) -> Token {
        let save = self.pos;
        self.tok_buf.clear();
        self.tok_buf.push('#');
        while let Some(&b) = self.src.get(self.pos) {
            if !is_ident_continue(b) { break; }
            self.tok_buf.push(b as char);
            self.pos += 1;
        }
        if let Some(tok) = self.hash_directive_token(start) {
            return tok;
        }
        self.pos = save;
        if let Some(tok) = self.scan_hash_date_literal(start) {
            return tok;
        }
        self.pos = save;
        Token::op(Kw::Hash, self.span(start))
    }

    /// If `tok_buf` (a `#`-prefixed word) names a conditional-compile keyword,
    /// produce that token.
    fn hash_directive_token(&mut self, start: usize) -> Option<Token> {
        if self.tok_buf.len() <= 1 {
            return None;
        }
        let key = self.tok_buf.to_ascii_lowercase();
        let &sym_id = self.ctx.by_key.get(&key)?;
        let token = self.ctx.symbols[sym_id].token;
        if token <= 0 {
            return None;
        }
        let kw = Kw::from_token_id(token as u16)?;
        Some(Token::kw(kw, sym_id, self.span(start)))
    }

    /// Scan a `#...#` date literal from the current position, if it closes.
    fn scan_hash_date_literal(&mut self, start: usize) -> Option<Token> {
        let date_content_start = self.pos;
        let found_close = loop {
            match self.src.get(self.pos) {
                None | Some(&0x00) | Some(&b'\n') | Some(&b'\r') => break false,
                Some(&b'#') => { self.pos += 1; break true; }
                Some(_) => { self.pos += 1; }
            }
        };
        if !found_close {
            return None;
        }
        let content_bytes = &self.src[date_content_start..self.pos - 1];
        let span = self.span(start);
        let s = std::str::from_utf8(content_bytes).ok()?;
        let serial = scan_vba_date_literal(s.trim());
        Some(Token::lit(TokenKind::DateLit, Lit::Date(serial), span))
    }

    /// Handle `'.'`: float literal or Dot.
    fn scan_dot(&mut self, start: usize) -> Token {
        if self.peek().map_or(false, |b| b.is_ascii_digit()) {
            self.scan_number(b'.', start)
        } else {
            Token::op(Kw::Dot, self.span(start))
        }
    }

    /// Handle `&`: Amp operator, or `&H`/`&O` numeric prefixes.
    fn scan_amp(&mut self, start: usize) -> Token {
        match self.peek() {
            Some(b'H') | Some(b'h') => {
                self.pos += 1;
                let val = self.scan_hex_digits();
                self.finish_radix_literal(val, start)
            }
            Some(b'O') | Some(b'o') => {
                self.pos += 1;
                let val = self.scan_octal_digits();
                self.finish_radix_literal(val, start)
            }
            _ => Token::op(Kw::Amp, self.span(start)),
        }
    }

    /// Accumulate hexadecimal digits from the current position.
    fn scan_hex_digits(&mut self) -> i64 {
        let mut val: i64 = 0;
        while let Some(&b) = self.src.get(self.pos) {
            let d = match b {
                b'0'..=b'9' => (b - b'0') as i64,
                b'a'..=b'f' => (b - b'a' + 10) as i64,
                b'A'..=b'F' => (b - b'A' + 10) as i64,
                _ => break,
            };
            val = val * 16 + d;
            self.pos += 1;
        }
        val
    }

    /// Accumulate octal digits from the current position.
    fn scan_octal_digits(&mut self) -> i64 {
        let mut val: i64 = 0;
        while let Some(&b) = self.src.get(self.pos) {
            match b {
                b'0'..=b'7' => {
                    val = val * 8 + (b - b'0') as i64;
                    self.pos += 1;
                }
                _ => break,
            }
        }
        val
    }

    /// Finish a `&H`/`&O` literal: honor a trailing `&` (Long) or size by range.
    fn finish_radix_literal(&mut self, val: i64, start: usize) -> Token {
        let span = self.span(start);
        if self.src.get(self.pos) == Some(&b'&') {
            self.pos += 1;
            return Token::lit(TokenKind::LongLit, Lit::Long(val as i32), span);
        }
        if val >= i16::MIN as i64 && val <= i16::MAX as i64 {
            Token::lit(TokenKind::IntLit, Lit::Int(val as i32), span)
        } else {
            Token::lit(TokenKind::LongLit, Lit::Long(val as i32), span)
        }
    }

    /// Handle `':'`: Colon or `:=`.
    fn scan_colon(&mut self, start: usize) -> Token {
        if self.peek() == Some(b'=') {
            self.pos += 1;
            Token::op(Kw::ColonEq, self.span(start))
        } else {
            Token::op(Kw::Colon, self.span(start))
        }
    }

    /// Handle `'<'`: Lt, Le, or Ne.
    fn scan_lt(&mut self, start: usize) -> Token {
        match self.peek() {
            Some(b'=') => { self.pos += 1; Token::op(Kw::Le,  self.span(start)) }
            Some(b'>') => { self.pos += 1; Token::op(Kw::Ne,  self.span(start)) }
            _           =>                 Token::op(Kw::Lt,  self.span(start)),
        }
    }

    /// Handle `'='`: Eq, EqLt, or EqGt.
    fn scan_eq(&mut self, start: usize) -> Token {
        match self.peek() {
            Some(b'<') => { self.pos += 1; Token::op(Kw::EqLt, self.span(start)) }
            Some(b'>') => { self.pos += 1; Token::op(Kw::EqGt, self.span(start)) }
            _           =>                 Token::op(Kw::Eq,   self.span(start)),
        }
    }

    /// Handle `'>'`: Gt, Ge, or GtLt.
    fn scan_gt(&mut self, start: usize) -> Token {
        match self.peek() {
            Some(b'=') => { self.pos += 1; Token::op(Kw::Ge,   self.span(start)) }
            Some(b'<') => { self.pos += 1; Token::op(Kw::GtLt, self.span(start)) }
            _           =>                 Token::op(Kw::Gt,   self.span(start)),
        }
    }

    /// Handle `'_'`: line-continuation marker or identifier start.
    fn scan_underscore(&mut self, first: u8, start: usize) -> Token {
        if self.peek().map_or(true, |b| !is_ident_continue(b)) {
            Token::op(Kw::LineCont, self.span(start))
        } else {
            self.scan_ident(first, start)
        }
    }

    /// Collect an identifier and resolve it to a keyword or user symbol.
    fn scan_ident(&mut self, first: u8, start: usize) -> Token {
        self.tok_buf.clear();
        self.tok_buf.push(first as char);
        while let Some(&b) = self.src.get(self.pos) {
            if !is_ident_continue(b) { break; }
            self.tok_buf.push(b as char);
            self.pos += 1;
        }

        let type_suffix = self.resolve_type_suffix();
        let span = self.span(start);
        let name: String = self.tok_buf.clone();
        match self.ctx.intern_ident(&name) {
            Some(sym_id) => {
                let mut tok = self.ident_token(sym_id, span);
                tok.type_suffix = type_suffix;
                tok
            }
            None => Token::error(span),
        }
    }

    /// Resolve a trailing type-suffix byte for the identifier in `tok_buf`,
    /// consuming it from the source. A suffix that folds the name into a
    /// keyword is appended to `tok_buf` instead and reported as `None`.
    fn resolve_type_suffix(&mut self) -> TypeSuffix {
        let Some(&suffix) = self.src.get(self.pos) else {
            return TypeSuffix::None;
        };
        let folds_to_keyword = matches!(suffix, b'$' | b'%' | b'!' | b'#' | b'@') && {
            let mut with_suffix = self.tok_buf.clone();
            with_suffix.push(suffix as char);
            self.ctx.lookup(&with_suffix).map_or(false, |s| s.token > 0)
        };
        if folds_to_keyword {
            self.tok_buf.push(suffix as char);
            self.pos += 1;
            return TypeSuffix::None;
        }
        let is_bang_member = suffix == b'!'
            && self.src.get(self.pos + 1).map_or(false, |&n| is_ident_start(n));
        let ts = TypeSuffix::from_byte(suffix);
        if ts != TypeSuffix::None && !is_bang_member {
            self.pos += 1;
            ts
        } else {
            TypeSuffix::None
        }
    }

    /// Build the token for an interned identifier symbol, mapping to a keyword
    /// token when the symbol carries one.
    fn ident_token(&self, sym_id: SymId, span: Span) -> Token {
        let token = self.ctx.symbol(sym_id).token;
        if token > 0 {
            Kw::from_token_id(token as u16)
                .map(|kw| Token::kw(kw, sym_id, span))
                .unwrap_or_else(|| Token::ident(sym_id, span))
        } else {
            Token::ident(sym_id, span)
        }
    }

    /// Scan a bracketed (verbatim) identifier `[name]`.
    fn scan_bracketed_ident(&mut self, start: usize) -> Token {
        self.tok_buf.clear();
        loop {
            match self.advance() {
                None | Some(0x00) => return Token::error(self.span(start)),
                Some(b']') => break,
                Some(b) => self.tok_buf.push(b as char),
            }
        }
        let span = self.span(start);
        let name: String = self.tok_buf.clone();
        match self.ctx.intern_ident(&name) {
            Some(sym_id) => {
                self.ctx.set_verbatim_flag(sym_id);
                Token::ident(sym_id, span)
            }
            None => Token::error(span),
        }
    }

    // ── low-level helpers ────────────────────────────────────────────────────

    /// Read the next byte and advance the position.
    #[inline]
    fn advance(&mut self) -> Option<u8> {
        let b = *self.src.get(self.pos)?;
        self.pos += 1;
        Some(b)
    }

    /// Peek at the next byte without consuming it.
    #[inline]
    fn peek(&self) -> Option<u8> {
        self.src.get(self.pos).copied()
    }

    /// Build a [`Span`] from `start` to the current position.
    #[inline]
    fn span(&self, start: usize) -> Span {
        Span { start: start as u32, len: (self.pos - start) as u32 }
    }
}
