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
use super::line_reader::{self, LineLimits, ReadProblem};
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
    /// A name could not be added because the symbol table is full.
    pub symbol_table_full: bool,
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
            symbol_table_full: false,
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

    /// The code page the module's language implies when it is a double-byte one (Japanese 932,
    /// Chinese 936 or 950, Korean 949).
    pub fn dbcs_code_page(&self) -> Option<u16> {
        match self.lcid & 0x3ff {
            0x11 => Some(932),
            0x12 => Some(949),
            0x04 => Some(if matches!(self.lcid, 0x0804 | 0x1004) { 936 } else { 950 }),
            _ => None,
        }
    }

    /// Whether `b` leads a two-byte character in the module's code page.
    pub fn is_dbcs_lead_byte(&self, b: u8) -> bool {
        match self.dbcs_code_page() {
            Some(932) => matches!(b, 0x81..=0x9F | 0xE0..=0xFC),
            Some(_) => matches!(b, 0x81..=0xFE),
            None => false,
        }
    }

    /// Look up an interned symbol by name (case-insensitive).
    pub fn lookup(&self, name: &str) -> Option<&Symbol> {
        self.by_key
            .get(&fold_name(name))
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
        let key = fold_name(name);

        // --- found path ---
        if let Some(&existing) = self.by_key.get(&key) {
            return Some(existing);
        }

        // --- not found: allocate ---
        if self.mode_flags != 0 && self.counter + HOST_SYMBOLS > 0x7ffe {
            self.symbol_table_full = true;
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
        let key = fold_name(name);
        if let Some(&existing) = self.by_key.get(&key) {
            return Some(existing);
        }
        if self.mode_flags != 0 && self.counter + HOST_SYMBOLS > 0x7ffe {
            self.symbol_table_full = true;
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

/// A name as symbols are keyed: letters fold to lower case, the accented Latin-1 letters
/// along with the ASCII ones.
pub fn fold_name(name: &str) -> String {
    name.chars()
        .map(|c| match c {
            '\u{c0}'..='\u{de}' if c != '\u{d7}' => char::from_u32(c as u32 + 0x20).unwrap_or(c),
            _ => c.to_ascii_lowercase(),
        })
        .collect()
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
/// The character a source byte stands for in the Windows ANSI code page (1252): the bytes
/// 0x80-0x9F that name no character stand for themselves.
fn ansi_char(b: u8) -> char {
    const HIGH: [u32; 32] = [
        0x20AC, 0x81, 0x201A, 0x0192, 0x201E, 0x2026, 0x2020, 0x2021, 0x02C6, 0x2030, 0x0160, 0x2039, 0x0152, 0x8D, 0x017D,
        0x8F, 0x90, 0x2018, 0x2019, 0x201C, 0x201D, 0x2022, 0x2013, 0x2014, 0x02DC, 0x2122, 0x0161, 0x203A, 0x0153, 0x9D,
        0x017E, 0x0178,
    ];
    match b {
        0x80..=0x9F => char::from_u32(HIGH[(b - 0x80) as usize]).unwrap_or(b as char),
        _ => b as char,
    }
}

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

/// Expand a two-digit year: 00–49 → 2000–2049, 50–99 → 1950–1999.
fn expand_two_digit_year(yy: i32) -> i32 {
    if yy < 50 { 2000 + yy } else { 1900 + yy }
}

const MONTH_NAMES: [&str; 12] = [
    "january", "february", "march", "april", "may", "june",
    "july", "august", "september", "october", "november", "december",
];
const WEEKDAY_NAMES: [&str; 7] = [
    "sunday", "monday", "tuesday", "wednesday", "thursday", "friday", "saturday",
];

/// One lexical piece of the text between the `#` marks of a date literal.
#[derive(Clone, Copy, PartialEq, Debug)]
enum DateTok<'a> {
    Num { value: i32, digits: usize },
    Word(&'a str),
    Colon,
    Slash,
    Dash,
}

/// Split the body of a date literal; `None` when it holds any other character.
fn lex_date_text(s: &str) -> Option<Vec<DateTok<'_>>> {
    let bytes = s.as_bytes();
    let mut out = Vec::new();
    let mut i = 0;
    // Two separators in a row (`//`, `--`, `,,`) are not a date.
    let mut last_sep: Option<u8> = None;
    while i < bytes.len() {
        let b = bytes[i];
        if matches!(b, b'/' | b'-' | b',') {
            if last_sep == Some(b) {
                return None;
            }
            last_sep = Some(b);
        } else if !matches!(b, b' ' | b'\t') {
            last_sep = None;
        }
        match b {
            b' ' | b'\t' | b',' => i += 1,
            // A period between numbers separates the parts of a time; after a word it is dropped.
            b'.' => {
                match out.last() {
                    Some(DateTok::Num { .. }) if bytes.get(i + 1).is_some_and(|d| d.is_ascii_digit()) => {
                        out.push(DateTok::Colon)
                    }
                    Some(DateTok::Word(_)) => {}
                    _ => return None,
                }
                i += 1;
            }
            b':' => { out.push(DateTok::Colon); i += 1; }
            b'/' => { out.push(DateTok::Slash); i += 1; }
            b'-' => { out.push(DateTok::Dash); i += 1; }
            b'0'..=b'9' => {
                let start = i;
                while i < bytes.len() && bytes[i].is_ascii_digit() { i += 1; }
                let value = s[start..i].parse::<i32>().ok()?;
                out.push(DateTok::Num { value, digits: i - start });
            }
            b'A'..=b'Z' | b'a'..=b'z' => {
                let start = i;
                while i < bytes.len() && bytes[i].is_ascii_alphabetic() { i += 1; }
                out.push(DateTok::Word(&s[start..i]));
            }
            _ => return None,
        }
    }
    Some(out)
}

fn month_from_word(w: &str) -> Option<i32> {
    // Any prefix of a month's name of at least three letters.
    let w = w.to_ascii_lowercase();
    if w.len() < 3 {
        return None;
    }
    MONTH_NAMES.iter().position(|m| m.starts_with(&w)).map(|i| i as i32 + 1)
}

fn is_weekday_word(w: &str) -> bool {
    let w = w.to_ascii_lowercase();
    w.len() >= 3
        && WEEKDAY_NAMES.iter().any(|d| *d == w || (w.len() == 3 && d[..3] == w))
}

fn days_in_month(y: i32, m: i32) -> i32 {
    match m {
        1 | 3 | 5 | 7 | 8 | 10 | 12 => 31,
        4 | 6 | 9 | 11 => 30,
        _ => if (y % 4 == 0 && y % 100 != 0) || y % 400 == 0 { 29 } else { 28 },
    }
}

/// A year as written: two digits expand, the result must lie in 100..=9999.
fn full_year(value: i32, digits: usize) -> Option<i32> {
    // A year below 100 expands, however many zeros pad it (`050` is 1950).
    let y = if value < 100 && digits <= 3 { expand_two_digit_year(value) } else { value };
    (100..=9999).contains(&y).then_some(y)
}

fn valid_serial(y: i32, m: i32, d: i32) -> Option<f64> {
    ((1..=12).contains(&m) && d >= 1 && d <= days_in_month(y, m))
        .then(|| vba_date_serial(y, m, d))
}

/// The year of the system clock, which a date written without a year takes.
pub fn current_year() -> i32 {
    let secs = std::time::SystemTime::now()
        .duration_since(std::time::UNIX_EPOCH)
        .map_or(0, |d| d.as_secs() as i64);
    let days = (secs / 86400) as i32 + 719468; // days since 0000-03-01
    let era = days / 146097;
    let doe = days - era * 146097;
    let yoe = (doe - doe / 1460 + doe / 36524 - doe / 146096) / 365;
    let doy = doe - (365 * yoe + yoe / 4 - yoe / 100);
    let m = (5 * doy + 2) / 153;
    let year = yoe + era * 400;
    if m >= 10 { year + 1 } else { year }
}

/// Date part made of numbers only: `m/d/y`, or `y/m/d` for a leading year; with two numbers,
/// `m/d` of the current year or `y/m` for a leading year.
fn numeric_date(nums: &[(i32, usize)]) -> Option<f64> {
    if let [(a, ad), (b, _)] = *nums {
        return if ad >= 3 || a > 31 {
            valid_serial(full_year(a, ad)?, b, 1)
        } else {
            valid_serial(current_year(), a, b)
        };
    }
    let [(a, ad), (b, _), (c, cd)] = *nums else { return None };
    if ad >= 3 || a > 31 {
        valid_serial(full_year(a, ad)?, b, c)
    } else if a > 12 && b <= 12 {
        // A first field that cannot be a month is the day: `13/1/2000` is the 13th of January.
        valid_serial(full_year(c, cd)?, b, a)
    } else {
        valid_serial(full_year(c, cd)?, a, b)
    }
}

/// Date part with a month name and one or two numbers. A date needs a year,
/// so a lone day is not a literal; a lone year means the first of the month.
fn named_date(month: i32, nums: &[(i32, usize)]) -> Option<f64> {
    let year_like = |v: i32, d: usize| d >= 3 || v > 31;
    match *nums {
        [(v, d)] if year_like(v, d) => valid_serial(full_year(v, d)?, month, 1),
        // A day alone is a day of the current year.
        [(v, _)] => valid_serial(current_year(), month, v),
        [(a, ad), (b, bd)] => {
            if year_like(a, ad) {
                valid_serial(full_year(a, ad)?, month, b)
            } else {
                valid_serial(full_year(b, bd)?, month, a)
            }
        }
        _ => None,
    }
}

/// The time of day as a fraction of a day: `H:M[:S]` with optional AM/PM, or
/// an hour followed by AM/PM.
fn time_fraction(hour: i32, minute: i32, second: i32, meridiem: Option<bool>) -> Option<f64> {
    if !(0..=23).contains(&hour) {
        return None;
    }
    // A meridiem shifts an hour of 12 or less; a later hour stands as written.
    let hour = match meridiem {
        Some(false) if hour == 12 => 0,
        Some(true) if hour < 12 => hour + 12,
        _ => hour,
    };
    if !(0..=59).contains(&minute) || !(0..=59).contains(&second) {
        return None;
    }
    Some((hour * 3600 + minute * 60 + second) as f64 / 86400.0)
}

/// Parse a VBA date/time literal's inner text into a serial, or `None` if the
/// text is not a valid date/time literal.
///
/// The `None` case is essential: a `#…#` run is only a literal when its body is
/// a real date or time. `Close #1, #2` puts two `#` on one line, and without
/// this check the scanner would greedily fuse `#1, #` into a bogus date literal.
fn parse_vba_date_literal(s: &str) -> Option<f64> {
    let toks = lex_date_text(s)?;
    if toks.is_empty() {
        return None;
    }
    let mut meridiem = None;
    let mut meridiem_at = 0;
    let mut rest: Vec<DateTok> = Vec::new();
    for t in toks {
        match t {
            DateTok::Word(w) if matches!(w.to_ascii_lowercase().as_str(), "a" | "am" | "p" | "pm") => {
                // One meridiem at most.
                if meridiem.is_some() {
                    return None;
                }
                meridiem = Some(w.to_ascii_lowercase().starts_with('p'));
                meridiem_at = rest.len();
            }
            // A weekday name makes the text invalid.
            DateTok::Word(w) if is_weekday_word(w) => return None,
            other => rest.push(other),
        }
    }
    // A meridiem opens the text or follows the number that ends the time.
    let leading_meridiem = meridiem.is_some() && meridiem_at == 0;
    if meridiem.is_some() && !leading_meridiem && !matches!(rest.get(meridiem_at - 1), Some(DateTok::Num { .. })) {
        return None;
    }
    // The time is `Num (: Num)+`, or the last number before AM/PM.
    let mut time: Option<f64> = None;
    if let Some(c) = rest.iter().position(|t| *t == DateTok::Colon) {
        let num_at = |i: usize| match rest.get(i) {
            Some(DateTok::Num { value, .. }) => Some(*value),
            _ => None,
        };
        let start = c.checked_sub(1)?;
        let hour = num_at(start)?;
        let minute = num_at(c + 1)?;
        let (second, end) = if rest.get(c + 2) == Some(&DateTok::Colon) {
            (num_at(c + 3)?, c + 3)
        } else {
            (0, c + 1)
        };
        if meridiem.is_some() && !leading_meridiem && end + 1 != meridiem_at {
            return None;
        }
        time = Some(time_fraction(hour, minute, second, meridiem)?);
        rest.drain(start..=end);
    } else if let Some(pm) = meridiem {
        if leading_meridiem {
            // `AM 1/1/2000`: nothing for it to qualify.
        } else {
            // An hour alone before the meridiem.
            let DateTok::Num { value, .. } = rest[meridiem_at - 1] else { return None };
            time = Some(time_fraction(value, 0, 0, Some(pm))?);
            rest.remove(meridiem_at - 1);
        }
    }
    let mut nums: Vec<(i32, usize)> = Vec::new();
    let mut month: Option<i32> = None;
    let mut has_sep = false;
    for t in &rest {
        match *t {
            DateTok::Num { value, digits } => nums.push((value, digits)),
            DateTok::Word(w) => {
                if month.is_some() { return None; }
                month = Some(month_from_word(w)?);
            }
            DateTok::Slash | DateTok::Dash => has_sep = true,
            DateTok::Colon => return None,
        }
    }
    let date = match (month, nums.is_empty()) {
        (None, true) => {
            if has_sep { return None; }
            return time;
        }
        (None, false) => numeric_date(&nums)?,
        (Some(m), _) => named_date(m, &nums)?,
    };
    // Before the epoch the fraction of a date is the time of day counted as a magnitude.
    let time = time.unwrap_or(0.0);
    Some(if date < 0.0 { date - time } else { date + time })
}

// ── Scanner ───────────────────────────────────────────────────────────────────

/// Per-statement tokenizer.
pub struct Scanner<'a> {
    ctx: &'a mut ScannerContext,
    /// The text as read: the source with a NUL where a piece starts inside its physical line.
    src: std::borrow::Cow<'a, [u8]>,
    /// Where in `src` those NULs are.
    separators: Vec<u32>,
    /// The first limit the text breaks.
    read_problem: Option<ReadProblem>,
    /// What the limits on the text are measured on.
    limits: LineLimits,
    /// Current byte offset into `src`.
    pos: usize,
    /// Scratch buffer for identifier and literal accumulation.
    tok_buf: String,
    /// The token returned last was a `.` or `!` that introduces a member name.
    after_member_operator: bool,
}

impl<'a> Scanner<'a> {
    /// Return the canonical name for a symbol by id.
    pub fn sym_name(&self, id: u32) -> &str {
        self.ctx.sym_name(id as usize)
    }

    /// Whether a name was refused because the symbol table is full.
    pub fn symbol_table_full(&self) -> bool {
        self.ctx.symbol_table_full
    }

    /// True if the symbol is a reserved word (which can never name a
    /// variable or a type); false for identifiers and contextual keywords.
    pub fn is_reserved(&self, id: u32) -> bool {
        self.ctx.symbols.get(id as usize).map_or(false, |s| s.flags & 4 != 0)
    }

    /// The keyword a symbol spells, or `None` for a name that is no keyword.
    pub fn symbol_keyword(&self, id: u32) -> Option<Kw> {
        let token = self.ctx.symbols.get(id as usize)?.token;
        if token > 0 { Kw::from_token_id(token as u16) } else { None }
    }

    /// True if the symbol is a keyword spelled with a trailing `$` (`Mid$`, `Date$`): a
    /// name of that kind is a keyword word, never a name with a type-declaration character.
    pub fn is_dollar_word(&self, id: u32) -> bool {
        self.ctx.symbols.get(id as usize).map_or(false, |s| s.flags & 1 != 0)
    }

    /// The first limit the text breaks (too many pieces in a logical line, too many
    /// lines in the module), if any.
    pub fn read_problem(&self) -> Option<ReadProblem> {
        self.read_problem
    }

    /// Declares the text before `header_end` to be header lines that are not code, such as a
    /// designer section that was blanked: they do not count toward the limit on the lines of a
    /// module.
    pub fn set_header_end(&mut self, header_end: u32) {
        self.read_problem = self.limits.problem(header_end);
    }

    /// Whether the NUL at `pos` is one that stands between two pieces of a line.
    fn is_separator(&self, pos: usize) -> bool {
        self.separators.binary_search(&(pos as u32)).is_ok()
    }

    /// Create a scanner for `src`, borrowing `ctx` for symbol interning.
    pub fn new(ctx: &'a mut ScannerContext, src: &'a [u8]) -> Self {
        let read = line_reader::read(src);
        Self {
            ctx,
            src: read.text,
            separators: read.separators,
            read_problem: read.problem,
            limits: read.limits,
            pos: 0,
            tok_buf: String::with_capacity(64),
            after_member_operator: false,
        }
    }

    /// Return the next token from the source, advancing the position.
    pub fn next_token(&mut self) -> Token {
        let mut tok = self.scan_token();
        if !self.separators.is_empty() {
            let start = line_reader::source_offset(&self.separators, tok.span.start as usize);
            let end = line_reader::source_offset(&self.separators, (tok.span.start + tok.span.len) as usize);
            tok.span = Span { start, len: end - start };
        }
        self.after_member_operator = matches!(tok.kind, TokenKind::Kw(Kw::Dot) | TokenKind::Kw(Kw::Bang));
        tok
    }

    fn scan_token(&mut self) -> Token {
        loop {
            let start = self.pos;
            let b = match self.advance() {
                None => return Token::eof(self.pos as u32),
                Some(b) => b,
            };
            // A tab in a source file is whitespace between tokens.
            if b == b'\t' {
                continue;
            }
            match b {
                // A NUL ends the text of its piece: what follows it in the piece belongs to no
                // token. The NUL that stands between two pieces of a line ends nothing more.
                0x00 => {
                    if !self.is_separator(start) {
                        while let Some(next) = self.peek() {
                            if matches!(next, b'\r' | b'\n') || (next == 0 && self.is_separator(self.pos)) {
                                break;
                            }
                            self.pos += 1;
                        }
                    }
                    continue;
                }
                // 0x19 is a blank.
                0x19 => continue,
                // Every other control character is invalid.
                0x01..=0x08 | 0x0b | 0x0c | 0x0e..=0x18 | 0x1a..=0x1f | 0x7f => {
                    return Token::error(self.span(start));
                }
                _ => {}
            }
            // A full-width form of an ASCII character is read as that character.
            let b = match self.full_width_at(start) {
                Some((ascii, _)) => {
                    self.pos += 1;
                    ascii
                }
                None => b,
            };
            // Another double-byte character starts a name.
            if b == self.src[start] && self.double_byte_at(start) {
                return self.scan_ident(b, start);
            }
            let cls = CHAR_CLASSES[b as usize];
            match cls {
                0x00 => {
                    // CR LF is one line break, so it is one end-of-line token.
                    if b == b'\r' && self.peek() == Some(b'\n') {
                        self.pos += 1;
                    }
                    return Token::eol(start as u32);
                }

                0x01..=0x04 => { /* skip */ }

                0x05 => { /* skip */ }

                0x06 => return Token::op(Kw::Bang, self.span(start)),

                0x07 => return self.scan_string(start),

                0x08 => return self.scan_hash(start),

                0x09 | 0x0a | 0x25 => return Token::error(self.span(start)),

                0x0b => return self.scan_amp(start),

                0x0c => return self.consume_comment(start, Kw::Apos),

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

                _ if matches!(b, b'~' | b'`' | b'|' | b'{' | b'}') => {
                    return Token::error(self.span(start));
                }
                _ => {
                    let tok = self.scan_ident(b, start);
                    // After `.` or `!` the word names a member, even `Rem`.
                    if matches!(tok.kind, TokenKind::Kw(Kw::Rem)) && !self.after_member_operator {
                        return self.consume_comment(start, Kw::Rem);
                    }
                    return tok;
                }
            }
        }
    }

    /// Consume the rest of the line as comment text. Tabs are part of the
    /// text. A comment whose last non-blank character is a line-continuation
    /// `_` (one not ending a word) continues on the next physical line.
    fn consume_comment(&mut self, start: usize, kind: Kw) -> Token {
        loop {
            let line_start = self.pos;
            while let Some(b) = self.peek() {
                // A comment runs to the end of its line whatever it holds.
                if b == b'\r' || b == b'\n' {
                    break;
                }
                self.pos += 1;
            }
            let prev = line_start.checked_sub(1).map(|i| self.src[i]);
            if !Self::ends_with_line_cont(&self.src[line_start..self.pos], prev) {
                break;
            }
            match (self.peek(), self.src.get(self.pos + 1).copied()) {
                (Some(b'\r'), Some(b'\n')) => self.pos += 2,
                (Some(b'\r') | Some(b'\n'), _) => self.pos += 1,
                _ => break,
            }
        }
        Token::op(kind, self.span(start))
    }

    /// True if `text` ends (ignoring trailing spaces and tabs) with a `_` that
    /// is not the last character of a word.
    fn ends_with_line_cont(text: &[u8], prev: Option<u8>) -> bool {
        let trimmed = match text.iter().rposition(|&b| b != b' ' && b != b'\t') {
            Some(i) => &text[..=i],
            None => return false,
        };
        match trimmed.split_last() {
            // The mark must follow a blank.
            Some((b'_', before)) => match before.last().copied().or(prev) {
                Some(b) => matches!(b, b' ' | b'\t' | b'\r' | b'\n'),
                None => true,
            },
            _ => false,
        }
    }

    /// The column (from 0) the byte at `at` takes on its physical line once tabs are
    /// expanded to stops every 8 columns.
    fn expanded_column(&self, at: usize) -> usize {
        let mut line_start = self.src[..at].iter().rposition(|b| matches!(b, b'\r' | b'\n')).map_or(0, |i| i + 1);
        // Columns count from the start of the piece.
        let before = self.separators.partition_point(|&s| (s as usize) < at);
        if before > 0 {
            line_start = line_start.max(self.separators[before - 1] as usize + 1);
        }
        let mut column = 0;
        for &b in &self.src[line_start..at] {
            column += if b == b'\t' { 8 - column % 8 } else { 1 };
        }
        column
    }

    /// Scan a string literal.
    fn scan_string(&mut self, start: usize) -> Token {
        self.tok_buf.clear();
        loop {
            // A full-width double quote closes (or doubles) a string like a plain one.
            let step = match self.full_width_at(self.pos) {
                Some((b'"', width)) => {
                    self.pos += width;
                    Some(b'"')
                }
                _ => self.advance(),
            };
            if step == Some(b'"') {
                let wide_next = matches!(self.full_width_at(self.pos), Some((b'"', _)));
                if self.peek() == Some(b'"') || wide_next {
                    self.pos += if wide_next { 2 } else { 1 };
                    self.tok_buf.push('"');
                } else {
                    break;
                }
                continue;
            }
            match step {
                None => break,
                // A NUL ends the text of the piece, and so the string: it is not closed, and
                // the NUL is left for the next token to skip the rest of the piece.
                Some(0x00) => {
                    self.pos -= 1;
                    break;
                }
                // A tab inside a string is expanded to the next tab stop (every 8 columns,
                // counted on the expanded line).
                Some(b'\t') => {
                    let column = self.expanded_column(self.pos - 1);
                    for _ in 0..(8 - column % 8) {
                        self.tok_buf.push(' ');
                    }
                }
                Some(b'\r') | Some(b'\n') => {
                    // A string stops at the end of its line. When the line ends in a
                    // continuation mark, the string stops before the blanks that lead to
                    // the mark and the mark joins the next line.
                    self.pos -= 1;
                    if let Some(mark) = self.continuation_mark_at_end(start) {
                        for _ in mark..self.pos {
                            self.tok_buf.pop();
                        }
                        while self.tok_buf.ends_with([' ', '\t']) {
                            self.tok_buf.pop();
                        }
                        self.pos = mark;
                    }
                    break;
                }
                Some(b'"') => {
                    if self.peek() == Some(b'"') {
                        self.pos += 1;
                        self.tok_buf.push('"');
                    } else {
                        break;
                    }
                }
                Some(b) => self.tok_buf.push(ansi_char(b)),
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
            let Some(&raw) = self.src.get(self.pos) else { break };
            // A full-width digit or point is the ASCII one.
            let (b, width) = self.full_width_at(self.pos).unwrap_or((raw, 1));
            match b {
                b'0'..=b'9' => {
                    self.tok_buf.push(b as char);
                    self.pos += width;
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
                        // An exponent mark must be followed by digits (after an optional sign).
                        while matches!(self.src.get(self.pos), Some(b) if is_ident_continue(*b)) {
                            self.pos += 1;
                        }
                        return self.number_error(self.span(start));
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
                let span = Span { len: span.len + 1, ..span };
                match self.tok_buf.parse::<f64>() {
                    Ok(v) if !is_float && v >= i16::MIN as f64 && v <= i16::MAX as f64 => {
                        Token::lit(TokenKind::IntLit, Lit::Int(v as i32), span)
                    }
                    _ => self.number_error(span),
                }
            }
            Some(b'&') if !is_float => {
                match self.tok_buf.parse::<i64>() {
                    Ok(v) if v >= i32::MIN as i64 && v <= i32::MAX as i64 => {
                        self.pos += 1;
                        let span = Span { len: span.len + 1, ..span };
                        Token::lit(TokenKind::LongLit, Lit::Long(v as i32), span)
                    }
                    // A whole number too large for a Long is a Double, and the `&` after it
                    // is the operator.
                    _ => self.finish_double(span),
                }
            }
            Some(b'!') => {
                self.pos += 1;
                let span = Span { len: span.len + 1, ..span };
                // The text becomes a Double first, and the Double a Single.
                match self.tok_buf.parse::<f64>().map(|d| d as f32) {
                    Ok(v) if v.is_finite() => Token::lit(TokenKind::SngLit, Lit::Single(v), span),
                    _ => self.number_error(span),
                }
            }
            Some(b'#') => {
                self.pos += 1;
                let span = Span { len: span.len + 1, ..span };
                self.finish_double(span)
            }
            Some(b'@') => {
                self.pos += 1;
                let span = Span { len: span.len + 1, ..span };
                match currency_units(&self.tok_buf) {
                    Some(v) => Token::lit(TokenKind::CurLit, Lit::Currency(v), span),
                    None => self.number_error(span),
                }
            }
            _ if is_float => self.finish_double(span),
            _ => match self.tok_buf.parse::<i64>() {
                Ok(v) if v >= i16::MIN as i64 && v <= i16::MAX as i64 => {
                    Token::lit(TokenKind::IntLit, Lit::Int(v as i32), span)
                }
                Ok(v) if v >= i32::MIN as i64 && v <= i32::MAX as i64 => {
                    Token::lit(TokenKind::LongLit, Lit::Long(v as i32), span)
                }
                // A whole number too large for a Long is a Double.
                _ => self.finish_double(span),
            },
        }
    }

    /// A Double literal from `tok_buf`; a magnitude beyond the Double range
    /// is an error token.
    fn finish_double(&mut self, span: Span) -> Token {
        match self.tok_buf.parse::<f64>() {
            Ok(v) if v.is_finite() => Token::lit(TokenKind::DblLit, Lit::Double(v), span),
            _ => self.number_error(span),
        }
    }

    /// A numeric literal that does not fit its type is a token error.
    fn number_error(&mut self, span: Span) -> Token {
        Token::error(span)
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
        // Blanks may separate the `#` from the directive word (`# If`).
        if self.tok_buf.len() == 1 {
            let mut p = self.pos;
            while matches!(self.src.get(p), Some(b' ') | Some(b'\t')) {
                p += 1;
            }
            let word_start = p;
            while let Some(&b) = self.src.get(p) {
                if !is_ident_continue(b) { break; }
                p += 1;
            }
            if p > word_start && word_start > self.pos {
                let word: String = self.src[word_start..p].iter().map(|&b| b as char).collect();
                let key = format!("#{}", word.to_ascii_lowercase());
                if key == "#endif" || self.ctx.by_key.contains_key(&key) {
                    self.tok_buf.push_str(&word);
                    self.pos = p;
                }
            }
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
        let mut key = self.tok_buf.to_ascii_lowercase();
        // `#EndIf` is the one-word spelling of `#End If`; the token carries a
        // marker payload so the parser needs no `If` after it.
        let one_word = key == "#endif";
        if one_word {
            key = "#end".to_string();
        }
        let &sym_id = self.ctx.by_key.get(&key)?;
        let token = self.ctx.symbols[sym_id].token;
        if token <= 0 {
            return None;
        }
        let kw = Kw::from_token_id(token as u16)?;
        // Like any word, a directive word may carry one type-declaration character.
        if matches!(self.src.get(self.pos), Some(b'%' | b'&' | b'!' | b'#' | b'@' | b'$')) {
            self.pos += 1;
        }
        let mut tok = Token::kw(kw, sym_id, self.span(start));
        if one_word {
            tok.lit = Some(Lit::Int(1));
        }
        Some(tok)
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
        // Only commit to a date literal when the body is actually a date/time;
        // otherwise return `None` so the caller falls back to the `#` operator
        // (e.g. the two channels in `Close #1, #2`).
        let serial = parse_vba_date_literal(s.trim())?;
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
                let val = self.scan_radix_digits(16);
                self.radix_literal_or_amp(val, start)
            }
            Some(b'O') | Some(b'o') => {
                self.pos += 1;
                let val = self.scan_radix_digits(8);
                self.radix_literal_or_amp(val, start)
            }
            // `&` and an octal digit start an octal literal without the `O`, wherever they stand.
            Some(b'0'..=b'7') => {
                let val = self.scan_radix_digits(8);
                self.radix_literal_or_amp(val, start)
            }
            _ => Token::op(Kw::Amp, self.span(start)),
        }
    }

    /// The literal read after `&`, or, when the text after the `&` is not a whole literal
    /// (no digits, too many digits, a stray `H` or `O` among hexadecimal digits, an 8 or 9
    /// in an octal one), the `&` operator alone, with the text after it left to be read
    /// as the next token.
    fn radix_literal_or_amp(&mut self, val: Option<u32>, start: usize) -> Token {
        if val.is_none() {
            self.pos = start + 1;
            return Token::op(Kw::Amp, self.span(start));
        }
        self.finish_radix_literal(val, start)
    }

    /// Accumulate digits of the given radix from the current position.
    /// `None` when there is no digit or the value exceeds 32 bits.
    fn scan_radix_digits(&mut self, radix: u32) -> Option<u32> {
        let mut val: u64 = 0;
        let mut count = 0usize;
        let mut overflow = false;
        while let Some(&b) = self.src.get(self.pos) {
            let Some(d) = (b as char).to_digit(radix) else {
                // Hexadecimal digits swallow an `H` or `O`, which spoils the literal; an 8 or
                // a 9 spoils an octal one.
                if (radix == 16 && matches!(b, b'h' | b'H' | b'o' | b'O')) || (radix == 8 && matches!(b, b'8' | b'9')) {
                    overflow = true;
                    self.pos += 1;
                    continue;
                }
                break;
            };
            val = val * radix as u64 + d as u64;
            if val > u32::MAX as u64 {
                overflow = true;
                val &= u32::MAX as u64;
            }
            count += 1;
            self.pos += 1;
        }
        if count == 0 || overflow { None } else { Some(val as u32) }
    }

    /// Finish a `&H`/`&O` literal. Without a suffix a value that fits 16 bits
    /// is an Integer (its bit pattern, so `&HFFFF` is -1) and a larger one is
    /// a Long; `&` makes it a Long; `%` requires 16 bits.
    fn finish_radix_literal(&mut self, val: Option<u32>, start: usize) -> Token {
        let suffix = self.src.get(self.pos).copied();
        if suffix == Some(b'&') || suffix == Some(b'%') {
            self.pos += 1;
        }
        let span = self.span(start);
        let Some(val) = val else { return self.number_error(span) };
        match suffix {
            Some(b'&') => Token::lit(TokenKind::LongLit, Lit::Long(val as i32), span),
            Some(b'%') if val > 0xFFFF => self.number_error(span),
            _ if val <= 0xFFFF => {
                Token::lit(TokenKind::IntLit, Lit::Int(val as u16 as i16 as i32), span)
            }
            _ => Token::lit(TokenKind::LongLit, Lit::Long(val as i32), span),
        }
    }

    /// The next byte after any blanks (spaces and tabs) and the offset just
    /// past it, without consuming anything. Two-character operators may have
    /// blanks between their characters.
    fn peek_past_blanks(&self) -> Option<(u8, usize)> {
        let mut i = self.pos;
        loop {
            match self.full_width_at(i) {
                Some((b' ', width)) => i += width,
                Some((b, width)) => return Some((b, i + width)),
                None if matches!(self.src.get(i), Some(b' ') | Some(b'\t') | Some(0x19)) => i += 1,
                None => return self.src.get(i).map(|&b| (b, i + 1)),
            }
        }
    }

    /// Consume the operator's second character when it is one of `wanted`
    /// (blanks may precede it); returns which one.
    fn eat_operator_tail(&mut self, wanted: &[u8]) -> Option<u8> {
        match self.peek_past_blanks() {
            Some((b, end)) if wanted.contains(&b) => {
                self.pos = end;
                Some(b)
            }
            _ => None,
        }
    }

    /// Handle `':'`: Colon or `:=`.
    fn scan_colon(&mut self, start: usize) -> Token {
        if self.eat_operator_tail(b"=").is_some() {
            Token::op(Kw::ColonEq, self.span(start))
        } else {
            Token::op(Kw::Colon, self.span(start))
        }
    }

    /// Handle `'<'`: Lt, Le, or Ne.
    fn scan_lt(&mut self, start: usize) -> Token {
        match self.eat_operator_tail(b"=>") {
            Some(b'=') => Token::op(Kw::Le, self.span(start)),
            Some(_) => Token::op(Kw::Ne, self.span(start)),
            None => Token::op(Kw::Lt, self.span(start)),
        }
    }

    /// Handle `'='`: Eq, or the reversed spellings `=<` (Le) and `=>` (Ge).
    fn scan_eq(&mut self, start: usize) -> Token {
        match self.eat_operator_tail(b"<>") {
            Some(b'<') => Token::op(Kw::Le, self.span(start)),
            Some(_) => Token::op(Kw::Ge, self.span(start)),
            None => Token::op(Kw::Eq, self.span(start)),
        }
    }

    /// Handle `'>'`: Gt, Ge, or the reversed `><` (Ne).
    fn scan_gt(&mut self, start: usize) -> Token {
        match self.eat_operator_tail(b"=<") {
            Some(b'=') => Token::op(Kw::Ge, self.span(start)),
            Some(_) => Token::op(Kw::Ne, self.span(start)),
            None => Token::op(Kw::Gt, self.span(start)),
        }
    }

    /// The position of the continuation mark that ends the text from `start` to the current
    /// position: an underscore after a blank, followed by blanks only.
    fn continuation_mark_at_end(&self, start: usize) -> Option<usize> {
        let mut end = self.pos;
        while end > start && matches!(self.src[end - 1], b' ' | b'\t') {
            end -= 1;
        }
        (end > start + 1 && self.src[end - 1] == b'_' && matches!(self.src[end - 2], b' ' | b'\t')).then(|| end - 1)
    }

    /// Handle `'_'`: line-continuation marker or identifier start.
    fn scan_underscore(&mut self, first: u8, start: usize) -> Token {
        if self.peek().map_or(true, |b| !is_ident_continue(b)) {
            // A continuation mark must follow a blank; at the start of a line or against another
            // token it is an invalid character.
            let after_blank = start > 0 && matches!(self.src[start - 1], b' ' | b'\t');
            if after_blank {
                Token::op(Kw::LineCont, self.span(start))
            } else {
                Token::error(self.span(start))
            }
        } else {
            // A name cannot begin with an underscore.
            self.scan_ident(first, start);
            Token::error(self.span(start))
        }
    }

    /// Collect an identifier and resolve it to a keyword or user symbol.
    fn scan_ident(&mut self, first: u8, start: usize) -> Token {
        self.tok_buf.clear();
        self.tok_buf.push(ansi_char(first));
        if self.src[start] == first && self.double_byte_at(start) {
            // A lead byte and the trail byte after it are one character of the name.
            self.tok_buf.push(ansi_char(self.src[self.pos]));
            self.pos += 1;
        }
        while let Some(&b) = self.src.get(self.pos) {
            if let Some((ascii, width)) = self.full_width_at(self.pos) {
                // Full-width letters, digits and the underscore continue the name as themselves.
                if is_ident_continue(ascii) {
                    self.tok_buf.push(ascii as char);
                    self.pos += width;
                    continue;
                }
                break;
            }
            if self.double_byte_at(self.pos) {
                self.tok_buf.push(ansi_char(b));
                self.tok_buf.push(ansi_char(self.src[self.pos + 1]));
                self.pos += 2;
                continue;
            }
            if !is_ident_continue(b) { break; }
            self.tok_buf.push(ansi_char(b));
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
        // The file statements take `#` as the channel mark, never as a suffix.
        if suffix == b'#'
            && ["Close", "Get", "Input", "Lock", "Open", "Print", "Put", "Seek", "Unlock", "Write"]
                .iter()
                .any(|w| self.tok_buf.eq_ignore_ascii_case(w))
        {
            return TypeSuffix::None;
        }
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
            && self.src.get(self.pos + 1).map_or(false, |&n| is_ident_start(n) && n != b'_');
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
        // Brackets inside the name nest: the name ends at the `]` that closes the first.
        let mut depth = 0u32;
        loop {
            match self.advance() {
                None | Some(0x00) => return Token::error(self.span(start)),
                // A bracketed name stops at the end of its line.
                Some(b'\r') | Some(b'\n') => {
                    self.pos -= 1;
                    return Token::error(self.span(start));
                }
                Some(b']') if depth == 0 => break,
                Some(b']') => {
                    depth -= 1;
                    self.tok_buf.push(']');
                }
                Some(b'[') => {
                    depth += 1;
                    self.tok_buf.push('[');
                }
                Some(b) => self.tok_buf.push(ansi_char(b)),
            }
        }
        let span = self.span(start);
        let name: String = self.tok_buf.clone();
        match self.ctx.intern_ident(&name) {
            Some(sym_id) => {
                self.ctx.set_verbatim_flag(sym_id);
                let mut tok = Token::ident(sym_id, span);
                tok.bracketed = true;
                tok
            }
            None => Token::error(span),
        }
    }

    // ── low-level helpers ────────────────────────────────────────────────────

    /// The ASCII character the full-width pair of bytes at `at` stands for, and how many
    /// bytes it takes (always 2).
    fn full_width_at(&self, at: usize) -> Option<(u8, usize)> {
        if !self.double_byte_at(at) {
            return None;
        }
        let column = match self.ctx.dbcs_code_page()? {
            932 => 0,
            936 => 1,
            950 => 2,
            _ => 3,
        };
        let pair = (u16::from(self.src[at]) << 8) | u16::from(self.src[at + 1]);
        FULL_WIDTH.iter().find_map(|&(ascii, count, codes)| {
            let first = codes[column];
            (first != 0 && pair >= first && pair < first + u16::from(count)).then(|| (ascii + (pair - first) as u8, 2))
        })
    }

    /// Whether the byte at `at` leads a double-byte character of the module's code page and a
    /// trail byte follows it on the same line.
    fn double_byte_at(&self, at: usize) -> bool {
        self.src.get(at).is_some_and(|&b| self.ctx.is_dbcs_lead_byte(b))
            && !matches!(self.src.get(at + 1), None | Some(0) | Some(b'\r') | Some(b'\n'))
    }

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

/// The value of a decimal literal in ten-thousandths (Currency units), exact
/// for any digits and exponent; `None` when it does not fit 64 bits.
fn currency_units(text: &str) -> Option<i64> {
    let (mantissa, exp) = match text.split_once('e') {
        Some((m, e)) => (m, e.parse::<i32>().ok()?),
        None => (text, 0),
    };
    let (int_part, frac_part) = mantissa.split_once('.').unwrap_or((mantissa, ""));
    let mut digits: i128 = 0;
    for c in int_part.chars().chain(frac_part.chars()) {
        digits = digits.checked_mul(10)?.checked_add(c.to_digit(10)? as i128)?;
    }
    let shift = 4 + exp - frac_part.len() as i32;
    let units = if shift >= 0 {
        digits.checked_mul(10i128.checked_pow(shift as u32)?)?
    } else {
        let div = 10i128.checked_pow((-shift) as u32)?;
        // Halves round to the even unit.
        let (q, r) = (digits / div, digits % div);
        if 2 * r > div || (2 * r == div && q % 2 == 1) { q + 1 } else { q }
    };
    i64::try_from(units).ok()
}


/// Names the compiler's table holds that are not keywords or names of the module being read,
/// measured as the number of names a module can add before the table is full (32436 members of
/// user types beside 65 type names and one procedure name).
const HOST_SYMBOLS: u32 = 10;

/// The full-width spelling of ASCII characters in each double-byte code page, as
/// (ASCII character, how many consecutive characters follow it, the code of its full-width
/// form in code pages 932, 936, 950 and 949). A pair of source bytes with one of these codes
/// stands for the ASCII character.
const FULL_WIDTH: [(u8, u8, [u16; 4]); 32] = [
    (b'>', 1, [0x8184, 0xa3be, 0xa1d6, 0xa3be]),
    (b'<', 1, [0x8183, 0xa3bc, 0xa1d5, 0xa3bc]),
    (b'=', 1, [0x8181, 0xa3bd, 0xa1d7, 0xa3bd]),
    (b':', 1, [0x8146, 0xa3ba, 0xa147, 0xa3ba]),
    (b' ', 1, [0x8140, 0xa1a1, 0xa140, 0xa1a1]),
    (b'%', 1, [0x8193, 0xa3a5, 0xa248, 0xa3a5]),
    (b'&', 1, [0x8195, 0xa3a6, 0xa1ae, 0xa3a6]),
    (b'!', 1, [0x8149, 0xa3a1, 0xa149, 0xa3a1]),
    (b'#', 1, [0x8194, 0xa3a3, 0xa1ad, 0xa3a3]),
    (b'$', 1, [0x8190, 0xa3a4, 0xa243, 0xa3a4]),
    (b'@', 1, [0x8197, 0xa3c0, 0xa249, 0xa3c0]),
    (b'[', 1, [0x816d, 0xa3db, 0xa165, 0xa3db]),
    (b']', 1, [0x816e, 0xa3dd, 0xa166, 0xa3dd]),
    (b',', 1, [0x8143, 0xa3ac, 0xa141, 0xa3a7]),
    (b'.', 1, [0x8144, 0xa3ae, 0xa144, 0xa3ae]),
    (b'+', 1, [0x817b, 0xa3ab, 0xa1cf, 0xa3ab]),
    (b'-', 1, [0x817c, 0xa3ad, 0xa1d0, 0xa3ad]),
    (b'"', 1, [0x8167, 0xa3a2, 0xa1a7, 0xa3a2]),
    (b'"', 1, [0x8168, 0xa3a2, 0xa1a8, 0xa3a2]),
    (b'\'', 1, [0x8166, 0xa3a7, 0xa141, 0xa1af]),
    (b'(', 1, [0x8169, 0xa3a8, 0xa15d, 0xa3a8]),
    (b')', 1, [0x816a, 0xa3a9, 0xa15e, 0xa3a9]),
    (b'/', 1, [0x815e, 0xa3af, 0xa1fe, 0xa3af]),
    (b'?', 1, [0x8148, 0xa3bf, 0xa1ec, 0xa3bf]),
    (b'?', 1, [0x8148, 0xa3bf, 0xa148, 0xa3bf]),
    (b'^', 1, [0x814f, 0xa3de, 0x0000, 0xa3de]),
    (b'*', 1, [0x8196, 0xa3aa, 0xa1af, 0xa3aa]),
    (b'A', 26, [0x8260, 0xa3c1, 0xa2cf, 0xa3c1]),
    (b'a', 26, [0x8281, 0xa3e1, 0xa2e9, 0xa3e1]),
    (b'e', 1, [0x8285, 0xa3e5, 0xa2ed, 0xa3e5]),
    (b'E', 1, [0x8264, 0xa3c5, 0xa2d3, 0xa3c5]),
    (b'0', 10, [0x824f, 0xa3b0, 0xa2af, 0xa3b0]),
];
