//! Reading of a module's text into physical lines, pieces and logical lines.
//!
//! A physical line ends at a line feed, a carriage return, or a carriage return and
//! line feed together. Each physical line is read in pieces of at most
//! [`PIECE_COLUMNS`] columns (a tab fills the columns up to the next multiple of 8,
//! counted from the start of its piece); a tab that would pass the limit starts the
//! next piece. The pieces of a line are joined for the scanner, but no token spans
//! two pieces, and a string or a name ends at the end of its piece.
//!
//! A logical line is the run of pieces that ends in a piece that is the last of its
//! physical line and does not end in a continuation mark (a blank and an underscore,
//! with blanks allowed after the underscore). A logical line holds at most
//! [`MAX_PIECES`] pieces and a module at most [`MAX_LINES`] logical lines, not
//! counting the header lines (`VERSION`, the `BEGIN` ... `END` block and
//! `Attribute` lines). A NUL ends the text of the piece it is in. A last byte of
//! 0x1A is read as a NUL.

use std::borrow::Cow;

use super::diagnostics::{ERR_OUT_OF_MEMORY, ERR_TOO_MANY_LINE_CONTINUATIONS};

/// Most columns of one piece.
pub const PIECE_COLUMNS: usize = 1023;
/// Most pieces of one logical line.
pub const MAX_PIECES: usize = 25;
/// Most logical lines of a module, not counting header lines.
pub const MAX_LINES: usize = 65530;

/// One physical line of a text: its text is `start..end`, the line break (if any)
/// is `end..next`.
#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub struct PhysicalLine {
    pub start: usize,
    pub end: usize,
    pub next: usize,
}

/// The physical lines of `src`. A last line without a line break is a line; the
/// empty text after a final line break is not.
pub fn physical_lines(src: &[u8]) -> impl Iterator<Item = PhysicalLine> + '_ {
    let mut pos = 0;
    std::iter::from_fn(move || {
        if pos >= src.len() {
            return None;
        }
        let start = pos;
        let end = src[start..].iter().position(|b| matches!(b, b'\r' | b'\n')).map_or(src.len(), |i| start + i);
        let next = if end == src.len() {
            end
        } else if src[end] == b'\r' && src.get(end + 1) == Some(&b'\n') {
            end + 2
        } else {
            end + 1
        };
        pos = next;
        Some(PhysicalLine { start, end, next })
    })
}

/// The offset at which each line of `src` starts, counting every line feed, carriage
/// return and carriage return with line feed as one break. The first entry is 0 and
/// a text that ends in a break has a last, empty line.
pub fn physical_line_starts(src: &[u8]) -> Vec<u32> {
    let mut starts = vec![0u32];
    for line in physical_lines(src) {
        if line.next > line.end {
            starts.push(line.next as u32);
        }
    }
    starts
}

/// Why the text cannot be read in full.
#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub struct ReadProblem {
    /// The error code of the problem.
    pub code: u32,
    /// Where the piece or line that breaks the limit starts.
    pub offset: u32,
}

/// The lines of a text as the compiler numbers them: every piece of a physical line
/// is a line of its own, and the lines of a logical line are numbered one after the
/// other. A text that ends in a line break has an empty last line after it.
#[derive(Clone, Debug, Default, PartialEq, Eq)]
pub struct LineTable {
    /// The offset at which each piece starts.
    piece_starts: Vec<u32>,
    /// For each piece, the index of the first piece of its logical line.
    logical_first: Vec<u32>,
    /// The length of the text, when it ends in a line break (the empty last line starts there).
    empty_last_line_at: Option<u32>,
}

impl LineTable {
    /// The table of `src`.
    pub fn new(src: &[u8]) -> Self {
        read(src).table
    }

    /// The number of pieces read.
    pub fn len(&self) -> usize {
        self.piece_starts.len()
    }

    pub fn is_empty(&self) -> bool {
        self.piece_starts.is_empty()
    }

    fn piece_at(&self, offset: u32) -> usize {
        self.piece_starts.partition_point(|&s| s <= offset).saturating_sub(1)
    }

    /// Whether `offset` is on the empty last line after a final line break.
    fn on_empty_last_line(&self, offset: u32) -> bool {
        self.empty_last_line_at.is_some_and(|end| offset >= end)
    }

    /// The number (from 1) of the line the piece holding `offset` is.
    pub fn line_of(&self, offset: u32) -> u32 {
        if self.on_empty_last_line(offset) {
            return self.piece_starts.len() as u32 + 1;
        }
        self.piece_at(offset) as u32 + 1
    }

    /// The offset at which the logical line holding `offset` starts.
    pub fn logical_line_start(&self, offset: u32) -> u32 {
        if let Some(end) = self.empty_last_line_at.filter(|&end| offset >= end) {
            return end;
        }
        match self.piece_starts.is_empty() {
            true => 0,
            false => self.piece_starts[self.logical_first[self.piece_at(offset)] as usize],
        }
    }

    /// The number (from 1) of the first line of the logical line holding `offset`:
    /// the line on which an error found in that logical line is reported.
    pub fn logical_line_of(&self, offset: u32) -> u32 {
        if self.on_empty_last_line(offset) {
            return self.piece_starts.len() as u32 + 1;
        }
        match self.piece_starts.is_empty() {
            true => 1,
            false => self.logical_first[self.piece_at(offset)] + 1,
        }
    }
}

/// A text as the scanner reads it.
#[derive(Debug)]
pub struct ReadText<'a> {
    /// The text: the source up to the end of what is read, with a NUL inserted where
    /// a piece starts inside its physical line, and a last 0x1A turned into a NUL.
    pub text: Cow<'a, [u8]>,
    /// Where in `text` the inserted NULs are, in ascending order.
    pub separators: Vec<u32>,
    /// The first limit the text breaks, if any.
    pub problem: Option<ReadProblem>,
    /// What the limits are measured on.
    pub limits: LineLimits,
    /// The lines as the compiler numbers them.
    pub table: LineTable,
}

/// What the limits on a text are measured on: where the first logical line with too many
/// pieces breaks the limit, and where each logical line that counts toward the line limit
/// starts.
#[derive(Clone, Debug, Default, PartialEq, Eq)]
pub struct LineLimits {
    piece_problem: Option<ReadProblem>,
    /// For each: where it starts, and where its last piece starts.
    counted: Vec<(u32, u32)>,
}

impl LineLimits {
    /// The first limit the text breaks, when the lines that start before `header_end` are
    /// header lines of a kind the text cannot show (a designer section that was blanked).
    pub fn problem(&self, header_end: u32) -> Option<ReadProblem> {
        let first = self.counted.partition_point(|&(start, _)| start < header_end);
        let line_problem = self
            .counted
            .get(first + MAX_LINES)
            .map(|&(start, last_piece)| (ReadProblem { code: ERR_OUT_OF_MEMORY, offset: start }, last_piece));
        // A line is stored when its last piece has been read; a piece found to be one too many
        // before that is found first.
        match (self.piece_problem, line_problem) {
            (Some(p), Some((l, last_piece))) => Some(if p.offset <= last_piece { p } else { l }),
            (Some(p), None) => Some(p),
            (None, l) => l.map(|(l, _)| l),
        }
    }
}

/// The offset in the source of the character at `pos` in a text with NULs inserted at
/// `separators`.
pub fn source_offset(separators: &[u32], pos: usize) -> u32 {
    (pos - separators.partition_point(|&s| (s as usize) < pos)) as u32
}

fn is_blank(b: u8) -> bool {
    matches!(b, b' ' | b'\t')
}

/// The text of a piece: what comes before its first NUL.
fn piece_text(piece: &[u8]) -> &[u8] {
    piece.split(|&b| b == 0).next().unwrap_or(&[])
}

/// Whether the text of a piece ends in a continuation mark: a blank, an underscore
/// and then only blanks.
fn ends_in_continuation_mark(piece: &[u8]) -> bool {
    let text = piece_text(piece);
    let mut end = text.len();
    while end > 0 && is_blank(text[end - 1]) {
        end -= 1;
    }
    end >= 2 && text[end - 1] == b'_' && is_blank(text[end - 2])
}

/// What the header of a module is made of, line by line.
#[derive(Default)]
struct HeaderState {
    /// Nesting depth of the open `BEGIN` ... `END` block.
    depth: usize,
    /// Whether a line other than a header, blank or comment line was seen.
    started: bool,
}

impl HeaderState {
    /// Whether the logical line starting with `first_piece` is a header line.
    fn is_header_line(&mut self, first_piece: &[u8]) -> bool {
        let text = piece_text(first_piece);
        let body = &text[text.iter().position(|&b| !is_blank(b)).unwrap_or(text.len())..];
        let word_len = body.iter().position(|b| !(b.is_ascii_alphanumeric() || *b == b'_')).unwrap_or(body.len());
        let word = body[..word_len].to_ascii_lowercase();
        if self.depth > 0 {
            match word.as_slice() {
                b"begin" | b"beginproperty" => self.depth += 1,
                b"end" | b"endproperty" => self.depth -= 1,
                _ => {}
            }
            return true;
        }
        if word == b"attribute" {
            return true;
        }
        if !self.started {
            match word.as_slice() {
                b"version" => return true,
                b"begin" => {
                    self.depth = 1;
                    return true;
                }
                _ => {}
            }
        }
        if !body.is_empty() && body[0] != b'\'' {
            self.started = true;
        }
        false
    }
}

struct Piece {
    start: usize,
    end: usize,
    /// The piece is not the last of its physical line.
    inner: bool,
    /// The piece ends in a continuation mark.
    marked: bool,
}

impl Piece {
    fn continues(&self) -> bool {
        self.inner || self.marked
    }
}

/// The text with a 0x1A that ends its last line read as a NUL. A carriage return that is the
/// last byte leaves its line open (a line feed might follow), so that line is the last.
fn final_substitute_as_nul(src: &[u8]) -> Cow<'_, [u8]> {
    let last_line_end = match src.last() {
        Some(b'\r') => src.len() - 1,
        _ => src.len(),
    };
    if last_line_end > 0 && src[last_line_end - 1] == 0x1a {
        let mut v = src.to_vec();
        v[last_line_end - 1] = 0;
        Cow::Owned(v)
    } else {
        Cow::Borrowed(src)
    }
}

/// Splits one physical line into pieces of at most [`PIECE_COLUMNS`] columns.
fn split_physical_line(bytes: &[u8], line: PhysicalLine, pieces: &mut Vec<Piece>) {
    let mut start = line.start;
    let mut column = 0usize;
    for i in line.start..line.end {
        let mut width = if bytes[i] == b'\t' { 8 - column % 8 } else { 1 };
        if column + width > PIECE_COLUMNS {
            pieces.push(Piece { start, end: i, inner: true, marked: false });
            start = i;
            column = 0;
            width = if bytes[i] == b'\t' { 8 } else { 1 };
        }
        column += width;
    }
    let marked = ends_in_continuation_mark(&bytes[start..line.end]);
    pieces.push(Piece { start, end: line.end, inner: false, marked });
}

/// The logical lines of the pieces, and the limits on them. Returns the table, the limits and
/// the index of the first piece of the last logical line.
fn build_logical_lines(bytes: &[u8], pieces: &[Piece]) -> (LineTable, LineLimits, usize) {
    let mut table = LineTable::default();
    let mut limits = LineLimits::default();
    let mut header = HeaderState::default();
    let mut first = 0usize;
    for (n, piece) in pieces.iter().enumerate() {
        if n == 0 || !pieces[n - 1].continues() {
            first = n;
        }
        table.piece_starts.push(piece.start as u32);
        table.logical_first.push(first as u32);
        if n - first == MAX_PIECES && limits.piece_problem.is_none() {
            limits.piece_problem =
                Some(ReadProblem { code: ERR_TOO_MANY_LINE_CONTINUATIONS, offset: piece.start as u32 });
        }
        if !piece.continues() && !header.is_header_line(&bytes[pieces[first].start..pieces[first].end]) {
            limits.counted.push((pieces[first].start as u32, piece.start as u32));
        }
    }
    (table, limits, first)
}

/// Reads `src`.
pub fn read(src: &[u8]) -> ReadText<'_> {
    let base = final_substitute_as_nul(src);
    let bytes: &[u8] = &base;

    let mut pieces: Vec<Piece> = Vec::new();
    for line in physical_lines(bytes) {
        split_physical_line(bytes, line, &mut pieces);
    }

    let (mut table, limits, first) = build_logical_lines(bytes, &pieces);
    let problem = limits.problem(0);

    // A logical line that ends in a continuation mark at the end of the text is not read.
    let mut read_end = bytes.len();
    if pieces.last().is_some_and(|p| p.continues()) {
        read_end = pieces[first].start;
        pieces.truncate(first);
        table.piece_starts.truncate(first);
        table.logical_first.truncate(first);
    }

    if read_end == bytes.len() && bytes.last().is_some_and(|b| matches!(b, b'\r' | b'\n')) {
        table.empty_last_line_at = Some(bytes.len() as u32);
    }

    // The text: a NUL where a piece starts inside its physical line.
    let splits: Vec<usize> =
        pieces.iter().enumerate().filter(|(n, _)| *n > 0 && pieces[n - 1].inner).map(|(_, p)| p.start).collect();
    if splits.is_empty() && read_end == bytes.len() {
        return ReadText { text: base, separators: Vec::new(), problem, limits, table };
    }
    let (out, separators) = insert_separators(bytes, read_end, &splits);
    ReadText { text: Cow::Owned(out), separators, problem, limits, table }
}

/// The text up to `read_end` with a NUL inserted at each of `splits`, and where the NULs are.
fn insert_separators(bytes: &[u8], read_end: usize, splits: &[usize]) -> (Vec<u8>, Vec<u32>) {
    let mut out = Vec::with_capacity(read_end + splits.len());
    let mut separators = Vec::with_capacity(splits.len());
    let mut copied = 0;
    for &s in splits {
        out.extend_from_slice(&bytes[copied..s]);
        separators.push(out.len() as u32);
        out.push(0);
        copied = s;
    }
    out.extend_from_slice(&bytes[copied..read_end]);
    (out, separators)
}
