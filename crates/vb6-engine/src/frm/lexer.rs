/// Line-oriented lexer for VB6 .frm / .cls text files.
///
/// A character-level state machine: skips whitespace, strips `'` line
/// comments, handles `"..."` quoted strings, `$"..."` dollar-prefixed strings,
/// and bare identifier tokens that stop at `=` / `;` / `:` / whitespace / `'`.
use std::fmt;

/// A single lexed token from one logical line.
#[derive(Debug, Clone, PartialEq)]
pub enum Token<'a> {
    /// Bare keyword / identifier / number (e.g. `Begin`, `BackColor`, `123`).
    Word(&'a str),
    /// `"quoted string"` — outer quotes stripped, `""` escapes preserved as-is.
    Quoted(&'a str),
    /// `$"..."` dollar-prefixed BSTR string — outer quotes stripped.
    DollarStr(&'a str),
    /// `=` assignment operator.
    Equals,
    /// `;` separator (used in `Object =` and `Reference =` lines).
    Semi,
    /// `:` separator.
    Colon,
    /// `{...}` curly-brace literal (GUID / ProgID).
    Curly(&'a str),
}

impl fmt::Display for Token<'_> {
    fn fmt(&self, f: &mut fmt::Formatter<'_>) -> fmt::Result {
        match self {
            Token::Word(s) | Token::Quoted(s) | Token::DollarStr(s) | Token::Curly(s) => {
                write!(f, "{}", s)
            }
            Token::Equals => write!(f, "="),
            Token::Semi => write!(f, ";"),
            Token::Colon => write!(f, ":"),
        }
    }
}

/// Lex one source line into tokens.  Returns an empty vec for blank /
/// comment-only lines.  `'` starts a comment that consumes the rest of the
/// line.
pub fn lex_line(line: &str) -> Vec<Token<'_>> {
    let mut tokens = Vec::new();
    let b = line.as_bytes();
    let mut pos = 0usize;

    while pos < b.len() {
        match b[pos] {
            // Whitespace — skip
            b' ' | b'\t' | b'\r' | b'\n' => pos += 1,

            // Comment — remainder of line ignored
            b'\'' => break,

            b'=' => { tokens.push(Token::Equals); pos += 1; }
            b';' => { tokens.push(Token::Semi);   pos += 1; }
            b':' => { tokens.push(Token::Colon);  pos += 1; }

            // Dollar-prefixed string: $"..."
            b'$' if b.get(pos + 1) == Some(&b'"') => {
                pos += 2; // skip $"
                let (content, end) = read_quoted(line, pos);
                tokens.push(Token::DollarStr(content));
                pos = end;
            }

            // Quoted string: "..."
            b'"' => {
                pos += 1; // skip opening "
                let (content, end) = read_quoted(line, pos);
                tokens.push(Token::Quoted(content));
                pos = end;
            }

            // GUID / ProgID in curly braces: {xxxxxxxx-...}
            b'{' => {
                let (content, end) = read_curly(line, pos + 1);
                tokens.push(Token::Curly(content));
                pos = end;
            }

            // Word: identifier, keyword, number, hex literal, &H..., #...
            _ => {
                let (content, end) = read_word(line, pos);
                tokens.push(Token::Word(content));
                pos = end;
            }
        }
    }

    tokens
}

/// Read a brace-delimited body starting at `pos` (after the opening `{`).
/// Returns (content_slice, pos_after_closing_brace); an unterminated body
/// runs to the end of the line.
fn read_curly(line: &str, pos: usize) -> (&str, usize) {
    let b = line.as_bytes();
    let mut end = pos;
    while end < b.len() && b[end] != b'}' {
        end += 1;
    }
    let next = if end < b.len() { end + 1 } else { end };
    (&line[pos..end], next)
}

/// Read a bare word starting at `pos`; returns (word_slice, pos_after_word).
fn read_word(line: &str, pos: usize) -> (&str, usize) {
    let b = line.as_bytes();
    let mut end = pos;
    while end < b.len() {
        match b[end] {
            b' ' | b'\t' | b'\r' | b'\n' | b'=' | b';' | b'\'' | b'"' | b':' | b'{' => break,
            _ => end += 1,
        }
    }
    (&line[pos..end], end)
}

/// Read a quoted-string body starting at `pos` (after the opening `"`).
/// Returns (content_slice, pos_after_closing_quote).
/// VB6 uses `""` as the escape for a literal `"` inside a string.
fn read_quoted(line: &str, pos: usize) -> (&str, usize) {
    let b = line.as_bytes();
    let start = pos;
    let mut i = pos;
    while i < b.len() {
        if b[i] == b'"' {
            if b.get(i + 1) == Some(&b'"') {
                // escaped quote — include both in content, keep scanning
                i += 2;
            } else {
                // closing quote
                let content = &line[start..i];
                return (content, i + 1); // skip closing "
            }
        } else {
            i += 1;
        }
    }
    // unterminated string — return what we have
    (&line[start..i], i)
}

// ---------------------------------------------------------------------------
// Tests
// ---------------------------------------------------------------------------

#[cfg(test)]
#[path = "tests/lexer.rs"]
mod tests;
