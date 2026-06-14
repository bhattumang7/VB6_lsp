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
                pos += 1;
                let start = pos;
                while pos < b.len() && b[pos] != b'}' {
                    pos += 1;
                }
                tokens.push(Token::Curly(&line[start..pos]));
                if pos < b.len() { pos += 1; } // skip '}'
            }

            // Word: identifier, keyword, number, hex literal, &H..., #...
            _ => {
                let start = pos;
                while pos < b.len() {
                    match b[pos] {
                        b' ' | b'\t' | b'\r' | b'\n' | b'=' | b';' | b'\'' | b'"' | b':' | b'{' => break,
                        _ => pos += 1,
                    }
                }
                tokens.push(Token::Word(&line[start..pos]));
            }
        }
    }

    tokens
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
mod tests {
    use super::*;

    #[test]
    fn lex_simple_property() {
        let toks = lex_line("   BackColor       =   &H80000005&");
        assert_eq!(toks[0], Token::Word("BackColor"));
        assert_eq!(toks[1], Token::Equals);
        assert_eq!(toks[2], Token::Word("&H80000005&"));
    }

    #[test]
    fn lex_quoted_value() {
        let toks = lex_line(r#"Caption         =   "Hello World""#);
        assert_eq!(toks[0], Token::Word("Caption"));
        assert_eq!(toks[1], Token::Equals);
        assert_eq!(toks[2], Token::Quoted("Hello World"));
    }

    #[test]
    fn lex_comment_stripped() {
        let toks = lex_line("  Begin  ' this is a comment");
        assert_eq!(toks.len(), 1);
        assert_eq!(toks[0], Token::Word("Begin"));
    }

    #[test]
    fn lex_begin_line() {
        let toks = lex_line("   Begin VB.Form Form1");
        assert_eq!(toks[0], Token::Word("Begin"));
        assert_eq!(toks[1], Token::Word("VB.Form"));
        assert_eq!(toks[2], Token::Word("Form1"));
    }

    #[test]
    fn lex_object_line() {
        // In .frm files the progid is in double quotes: Object = "...#ver#lcid"
        let toks = lex_line(
            r#"Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX""#,
        );
        assert_eq!(toks[0], Token::Word("Object"));
        assert_eq!(toks[1], Token::Equals);
        // The entire progid is inside outer double-quotes → Quoted token
        match &toks[2] {
            Token::Quoted(s) => assert!(s.contains("831FDD16")),
            other => panic!("expected Quoted, got {:?}", other),
        }
        assert_eq!(toks[3], Token::Semi);
        assert_eq!(toks[4], Token::Quoted("MSCOMCTL.OCX"));
    }

    #[test]
    fn lex_object_line_unquoted() {
        // In .vbp files the progid is bare: Object={GUID}#ver#lcid ; filename
        let toks = lex_line(
            "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0; MSCOMCTL.OCX",
        );
        match &toks[0] {
            Token::Curly(s) => assert!(s.contains("831FDD16")),
            other => panic!("expected Curly, got {:?}", other),
        }
        // #2.0#0 follows as Word
        assert_eq!(toks[1], Token::Word("#2.0#0"));
    }

    #[test]
    fn lex_dollar_string() {
        let toks = lex_line(r#"Caption         =   $"Hello""#);
        assert_eq!(toks[2], Token::DollarStr("Hello"));
    }

    #[test]
    fn lex_escaped_quote_in_string() {
        let toks = lex_line(r#"Caption = "say ""hi""  ""#);
        assert_eq!(toks[2], Token::Quoted(r#"say ""hi""  "#));
    }
}
