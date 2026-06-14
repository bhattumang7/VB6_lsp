/// VB6 .frm / .cls / .ctl / .dob / .pag parser.
///
/// Entry point: [`parse_frm`].
use std::fmt;

use super::lexer::{Token, lex_line};

// ---------------------------------------------------------------------------
// Public types
// ---------------------------------------------------------------------------

/// A parsed VB6 designer file.
#[derive(Debug, Clone)]
pub struct FrmFile {
    /// Version string from the `VERSION` header (e.g. `"5.00"`).
    pub version: String,
    /// File kind determined from the header.
    pub kind: FileKind,
    /// `Object = "{progid}"; "name"` declarations (OCX control references).
    pub objects: Vec<ObjectRef>,
    /// `Attribute VB_Name = "value"` lines before the first `Begin`.
    pub attributes: Vec<Attribute>,
    /// Top-level `Begin...End` block (the form or user-control itself).
    /// `None` for `.cls` class modules which have no designer block.
    pub root: Option<BeginBlock>,
    /// Number of source lines consumed by the designer header (VERSION + Object
    /// lines + root Begin/End block). The VB code section starts at line
    /// `designer_lines + 1` (1-based). Zero means no designer block was found.
    pub designer_lines: usize,
}

/// Discriminates the file type identified in the header.
#[derive(Debug, Clone, PartialEq)]
pub enum FileKind {
    /// `.frm` — standard Form
    Form,
    /// `.frm` — MDI Form
    MdiForm,
    /// `.cls` — Class Module (no designer block)
    Class,
    /// `.ctl` — UserControl
    UserControl,
    /// `.dob` — UserDocument
    UserDocument,
    /// `.pag` — PropertyPage
    PropertyPage,
    /// Version 1.x or unknown header body
    Unknown,
}

/// `Object = "{progid}"; "filename"` declaration: reads the progid plus an
/// optional `; filename`.
#[derive(Debug, Clone)]
pub struct ObjectRef {
    /// Full ProgID string, e.g. `{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0`
    pub progid: String,
    /// Optional filename / registered name after `;`.
    pub filename: Option<String>,
}

/// An `Attribute` line: `Attribute VB_Name = "value"`.
#[derive(Debug, Clone)]
pub struct Attribute {
    pub name: String,
    pub value: String,
}

/// A `Begin ... End` control block.
#[derive(Debug, Clone)]
pub struct BeginBlock {
    /// Control type, e.g. `VB.Form`, `VB.CommandButton`, `MSComctlLib.TabStrip`.
    pub control_type: String,
    /// Control instance name, e.g. `Form1`, `cmdOK`.
    pub name: String,
    /// Properties set on this control.
    pub properties: Vec<Property>,
    /// Nested child controls (each has its own `Begin...End` block).
    pub children: Vec<BeginBlock>,
}

/// A single property entry inside a `Begin...End` block.
#[derive(Debug, Clone)]
pub struct Property {
    pub name: String,
    pub kind: PropKind,
    pub value: PropValue,
}

/// How the property value was expressed in the source.
#[derive(Debug, Clone, PartialEq)]
pub enum PropKind {
    /// `PropName = value` — inline scalar assignment.
    Simple,
    /// `BeginProperty PropName {GUID}` … `EndProperty` — sub-property bag.
    /// `guid` is the optional type-library GUID that follows the property name
    /// (e.g. `BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851}`).
    BeginProperty { guid: Option<String> },
}

/// The value side of a property assignment.
#[derive(Debug, Clone)]
pub enum PropValue {
    /// Scalar value — stored verbatim as it appeared in the source
    /// (e.g. `"Hello"`, `&H80000005&`, `-1  'True`, `1`).
    Scalar(String),
    /// Sub-properties inside a `BeginProperty ... EndProperty` block.
    Bag(Vec<Property>),
}

// ---------------------------------------------------------------------------
// Error type
// ---------------------------------------------------------------------------

#[derive(Debug)]
pub struct FrmError {
    pub line: usize,
    pub msg: String,
}

impl fmt::Display for FrmError {
    fn fmt(&self, f: &mut fmt::Formatter<'_>) -> fmt::Result {
        write!(f, "frm parse error at line {}: {}", self.line, self.msg)
    }
}

impl std::error::Error for FrmError {}

// ---------------------------------------------------------------------------
// Parser state
// ---------------------------------------------------------------------------

struct Parser<'src> {
    lines: Vec<(usize, &'src str)>, // (1-based line number, content)
    pos: usize,
}

impl<'src> Parser<'src> {
    fn new(src: &'src str) -> Self {
        let lines: Vec<(usize, &'src str)> = src
            .lines()
            .enumerate()
            .map(|(i, l)| (i + 1, l))
            .collect();
        Parser { lines, pos: 0 }
    }

    fn next_tokens(&mut self) -> Option<(usize, Vec<Token<'src>>)> {
        while self.pos < self.lines.len() {
            let (ln, text) = self.lines[self.pos];
            self.pos += 1;
            let toks = lex_line(text);
            if !toks.is_empty() {
                return Some((ln, toks));
            }
        }
        None
    }

    fn err(&self, ln: usize, msg: impl Into<String>) -> FrmError {
        FrmError { line: ln, msg: msg.into() }
    }
}

// ---------------------------------------------------------------------------
// Public entry point
// ---------------------------------------------------------------------------

/// Parse a VB6 designer file from `src` text.
///
/// Handles the VERSION header and file-type detection, then the `Object=`
/// lines, Begin/End block recursion, and property loop.
pub fn parse_frm(src: &str) -> Result<FrmFile, FrmError> {
    let mut p = Parser::new(src);

    // --- VERSION line -------------------------------------------------------
    // Require the "VERSION" keyword then a number.
    let (ln, toks) = p
        .next_tokens()
        .ok_or_else(|| FrmError { line: 0, msg: "empty file".into() })?;

    let (version, is_class) = parse_version_line(ln, &toks)?;

    // --- Determine file kind from next token --------------------------------
    // If the next token == "Class" → .cls; otherwise expect end of line → form.
    // We peek at the next non-blank line.
    let (ln2, toks2) = p
        .next_tokens()
        .ok_or_else(|| FrmError { line: ln, msg: "unexpected EOF after VERSION".into() })?;

    // Collect Object=, Attribute, Begin, and detect file kind
    let mut objects: Vec<ObjectRef> = Vec::new();
    let mut attributes: Vec<Attribute> = Vec::new();
    let mut root: Option<BeginBlock> = None;
    let mut kind = if is_class { FileKind::Class } else { FileKind::Unknown };
    let mut designer_lines = 0usize;

    // Process toks2 — may be a CLASS keyword line, Object=, Attribute, or Begin.
    // After any Begin...End block, keep reading for trailing Attribute lines
    // (class modules have Attribute lines *after* the END block).
    let consumed = process_header_line(ln2, &toks2, &mut kind, &mut objects, &mut attributes)?;
    if consumed == ConsumedLine::Begin {
        root = Some(parse_begin_block(&mut p, ln2, &toks2)?);
        designer_lines = p.pos;
    }

    // Continue reading: Object=, Attribute, and (if not yet found) Begin lines
    loop {
        match p.next_tokens() {
            None => break,
            Some((ln, toks)) => {
                let c = process_header_line(ln, &toks, &mut kind, &mut objects, &mut attributes)?;
                if c == ConsumedLine::Begin && root.is_none() {
                    root = Some(parse_begin_block(&mut p, ln, &toks)?);
                    designer_lines = p.pos;
                }
            }
        }
    }

    Ok(FrmFile { version, kind, objects, attributes, root, designer_lines })
}

// ---------------------------------------------------------------------------
// Header line processing
// ---------------------------------------------------------------------------

#[derive(PartialEq)]
enum ConsumedLine {
    Done,   // line processed (Object=, Attribute, Class, etc.)
    Begin,  // line starts a Begin block — caller should parse it
}

fn process_header_line<'src>(
    ln: usize,
    toks: &[Token<'src>],
    kind: &mut FileKind,
    objects: &mut Vec<ObjectRef>,
    attributes: &mut Vec<Attribute>,
) -> Result<ConsumedLine, FrmError> {
    match toks.first() {
        Some(Token::Word(kw)) => match kw.to_ascii_uppercase().as_str() {
            "CLASS" => {
                *kind = FileKind::Class;
                Ok(ConsumedLine::Done)
            }
            "OBJECT" => {
                objects.push(parse_object_line(ln, toks)?);
                Ok(ConsumedLine::Done)
            }
            "ATTRIBUTE" => {
                attributes.push(parse_attribute_line(ln, toks)?);
                Ok(ConsumedLine::Done)
            }
            "BEGIN" => {
                // Detect form type from control_type token (toks[1])
                if let Some(Token::Word(ct)) = toks.get(1) {
                    *kind = detect_file_kind(ct);
                }
                Ok(ConsumedLine::Begin)
            }
            _ => {
                // Unknown header line — tolerate and skip
                Ok(ConsumedLine::Done)
            }
        },
        _ => Ok(ConsumedLine::Done),
    }
}

fn detect_file_kind(control_type: &str) -> FileKind {
    match control_type {
        s if s.eq_ignore_ascii_case("VB.MDIForm") => FileKind::MdiForm,
        s if s.eq_ignore_ascii_case("VB.UserControl") => FileKind::UserControl,
        s if s.eq_ignore_ascii_case("VB.UserDocument") => FileKind::UserDocument,
        s if s.eq_ignore_ascii_case("VB.PropertyPage") => FileKind::PropertyPage,
        _ => FileKind::Form,
    }
}

// ---------------------------------------------------------------------------
// Line parsers
// ---------------------------------------------------------------------------

/// Returns `(version_string, is_class_module)`.
/// VB6 class files begin with `VERSION 1.0 CLASS` — the CLASS keyword
/// appears on the same line as VERSION. After reading VERSION, the next token
/// is read; if it equals "Class" the file is a class module.
fn parse_version_line(ln: usize, toks: &[Token<'_>]) -> Result<(String, bool), FrmError> {
    match toks.first() {
        Some(Token::Word(kw)) if kw.eq_ignore_ascii_case("VERSION") => {}
        _ => return Err(FrmError { line: ln, msg: "expected VERSION header".into() }),
    }
    let ver = match toks.get(1) {
        Some(Token::Word(v)) => v.to_string(),
        _ => return Err(FrmError { line: ln, msg: "expected version number after VERSION".into() }),
    };
    let is_class = matches!(toks.get(2), Some(Token::Word(kw)) if kw.eq_ignore_ascii_case("CLASS"));
    Ok((ver, is_class))
}

fn parse_object_line(ln: usize, toks: &[Token<'_>]) -> Result<ObjectRef, FrmError> {
    // .frm format: Object = "{GUID}#ver#lcid" ; "filename.ocx"
    // The entire progid is enclosed in outer double-quotes:
    //   toks: [Word("Object"), Equals, Quoted("{GUID}#2.0#0"), Semi, Quoted("MSCOMCTL.OCX")]
    // We also handle the bare form (no outer quotes) for robustness.

    let mut progid: Option<String> = None;
    let mut filename: Option<String> = None;
    let mut seen_eq = false;
    let mut after_semi = false;

    for tok in toks {
        match tok {
            Token::Word(kw) if kw.eq_ignore_ascii_case("Object") && !seen_eq => continue,
            Token::Equals => { seen_eq = true; }
            Token::Semi => { after_semi = true; }
            Token::Quoted(s) if after_semi => { filename = Some(s.to_string()); }
            // Progid as quoted string (normal .frm format)
            Token::Quoted(s) if seen_eq && !after_semi => {
                progid = Some(s.to_string());
            }
            // Progid as bare tokens (fallback)
            Token::Curly(s) if seen_eq && !after_semi && progid.is_none() => {
                progid = Some(format!("{{{}}}", s));
            }
            Token::Word(s) if seen_eq && !after_semi && progid.is_none() => {
                progid = Some(s.to_string());
            }
            _ => {}
        }
    }

    let progid = progid
        .ok_or_else(|| FrmError { line: ln, msg: "malformed Object= line".into() })?;

    Ok(ObjectRef { progid, filename })
}

fn parse_attribute_line(ln: usize, toks: &[Token<'_>]) -> Result<Attribute, FrmError> {
    // Attribute VB_Name = "value"
    // toks: [Word("Attribute"), Word("VB_Name"), Equals, Quoted("value")]
    let name = match toks.get(1) {
        Some(Token::Word(n)) => n.to_string(),
        _ => return Err(FrmError { line: ln, msg: "malformed Attribute line".into() }),
    };
    // Find value after '='
    let value = toks
        .iter()
        .skip_while(|t| !matches!(t, Token::Equals))
        .skip(1)
        .find_map(|t| match t {
            Token::Quoted(s) | Token::DollarStr(s) | Token::Word(s) => Some(s.to_string()),
            _ => None,
        })
        .unwrap_or_default();

    Ok(Attribute { name, value })
}

// ---------------------------------------------------------------------------
// Begin ... End block parser
// ---------------------------------------------------------------------------

/// Parse a `Begin TypeName CtrlName` block.  `toks` is already the `Begin`
/// line; the parser is positioned at the line *after* Begin.
///
/// VB6 class modules use a bare `BEGIN` / `END` block (no type or name) for
/// class-module attributes — this is handled by making type and name optional.
fn parse_begin_block<'src>(
    p: &mut Parser<'src>,
    begin_ln: usize,
    begin_toks: &[Token<'src>],
) -> Result<BeginBlock, FrmError> {
    // begin_toks: [Word("Begin"), Word(control_type), Word(name)]
    // For .cls class modules: just [Word("BEGIN")] — no type or name.
    let control_type = match begin_toks.get(1) {
        Some(Token::Word(s)) => s.to_string(),
        _ => String::new(),
    };
    let name = match begin_toks.get(2) {
        Some(Token::Word(s)) => s.to_string(),
        _ => String::new(),
    };
    let _ = begin_ln; // suppress unused warning

    let mut properties: Vec<Property> = Vec::new();
    let mut children: Vec<BeginBlock> = Vec::new();

    loop {
        let (ln, toks) = match p.next_tokens() {
            Some(v) => v,
            None => return Err(p.err(begin_ln, "unexpected EOF inside Begin block")),
        };
        if parse_begin_block_line(p, ln, &toks, &mut properties, &mut children)? {
            break;
        }
    }

    Ok(BeginBlock { control_type, name, properties, children })
}

/// Handle one line inside a `Begin` block, appending to `properties`/`children`.
/// Returns `Ok(true)` when the line is the closing `END` (caller should stop).
fn parse_begin_block_line<'src>(
    p: &mut Parser<'src>,
    ln: usize,
    toks: &[Token<'src>],
    properties: &mut Vec<Property>,
    children: &mut Vec<BeginBlock>,
) -> Result<bool, FrmError> {
    let keyword = match toks.first() {
        Some(Token::Word(kw)) => kw.to_ascii_uppercase(),
        // Non-keyword line — try as simple property.
        _ => {
            if let Some(prop) = parse_simple_property(ln, toks)? {
                properties.push(prop);
            }
            return Ok(false);
        }
    };

    match keyword.as_str() {
        "END" => return Ok(true),
        "BEGIN" => {
            // Nested control
            children.push(parse_begin_block(p, ln, toks)?);
        }
        "BEGINPROPERTY" => {
            properties.push(parse_begin_property(p, ln, toks)?);
        }
        "ATTRIBUTE" => {
            // Attribute lines can appear inside Begin blocks too. Store as a
            // simple property with name "Attribute.<attr.name>".
            let attr = parse_attribute_line(ln, toks)?;
            properties.push(Property {
                name: format!("Attribute.{}", attr.name),
                kind: PropKind::Simple,
                value: PropValue::Scalar(attr.value),
            });
        }
        // Simple property: PropName = value
        _ => {
            if let Some(prop) = parse_simple_property(ln, toks)? {
                properties.push(prop);
            }
        }
    }
    Ok(false)
}

/// Parse a `BeginProperty ... EndProperty` block.
fn parse_begin_property<'src>(
    p: &mut Parser<'src>,
    begin_ln: usize,
    begin_toks: &[Token<'src>],
) -> Result<Property, FrmError> {
    // begin_toks: [Word("BeginProperty"), Word(name), optional Curly(guid)]
    let name = match begin_toks.get(1) {
        Some(Token::Word(s)) => s.to_string(),
        _ => return Err(p.err(begin_ln, "BeginProperty requires property name")),
    };
    let guid = begin_toks.get(2).and_then(|t| match t {
        Token::Curly(s) => Some(s.to_string()),
        _ => None,
    });

    let mut sub: Vec<Property> = Vec::new();

    loop {
        let (ln, toks) = match p.next_tokens() {
            Some(v) => v,
            None => return Err(p.err(begin_ln, "unexpected EOF inside BeginProperty")),
        };

        match toks.first() {
            Some(Token::Word(kw)) if kw.eq_ignore_ascii_case("EndProperty") => break,
            Some(Token::Word(kw)) if kw.eq_ignore_ascii_case("BeginProperty") => {
                // Nested BeginProperty
                let nested = parse_begin_property(p, ln, &toks)?;
                sub.push(nested);
            }
            _ => {
                if let Some(prop) = parse_simple_property(ln, &toks)? {
                    sub.push(prop);
                }
            }
        }
    }

    Ok(Property { name, kind: PropKind::BeginProperty { guid }, value: PropValue::Bag(sub) })
}

/// Parse a `PropName = value` line.  Returns `None` for lines that don't
/// match.
fn parse_simple_property<'src>(
    ln: usize,
    toks: &[Token<'src>],
) -> Result<Option<Property>, FrmError> {
    // Need at least [Word(name), Equals, <value>]
    let name = match toks.first() {
        Some(Token::Word(n)) => n.to_string(),
        _ => return Ok(None),
    };

    if !matches!(toks.get(1), Some(Token::Equals)) {
        return Ok(None);
    }

    // Collect everything after '=' as the scalar value string
    let value = tokens_to_value_string(&toks[2..], ln);
    Ok(Some(Property { name, kind: PropKind::Simple, value: PropValue::Scalar(value) }))
}

/// Reconstitute a value from token slices, preserving the original text
/// as closely as possible (VB6 IDE round-trips matter for property values).
fn tokens_to_value_string(toks: &[Token<'_>], _ln: usize) -> String {
    let mut parts: Vec<String> = Vec::new();
    for tok in toks {
        match tok {
            Token::Word(s) => parts.push(s.to_string()),
            Token::Quoted(s) => parts.push(format!("\"{}\"", s)),
            Token::DollarStr(s) => parts.push(format!("$\"{}\"", s)),
            Token::Curly(s) => parts.push(format!("{{{}}}", s)),
            Token::Equals => parts.push("=".into()),
            Token::Semi => parts.push(";".into()),
            Token::Colon => parts.push(":".into()),
        }
    }
    parts.join(" ")
}

// ---------------------------------------------------------------------------
// Tests
// ---------------------------------------------------------------------------

#[cfg(test)]
mod tests {
    use super::*;

    const MINIMAL_FRM: &str = r#"VERSION 5.00
Begin VB.Form Form1
   Caption         =   "Hello"
   ClientHeight    =   3600
   BackColor       =   &H80000005&
End
"#;

    #[test]
    fn parse_minimal_form() {
        let f = parse_frm(MINIMAL_FRM).unwrap();
        assert_eq!(f.version, "5.00");
        assert_eq!(f.kind, FileKind::Form);
        let root = f.root.unwrap();
        assert_eq!(root.control_type, "VB.Form");
        assert_eq!(root.name, "Form1");
        assert_eq!(root.properties.len(), 3);
        assert_eq!(root.properties[0].name, "Caption");
        match &root.properties[0].value {
            PropValue::Scalar(s) => assert_eq!(s, "\"Hello\""),
            _ => panic!("expected scalar"),
        }
    }

    #[test]
    fn parse_nested_controls() {
        let src = r#"VERSION 5.00
Begin VB.Form Form1
   Caption = "Test"
   Begin VB.CommandButton cmdOK
      Caption = "OK"
      Height  = 375
   End
End
"#;
        let f = parse_frm(src).unwrap();
        let root = f.root.unwrap();
        assert_eq!(root.children.len(), 1);
        assert_eq!(root.children[0].name, "cmdOK");
    }

    #[test]
    fn parse_object_ref() {
        // In .frm files the progid is in outer double-quotes
        let src = "VERSION 5.00\r\nObject = \"{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0\"; \"MSCOMCTL.OCX\"\r\nBegin VB.Form Form1\r\nEnd\r\n";
        let f = parse_frm(src).unwrap();
        assert_eq!(f.objects.len(), 1);
        assert!(f.objects[0].progid.contains("831FDD16"));
        assert_eq!(f.objects[0].filename, Some("MSCOMCTL.OCX".into()));
    }

    #[test]
    fn parse_begin_property() {
        let src = r#"VERSION 5.00
Begin VB.Form Form1
   BeginProperty Font
      Name = "MS Sans Serif"
      Size = 8.25
   EndProperty
End
"#;
        let f = parse_frm(src).unwrap();
        let root = f.root.unwrap();
        assert_eq!(root.properties.len(), 1);
        assert_eq!(root.properties[0].name, "Font");
        assert_eq!(root.properties[0].kind, PropKind::BeginProperty { guid: None });
        match &root.properties[0].value {
            PropValue::Bag(sub) => {
                assert_eq!(sub.len(), 2);
                assert_eq!(sub[0].name, "Name");
            }
            _ => panic!("expected bag"),
        }
    }

    #[test]
    fn parse_class_file() {
        let src = r#"VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
END
Attribute VB_Name = "Class1"
"#;
        let f = parse_frm(src).unwrap();
        assert_eq!(f.kind, FileKind::Class);
        assert_eq!(f.attributes.len(), 1);
        assert_eq!(f.attributes[0].name, "VB_Name");
        assert_eq!(f.attributes[0].value, "Class1");
    }
}
