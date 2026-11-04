//! VB6 recursive-descent parser.
//!
//! Implements the source-to-AST stage for Visual Basic 6.
//!
//! **Representation choices:**
//! * Global error-flag + continue → `Diagnostics` collector.
//! * Process-global scanner state → owned `Scanner` threaded via `Parser`.
//! * List chain nodes → `Vec<NodeId>` in typed `ArgList`/`Block` variants.

use crate::frontend::cond_compile::{CcOp, CcValue};
use crate::frontend::ast::{
    AstLit, BinOpKind, DoKind, ExitKind, ExprArena, ExprNode, FileIoKind, LabelRef, NodeId,
    NodeSpans, OnErrorKind, ProcKind, ResumeTarget, UnOpKind,
};
use crate::frontend::diagnostics::*;
use crate::frontend::scanner::{Scanner, ScannerContext};
use crate::frontend::token::{Kw, Lit, Span, Token, TokenKind, TypeSuffix};


// ── Binding-power table (Pratt) ───────────────────────────────────────────────

/// Infix binding powers for VB6 binary operators.
///
/// Returns `(left_bp, right_bp, BinOpKind)` or `None`.
///
/// Precedence levels are: `Imp`=1, `Eqv`=2, `Xor`=3, `Or`=4, `And`=5,
/// comparison/`Like`/`Is`=6, `&`=7, `+`/`-`=8, `Mod`=9, `\`=10, `*`/`/`=11,
/// `^`=12.  Member access (`.`/`!`) is not an operator: it is postfix, and
/// only when no blank precedes it (see `parse_postfix`).
///
/// Each `(left_bp, right_bp)` pair below encodes level `N` as
/// `(2N, 2N+1)` (left-associative) so the relative ordering matches VB6
/// exactly. `^` is left-associative too: `2 ^ 3 ^ 2` is `(2 ^ 3) ^ 2`.
/// The operand of a unary plus takes operators of level 9 (`Mod`) and above, whose left binding
/// power is 18 or more.
const PLUS_OPERAND_MIN_BP: u8 = 17;

fn infix_bp(kind: &TokenKind) -> Option<(u8, u8, BinOpKind)> {
    use BinOpKind as B;
    let (l, r, o) = match kind {
        // ^ exponentiation, level 12, left-associative
        TokenKind::Kw(Kw::Caret)     => (25, 26, B::Pow),
        // * / multiplicative, level 11
        TokenKind::Kw(Kw::Star)      => (22, 23, B::Mul),
        TokenKind::Kw(Kw::Slash)     => (22, 23, B::Div),
        // \ integer division, level 10
        TokenKind::Kw(Kw::Backslash) => (20, 21, B::IDiv),
        // Mod, level 9
        TokenKind::Kw(Kw::Mod)       => (18, 19, B::Mod),
        // + - additive, level 8
        TokenKind::Kw(Kw::Plus)      => (16, 17, B::Add),
        TokenKind::Kw(Kw::Minus)     => (16, 17, B::Sub),
        // & string concat, level 7
        TokenKind::Kw(Kw::Amp)       => (14, 15, B::Cat),
        // comparison / Like / Is, level 6 (all same precedence)
        TokenKind::Kw(Kw::Eq)        => (12, 13, B::Eq),
        TokenKind::Kw(Kw::Ne)        => (12, 13, B::Ne),
        TokenKind::Kw(Kw::Lt)        => (12, 13, B::Lt),
        TokenKind::Kw(Kw::Gt)        => (12, 13, B::Gt),
        TokenKind::Kw(Kw::Le)        => (12, 13, B::Le),
        TokenKind::Kw(Kw::Ge)        => (12, 13, B::Ge),
        TokenKind::Kw(Kw::Like)      => (12, 13, B::Like),
        TokenKind::Kw(Kw::Is)        => (12, 13, B::Is),
        // And, level 5
        TokenKind::Kw(Kw::And)       => (10, 11, B::And),
        // Or, level 4
        TokenKind::Kw(Kw::Or)        => (8,  9,  B::Or),
        // Xor, level 3
        TokenKind::Kw(Kw::Xor)       => (6,  7,  B::Xor),
        // Eqv, level 2
        TokenKind::Kw(Kw::Eqv)       => (4,  5,  B::Eqv),
        // Imp, level 1
        TokenKind::Kw(Kw::Imp)       => (2,  3,  B::Imp),
        _ => return None,
    };
    Some((l, r, o))
}

/// Prefix binding power for VB6 unary operators.
///
/// Unary `-`/`+` bind below `^` (so `-2^2` == `-(2^2)`) but above `*`/`/`.
/// Unary `Not` is a logical operator and binds below comparison,
/// so `Not a = b` parses as `Not (a = b)`.
fn prefix_bp(kind: &TokenKind) -> Option<(u8, UnOpKind)> {
    use Kw::*;
    match kind {
        TokenKind::Kw(Minus) => Some((24, UnOpKind::Neg)),
        TokenKind::Kw(Not)   => Some((11, UnOpKind::Not)),
        _ => None,
    }
}


/// If `k` is a `Def<Type>` keyword (`DefBool`..`DefVar`, token ids 0x31–0x3c),
/// return its token id; otherwise `None`.  Used by module-level dispatch.
/// Whether `k` begins an executable statement (one that is only valid inside a
/// procedure).
fn starts_executable_statement(k: Kw) -> bool {
    use Kw::*;
    matches!(
        k,
        If | For | Do | While | Select | With | Let | Set | Call | GoTo | GoSub | Go | Exit | ReDim
            | Return | Resume | On | Error | Stop | Erase | RaiseEvent | Debug | Open | Close
            | Print | Write | Input | Line | Get | Put | Seek | Lock | Unlock | Width | Name
            | Mid | MidS | MidB | MidBS | LSet | RSet | Question | InputS | InputB | InputBS
    )
}

fn def_type_kw(k: Kw) -> Option<u16> {
    use Kw::*;
    matches!(
        k,
        DefBool | DefByte | DefCur | DefDate | DefDec | DefDbl
            | DefInt | DefLng | DefObj | DefSng | DefStr | DefVar
    )
    .then_some(k as u16)
}

/// Build the [`AstLit`] carried by a literal token. The token's `kind` selects
/// the variant; a missing/mismatched `lit` payload falls back to a zero value
/// (preserving the parser's lenient literal handling).
fn ast_lit_from_token(tok: &Token) -> AstLit {
    // Each arm reads the matching payload (defaulting a missing/mismatched one to
    // zero) via a tiny extractor, so the dispatch itself stays a flat table.
    match tok.kind {
        TokenKind::IntLit  => AstLit::Int(lit_i32(&tok.lit)),
        TokenKind::LongLit => AstLit::Long(lit_long(&tok.lit)),
        TokenKind::SngLit  => AstLit::Single(lit_single(&tok.lit)),
        TokenKind::DblLit  => AstLit::Double(lit_double(&tok.lit)),
        TokenKind::CurLit  => AstLit::Currency(lit_currency(&tok.lit)),
        TokenKind::StrLit  => AstLit::Str(lit_str(&tok.lit)),
        TokenKind::DateLit => AstLit::Date(lit_double(&tok.lit)),
        _ => AstLit::Int(0),
    }
}

/// Line-label value carried by a numeric line label / line target.
///
/// A line number can lex as either `Int` (≤ i16) or `Long`, so both are
/// accepted; the value is preserved exactly rather than folded to a sentinel.
fn line_label_value(lit: &Option<Lit>) -> i32 {
    match lit {
        Some(Lit::Int(n)) | Some(Lit::Long(n)) => *n,
        _ => 0,
    }
}

fn lit_i32(lit: &Option<Lit>) -> i32 { if let Some(Lit::Int(n)) = lit { *n } else { 0 } }
fn lit_long(lit: &Option<Lit>) -> i32 { if let Some(Lit::Long(n)) = lit { *n } else { 0 } }
fn lit_single(lit: &Option<Lit>) -> f32 { if let Some(Lit::Single(f)) = lit { *f } else { 0.0 } }
fn lit_double(lit: &Option<Lit>) -> f64 {
    match lit {
        Some(Lit::Double(f)) | Some(Lit::Date(f)) => *f,
        _ => 0.0,
    }
}
fn lit_currency(lit: &Option<Lit>) -> i64 { if let Some(Lit::Currency(c)) = lit { *c } else { 0 } }
fn lit_str(lit: &Option<Lit>) -> Box<str> {
    if let Some(Lit::Str(b)) = lit { b.clone() } else { Box::from("") }
}

// ── Module kind ───────────────────────────────────────────────────────────────

/// The kind of VBA module being parsed.
///
/// Used to enforce class-module-only syntax (`WithEvents`, `Implements`).
#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub enum ModuleKind {
    Standard,
    Class,
    Form,
}

// ── Parser struct ─────────────────────────────────────────────────────────────

/// VB6 recursive-descent parser.
///
/// Wraps a [`Scanner`] with a 1-token lookahead buffer.  Produces AST nodes
/// in a caller-supplied [`ExprArena`].  Syntax errors are pushed to an internal
/// [`Diagnostics`] collector; parsing continues after each error.
pub struct Parser<'src> {
    scanner:     Scanner<'src>,
    /// The source text being parsed.
    src:         &'src [u8],
    peeked:      Option<Token>,
    /// Second lookahead slot (for `peek2`), filled lazily after `peeked`.
    peeked2:     Option<Token>,
    /// Token read past a `_` to check whether it ended the line; returned by
    /// the next [`Parser::next_logical`] call.
    after_line_cont: Option<Token>,
    pub diagnostics: Diagnostics,
    /// Source-span side table for symbol-naming nodes (`NameRef` uses and
    /// declaration-name identifiers).  Consumed by the binder and LSP layer.
    pub node_spans: NodeSpans,
    /// Explicit visibility per declaration node (`NodeId.0` → is_public), set
    /// when an access modifier (`Public`/`Private`/`Friend`/`Global`) is present.
    /// Declarations with no modifier are absent; the binder applies VB6 defaults.
    pub decl_public: std::collections::HashMap<u32, bool>,
    /// Declaration nodes written with `Friend`: public inside the project
    /// (they are also in `decl_public` as public) but not part of the class's
    /// interface to other projects.
    pub decl_friend: std::collections::HashSet<u32>,
    /// Procedures declared `Static` (`Static Sub`, `Private Static Function`):
    /// all of their locals are static.
    pub decl_static: std::collections::HashSet<u32>,
    /// Module kind — controls accept/reject for class-only syntax.
    pub module_kind: ModuleKind,
    /// True while a procedure body is being parsed.
    in_procedure: bool,
    /// True while the member of a `Type` block is being parsed: a member may
    /// be named with a reserved word.
    reserved_names_allowed: bool,
    /// The declaration item parsed last was given a type with `As`.
    last_item_had_as: bool,
    /// An `AddressOf` is allowed where the next primary starts: the start of a call argument.
    address_of_ok: bool,
    /// The statements of a single-line `If` are being parsed: a `Rem` comment may follow one.
    in_inline_if: bool,
    /// A statement's target is being parsed, where built-in functions are plain names.
    in_target: bool,
    /// The target about to be parsed is at the start of a statement, where a few reserved
    /// words (`Date`, `DoEvents`, `CDec`) may name what is assigned or called.
    target_at_statement_start: bool,
    /// Open block statements of the procedure being parsed, innermost last.
    blocks: Vec<Block>,
    /// The last token consumed was the end of a line.
    last_was_eol: bool,
    /// A `:` statement separator has been consumed on the current line.
    after_colon: bool,
    /// A module-level variable, procedure or `Declare` has been parsed.
    module_decl_seen: bool,
    /// A module-level array variable has been declared.
    module_array_seen: bool,
    /// Byte offset just past the most recently consumed token.
    last_token_end: u32,
    /// `Option Base`, `Option Compare`, `Option Explicit`, `Option Private`
    /// already given.
    seen_option_base: bool,
    seen_option_compare: bool,
    seen_option_explicit: bool,
    seen_option_private: bool,
    /// Letters (bit 0 = A) already given a type by a `DefType` statement.
    deftype_letters: u32,
    /// Labels defined, and label references made, in the procedure being parsed.
    labels_defined: Vec<LabelRef>,
    label_refs: Vec<(LabelRef, Span)>,
    /// Kind of the procedure being parsed.
    proc_kind: Option<ProcKind>,
    /// An unclosed block has been reported in the current procedure; VB6
    /// reports only the innermost one.
    block_error_reported: bool,
    /// Control-variable names of a `Next a, b, …` still to be matched by the
    /// enclosing `For` blocks, with the span of each name.
    next_carry: Vec<(Option<u32>, Span)>,
    /// A variable named `Error` is declared at module level / in the current
    /// procedure (parameters included). It shadows the `Error` statement, so a
    /// line that starts with `Error` is then an ordinary target expression.
    error_var_module: bool,
    error_var_local: bool,
    /// Conditional-compilation constant table: `(lowercase_name, value)`.
    /// Pre-seeded with VB6 predefined constants.
    cc_defs: Vec<(String, CcValue)>,
    /// Open `#If` blocks, innermost last.
    cc_stack: Vec<CcFrame>,
    /// One-token pushback for directive processing.
    cc_pending: Option<Token>,
}

/// An open block-structured statement. The parser keeps these on a stack, the
/// way VB6's compiler does: a terminator is valid only when it matches the
/// innermost open block.
#[derive(Clone, Copy, PartialEq, Eq)]
enum Block {
    /// `If … Then` before any `Else`.
    If,
    /// `If … Then` after its `Else`: only `End If` may follow.
    IfElse,
    /// `For` / `For Each`, with the control variable's symbol when it is a plain name.
    For(Option<u32>),
    /// `Do`; `pre_test` is true for `Do While` / `Do Until`.
    Do { pre_test: bool },
    While,
    With,
    Select,
}

/// A keyword (or keyword pair) that closes or continues a block.
#[derive(Clone, Copy, PartialEq, Eq)]
enum Term { Else, ElseIf, EndIf, Next, Loop, Wend, Case, EndWith, EndSelect }

/// Whether `term` ends the body of the open block `top`.
fn block_closed_by(top: Block, term: Term) -> bool {
    match top {
        Block::If => matches!(term, Term::Else | Term::ElseIf | Term::EndIf),
        Block::IfElse => term == Term::EndIf,
        Block::For(_) => term == Term::Next,
        Block::Do { .. } => term == Term::Loop,
        Block::While => term == Term::Wend,
        Block::With => term == Term::EndWith,
        Block::Select => matches!(term, Term::Case | Term::EndSelect),
    }
}

/// Which kind of declaration owns a parameter list (the rules differ).
#[derive(Clone, Copy, PartialEq, Eq)]
enum ParamOwner {
    Procedure,
    /// `Property Let` / `Property Set`: the last parameter is the value.
    PropertyValue,
    Event,
}

/// State of one open `#If … #End If` block.
#[derive(Clone, Copy)]
struct CcFrame {
    /// Where the `#If` is.
    open: Span,
    /// An `#Else` has been seen: only `#End If` may follow.
    else_seen: bool,
    /// The enclosing region is active.
    parent_active: bool,
    /// The current branch is active.
    active: bool,
    /// Some branch of this block has already been taken.
    taken: bool,
}

impl<'src> Parser<'src> {
    /// Construct a parser over `src` using the provided scanner context.
    pub fn new(ctx: &'src mut ScannerContext, src: &'src [u8]) -> Self {
        Self {
            scanner:     Scanner::new(ctx, src),
            src,
            peeked:      None,
            peeked2:     None,
            after_line_cont: None,
            diagnostics: Diagnostics::new(),
            node_spans:  NodeSpans::new(),
            decl_public: std::collections::HashMap::new(),
            decl_friend: std::collections::HashSet::new(),
            decl_static: std::collections::HashSet::new(),
            in_procedure: false,
            reserved_names_allowed: false,
            last_item_had_as: false,
            address_of_ok: false,
            in_inline_if: false,
            in_target: false,
            target_at_statement_start: false,
            blocks: Vec::new(),
            last_was_eol: false,
            after_colon: false,
            module_decl_seen: false,
            module_array_seen: false,
            last_token_end: 0,
            seen_option_base: false,
            seen_option_compare: false,
            seen_option_explicit: false,
            seen_option_private: false,
            deftype_letters: 0,
            labels_defined: Vec::new(),
            label_refs: Vec::new(),
            proc_kind: None,
            block_error_reported: false,
            next_carry: Vec::new(),
            error_var_module: false,
            error_var_local: false,
            module_kind: ModuleKind::Standard,
            cc_stack: Vec::new(),
            cc_pending: None,
            cc_defs: vec![
                ("win32".into(), CcValue::TRUE),
                ("win64".into(), CcValue::FALSE),
                ("win16".into(), CcValue::FALSE),
                ("vba6".into(), CcValue::TRUE),
                ("vba7".into(), CcValue::FALSE),
            ],
        }
    }

    /// Construct a parser with a specific module kind.
    pub fn with_module_kind(ctx: &'src mut ScannerContext, src: &'src [u8], kind: ModuleKind) -> Self {
        let mut p = Self::new(ctx, src);
        p.module_kind = kind;
        p
    }

    // ── Internal token helpers ────────────────────────────────────────────────

    fn peek(&mut self) -> &Token {
        if self.peeked.is_none() {
            self.peeked = Some(
                self.peeked2.take().unwrap_or_else(|| self.next_logical()),
            );
        }
        self.peeked.as_ref().unwrap()
    }

    /// Two-token lookahead.  Used for label detection (`name :`) and other
    /// cases where one token of lookahead is insufficient.
    fn peek2(&mut self) -> &Token {
        self.peek(); // ensure `peeked` is filled first
        if self.peeked2.is_none() {
            self.peeked2 = Some(self.next_logical());
        }
        self.peeked2.as_ref().unwrap()
    }

    fn advance(&mut self) -> Token {
        let tok = match self.peeked.take() {
            Some(t) => t,
            None    => self.peeked2.take().unwrap_or_else(|| self.next_logical()),
        };
        self.last_was_eol = tok.kind == TokenKind::Eol;
        self.last_token_end = tok.span.start + tok.span.len;
        match tok.kind {
            TokenKind::Kw(Kw::Colon) => self.after_colon = true,
            TokenKind::Eol => self.after_colon = false,
            _ => {}
        }
        tok
    }

    fn current_span(&mut self) -> Span {
        self.peek().span
    }

    /// True when the next token follows the previous one after blanks on the
    /// same line. A line continuation between them does not count: the
    /// continued line joins the statement as if no blanks were there.
    fn blank_before_next(&mut self) -> bool {
        let start = self.peek().span.start as usize;
        let end = self.last_token_end as usize;
        start > end
            && start <= self.src.len()
            && !self.src[end..start].iter().any(|b| *b == b'\n' || *b == b'\r')
    }

    /// Record the source span of a symbol-naming node — the identifier span of
    /// a `NameRef` use site, or the declared-name identifier of a declaration.
    fn set_span(&mut self, id: NodeId, span: Span) {
        self.node_spans.set(id, span);
    }

    /// Consume the next token if its kind matches; return whether consumed.
    fn eat(&mut self, expected: TokenKind) -> bool {
        if self.peek().kind == expected {
            self.advance();
            true
        } else {
            false
        }
    }

    /// The next token of the logical line. A line continuation (`_` followed
    /// by the end of the physical line) joins two physical lines and is
    /// whitespace to the grammar, so the `_` and its end-of-line are dropped
    /// here, wherever they occur. A `_` that is not followed by an end of line
    /// is returned as a `LineCont` token for the grammar to reject.
    fn next_logical(&mut self) -> Token {
        loop {
            let tok = self.next_physical();
            if let TokenKind::Kw(kw @ (Kw::CcIf | Kw::CcElseIf | Kw::CcElse | Kw::CcEnd | Kw::CcConst)) =
                tok.kind
            {
                self.cc_directive(kw, tok.span, tok.lit.is_some());
                continue;
            }
            if tok.kind == TokenKind::Eof {
                self.report_open_cc_blocks();
            }
            if !self.cc_active() {
                if tok.kind == TokenKind::Eof {
                    return tok;
                }
                continue;
            }
            if tok.kind == TokenKind::Ident {
                let too_long = tok
                    .sym
                    .is_some_and(|id| self.scanner.sym_name(id as u32).len() > MAX_IDENTIFIER_LEN);
                if too_long {
                    self.diagnostics.push(ERR_IDENTIFIER_TOO_LONG, tok.span);
                }
            }
            return tok;
        }
    }

    /// The next token of the logical line, before conditional compilation.
    fn next_physical(&mut self) -> Token {
        if let Some(t) = self.cc_pending.take() {
            return t;
        }
        loop {
            let tok = self.after_line_cont.take().unwrap_or_else(|| self.scanner.next_token());
            if tok.kind == TokenKind::Kw(Kw::LineCont) {
                let next = self.scanner.next_token();
                if next.kind == TokenKind::Eol {
                    continue;
                }
                self.after_line_cont = Some(next);
            }
            return tok;
        }
    }

    /// Consume and return the next token if it matches `expected`; otherwise
    /// push diagnostic `code` (with an auto-derived label for `0x9c6f` errors)
    /// and return a synthetic placeholder token without consuming.
    fn expect(&mut self, expected: TokenKind, code: u32) -> Token {
        if self.peek().kind == expected {
            self.advance()
        } else {
            let span = self.current_span();
            if code == 0x9c6f {
                let label = token_kind_label(&expected);
                self.diagnostics.push_labeled(code, span, label);
            } else {
                self.diagnostics.push(code, span);
            }
            // Return a synthetic placeholder so callers can continue.
            Token { kind: expected, sym: None, lit: None, type_suffix: TypeSuffix::None, span }
        }
    }

    fn expect_ident(&mut self) -> Token {
        // Accept plain identifiers AND keywords that were scanned from identifier-like
        // text (they have sym: Some(_)).  In VBA, keywords are valid member names,
        // parameter names, and variable names in many contexts.
        let accept = {
            let t = self.peek();
            t.kind == TokenKind::Ident
                || (matches!(t.kind, TokenKind::Kw(_)) && t.sym.is_some())
        };
        if accept {
            self.advance()
        } else {
            let span = self.current_span();
            self.diagnostics.push_labeled(ERR_EXPECTED_IDENT, span, "Identifier");
            Token { kind: TokenKind::Ident, sym: None, lit: None, type_suffix: TypeSuffix::None, span }
        }
    }

    /// A name being declared. A reserved word cannot name anything, apart from
    /// the words that only mean something in a module header (`Attribute`,
    /// `VB_Name`, ...), which stay free as names.
    fn expect_decl_name(&mut self) -> Token {
        // A `Type` member may take any reserved word but `Me`.
        let me = self.peek().kind == TokenKind::Kw(Kw::Me) && self.peek2().kind != TokenKind::Kw(Kw::LParen);
        if (!self.reserved_names_allowed || me) && self.at_reserved_word() {
            return self.reject_reserved_word();
        }
        self.expect_ident()
    }

    /// True when the next token is a reserved word, which cannot be used as a name.
    fn at_reserved_word(&mut self) -> bool {
        let t = self.peek();
        match (&t.kind, t.sym) {
            (TokenKind::Kw(_), Some(sym)) => {
                let sym = sym as u32;
                self.scanner.is_reserved(sym) && !is_header_word(self.scanner.sym_name(sym))
            }
            _ => false,
        }
    }

    /// Reports a missing or unexpected `what` at `span`: inside a procedure body the
    /// compiler gives only the general syntax error, elsewhere it names what it expected.
    fn report_expected(&mut self, span: Span, what: &'static str) {
        if self.in_procedure {
            self.diagnostics.push(ERR_SYNTAX_ERROR, span);
        } else {
            self.diagnostics.push_labeled(0x9c6f, span, what);
        }
    }

    /// The name of a `Type`, `Enum` or `Event` carries no type-declaration character.
    fn reject_name_suffix(&mut self, name: &Token) {
        if name.type_suffix != TypeSuffix::None {
            self.report_expected(name.span, "identifier");
        }
    }

    /// A name that carries a type-declaration character cannot also be given a type with `As`.
    fn reject_suffix_with_as(&mut self, name: &Token, what: &'static str) {
        if name.type_suffix != TypeSuffix::None && self.peek().kind == TokenKind::Kw(Kw::As) {
            let span = self.current_span();
            self.report_expected(span, what);
        }
    }

    /// True when the next token is a keyword word (not punctuation or an
    /// operator) that can begin a statement made of a name: a keyword that is
    /// not reserved, or one of the few reserved words that are also names of
    /// callable things.
    fn starts_name_statement(&mut self) -> bool {
        let t = self.peek();
        let Some(sym) = t.sym else { return false };
        let name = self.scanner.sym_name(sym as u32);
        if !name.starts_with(|c: char| c.is_ascii_alphabetic()) {
            return false;
        }
        if ["CDec", "Date", "DoEvents"].iter().any(|n| name.eq_ignore_ascii_case(n)) {
            return true;
        }
        !self.at_reserved_word()
    }

    /// True when the next token is punctuation or an operator that is not the
    /// member dot or bang.
    fn at_stray_punctuation(&mut self) -> bool {
        let TokenKind::Kw(k) = self.peek().kind.clone() else { return false };
        if matches!(k, Kw::Dot | Kw::Bang) {
            return false;
        }
        match super::keyword_table::KEYWORD_TABLE.get(k.token_id() as usize) {
            Some(entry) => !entry.name.starts_with(|c: char| c.is_ascii_alphabetic()),
            None => false,
        }
    }

    /// True when the next token is `Me` or one of the reserved words that also name
    /// callable things (`Date`, `DoEvents`, `CDec`).
    fn at_callable_reserved_word(&mut self) -> bool {
        let Some(sym) = self.peek().sym else { return false };
        let name = self.scanner.sym_name(sym as u32);
        ["Me", "Date", "DoEvents", "CDec"].iter().any(|n| name.eq_ignore_ascii_case(n))
    }

    /// True when the next token is a keyword that begins a statement of its own
    /// and also names things (`Name`, `Width`, `Line`, ...).
    fn at_statement_word(&mut self) -> bool {
        let Some(sym) = self.peek().sym else { return false };
        let name = self.scanner.sym_name(sym as u32);
        STATEMENT_WORDS.iter().any(|n| name.eq_ignore_ascii_case(n))
    }

    /// True when the next keyword can name an `Enum` member.
    fn names_enum_member(&mut self) -> bool {
        let Some(sym) = self.peek().sym else { return false };
        let name = self.scanner.sym_name(sym as u32).to_string();
        if ["Date", "DoEvents", "CDec"].iter().any(|n| name.eq_ignore_ascii_case(n)) {
            return true;
        }
        let valued = self.peek2().kind == TokenKind::Kw(Kw::Eq);
        self.at_contextual_keyword()
            && (valued || !STATEMENT_WORDS.iter().any(|n| name.eq_ignore_ascii_case(n)))
    }

    /// The call of a built-in function whose arguments have a grammar of their own:
    /// brackets (unless the function may stand bare), a fixed number of arguments,
    /// none named or left out, and nothing joined to the result.
    fn parse_shaped_intrinsic_call(
        &mut self,
        arena: &mut ExprArena,
        base: NodeId,
        span: Span,
        shape: &IntrinsicShape,
    ) -> NodeId {
        if self.peek().kind != TokenKind::Kw(Kw::LParen) {
            if !shape.bare_ok {
                self.report_expected(span, "(");
            }
            return base;
        }
        self.advance(); // (
        let mut args = Vec::new();
        let mut bad = false;
        if self.peek().kind != TokenKind::Kw(Kw::RParen) {
            loop {
                if matches!(self.peek().kind, TokenKind::Kw(Kw::Comma) | TokenKind::Kw(Kw::RParen)) {
                    bad = true; // an argument left out
                    break;
                }
                if let Some(named) = self.try_parse_named_arg(arena) {
                    bad = true; // a named argument
                    args.push(named);
                } else {
                    self.eat(TokenKind::Kw(Kw::ByVal));
                    args.push(self.parse_expr(arena, 0));
                }
                if !self.eat(TokenKind::Kw(Kw::Comma)) {
                    break;
                }
            }
        }
        self.expect(TokenKind::Kw(Kw::RParen), ERR_EXPECTED_RPAREN);
        if args.len() < shape.min || args.len() > shape.max {
            bad = true;
        }
        if shape.name_first {
            // The first argument names the array: a variable, not a value.
            bad |= args.first().is_some_and(|a| {
                !matches!(
                    arena.get(*a),
                    ExprNode::NameRef { .. } | ExprNode::MemberAccess { .. } | ExprNode::Call { .. } | ExprNode::Me
                )
            });
        }
        if bad {
            self.report_expected(span, "expression");
        }
        let args = arena.alloc(ExprNode::ArgList { args });
        arena.alloc(ExprNode::Call { func: base, args })
    }

    /// A keyword that begins a statement of its own, met where that statement's
    /// operands are missing. Some keywords also name things: such a keyword
    /// is then a name only when something is assigned to it (`Width = 1`) or one
    /// of its members is used (`Name.Length`). Anything else is the statement
    /// without its operands.
    fn keyword_as_name_stmt(&mut self, arena: &mut ExprArena, t: Token) -> NodeId {
        if matches!(
            self.peek().kind,
            TokenKind::Kw(Kw::Eq) | TokenKind::Kw(Kw::Dot) | TokenKind::Kw(Kw::Bang)
        ) {
            return self.parse_ident_stmt_with_tok(arena, t);
        }
        self.statement_without_operands(arena, t)
    }

    /// After `New`: a class name. A built-in type word is reported as a missing type name.
    fn require_class_name_after_new(&mut self) {
        if self.at_reserved_word() || self.peek().kind == TokenKind::Kw(Kw::Object) {
            let span = self.current_span();
            self.report_expected(span, "type name");
        }
    }

    /// `Scale` or `Line` followed by a coordinate pair: the points are two values
    /// in brackets, so a bracket holding anything else is not that statement.
    fn graphics_point_stmt(&mut self, arena: &mut ExprArena, t: Token) -> NodeId {
        let span = t.span;
        let node = self.parse_ident_stmt_with_tok(arena, t);
        let pair = match arena.get(node) {
            ExprNode::CallStmt { .. } => true,
            ExprNode::Call { args, .. } => {
                matches!(arena.get(*args), ExprNode::ArgList { args } if args.len() == 2)
            }
            _ => false,
        };
        if !pair {
            self.diagnostics.push(ERR_SYNTAX_ERROR, span);
        }
        node
    }

    /// Reports a statement keyword whose operands are missing, and skips the line.
    fn statement_without_operands(&mut self, arena: &mut ExprArena, t: Token) -> NodeId {
        self.diagnostics.push(ERR_SYNTAX_ERROR, t.span);
        self.skip_line();
        let sym = t.sym.map(|s| s as u32).unwrap_or(0);
        let node = arena.alloc(ExprNode::NameRef { sym, suffix: t.type_suffix });
        self.set_span(node, t.span);
        node
    }

    /// `Array(` : the one reserved built-in function that a target expression may call.
    fn at_array_call(&mut self) -> bool {
        let Some(sym) = self.peek().sym else { return false };
        self.scanner.sym_name(sym as u32).eq_ignore_ascii_case("Array")
            && self.peek2().kind == TokenKind::Kw(Kw::LParen)
    }

    /// A reserved word where a target belongs: reported, consumed and replaced by a placeholder.
    fn reject_reserved_word_target(&mut self, arena: &mut ExprArena) -> NodeId {
        self.reject_reserved_word();
        self.parse_primary_placeholder(arena)
    }

    /// Reports the reserved word at the next token where a name belongs and
    /// consumes it, returning it as if it were a name so parsing goes on. Inside a
    /// procedure the compiler gives only the general syntax error.
    fn reject_reserved_word(&mut self) -> Token {
        let span = self.current_span();
        if self.in_procedure {
            self.diagnostics.push(ERR_SYNTAX_ERROR, span);
        } else {
            self.diagnostics.push_labeled(ERR_EXPECTED_IDENT, span, "identifier");
        }
        let tok = self.advance();
        Token { kind: TokenKind::Ident, ..tok }
    }

    /// True if the next token is a contextual (non-reserved) keyword such as
    /// `Name`, `Text` or `Error`, which VB6 also accepts as a variable or
    /// type name.
    fn at_contextual_keyword(&mut self) -> bool {
        let t = self.peek();
        match (&t.kind, t.sym) {
            (TokenKind::Kw(_), Some(sym)) => {
                let sym = sym as u32;
                !self.scanner.is_reserved(sym)
            }
            _ => false,
        }
    }

    /// Skip tokens until we hit a statement end (Eol / Eof / Apos / Rem).
    fn skip_to_stmt_end(&mut self) {
        loop {
            if self.peek().is_stmt_end() { break; }
            self.advance();
        }
    }

    /// Skip to the end of the current line, consuming the terminator (including
    /// any trailing comment + its Eol).  Safe to call even when the line has an
    /// inline `'comment` that would otherwise leave `skip_to_stmt_end` stopped
    /// on an Apos token forever.
    fn skip_line(&mut self) {
        self.skip_to_stmt_end();
        if matches!(self.peek().kind, TokenKind::Kw(Kw::Apos) | TokenKind::Kw(Kw::Rem)) {
            self.advance(); // consume the comment token itself
        }
        self.eat(TokenKind::Eol);
    }

    /// Consume Eol/Eof (statement terminator).  For single-line `If` etc. the
    /// `:` colon is also accepted.
    fn eat_eol(&mut self) {
        if self.peek().is_stmt_end() {
            self.advance();
        }
    }

    /// End of a complete statement: the line's terminator, or the `:` that
    /// begins the next statement, or the `Else` that continues a single-line
    /// `If`. Anything else is an error: the rest of the line is skipped so
    /// parsing resumes at the next statement.
    fn end_stmt(&mut self) {
        if self.peek().is_stmt_end() || (self.in_inline_if && self.peek().kind == TokenKind::Kw(Kw::Rem)) {
            self.advance();
        } else if !matches!(
            self.peek().kind,
            TokenKind::Kw(Kw::Colon) | TokenKind::Kw(Kw::Else)
        ) {
            let span = self.current_span();
            self.diagnostics.push_labeled(0x9c6f, span, "end of statement");
            self.skip_line();
        }
    }

    /// Remember that a variable called `Error` is declared (see
    /// `error_var_module`).
    fn note_declared_name(&mut self, name: &Token, module_level: bool) {
        let is_error = name
            .sym
            .map_or(false, |s| self.scanner.sym_name(s as u32).eq_ignore_ascii_case("error"));
        if is_error {
            if module_level {
                self.error_var_module = true;
            } else {
                self.error_var_local = true;
            }
        }
    }

    /// True when a line starting with the `Error` keyword is a target
    /// expression (`Error = 5`, `Error.Foo`, `Error(i) = 5`) rather than the
    /// `Error <number>` statement.
    fn error_starts_target(&mut self) -> bool {
        match self.peek2().kind {
            TokenKind::Kw(Kw::Eq) | TokenKind::Kw(Kw::Dot) | TokenKind::Kw(Kw::Bang) => true,
            TokenKind::Kw(Kw::LParen) => self.error_var_module || self.error_var_local,
            _ => false,
        }
    }

    /// The block terminator at the current token, if any: `Else`, `ElseIf`,
    /// `EndIf`, `Next`, `Loop`, `Wend`, `Case`, `End If`, `End With`,
    /// `End Select`. A bare `End` (terminate the program) is not one.
    fn terminator(&mut self) -> Option<Term> {
        match self.peek().kind {
            TokenKind::Kw(Kw::Else) => Some(Term::Else),
            TokenKind::Kw(Kw::ElseIf) => Some(Term::ElseIf),
            TokenKind::Kw(Kw::EndIf) => Some(Term::EndIf),
            TokenKind::Kw(Kw::Next) => Some(Term::Next),
            TokenKind::Kw(Kw::Loop) => Some(Term::Loop),
            TokenKind::Kw(Kw::Wend) => Some(Term::Wend),
            TokenKind::Kw(Kw::Case) => Some(Term::Case),
            TokenKind::Kw(Kw::End) => match self.peek2().kind {
                TokenKind::Kw(Kw::If) => Some(Term::EndIf),
                TokenKind::Kw(Kw::With) => Some(Term::EndWith),
                TokenKind::Kw(Kw::Select) => Some(Term::EndSelect),
                _ => None,
            },
            _ => None,
        }
    }

    /// The current token ends the body of the innermost open block: its own
    /// terminator, the end of the procedure, or the end of the file.
    fn at_block_end(&mut self) -> bool {
        if self.peek().kind == TokenKind::Eof || self.at_proc_end() {
            return true;
        }
        let Some(top) = self.blocks.last().copied() else { return false };
        if matches!(top, Block::For(_)) && !self.next_carry.is_empty() {
            return true;
        }
        self.terminator().is_some_and(|t| block_closed_by(top, t))
    }

    /// Consume the current block terminator when it closes the innermost open
    /// block; otherwise report the block as unclosed with `unclosed`. Returns
    /// whether a terminator was consumed.
    fn close_block(&mut self, unclosed: u32) -> bool {
        let closes = match (self.blocks.last().copied(), self.terminator()) {
            (Some(top), Some(t)) => block_closed_by(top, t),
            _ => false,
        };
        if !closes {
            self.report_unclosed(unclosed);
            return false;
        }
        if self.peek().kind == TokenKind::Kw(Kw::End) {
            self.advance();
        }
        self.advance();
        true
    }

    /// Report an unclosed block at the `End Sub`/`End Function`/`End Property`
    /// that ends the procedure. VB6 reports only the innermost one, and at the
    /// end of the file it reports the missing `End` of the procedure instead.
    fn report_unclosed(&mut self, code: u32) {
        if self.peek().kind == TokenKind::Eof || self.block_error_reported {
            return;
        }
        self.block_error_reported = true;
        let span = self.current_span();
        self.diagnostics.push(code, span);
    }

    /// A block terminator with no matching open block: report it and skip the line.
    fn stray_terminator(&mut self) -> Option<NodeId> {
        let code = match self.terminator() {
            Some(Term::Else) | Some(Term::ElseIf) => ERR_ELSE_WITHOUT_IF,
            Some(Term::EndIf) => ERR_END_IF_WITHOUT_BLOCK_IF,
            Some(Term::Next) => ERR_NEXT_WITHOUT_FOR,
            Some(Term::Loop) => ERR_LOOP_WITHOUT_DO,
            Some(Term::Wend) => ERR_WEND_WITHOUT_WHILE,
            Some(Term::Case) if self.peek2().kind == TokenKind::Kw(Kw::Else) => {
                ERR_CASE_ELSE_OUTSIDE_SELECT
            }
            Some(Term::Case) => ERR_CASE_WITHOUT_SELECT,
            Some(Term::EndWith) => ERR_END_WITH_WITHOUT_WITH,
            Some(Term::EndSelect) => ERR_END_SELECT_WITHOUT_SELECT,
            None => ERR_UNEXPECTED_TOKEN,
        };
        let span = self.current_span();
        self.diagnostics.push(code, span);
        self.skip_line();
        None
    }

    /// The current token is the `End Sub` / `End Function` / `End Property`
    /// that ends a procedure.
    fn at_proc_end(&mut self) -> bool {
        if !matches!(self.peek().kind, TokenKind::Kw(Kw::End)) {
            return false;
        }
        matches!(
            self.peek2().kind,
            TokenKind::Kw(Kw::Sub)
                | TokenKind::Kw(Kw::Function)
                | TokenKind::Kw(Kw::Property)
        )
    }

    /// Parse a procedure body: statements up to the procedure's `End`.
    fn parse_block_proc(&mut self, arena: &mut ExprArena) -> NodeId {
        let was_in_procedure = std::mem::replace(&mut self.in_procedure, true);
        let mut stmts: Vec<NodeId> = Vec::new();
        loop {
            while self.eat(TokenKind::Eol) {}
            if self.at_proc_end() || self.peek().kind == TokenKind::Eof { break; }
            if let Some(n) = self.parse_stmt(arena) {
                stmts.push(n);
            }
        }
        self.in_procedure = was_in_procedure;
        arena.alloc(ExprNode::Block { stmts })
    }

    // ── Module-level parse ────────────────────────────────────────────────────

    /// Parse a complete module: a sequence of declarations and statement blocks.
    ///
    /// Returns a list of top-level node ids.
    pub fn parse_module(&mut self, arena: &mut ExprArena) -> Vec<NodeId> {
        let mut nodes = Vec::new();
        loop {
            self.skip_module_blank_lines();
            if self.peek().kind == TokenKind::Eof { break; }

            // `Else` or a line number in front of an attribute line does not change it.
            if self.at_prefixed_attribute_line() {
                self.advance();
                self.eat(TokenKind::Kw(Kw::Colon));
                let n = self.parse_attribute_stmt(arena);
                nodes.push(n);
                continue;
            }

            // A colon that starts the line cannot precede the declaration of a procedure,
            // a `Declare` or an `Event`.
            if self.after_colon && self.starts_procedure_like_declaration() {
                let span = self.current_span();
                self.diagnostics.push_labeled(0x9c6f, span, "statement or end of statement");
            }

            match self.peek().kind.clone() {
                // Procedure declarations
                TokenKind::Kw(Kw::Sub)      => {
                    let n = self.parse_proc_decl(arena, ProcKind::Sub);
                    nodes.push(n);
                }
                TokenKind::Kw(Kw::Function) => {
                    let n = self.parse_proc_decl(arena, ProcKind::Function);
                    nodes.push(n);
                }
                TokenKind::Kw(Kw::Property) => {
                    let n = self.parse_property_decl(arena);
                    nodes.push(n);
                }
                // Declaration modifiers that precede Sub/Function/Property/Declare
                // or a variable/constant list. `Global` is `Public`.
                TokenKind::Kw(Kw::Public)
                | TokenKind::Kw(Kw::Global)
                | TokenKind::Kw(Kw::Private)
                | TokenKind::Kw(Kw::Friend)
                | TokenKind::Kw(Kw::Static) => {
                    let n = self.parse_access_decl(arena);
                    nodes.push(n);
                }
                // Type / Enum declarations
                TokenKind::Kw(Kw::Type)  => {
                    let n = self.parse_type_decl(arena);
                    nodes.push(n);
                }
                TokenKind::Kw(Kw::Enum)  => {
                    let n = self.parse_enum_decl(arena);
                    nodes.push(n);
                }
                // External declarations
                TokenKind::Kw(Kw::Declare) => {
                    let n = self.parse_declare_stmt(arena);
                    nodes.push(n);
                }
                // Event declarations — class-module-only
                TokenKind::Kw(Kw::Event) => {
                    let n = self.parse_event_decl(arena);
                    nodes.push(n);
                }
                // Implements — class-module-only
                TokenKind::Kw(Kw::Implements) => {
                    let n = self.parse_implements_stmt(arena);
                    nodes.push(n);
                }
                // Option statements
                TokenKind::Kw(Kw::Option) => {
                    let n = self.parse_option_stmt(arena);
                    nodes.push(n);
                }
                // Module-level Dim / Const
                TokenKind::Kw(Kw::Dim)
                | TokenKind::Kw(Kw::Const) => {
                    if let Some(n) = self.parse_stmt(arena) {
                        nodes.push(n);
                    }
                }
                // Def<Type> default-type declarations (DefInt A-Z, …).
                TokenKind::Kw(k) if def_type_kw(k).is_some() => {
                    let n = self.parse_def_stmt(arena);
                    nodes.push(n);
                }
                // Attribute metadata lines (hosts embed these in module text).
                TokenKind::Kw(Kw::Attribute) => {
                    let n = self.parse_attribute_stmt(arena);
                    nodes.push(n);
                }
                // `BEGIN ... END` designer-metadata block and `VERSION x.y CLASS`
                // header line embedded in .cls/.frm files.  Neither `BEGIN` nor
                // `VERSION` is a VB6 keyword; both appear as Ident at module level.
                // Skip them without emitting diagnostics.
                TokenKind::Ident => {
                    let name = self.peek().sym
                        .map(|id| self.scanner.sym_name(id as u32).to_ascii_lowercase())
                        .unwrap_or_default();
                    if name == "begin" {
                        self.skip_begin_end_block();
                    } else if name == "version" {
                        // VERSION x.y CLASS / VERSION 5.00 — file-format header; skip entire line.
                        self.skip_line();
                    } else {
                        self.skip_executable_statement(arena);
                    }
                }
                TokenKind::Kw(k) if starts_executable_statement(k) => {
                    self.skip_executable_statement(arena);
                }
                TokenKind::IntLit | TokenKind::LongLit => self.skip_executable_statement(arena),
                // `End Type` with no `Type` block open.
                TokenKind::Kw(Kw::End) if self.peek2().kind == TokenKind::Kw(Kw::Type) => {
                    let span = self.current_span();
                    self.diagnostics.push(ERR_END_TYPE_OUTSIDE_TYPE, span);
                    self.skip_line();
                }
                _ => self.recover_module_unexpected(),
            }
        }
        nodes
    }

    /// True when the next tokens begin the declaration of a procedure, a `Declare` or an
    /// `Event`, with or without an access modifier before it.
    fn starts_procedure_like_declaration(&mut self) -> bool {
        let is_decl_word = |k: &TokenKind| {
            matches!(
                k,
                TokenKind::Kw(Kw::Sub)
                    | TokenKind::Kw(Kw::Function)
                    | TokenKind::Kw(Kw::Property)
                    | TokenKind::Kw(Kw::Declare)
                    | TokenKind::Kw(Kw::Event)
            )
        };
        let first = self.peek().kind.clone();
        if is_decl_word(&first) {
            return true;
        }
        matches!(
            first,
            TokenKind::Kw(Kw::Public) | TokenKind::Kw(Kw::Private) | TokenKind::Kw(Kw::Friend) | TokenKind::Kw(Kw::Global) | TokenKind::Kw(Kw::Static)
        ) && is_decl_word(&self.peek2().kind.clone())
    }

    /// An executable statement outside any procedure: report it once, at its
    /// start, and consume the whole statement (a block statement included).
    fn skip_executable_statement(&mut self, arena: &mut ExprArena) {
        if self.line_has_trailing_attribute_line() {
            // The whole line is an attribute line.
            self.skip_line();
            return;
        }
        let span = self.current_span();
        let mark = self.diagnostics.len();
        let before = span.start;
        let _ = self.parse_stmt(arena);
        // The rest of the line (statements after `:` or a line number) goes too.
        if !self.last_was_eol {
            self.skip_line();
        }
        self.diagnostics.truncate(mark);
        self.diagnostics.push(ERR_INVALID_OUTSIDE_PROCEDURE, span);
        // Always make progress.
        if self.current_span().start == before && self.peek().kind != TokenKind::Eof {
            self.advance();
        }
    }

    /// Skip a `BEGIN ... END` designer-metadata block found in `.cls` files.
    ///
    /// VB6 writes this block into every class module; it holds IDE property
    /// values and is not executable code.  `BEGIN` is not a keyword — the
    /// caller has already confirmed the current token is the identifier "BEGIN".
    /// Nesting is tracked so that embedded `Begin ... End` sub-blocks (as in
    /// `.frm` form-designer sections) are also consumed correctly.
    fn skip_begin_end_block(&mut self) {
        self.advance(); // consume BEGIN identifier
        self.skip_line();

        let mut depth: u32 = 1;
        loop {
            if self.peek().kind == TokenKind::Eof { break; }
            // Check if this line opens a nested block (ident "Begin"/"BEGIN")
            // or closes one (`End` keyword — Kw::End).
            match self.peek().kind.clone() {
                TokenKind::Ident => {
                    let is_begin = self.peek().sym
                        .map(|id| self.scanner.sym_name(id as u32).eq_ignore_ascii_case("begin"))
                        .unwrap_or(false);
                    if is_begin { depth += 1; }
                    self.skip_line();
                }
                TokenKind::Kw(Kw::End) => {
                    self.advance(); // consume End
                    self.skip_line();
                    depth -= 1;
                    if depth == 0 { break; }
                }
                _ => {
                    self.skip_line();
                }
            }
        }
    }

    /// Skip blank lines / comment-only lines at module level. A comment is an
    /// `Apos`/`Rem` token followed by an `Eol`; consume both so the loop advances
    /// past it. These tokens are statement terminators, so `skip_to_stmt_end`
    /// (used by the module recovery path) stops *at* them without consuming —
    /// leaving a comment here would spin the loop forever.
    fn skip_module_blank_lines(&mut self) {
        loop {
            if self.eat(TokenKind::Eol) || self.eat(TokenKind::Kw(Kw::Colon)) { continue; }
            if matches!(self.peek().kind, TokenKind::Kw(Kw::Apos) | TokenKind::Kw(Kw::Rem)) {
                self.advance();
                self.eat(TokenKind::Eol);
                continue;
            }
            break;
        }
    }

    /// Recover from an unexpected token at module level — skip to next line.
    fn recover_module_unexpected(&mut self) {
        let span = self.current_span();
        self.diagnostics.push(ERR_UNEXPECTED_TOKEN, span);
        // `skip_to_stmt_end` halts *at* a statement terminator
        // (Eol/Eof/Apos/Rem) without consuming it. If the offending
        // token already is one (e.g. a stray comment), nothing is
        // consumed; force one token of progress so the loop can
        // never spin.
        let before = self.current_span().start;
        self.skip_to_stmt_end();
        if self.current_span().start == before && self.peek().kind != TokenKind::Eof {
            self.advance();
        }
    }

    // ── Statement block ───────────────────────────────────────────────────────

    /// Parses a sequence of statements terminated by a block-end keyword
    /// (`End`, `Else`, `ElseIf`, `Loop`, `Next`, `Wend`, `Case`) or EOF.
    /// Returns a `BLOCK` node wrapping the collected statements.
    pub fn parse_block(&mut self, arena: &mut ExprArena) -> NodeId {
        let mut stmts: Vec<NodeId> = Vec::new();
        loop {
            while self.eat(TokenKind::Eol) {}
            if self.at_block_end() || self.peek().kind == TokenKind::Eof { break; }
            if let Some(n) = self.parse_stmt(arena) {
                stmts.push(n);
            }
        }
        arena.alloc(ExprNode::Block { stmts })
    }

    // ── Statement dispatch helpers ────────────────────────────────────────────

    /// `Print`/`?`: `Print #n, …` writes to a file; bare `Print …` writes to the
    /// immediate window.
    fn parse_print_keyword_stmt(&mut self, arena: &mut ExprArena) -> NodeId {
        self.advance();
        if self.peek().kind == TokenKind::Kw(Kw::Hash) {
            self.parse_file_print_after_keyword(arena, FileIoKind::Print)
        } else {
            self.parse_print_list(arena, None)
        }
    }

    /// `Write #n, …` writes to a file; bare `Write …` is an implicit call.
    fn parse_write_keyword_stmt(&mut self, arena: &mut ExprArena) -> NodeId {
        let t = self.advance();
        if self.peek().kind == TokenKind::Kw(Kw::Hash) {
            self.parse_file_print_after_keyword(arena, FileIoKind::Write)
        } else {
            self.statement_without_operands(arena, t)
        }
    }

    /// `Input`/`Input$`/`InputB`/`InputB$`: `#n, vars` reads a file; otherwise
    /// an implicit call.
    fn parse_input_keyword_stmt(&mut self, arena: &mut ExprArena) -> NodeId {
        let t = self.advance();
        if self.peek().kind == TokenKind::Kw(Kw::Hash) {
            self.parse_file_input_after_keyword(arena)
        } else {
            self.statement_without_operands(arena, t)
        }
    }

    /// `Line Input #n, var`; `Line (x1,y1)-(x2,y2)[,color[,flags]]`; otherwise an implicit call.
    fn parse_line_keyword_stmt(&mut self, arena: &mut ExprArena) -> NodeId {
        let t = self.advance();
        if self.eat(TokenKind::Kw(Kw::Input)) {
            self.eat(TokenKind::Kw(Kw::Hash));
            let ch = self.parse_expr(arena, 0);
            self.expect(TokenKind::Kw(Kw::Comma), ERR_UNEXPECTED_TOKEN);
            let var = self.parse_expr(arena, 0);
            self.end_stmt();
            arena.alloc(ExprNode::FileIoStmt {
                kind: FileIoKind::LineInput,
                channel: Some(ch),
                args: vec![var],
                open_mode: 0,
                open_flags: 0,
            })
        } else {
            // `Line [Step](x1,y1)-(x2,y2)[,color[,B|BF]]` is the graphics method
            // of the enclosing form; otherwise an ordinary name.
            if matches!(self.peek().kind, TokenKind::Kw(Kw::LParen) | TokenKind::Kw(Kw::Minus) | TokenKind::Kw(Kw::Step)) {
                self.graphics_point_stmt(arena, t)
            } else {
                self.keyword_as_name_stmt(arena, t)
            }
        }
    }

    /// Parse `(x, y)` — a VB6 graphics coordinate pair.
    /// Returns `(x_node, y_node)`.
    fn parse_coord_pair(&mut self, arena: &mut ExprArena) -> (NodeId, NodeId) {
        self.expect(TokenKind::Kw(Kw::LParen), ERR_UNEXPECTED_TOKEN);
        let x = self.parse_expr(arena, 0);
        self.expect(TokenKind::Kw(Kw::Comma), ERR_UNEXPECTED_TOKEN);
        let y = self.parse_expr(arena, 0);
        self.expect(TokenKind::Kw(Kw::RParen), ERR_EXPECTED_RPAREN);
        (x, y)
    }

    /// The items of a `Print` method statement, separated by `;` or `,`
    /// (either may also end the statement).
    fn parse_print_method_items(&mut self, arena: &mut ExprArena, args: &mut Vec<NodeId>) {
        while !self.peek().is_stmt_end() {
            if self.eat(TokenKind::Kw(Kw::Semi)) || self.eat(TokenKind::Kw(Kw::Comma)) {
                continue;
            }
            let before = self.peek().span.start;
            match self.try_parse_print_item(arena) {
                Some(expr) => args.push(expr),
                None => break,
            }
            if self.peek().span.start == before {
                break;
            }
        }
    }

    /// One item of a `Print`-family list: an expression, or the `Tab[(n)]` /
    /// `Spc(n)` positioning directives.
    fn try_parse_print_item(&mut self, arena: &mut ExprArena) -> Option<NodeId> {
        if matches!(self.peek().kind, TokenKind::Kw(Kw::Tab) | TokenKind::Kw(Kw::Spc)) {
            let tok = self.advance();
            let sym = tok.sym.map(|s| s as u32).unwrap_or(0);
            let base = arena.alloc(ExprNode::NameRef { sym, suffix: tok.type_suffix });
            self.set_span(base, tok.span);
            return Some(self.parse_postfix(arena, base));
        }
        self.try_parse_expr(arena)
    }

    /// `Step (x, y)` as a single [`ExprNode::GraphicsPoint`].
    fn parse_step_point(&mut self, arena: &mut ExprArena) -> NodeId {
        self.advance(); // `Step`
        let (x, y) = self.parse_coord_pair(arena);
        arena.alloc(ExprNode::GraphicsPoint { step: true, x, y })
    }

    /// The member (or bare) name `lhs` ends in, lower-cased, when `lhs` is a
    /// name, a member access, or a call of one of those.
    fn graphics_target_is(&self, arena: &ExprArena, lhs: NodeId, wanted: &str) -> bool {
        let target = match arena.get(lhs) {
            ExprNode::Call { func, .. } => *func,
            _ => lhs,
        };
        let sym = match arena.get(target) {
            ExprNode::NameRef { sym, .. } => *sym,
            ExprNode::MemberAccess { member, .. } => *member,
            _ => return false,
        };
        self.scanner.sym_name(sym).eq_ignore_ascii_case(wanted)
    }

    /// True at the `-` that continues a graphics point pair: after
    /// `Obj.Line (x1, y1)` (a call), or at the start of `Obj.Line -(x, y)` /
    /// `Obj.Line -Step(x, y)` where the first point is omitted.
    fn at_graphics_continuation(&mut self, arena: &ExprArena, lhs: NodeId) -> bool {
        if self.peek().kind != TokenKind::Kw(Kw::Minus) {
            return false;
        }
        let follows = matches!(
            self.peek2().kind,
            TokenKind::Kw(Kw::LParen) | TokenKind::Kw(Kw::Step)
        );
        follows
            && (matches!(arena.get(lhs), ExprNode::Call { .. })
                || self.graphics_target_is(arena, lhs, "line")
                || self.graphics_target_is(arena, lhs, "scale"))
    }

    /// The `, arg, arg …` after the points of a graphics statement. An
    /// omitted argument is a `MissingArg`; in a `Line` statement the second
    /// argument may be the `B` / `BF` box flag.
    fn parse_graphics_tail(&mut self, arena: &mut ExprArena, args: &mut Vec<NodeId>, is_line: bool) {
        let mut position = 0;
        while self.eat(TokenKind::Kw(Kw::Comma)) {
            let arg = if matches!(self.peek().kind, TokenKind::Kw(Kw::Comma)) || self.peek().is_stmt_end() {
                arena.alloc(ExprNode::MissingArg)
            } else if let Some(flag) = self.try_parse_line_box(arena, is_line && position == 1) {
                flag
            } else if self.peek().kind == TokenKind::Kw(Kw::Step) && self.peek2().kind == TokenKind::Kw(Kw::LParen) {
                self.parse_step_point(arena)
            } else if let Some(expr) = self.try_parse_expr(arena) {
                expr
            } else {
                arena.alloc(ExprNode::MissingArg)
            };
            args.push(arg);
            position += 1;
        }
    }

    /// `B` / `BF` where a `Line` statement takes its box flag.
    fn try_parse_line_box(&mut self, arena: &mut ExprArena, allowed: bool) -> Option<NodeId> {
        if !allowed {
            return None;
        }
        let filled = match self.peek().kind {
            TokenKind::Kw(Kw::B) => false,
            TokenKind::Kw(Kw::BF) => true,
            _ => return None,
        };
        let span = self.current_span();
        self.advance();
        let id = arena.alloc(ExprNode::LineBox { filled });
        self.set_span(id, span);
        Some(id)
    }

    /// `Get`/`Put` random-access I/O: `#n, …`; otherwise an implicit call.
    fn parse_get_put_keyword_stmt(&mut self, arena: &mut ExprArena, kind: FileIoKind) -> NodeId {
        let t = self.advance();
        // The `#` before the channel is optional: `Get F, pos, value`.
        if self.starts_file_operand() {
            self.parse_file_get_put_after_keyword(arena, kind)
        } else {
            self.statement_without_operands(arena, t)
        }
    }

    /// True when the token after a reserved file-statement keyword (`Lock`,
    /// `Unlock`, `Seek`) begins its channel operand, rather than the keyword
    /// standing for a member name (`Seek = 1`, `Seek.x`, `Seek(2)`).
    fn starts_file_operand(&mut self) -> bool {
        !(self.peek().is_stmt_end()
            || matches!(
                self.peek().kind,
                TokenKind::Kw(Kw::Eq)
                    | TokenKind::Kw(Kw::Colon)
                    | TokenKind::Kw(Kw::Comma)
                    | TokenKind::Kw(Kw::Else)
                    | TokenKind::Eof
            ))
    }

    /// `Lock`/`Unlock`: `#n, …`; otherwise an implicit call.
    fn parse_lock_keyword_stmt(&mut self, arena: &mut ExprArena, kind: FileIoKind) -> NodeId {
        let t = self.advance();
        if self.starts_file_operand() {
            self.parse_file_lock_after_keyword(arena, kind)
        } else {
            self.statement_without_operands(arena, t)
        }
    }

    /// `Seek #n, pos` / `Width #n, w` — `#n, expr` channel statements; otherwise
    /// an implicit call.
    fn parse_channel_expr_keyword_stmt(
        &mut self,
        arena: &mut ExprArena,
        kind: FileIoKind,
    ) -> NodeId {
        let t = self.advance();
        // `Seek` is reserved, so its channel may omit the `#`; `Width` is an
        // ordinary name unless the `#` is written.
        let has_channel = if kind == FileIoKind::Seek {
            self.starts_file_operand()
        } else {
            self.peek().kind == TokenKind::Kw(Kw::Hash)
        };
        if has_channel {
            self.eat(TokenKind::Kw(Kw::Hash));
            let ch = self.parse_expr(arena, 0);
            self.expect(TokenKind::Kw(Kw::Comma), ERR_UNEXPECTED_TOKEN);
            let val = self.parse_expr(arena, 0);
            self.end_stmt();
            arena.alloc(ExprNode::FileIoStmt {
                kind,
                channel: Some(ch),
                args: vec![val],
                open_mode: 0,
                open_flags: 0,
            })
        } else if kind == FileIoKind::Seek {
            self.statement_without_operands(arena, t)
        } else {
            self.keyword_as_name_stmt(arena, t)
        }
    }

    /// `Name "old" As "new"`; or a `Name`-identifier assignment / call fallback.
    fn parse_name_keyword_stmt(&mut self, arena: &mut ExprArena) -> NodeId {
        let t = self.advance();
        // `Name` used as an identifier: assigned (`Name = x`) or with a member
        // (`Name.Length`). Alone it is a `Name` statement without operands.
        if self.peek().is_stmt_end()
            || matches!(
                self.peek().kind,
                TokenKind::Kw(Kw::Eq) | TokenKind::Kw(Kw::Dot) | TokenKind::Kw(Kw::Bang)
                    | TokenKind::Kw(Kw::Colon) | TokenKind::Kw(Kw::Else)
            )
        {
            return self.keyword_as_name_stmt(arena, t);
        }
        // Name "old" As "new"
        let old_path = self.parse_expr(arena, 0);
        if self.eat(TokenKind::Kw(Kw::As)) {
            let new_path = self.parse_expr(arena, 0);
            self.end_stmt();
            return arena.alloc(ExprNode::FileIoStmt {
                kind: FileIoKind::Name,
                channel: None,
                args: vec![old_path, new_path],
                open_mode: 0,
                open_flags: 0,
            });
        }
        // `Name <path>` must go on to `As`.
        let span = self.current_span();
        self.diagnostics.push(ERR_SYNTAX_ERROR, span);
        self.skip_line();
        old_path
    }

    /// `End`: bare `End` terminates the program; `End If`/`End With`/… are
    /// structural block terminators consumed by the parent block parser.
    fn parse_end_stmt(&mut self, arena: &mut ExprArena) -> Option<NodeId> {
        // "End" alone terminates the program.
        // End Sub / End Function etc. are handled in their parent.
        if self.terminator().is_some() {
            return self.stray_terminator();
        }
        self.advance();
        if self.peek().is_stmt_end()
            || matches!(
                self.peek().kind,
                TokenKind::Kw(Kw::Colon) | TokenKind::Kw(Kw::Else) | TokenKind::Eof
            )
        {
            self.end_stmt();
            return Some(arena.alloc(ExprNode::EndStmt));
        }
        // `End` takes no operand: anything else on the line is an error.
        let span = self.current_span();
        self.diagnostics.push_labeled(ERR_UNEXPECTED_TOKEN, span, "end of statement");
        self.skip_line();
        None
    }

    /// `LSet`/`RSet target = expr` — range-copy assignment preserving justification.
    fn parse_range_assign_stmt(&mut self, arena: &mut ExprArena) -> NodeId {
        let right_justify = matches!(self.peek().kind, TokenKind::Kw(Kw::RSet));
        self.advance();
        let target = self.parse_target_expr(arena);
        self.expect(TokenKind::Kw(Kw::Eq), ERR_EXPECTED_EQ);
        let value = self.parse_expr(arena, 0);
        self.end_stmt();
        arena.alloc(ExprNode::RangeAssign { right_justify, target, value })
    }

    // ── Statement dispatch ────────────────────────────────────────────────────

    /// Dispatches to the appropriate sub-parser based on the current token.
    /// Returns `None` on a pure blank/comment line.
    pub fn parse_stmt(&mut self, arena: &mut ExprArena) -> Option<NodeId> {
        // Skip statement ends
        while self.eat(TokenKind::Eol) {}
        if self.peek().kind == TokenKind::Eof { return None; }

        // `Else` or a line number in front of an attribute line does not change it.
        if self.at_prefixed_attribute_line() {
            self.advance();
            self.eat(TokenKind::Kw(Kw::Colon));
            return Some(self.parse_attribute_stmt(arena));
        }

        let kind = self.peek().kind.clone();
        let node = match &kind {
            // A keyword that is not reserved names a line label like any name;
            // a reserved word that starts a statement is that statement.
            TokenKind::Kw(_)
                if self.peek().sym.is_some()
                    && self.peek2().kind == TokenKind::Kw(Kw::Colon)
                    && !self.at_reserved_word() =>
            {
                let name = self.advance();
                self.advance(); // colon
                self.after_colon = false; // the colon belongs to the label
                let sym = name.sym.map(|s| s as u32).unwrap_or(0);
                self.labels_defined.push(LabelRef::Name(sym));
                let id = arena.alloc(ExprNode::Label { target: LabelRef::Name(sym) });
                self.set_span(id, name.span);
                id
            }
            // Variable / constant declarations
            TokenKind::Kw(Kw::Dim)
            | TokenKind::Kw(Kw::Static)
            | TokenKind::Kw(Kw::Public)
            | TokenKind::Kw(Kw::Private)
            | TokenKind::Kw(Kw::Global) => {
                let is_static = matches!(kind, TokenKind::Kw(Kw::Static));
                self.advance();
                self.parse_dim_list_static(arena, false, is_static)
            }
            TokenKind::Kw(Kw::Const) => {
                self.advance();
                self.parse_const_list(arena)
            }
            // ReDim
            TokenKind::Kw(Kw::ReDim) => self.parse_redim_stmt(arena),
            // Control flow
            TokenKind::Kw(Kw::If)     => self.parse_if_stmt(arena),
            TokenKind::Kw(Kw::For)    => self.parse_for_stmt(arena),
            TokenKind::Kw(Kw::Do)     => self.parse_do_stmt(arena),
            TokenKind::Kw(Kw::While)  => self.parse_while_stmt(arena),
            TokenKind::Kw(Kw::Select) => self.parse_select_stmt(arena),
            TokenKind::Kw(Kw::With)   => self.parse_with_stmt(arena),
            // Assignment keywords
            TokenKind::Kw(Kw::Let)    => self.parse_let_stmt(arena),
            TokenKind::Kw(Kw::Set)    => self.parse_set_stmt(arena),
            // Call / exit / goto
            TokenKind::Kw(Kw::Call)   => self.parse_call_stmt(arena),
            TokenKind::Kw(Kw::Exit)   => self.parse_exit_stmt(arena),
            TokenKind::Kw(Kw::Return) => {
                self.advance();
                self.end_stmt();
                arena.alloc(ExprNode::ReturnStmt)
            }
            TokenKind::Kw(Kw::GoTo)   => self.parse_goto_stmt(arena, false),
            TokenKind::Kw(Kw::GoSub)  => self.parse_goto_stmt(arena, true),
            // Space-separated `Go To` / `Go Sub` — the second keyword selects
            // the kind. VB6 accepts both spellings (oracle-confirmed).
            TokenKind::Kw(Kw::Go)
                if matches!(
                    self.peek2().kind,
                    TokenKind::Kw(Kw::Eq) | TokenKind::Kw(Kw::Dot) | TokenKind::Kw(Kw::Bang)
                ) =>
            {
                self.parse_ident_stmt(arena)
            }
            TokenKind::Kw(Kw::Go)     => self.parse_goto_stmt(arena, false),
            TokenKind::Kw(Kw::Resume) => self.parse_resume_stmt(arena),
            TokenKind::Kw(Kw::On)     => self.parse_on_stmt(arena),
            // Error
            TokenKind::Kw(Kw::Error) if self.error_starts_target() => self.parse_ident_stmt(arena),
            TokenKind::Kw(Kw::Error) => {
                self.advance();
                let expr = self.parse_expr(arena, 0);
                self.end_stmt();
                arena.alloc(ExprNode::ErrorStmt { expr })
            }
            // Misc
            TokenKind::Kw(Kw::Stop) => {
                self.advance(); self.end_stmt();
                arena.alloc(ExprNode::Stop)
            }
            TokenKind::Kw(Kw::End) => return self.parse_end_stmt(arena),
            // A block terminator that does not close the innermost open block.
            TokenKind::Kw(Kw::Else)
            | TokenKind::Kw(Kw::ElseIf)
            | TokenKind::Kw(Kw::EndIf)
            | TokenKind::Kw(Kw::Next)
            | TokenKind::Kw(Kw::Loop)
            | TokenKind::Kw(Kw::Wend)
            | TokenKind::Kw(Kw::Case) => return self.stray_terminator(),
            TokenKind::Kw(Kw::Erase)  => self.parse_erase_stmt(arena),
            TokenKind::Kw(Kw::RaiseEvent) => self.parse_raise_event(arena),
            TokenKind::Kw(Kw::Debug)  => self.parse_debug_stmt(arena),

            // Attribute metadata line
            TokenKind::Kw(Kw::Attribute) => self.parse_attribute_stmt(arena),
            // File I/O statements
            TokenKind::Kw(Kw::Open)   => self.parse_file_open(arena),
            TokenKind::Kw(Kw::Close)  => self.parse_file_close(arena),
            TokenKind::Kw(Kw::Print)  => self.parse_print_keyword_stmt(arena),
            TokenKind::Kw(Kw::Write)  => self.parse_write_keyword_stmt(arena),
            TokenKind::Kw(Kw::Input)  => self.parse_input_keyword_stmt(arena),
            TokenKind::Kw(Kw::Line)   => self.parse_line_keyword_stmt(arena),
            TokenKind::Kw(Kw::Get)    => self.parse_get_put_keyword_stmt(arena, FileIoKind::Get),
            TokenKind::Kw(Kw::Put)    => self.parse_get_put_keyword_stmt(arena, FileIoKind::Put),
            TokenKind::Kw(Kw::Seek)   => {
                self.parse_channel_expr_keyword_stmt(arena, FileIoKind::Seek)
            }
            TokenKind::Kw(Kw::Lock)   => self.parse_lock_keyword_stmt(arena, FileIoKind::Lock),
            TokenKind::Kw(Kw::Unlock) => self.parse_lock_keyword_stmt(arena, FileIoKind::Unlock),
            TokenKind::Kw(Kw::Width)  => {
                self.parse_channel_expr_keyword_stmt(arena, FileIoKind::Width)
            }
            TokenKind::Kw(Kw::Name)   => self.parse_name_keyword_stmt(arena),
            // Mid / Mid$ / MidB / MidB$ statement: `Mid(s, start[, len]) = value`.
            // The spelling selects the variant flag.
            TokenKind::Kw(Kw::Mid) | TokenKind::Kw(Kw::MidS)
            | TokenKind::Kw(Kw::MidB) | TokenKind::Kw(Kw::MidBS) => {
                if self.peek2().kind == TokenKind::Kw(Kw::LParen) {
                    self.parse_mid_assign(arena)
                } else {
                    let t = self.advance();
                    self.keyword_as_name_stmt(arena, t)
                }
            }
            // `LSet target = expr` / `RSet target = expr` — range-copy assignment.
            // The justification side is preserved.
            TokenKind::Kw(Kw::LSet) | TokenKind::Kw(Kw::RSet) => {
                self.parse_range_assign_stmt(arena)
            }
            // `?` is the `Print` shortcut.
            // It parses identically to `Print`: `? #n, args` writes
            // to a file, bare `? args` writes to the immediate window.
            TokenKind::Kw(Kw::Question) => self.parse_print_keyword_stmt(arena),
            // `Input$ #n, vars` / `InputB #n, vars` / `InputB$ #n, vars` —
            // same file-read semantics as `Input #n, vars`.
            TokenKind::Kw(Kw::InputS) | TokenKind::Kw(Kw::InputB) | TokenKind::Kw(Kw::InputBS) => {
                self.parse_input_keyword_stmt(arena)
            }
            // `Property` used as a name: assigned, or with a member.
            TokenKind::Kw(Kw::Property)
                if matches!(
                    self.peek2().kind,
                    TokenKind::Kw(Kw::Eq) | TokenKind::Kw(Kw::Dot) | TokenKind::Kw(Kw::Bang)
                ) =>
            {
                let t = self.advance();
                self.parse_ident_stmt_with_tok(arena, t)
            }
            // `Scale` alone, or with a coordinate pair, is the graphics method.
            TokenKind::Kw(Kw::Scale) => {
                let t = self.advance();
                if self.peek().is_stmt_end()
                    || matches!(
                        self.peek().kind,
                        TokenKind::Kw(Kw::Colon) | TokenKind::Kw(Kw::Else) | TokenKind::Eof
                    )
                {
                    self.end_stmt();
                    let sym = t.sym.map(|s| s as u32).unwrap_or(0);
                    let node = arena.alloc(ExprNode::NameRef { sym, suffix: t.type_suffix });
                    self.set_span(node, t.span);
                    node
                } else if self.peek().kind == TokenKind::Kw(Kw::LParen) {
                    self.graphics_point_stmt(arena, t)
                } else {
                    self.statement_without_operands(arena, t)
                }
            }
            // Sub / Function / Property — nested (error but recover)
            TokenKind::Kw(Kw::Sub) | TokenKind::Kw(Kw::Function) | TokenKind::Kw(Kw::Property)
            | TokenKind::Kw(Kw::Declare) | TokenKind::Kw(Kw::Event) | TokenKind::Kw(Kw::Implements)
            | TokenKind::Kw(Kw::Option) => {
                let span = self.current_span();
                self.diagnostics.push(ERR_INVALID_IN_PROCEDURE, span);
                self.skip_line();
                return None;
            }
            // A `Type` or `Enum` block cannot be declared inside a procedure; the
            // whole block is consumed so its members are not read as statements.
            TokenKind::Kw(Kw::Type) | TokenKind::Kw(Kw::Enum) => {
                let span = self.current_span();
                let mark = self.diagnostics.len();
                if kind == TokenKind::Kw(Kw::Type) {
                    self.parse_type_decl(arena);
                } else {
                    self.parse_enum_decl(arena);
                }
                self.diagnostics.truncate(mark);
                self.diagnostics.push(ERR_INVALID_IN_PROCEDURE, span);
                return None;
            }
            TokenKind::Kw(k) if def_type_kw(*k).is_some() => {
                let span = self.current_span();
                self.diagnostics.push(ERR_INVALID_IN_PROCEDURE, span);
                self.skip_line();
                return None;
            }
            // Comment tokens — consume and return None
            TokenKind::Kw(Kw::Apos) | TokenKind::Kw(Kw::Rem) => {
                self.advance(); // skip to EOL is implicit (scanner stops at Apos)
                return None;
            }
            // Identifier — could be a line label (`Foo:`), assignment, or implicit call.
            TokenKind::Ident if self.peek2().kind == TokenKind::Kw(Kw::Colon) => {
                let name = self.advance(); // identifier
                self.advance();            // colon
                self.after_colon = false;  // the colon belongs to the label
                let sym = name.sym.map(|s| s as u32).unwrap_or(0);
                self.labels_defined.push(LabelRef::Name(sym));
                let id = arena.alloc(ExprNode::Label { target: LabelRef::Name(sym) });
                self.set_span(id, name.span);
                id
            }
            TokenKind::Ident => self.parse_ident_stmt(arena),
            // Leading dot — With-block member access (.Member = expr or .Method args)
            TokenKind::Kw(Kw::Dot) => self.parse_ident_stmt(arena),
            // Colon — empty statement separator
            TokenKind::Kw(Kw::Colon) => {
                self.advance();
                return None;
            }
            // Numeric line label: `10  stmt` or `10: stmt`. The label is a
            // jump target in its own right (e.g. `GoTo 10`), so it is emitted
            // as a `Label` node carrying the line number; any statement that
            // follows on the same line is parsed by the enclosing block loop.
            TokenKind::IntLit | TokenKind::LongLit => {
                if self.after_colon {
                    // A line number only starts a line.
                    let span = self.current_span();
                    self.diagnostics.push(ERR_SYNTAX_ERROR, span);
                    self.skip_line();
                    return None;
                }
                let tok = self.advance(); // consume the line number
                if matches!(self.peek().kind, TokenKind::IntLit | TokenKind::LongLit) {
                    let span = self.current_span();
                    self.diagnostics.push(ERR_SYNTAX_ERROR, span);
                    self.skip_line();
                    return None;
                }
                let line = line_label_value(&tok.lit);
                if self.eat(TokenKind::Kw(Kw::Colon)) { // optional colon
                    self.after_colon = false;
                }
                self.labels_defined.push(LabelRef::Line(line));
                let id = arena.alloc(ExprNode::Label { target: LabelRef::Line(line) });
                self.set_span(id, tok.span);
                id
            }
            // `PSet` and `Circle` take a coordinate pair, `Step`-relative or not.
            TokenKind::Kw(Kw::PSet) | TokenKind::Kw(Kw::Circle)
                if matches!(self.peek2().kind, TokenKind::Kw(Kw::LParen) | TokenKind::Kw(Kw::Step)) =>
            {
                let t = self.advance();
                self.graphics_point_stmt(arena, t)
            }
            // `Me` starts a statement only as the object of a member, an element or an assignment.
            TokenKind::Kw(Kw::Me)
                if matches!(
                    self.peek2().kind,
                    TokenKind::Kw(Kw::Dot) | TokenKind::Kw(Kw::Bang) | TokenKind::Kw(Kw::LParen) | TokenKind::Kw(Kw::Eq)
                ) =>
            {
                self.parse_ident_stmt(arena)
            }
            // A leading `!` is a With-object key.
            TokenKind::Kw(Kw::Bang) => self.parse_ident_stmt(arena),
            // Punctuation, operators and reserved words that name no statement cannot
            // start one; any other keyword is an ordinary name.
            TokenKind::Kw(_) if !self.starts_name_statement() => {
                let span = self.current_span();
                self.diagnostics.push(ERR_SYNTAX_ERROR, span);
                self.skip_line();
                return None;
            }
            TokenKind::Kw(_) => self.parse_ident_stmt(arena),
            // Anything else is unexpected
            _ => {
                let span = self.current_span();
                self.diagnostics.push(ERR_UNEXPECTED_TOKEN, span);
                self.skip_to_stmt_end();
                return None;
            }
        };
        Some(node)
    }

    // ── Individual statement parsers ──────────────────────────────────────────

    /// Implicit assignment or call when the line starts with an identifier.
    ///
    /// Covers: `target = expr` (Let), `Call target(args)`, label definition,
    /// property call without `Set`.
    fn parse_ident_stmt(&mut self, arena: &mut ExprArena) -> NodeId {
        self.target_at_statement_start = true;
        let lhs = self.parse_target_expr(arena);
        self.finish_ident_stmt(arena, lhs)
    }

    /// The rest of a statement that starts with a target expression `lhs`:
    /// an assignment, a graphics coordinate call, or an implicit call.
    fn finish_ident_stmt(&mut self, arena: &mut ExprArena, lhs: NodeId) -> NodeId {
        if self.eat(TokenKind::Kw(Kw::Eq)) {
            let value = self.parse_expr(arena, 0);
            self.end_stmt();
            arena.alloc(ExprNode::Assign { target: lhs, value })
        } else if self.eat(TokenKind::Kw(Kw::ColonEq)) {
            // Named argument in a call — parse rest as expression
            let value = self.parse_expr(arena, 0);
            self.end_stmt();
            arena.alloc(ExprNode::Assign { target: lhs, value })
        } else if let Some((callee, arg)) = self.paren_first_argument(arena, lhs) {
            // `Foo (a) - (b), c` / `Foo (a), b`: in a call statement the
            // parenthesised group after the callee begins the first argument
            // expression; it is not the argument list.
            let first = arena.alloc(ExprNode::Paren { inner: arg });
            let first = self.parse_infix(arena, first, 0);
            let mut args = vec![first];
            if self.graphics_target_is(arena, callee, "print") {
                self.parse_print_method_items(arena, &mut args);
            } else if self.eat(TokenKind::Kw(Kw::Comma)) {
                args.extend(self.parse_stmt_arg_list(arena));
            }
            self.end_stmt();
            let args = arena.alloc(ExprNode::ArgList { args });
            arena.alloc(ExprNode::CallStmt { callee, args })
        } else if self.at_graphics_continuation(arena, lhs) {
            // `Obj.Line (x1,y1)-(x2,y2)` — postfix consumed `(x1,y1)` as a call
            // arg-list; the `-(x2,y2)` continuation signals a graphics coord-pair
            // call. `Obj.Line -(x2,y2)` omits the first point.
            let is_line = self.graphics_target_is(arena, lhs, "line");
            let mut args = Vec::new();
            if !matches!(arena.get(lhs), ExprNode::Call { .. }) {
                args.push(arena.alloc(ExprNode::MissingArg));
            }
            self.advance(); // consume `-`
            if self.peek().kind == TokenKind::Kw(Kw::Step) {
                args.push(self.parse_step_point(arena));
            } else if matches!(arena.get(lhs), ExprNode::Call { .. }) {
                let (x2, y2) = self.parse_coord_pair(arena);
                args.extend([x2, y2]);
            } else {
                let (x2, y2) = self.parse_coord_pair(arena);
                args.push(arena.alloc(ExprNode::GraphicsPoint { step: false, x: x2, y: y2 }));
            }
            self.parse_graphics_tail(arena, &mut args, is_line);
            self.end_stmt();
            let args = arena.alloc(ExprNode::ArgList { args });
            arena.alloc(ExprNode::CallStmt { callee: lhs, args })
        } else if self.has_empty_call(arena, lhs)
            || self.call_with_address_of_argument(arena, lhs)
            || self.call_with_several_arguments(arena, lhs)
        {
            // `Foo()`, `Foo(AddressOf Bar)` and `Foo(a, b)` are not statements: `Call` is needed for them.
            let span = self.current_span();
            self.report_expected(span, "end of statement");
            self.skip_line();
            lhs
        } else {
            // Implicit call with optional trailing arguments: `Foo a, b`. After
            // a parenthesised coordinate pair (`PSet (x, y), color`) the comma
            // only separates the pair from the remaining arguments.
            if matches!(arena.get(lhs), ExprNode::Call { .. }) {
                let at_end = self.peek().is_stmt_end()
                    || matches!(
                        self.peek().kind,
                        TokenKind::Kw(Kw::Colon) | TokenKind::Kw(Kw::Else) | TokenKind::Eof
                    );
                if !at_end && !self.eat(TokenKind::Kw(Kw::Comma)) {
                    let span = self.current_span();
                    self.report_expected(span, "end of statement");
                    self.skip_line();
                    return lhs;
                }
            }
            if self.graphics_target_is(arena, lhs, "print") {
                // `Obj.Print a; b, c;` — `;` and `,` separate the items.
                let mut args = Vec::new();
                self.parse_print_method_items(arena, &mut args);
                self.end_stmt();
                if args.is_empty() {
                    return lhs;
                }
                let args = arena.alloc(ExprNode::ArgList { args });
                return arena.alloc(ExprNode::CallStmt { callee: lhs, args });
            }
            if self.peek().kind == TokenKind::Kw(Kw::Step) && self.peek2().kind == TokenKind::Kw(Kw::LParen) {
                // `Obj.PSet Step(x, y)`, `Obj.Line Step(x1, y1)-Step(x2, y2)`.
                let is_line = self.graphics_target_is(arena, lhs, "line");
                let mut args = vec![self.parse_step_point(arena)];
                if self.peek().kind == TokenKind::Kw(Kw::Minus)
                    && matches!(self.peek2().kind, TokenKind::Kw(Kw::LParen) | TokenKind::Kw(Kw::Step))
                {
                    self.advance(); // `-`
                    if self.peek().kind == TokenKind::Kw(Kw::Step) {
                        args.push(self.parse_step_point(arena));
                    } else {
                        let (x, y) = self.parse_coord_pair(arena);
                        args.push(arena.alloc(ExprNode::GraphicsPoint { step: false, x, y }));
                    }
                }
                self.parse_graphics_tail(arena, &mut args, is_line);
                self.end_stmt();
                let args = arena.alloc(ExprNode::ArgList { args });
                return arena.alloc(ExprNode::CallStmt { callee: lhs, args });
            }
            let args = self.parse_stmt_arg_list(arena);
            self.end_stmt();
            if args.is_empty() {
                lhs
            } else {
                let args = arena.alloc(ExprNode::ArgList { args });
                arena.alloc(ExprNode::CallStmt { callee: lhs, args })
            }
        }
    }

    /// True when `lhs` is a call, or a call of a call, with empty brackets.
    fn has_empty_call(&self, arena: &ExprArena, lhs: NodeId) -> bool {
        let ExprNode::Call { func, args } = arena.get(lhs) else { return false };
        let empty = matches!(arena.get(*args), ExprNode::ArgList { args } if args.is_empty());
        empty || self.has_empty_call(arena, *func)
    }

    /// True when `lhs` is a bracketed call of two or more arguments, which only the
    /// graphics statements (`PSet (x, y)`, ...) may be without `Call`.
    fn call_with_several_arguments(&self, arena: &ExprArena, lhs: NodeId) -> bool {
        let ExprNode::Call { args, .. } = arena.get(lhs) else { return false };
        let several = matches!(arena.get(*args), ExprNode::ArgList { args } if args.len() >= 2);
        several
            && !["PSet", "Circle", "Line", "Scale"]
                .iter()
                .any(|name| self.graphics_target_is(arena, lhs, name))
    }

    /// True when `lhs` is a bracketed call that has an `AddressOf` argument.
    fn call_with_address_of_argument(&self, arena: &ExprArena, lhs: NodeId) -> bool {
        let ExprNode::Call { args, .. } = arena.get(lhs) else { return false };
        matches!(arena.get(*args), ExprNode::ArgList { args }
            if args.iter().any(|a| matches!(arena.get(*a), ExprNode::AddressOf { .. })))
    }

    /// If `lhs` is a call with exactly one plain argument and the statement
    /// goes on with an operator or another argument, return `(callee, arg)`:
    /// that parenthesised argument starts the first argument expression. A
    /// two-value group (`(x, y)`) is a graphics coordinate pair, never this.
    fn paren_first_argument(&mut self, arena: &ExprArena, lhs: NodeId) -> Option<(NodeId, NodeId)> {
        let ExprNode::Call { func, args } = arena.get(lhs) else { return None };
        let ExprNode::ArgList { args } = arena.get(*args) else { return None };
        let [arg] = args.as_slice() else { return None };
        if matches!(arena.get(*arg), ExprNode::NamedArg { .. } | ExprNode::MissingArg) {
            return None;
        }
        let kind = self.peek().kind.clone();
        let (func, arg) = (*func, *arg);
        // `Obj.Print (a); b`: `;` separates the items of a `Print` method.
        let semi = kind == TokenKind::Kw(Kw::Semi) && self.graphics_target_is(arena, func, "print");
        let continues = kind == TokenKind::Kw(Kw::Comma) || infix_bp(&kind).is_some() || semi;
        continues.then_some((func, arg))
    }

    /// Parses a `Let` assignment statement.
    fn parse_let_stmt(&mut self, arena: &mut ExprArena) -> NodeId {
        self.advance(); // consume Let
        let target = self.parse_target_expr(arena);
        self.expect(TokenKind::Kw(Kw::Eq), ERR_EXPECTED_EQ);
        let value = self.parse_expr(arena, 0);
        self.end_stmt();
        arena.alloc(ExprNode::Assign { target, value })
    }

    /// Parses a `Set` assignment statement.
    fn parse_set_stmt(&mut self, arena: &mut ExprArena) -> NodeId {
        self.advance(); // consume Set
        let target = self.parse_target_expr(arena);
        self.expect(TokenKind::Kw(Kw::Eq), ERR_EXPECTED_EQ);
        let value = self.parse_expr(arena, 0);
        self.end_stmt();
        arena.alloc(ExprNode::SetAssign { target, value })
    }

    /// `Mid`/`Mid$` string-replacement statement: `Mid(s, start[, len]) = value`.
    ///
    /// The spelling selects the variant flag: `MidB`/`MidB$` are byte-oriented;
    /// `Mid$`/`MidB$` carry the `$` bit.
    fn parse_mid_assign(&mut self, arena: &mut ExprArena) -> NodeId {
        let (byte_oriented, dollar) = match self.peek().kind {
            TokenKind::Kw(Kw::Mid)   => (false, false),
            TokenKind::Kw(Kw::MidS)  => (false, true),
            TokenKind::Kw(Kw::MidB)  => (true,  false),
            TokenKind::Kw(Kw::MidBS) => (true,  true),
            _ => (false, false),
        };
        self.advance(); // consume Mid / Mid$ / MidB / MidB$
        self.expect(TokenKind::Kw(Kw::LParen), ERR_EXPECTED_LPAREN);
        let args = self.parse_arg_list(arena);
        // `Mid(target, start [, length])`
        let arg_count = match arena.get(args) {
            ExprNode::ArgList { args } => args.len(),
            _ => 0,
        };
        if !(2..=3).contains(&arg_count) {
            let span = self.current_span();
            self.diagnostics.push(ERR_UNEXPECTED_TOKEN, span);
        }
        self.expect(TokenKind::Kw(Kw::RParen), ERR_EXPECTED_RPAREN);
        self.expect(TokenKind::Kw(Kw::Eq), ERR_EXPECTED_EQ);
        let value = self.parse_expr(arena, 0);
        self.end_stmt();
        arena.alloc(ExprNode::MidAssign { byte_oriented, dollar, args, value })
    }

    /// Parse a `Print`/`?` output list.
    fn parse_print_list(&mut self, arena: &mut ExprArena, channel: Option<NodeId>) -> NodeId {
        let mut args = Vec::new();
        self.parse_print_method_items(arena, &mut args);
        self.end_stmt();
        arena.alloc(ExprNode::FileIoStmt { kind: FileIoKind::Print, channel, args, open_mode: 0, open_flags: 0 })
    }

    /// Materialise the declared type implied by a type-declaration suffix.
    /// Returns `None` for `TypeSuffix::None`.
    fn type_node_from_suffix(&mut self, arena: &mut ExprArena, suffix: TypeSuffix) -> Option<NodeId> {
        match suffix {
            TypeSuffix::None => None,
            TypeSuffix::String => Some(arena.alloc(ExprNode::StringType { fixed_len: None })),
            other => other
                .builtin_kind()
                .map(|kind| arena.alloc(ExprNode::BuiltinType { kind })),
        }
    }

    /// Parse a Dim/Static/Public/Private variable list.
    fn parse_dim_list(&mut self, arena: &mut ExprArena, is_const: bool) -> NodeId {
        self.parse_dim_list_static(arena, is_const, false)
    }

    /// Parse a Dim/Static/Public/Private variable list, marking each item static
    /// when declared with the `Static` keyword.
    fn parse_dim_list_static(&mut self, arena: &mut ExprArena, is_const: bool, is_static: bool) -> NodeId {
        if is_static && self.in_procedure && self.peek().kind == TokenKind::Kw(Kw::Property) {
            // After `Static` a `Property` starts a property procedure, which a procedure body cannot hold.
            let span = self.current_span();
            self.diagnostics.push(ERR_SYNTAX_ERROR, span);
            self.skip_line();
            return arena.alloc(ExprNode::Block { stmts: Vec::new() });
        }
        let mut stmts = Vec::new();
        loop {
            // Variable declarations accept `As Type()`; `Const` does not.
            let item = self.parse_dim_item_with(arena, is_const, is_static, !is_const, is_const);
            stmts.push(item);
            if !self.eat(TokenKind::Kw(Kw::Comma)) { break; }
        }
        self.end_stmt();
        arena.alloc(ExprNode::Block { stmts })
    }

    /// Parse a Const declaration list.
    fn parse_const_list(&mut self, arena: &mut ExprArena) -> NodeId {
        self.parse_dim_list(arena, true)
    }

    /// Single variable: `[WithEvents] name [As type]` or `name(dims) [As type]` or `name = expr`.
    /// `WithEvents` may precede each variable of a list (not a constant).
    /// `array_suffix_ok` permits a trailing `As Type()` array marker.
    fn parse_dim_item(&mut self, arena: &mut ExprArena, is_const: bool, is_static: bool, array_suffix_ok: bool) -> NodeId {
        self.parse_dim_item_with(arena, is_const, is_static, array_suffix_ok, false)
    }

    /// As [`Self::parse_dim_item`]; `value_required` makes `= value` mandatory
    /// (a `Const` declaration, as opposed to an `Enum` member).
    fn parse_dim_item_with(
        &mut self,
        arena: &mut ExprArena,
        is_const: bool,
        is_static: bool,
        array_suffix_ok: bool,
        value_required: bool,
    ) -> NodeId {
        // A member named `WithEvents` is followed by its type, not by a name.
        let with_events = !is_const
            && self.peek().kind == TokenKind::Kw(Kw::WithEvents)
            && !matches!(self.peek2().kind, TokenKind::Kw(Kw::As) | TokenKind::Kw(Kw::LParen));
        if with_events {
            self.reject_in_standard_module();
            if self.in_procedure {
                let span = self.current_span();
                self.diagnostics.push(ERR_INVALID_IN_PROCEDURE, span);
            }
            self.advance();
        }
        let name = self.expect_decl_name();
        self.note_declared_name(&name, !self.in_procedure);
        let name_id = name.sym.map(|s| s as u32).unwrap_or(0);
        if is_const && self.peek().kind == TokenKind::Kw(Kw::LParen) {
            let span = self.current_span();
            self.report_expected(span, "As or =");
        }
        let bounds = if self.eat(TokenKind::Kw(Kw::LParen)) {
            let d = self.parse_array_bounds(arena);
            self.expect(TokenKind::Kw(Kw::RParen), ERR_EXPECTED_RPAREN);
            d
        } else {
            None
        };
        if !self.in_procedure && !is_const {
            self.module_decl_seen = true;
            self.module_array_seen |= bounds.is_some();
        }
        self.reject_suffix_with_as(&name, "end of statement");
        self.last_item_had_as = self.peek().kind == TokenKind::Kw(Kw::As);
        let (is_new, type_node) = if self.eat(TokenKind::Kw(Kw::As)) {
            let is_new = self.eat(TokenKind::Kw(Kw::New));
            // `Object` is not a creatable class.
            if is_new && self.peek().kind == TokenKind::Kw(Kw::Object) {
                let span = self.current_span();
                self.diagnostics.push_labeled(ERR_UNEXPECTED_TOKEN, span, "type name");
            } else if is_new && self.reserved_names_allowed && self.at_reserved_word() {
                // A `Type` member takes `New` before a class name only.
                let span = self.current_span();
                self.diagnostics.push_labeled(0x9c6f, span, "identifier");
            }
            (is_new, Some(self.parse_type_spec(arena, array_suffix_ok)))
        } else {
            // `Dim x%` ≡ `Dim x As Integer`
            (false, self.type_node_from_suffix(arena, name.type_suffix))
        };
        // Only a constant has an initial value.
        let init = if is_const {
            if value_required {
                self.expect(TokenKind::Kw(Kw::Eq), ERR_EXPECTED_EQ);
                Some(self.parse_expr(arena, 0))
            } else if self.eat(TokenKind::Kw(Kw::Eq)) {
                Some(self.parse_expr(arena, 0))
            } else { None }
        } else { None };
        let id = arena.alloc(ExprNode::DimItem { name: name_id, is_const, is_static, is_new, with_events, bounds, type_node, init });
        self.set_span(id, name.span);
        id
    }

    /// Parses array bounds. Returns the bounds-list `ArgList` node. A dynamic
    /// array (`a()`) yields an empty `ArgList` — `Some` so the declaration is
    /// still recognised as an array (distinct from a scalar), but with no
    /// dimensions until `ReDim`.
    fn parse_array_bounds(&mut self, arena: &mut ExprArena) -> Option<NodeId> {
        if self.peek().kind == TokenKind::Kw(Kw::RParen) {
            return Some(arena.alloc(ExprNode::ArgList { args: Vec::new() }));
        }
        let list_span = self.current_span();
        let mut args = Vec::new();
        loop {
            let lo = self.parse_expr(arena, 0);
            let node = if self.eat(TokenKind::Kw(Kw::To)) {
                let hi = self.parse_expr(arena, 0);
                arena.alloc(ExprNode::RangeTo { lo, hi })
            } else {
                lo
            };
            args.push(node);
            if !self.eat(TokenKind::Kw(Kw::Comma)) { break; }
        }
        if args.len() > MAX_DIMENSIONS {
            self.diagnostics.push(ERR_TOO_MANY_DIMENSIONS, list_span);
        }
        Some(arena.alloc(ExprNode::ArgList { args }))
    }

    /// Parses a ReDim statement.
    fn parse_redim_stmt(&mut self, arena: &mut ExprArena) -> NodeId {
        self.advance(); // consume ReDim
        let _shared = self.eat(TokenKind::Kw(Kw::Shared));
        let preserve = self.eat(TokenKind::Kw(Kw::Preserve));
        let mut stmts = Vec::new();
        loop {
            // `a(...)`, a member array `a.b.c(...)`, or an element's member
            // array `a(i).b(...)`: everything before the last name is the base
            // it belongs to. A parenthesised list followed by `.` indexes the
            // base; the last list holds the bounds.
            // A leading `.` names a member of the innermost `With` object.
            let mut base: Option<NodeId> = None;
            if self.peek().kind == TokenKind::Kw(Kw::Dot) {
                let dot_span = self.current_span();
                self.advance();
                let with_ctx = arena.alloc(ExprNode::WithContext);
                self.set_span(with_ctx, dot_span);
                base = Some(with_ctx);
            }
            let mut name = self.expect_decl_name();
            let mut bounds: Option<NodeId> = None;
            loop {
                let sym = name.sym.map(|s| s as u32).unwrap_or(0);
                if self.peek().kind == TokenKind::Kw(Kw::LParen) {
                    self.advance();
                    let list = self.parse_array_bounds(arena);
                    let bounds_given = list.is_some_and(|l| {
                        matches!(arena.get(l), ExprNode::ArgList { args }
                            if !args.is_empty() && args.iter().all(|a| !matches!(arena.get(*a), ExprNode::MissingArg)))
                    });
                    if !bounds_given && self.peek().kind == TokenKind::Kw(Kw::RParen) {
                        let span = self.current_span();
                        self.report_expected(span, "expression");
                    }
                    self.expect(TokenKind::Kw(Kw::RParen), ERR_EXPECTED_RPAREN);
                    if self.peek().kind != TokenKind::Kw(Kw::Dot) {
                        bounds = list;
                        break;
                    }
                    // An index into the element named so far.
                    let element = match base {
                        None => arena.alloc(ExprNode::NameRef { sym, suffix: name.type_suffix }),
                        Some(b) => arena.alloc(ExprNode::MemberAccess { base: b, member: sym, bang: false }),
                    };
                    self.set_span(element, name.span);
                    let args = list.unwrap_or_else(|| arena.alloc(ExprNode::ArgList { args: vec![] }));
                    base = Some(arena.alloc(ExprNode::Call { func: element, args }));
                    self.advance(); // `.`
                    name = self.expect_ident();
                } else if self.eat(TokenKind::Kw(Kw::Dot)) {
                    let node = match base {
                        None => arena.alloc(ExprNode::NameRef { sym, suffix: name.type_suffix }),
                        Some(b) => arena.alloc(ExprNode::MemberAccess { base: b, member: sym, bang: false }),
                    };
                    self.set_span(node, name.span);
                    base = Some(node);
                    name = self.expect_ident();
                } else {
                    self.expect(TokenKind::Kw(Kw::LParen), ERR_EXPECTED_LPAREN);
                    break;
                }
            }
            let name_id = name.sym.map(|s| s as u32).unwrap_or(0);
            self.reject_suffix_with_as(&name, "end of statement");
            let type_node = if self.eat(TokenKind::Kw(Kw::As)) {
                Some(self.parse_type_spec(arena, false))
            } else { None };
            let item = arena.alloc(ExprNode::ReDimItem { preserve, base, name: name_id, bounds, type_node });
            self.set_span(item, name.span);
            stmts.push(item);
            if !self.eat(TokenKind::Kw(Kw::Comma)) { break; }
        }
        self.end_stmt();
        arena.alloc(ExprNode::Block { stmts })
    }

    /// Parses a For loop.
    fn parse_for_stmt(&mut self, arena: &mut ExprArena) -> NodeId {
        self.advance(); // consume For
        if self.eat(TokenKind::Kw(Kw::Each)) {
            return self.parse_for_each_stmt(arena);
        }
        let var_span = self.current_span();
        if self.at_reserved_word() && self.peek().kind != TokenKind::Kw(Kw::Me) {
            self.reject_reserved_word();
        }
        let var   = self.parse_target_expr(arena);
        let var_sym = self.for_variable(arena, var, var_span);
        self.expect(TokenKind::Kw(Kw::Eq), ERR_EXPECTED_EQ);
        let start = self.parse_expr(arena, 0);
        self.expect(TokenKind::Kw(Kw::To), ERR_EXPECTED_TO);
        let end = self.parse_expr(arena, 0);
        let step = if self.eat(TokenKind::Kw(Kw::Step)) {
            Some(self.parse_expr(arena, 0))
        } else { None };
        self.end_stmt();
        self.blocks.push(Block::For(var_sym));
        let body = self.parse_block(arena);
        self.finish_next(arena);
        self.blocks.pop();
        arena.alloc(ExprNode::For { var, start, end, step, body })
    }

    fn parse_for_each_stmt(&mut self, arena: &mut ExprArena) -> NodeId {
        let var_span = self.current_span();
        if self.at_reserved_word() {
            self.reject_reserved_word();
        }
        let var = self.parse_target_expr(arena);
        let var_sym = self.for_variable(arena, var, var_span);
        self.expect(TokenKind::Kw(Kw::In), ERR_EXPECTED_IN);
        // The collection is a target expression: a name or a chain of members, indexes and calls.
        let collection = self.parse_target_expr(arena);
        self.end_stmt();
        self.blocks.push(Block::For(var_sym));
        let body = self.parse_block(arena);
        self.finish_next(arena);
        self.blocks.pop();
        arena.alloc(ExprNode::ForEach { var, collection, body })
    }

    /// The symbol of a `For` control variable written as a plain name; reports
    /// a variable already controlling an enclosing `For`.
    fn for_variable(&mut self, arena: &ExprArena, var: NodeId, span: Span) -> Option<u32> {
        let ExprNode::NameRef { sym, .. } = arena.get(var) else { return None };
        let sym = *sym;
        if self.blocks.contains(&Block::For(Some(sym))) {
            self.diagnostics.push(ERR_FOR_VARIABLE_IN_USE, span);
        }
        Some(sym)
    }

    /// End of a `For` body: `Next [var [, var]…]`, or a variable carried over
    /// from an inner `Next a, b`. Names are matched innermost loop first.
    fn finish_next(&mut self, arena: &mut ExprArena) {
        if !self.next_carry.is_empty() {
            let (name, span) = self.next_carry.remove(0);
            self.check_next_variable(name, span);
            return;
        }
        if self.peek().kind != TokenKind::Kw(Kw::Next) {
            self.report_unclosed(ERR_FOR_WITHOUT_NEXT);
            return;
        }
        self.advance(); // Next
        let mut names: Vec<(Option<u32>, Span)> = Vec::new();
        while !self.peek().is_stmt_end() {
            if matches!(self.peek().kind, TokenKind::Kw(Kw::Rem) | TokenKind::Kw(Kw::Colon) | TokenKind::Kw(Kw::Else)) {
                break;
            }
            // Each variable is a target expression; only a plain name can be a control variable.
            let span = self.current_span();
            let var = self.parse_target_expr(arena);
            let name = match arena.get(var) {
                ExprNode::NameRef { sym, .. } => Some(*sym),
                _ => None,
            };
            names.push((name, span));
            if !self.eat(TokenKind::Kw(Kw::Comma)) { break; }
            if self.peek().is_stmt_end() {
                // A comma needs another variable after it.
                let span = self.current_span();
                self.report_expected(span, "variable");
            }
        }
        if names.is_empty() && self.peek().kind == TokenKind::Kw(Kw::Else) {
            // `Next` alone does not end a single-line `If` part; `Next i` does.
            let span = self.current_span();
            self.diagnostics.push(ERR_SYNTAX_ERROR, span);
        }
        self.end_stmt();
        // Each further name closes the next enclosing `For`, provided the
        // `For` blocks are directly nested.
        let directly_nested = self.blocks.iter().rev().take_while(|b| matches!(b, Block::For(_))).count();
        let mut names = names.into_iter();
        if let Some((name, span)) = names.next() {
            self.check_next_variable(name, span);
        }
        for (i, (name, span)) in names.enumerate() {
            if i + 1 < directly_nested {
                self.next_carry.push((name, span));
            } else {
                self.diagnostics.push(ERR_NEXT_WITHOUT_FOR, span);
            }
        }
    }

    /// A name after `Next` must be the control variable of the loop it closes.
    fn check_next_variable(&mut self, name: Option<u32>, span: Span) {
        if let Some(Block::For(Some(var))) = self.blocks.last().copied() {
            if name != Some(var) {
                self.diagnostics.push(ERR_INVALID_NEXT_VARIABLE, span);
            }
        }
    }

    /// Parses a Do loop.
    fn parse_do_stmt(&mut self, arena: &mut ExprArena) -> NodeId {
        self.advance(); // consume Do
        let pre = match self.peek().kind.clone() {
            TokenKind::Kw(Kw::While) => { self.advance(); Some((DoKind::PreWhile, self.parse_expr(arena, 0))) }
            TokenKind::Kw(Kw::Until) => { self.advance(); Some((DoKind::PreUntil, self.parse_expr(arena, 0))) }
            _ => None,
        };
        self.end_stmt();
        self.blocks.push(Block::Do { pre_test: pre.is_some() });
        let body = self.parse_block(arena);
        let loop_span = self.current_span();
        let post = if self.close_block(ERR_DO_WITHOUT_LOOP) {
            let p = match self.peek().kind.clone() {
                TokenKind::Kw(Kw::While) => { self.advance(); Some((DoKind::PostWhile, self.parse_expr(arena, 0))) }
                TokenKind::Kw(Kw::Until) => { self.advance(); Some((DoKind::PostUntil, self.parse_expr(arena, 0))) }
                _ => None,
            };
            // A loop tests its condition at the top or at the bottom, not both.
            if p.is_some() && pre.is_some() {
                self.diagnostics.push(ERR_LOOP_WITHOUT_DO, loop_span);
            }
            self.end_stmt();
            p
        } else {
            None
        };
        self.blocks.pop();
        let (kind, cond) = if let Some((k, c)) = post {
            (k, Some(c))
        } else if let Some((k, c)) = pre {
            (k, Some(c))
        } else {
            (DoKind::Inf, None)
        };
        arena.alloc(ExprNode::Do { kind, cond, body })
    }

    /// Parses a While/Wend loop.
    fn parse_while_stmt(&mut self, arena: &mut ExprArena) -> NodeId {
        self.advance(); // consume While
        let cond = self.parse_expr(arena, 0);
        self.end_stmt();
        self.blocks.push(Block::While);
        let body = self.parse_block(arena);
        if self.close_block(ERR_WHILE_WITHOUT_WEND) {
            self.end_stmt();
        }
        self.blocks.pop();
        arena.alloc(ExprNode::While { cond, body })
    }

    /// Parses a With statement.
    fn parse_with_stmt(&mut self, arena: &mut ExprArena) -> NodeId {
        self.advance(); // consume With
        let obj = self.parse_expr(arena, 0);
        self.end_stmt();
        self.blocks.push(Block::With);
        let body = self.parse_block(arena);
        if self.close_block(ERR_EXPECTED_END_WITH) {
            self.end_stmt();
        }
        self.blocks.pop();
        arena.alloc(ExprNode::With { obj, body })
    }

    /// Parses an If statement.
    fn parse_if_stmt(&mut self, arena: &mut ExprArena) -> NodeId {
        let if_span = self.current_span();
        let not_first_on_line = self.after_colon;
        self.advance(); // consume If
        let cond = self.parse_expr(arena, 0);

        // Legacy single-line form `If <cond> GoTo|GoSub <label>` (no `Then`).
        let legacy_goto = matches!(
            self.peek().kind,
            TokenKind::Kw(Kw::GoTo) | TokenKind::Kw(Kw::Go)
        );
        let has_then = if legacy_goto {
            false
        } else {
            self.expect(TokenKind::Kw(Kw::Then), ERR_EXPECTED_THEN);
            true
        };

        // In the legacy form `Go` is followed by `To`.
        if legacy_goto && self.peek().kind == TokenKind::Kw(Kw::Go) && self.peek2().kind != TokenKind::Kw(Kw::To) {
            let span = self.peek2().span;
            self.report_expected(span, "To");
        }
        // A `Dim` cannot be the statement after `Then`.
        if has_then && self.peek().kind == TokenKind::Kw(Kw::Dim) {
            let span = self.current_span();
            self.report_expected(span, "statement");
        }
        // Single-line If: a `Then` followed immediately by a statement (not Eol),
        // or the legacy `GoTo` form.
        if has_then && self.peek().kind == TokenKind::Kw(Kw::Rem) {
            // A comment cannot stand in for the statement after `Then`.
            let span = self.current_span();
            self.diagnostics.push(ERR_SYNTAX_ERROR, span);
        }
        if legacy_goto || (has_then && !self.peek().is_stmt_end()) {
            let stmts = self.parse_inline_statements(arena);
            let then_body = arena.alloc(ExprNode::Block { stmts });
            let else_body = if !self.last_was_eol && self.eat(TokenKind::Kw(Kw::Else)) {
                let stmts = self.parse_inline_statements(arena);
                Some(arena.alloc(ExprNode::Block { stmts }))
            } else { None };
            return arena.alloc(ExprNode::If { cond, then_body, else_body });
        }
        // A block `If` (`Then` ends the line) must start its line.
        if not_first_on_line {
            self.diagnostics.push(ERR_MUST_BE_FIRST_ON_LINE, if_span);
        }
        self.end_stmt();
        self.blocks.push(Block::If);
        let then_body = self.parse_block(arena);
        let else_body = self.parse_else_clause(arena);
        if self.close_block(ERR_BLOCK_IF_WITHOUT_END_IF) { // `End If` or the one-word `EndIf`
            self.end_stmt();
        }
        self.blocks.pop();
        arena.alloc(ExprNode::If { cond, then_body, else_body })
    }

    /// The statements of one part of a single-line `If`: every statement
    /// (separated by `:`) up to the `Else` or the end of the line.
    fn parse_inline_statements(&mut self, arena: &mut ExprArena) -> Vec<NodeId> {
        let was_inline = std::mem::replace(&mut self.in_inline_if, true);
        let stmts = self.parse_inline_statement_list(arena);
        self.in_inline_if = was_inline;
        stmts
    }

    fn parse_inline_statement_list(&mut self, arena: &mut ExprArena) -> Vec<NodeId> {
        let mut stmts = Vec::new();
        // `If c Then 10`: a line number straight after `Then` / `Else` is a jump.
        if matches!(self.peek().kind, TokenKind::IntLit | TokenKind::LongLit) {
            let tok = self.advance();
            let target = LabelRef::Line(line_label_value(&tok.lit));
            self.label_refs.push((target, tok.span));
            let jump = arena.alloc(ExprNode::GoTo { target });
            self.set_span(jump, tok.span);
            stmts.push(jump);
        }
        loop {
            while self.eat(TokenKind::Kw(Kw::Colon)) {}
            if self.last_was_eol
                || self.peek().is_stmt_end()
                || self.peek().kind == TokenKind::Kw(Kw::Else)
                || (!stmts.is_empty() && self.peek().kind == TokenKind::Kw(Kw::Rem))
            {
                break;
            }
            if let Some(n) = self.parse_stmt(arena) {
                stmts.push(n);
            }
            if self.last_was_eol {
                break;
            }
        }
        stmts
    }

    fn parse_else_clause(&mut self, arena: &mut ExprArena) -> Option<NodeId> {
        match self.peek().kind.clone() {
            TokenKind::Kw(Kw::Else) => {
                if self.after_colon {
                    let span = self.current_span();
                    self.diagnostics.push(ERR_MUST_BE_FIRST_ON_LINE, span);
                }
                self.advance();
                // A statement on the `Else` line does not start a line of its own.
                if !self.peek().is_stmt_end() {
                    self.after_colon = true;
                }
                self.eat_eol();
                if let Some(top) = self.blocks.last_mut() {
                    *top = Block::IfElse;
                }
                Some(self.parse_block(arena))
            }
            TokenKind::Kw(Kw::ElseIf) => {
                if self.after_colon {
                    let span = self.current_span();
                    self.diagnostics.push(ERR_MUST_BE_FIRST_ON_LINE, span);
                }
                self.advance();
                let cond = self.parse_expr(arena, 0);
                self.expect(TokenKind::Kw(Kw::Then), ERR_EXPECTED_THEN);
                if self.peek().kind == TokenKind::Kw(Kw::Rem) {
                    let span = self.current_span();
                    self.diagnostics.push(ERR_SYNTAX_ERROR, span);
                }
                self.eat_eol();
                let then_body = self.parse_block(arena);
                let else_body = self.parse_else_clause(arena);
                Some(arena.alloc(ExprNode::If { cond, then_body, else_body }))
            }
            _ => None,
        }
    }

    /// Parses a Select Case statement.
    fn parse_select_stmt(&mut self, arena: &mut ExprArena) -> NodeId {
        self.advance(); // consume Select
        self.expect(TokenKind::Kw(Kw::Case), ERR_EXPECTED_CASE);
        let subject = self.parse_expr(arena, 0);
        self.end_stmt();
        self.blocks.push(Block::Select);
        let mut pre = Vec::new();
        let mut cases = Vec::new();
        let mut seen_else = false;
        loop {
            while self.eat(TokenKind::Eol) {}
            match self.peek().kind.clone() {
                TokenKind::Kw(Kw::Case) => {
                    let case_span = self.current_span();
                    self.advance();
                    if seen_else {
                        // `Case Else` is the last clause.
                        self.diagnostics.push(ERR_CASE_WITHOUT_SELECT, case_span);
                    }
                    let clause = self.parse_case_clause(arena);
                    seen_else |= matches!(arena.get(clause), ExprNode::CaseElse { .. });
                    cases.push(clause);
                }
                TokenKind::Kw(Kw::End) if self.terminator() == Some(Term::EndSelect) => {
                    self.close_block(ERR_SELECT_WITHOUT_END_SELECT);
                    self.end_stmt();
                    break;
                }
                TokenKind::Eof => break,
                _ if self.at_proc_end() => {
                    self.report_unclosed(ERR_SELECT_WITHOUT_END_SELECT);
                    break;
                }
                _ => {
                    // Statements before the first Case clause.
                    if self.peek().is_stmt_end() {
                        self.advance();
                    } else {
                        let span = self.current_span();
                        if let Some(stmt) = self.parse_stmt(arena) {
                            self.diagnostics.push(ERR_STATEMENT_BEFORE_FIRST_CASE, span);
                            pre.push(stmt);
                        }
                    }
                }
            }
        }
        self.blocks.pop();
        arena.alloc(ExprNode::SelectCase { subject, pre, cases })
    }

    fn parse_case_clause(&mut self, arena: &mut ExprArena) -> NodeId {
        if self.eat(TokenKind::Kw(Kw::Else)) {
            self.end_stmt();
            let body = self.parse_block(arena);
            arena.alloc(ExprNode::CaseElse { body })
        } else {
            let test = self.parse_case_test_expr(arena);
            self.end_stmt();
            let body = self.parse_block(arena);
            arena.alloc(ExprNode::CaseBlock { test, body })
        }
    }

    /// Parses Case test expressions like `Is op expr`, `lo To hi`, or bare `expr`.
    fn parse_case_test_expr(&mut self, arena: &mut ExprArena) -> NodeId {
        let mut args = Vec::new();
        loop {
            let implicit_is = matches!(
                self.peek().kind,
                TokenKind::Kw(Kw::Eq) | TokenKind::Kw(Kw::Ne) | TokenKind::Kw(Kw::Lt)
                    | TokenKind::Kw(Kw::Gt) | TokenKind::Kw(Kw::Le) | TokenKind::Kw(Kw::Ge)
            );
            let t = if implicit_is || self.eat(TokenKind::Kw(Kw::Is)) {
                let op_tok = self.advance();
                let op = match &op_tok.kind {
                    TokenKind::Kw(Kw::Eq) => BinOpKind::Eq,
                    TokenKind::Kw(Kw::Ne) => BinOpKind::Ne,
                    TokenKind::Kw(Kw::Lt) => BinOpKind::Lt,
                    TokenKind::Kw(Kw::Gt) => BinOpKind::Gt,
                    TokenKind::Kw(Kw::Le) => BinOpKind::Le,
                    TokenKind::Kw(Kw::Ge) => BinOpKind::Ge,
                    _ => { self.diagnostics.push(ERR_UNEXPECTED_TOKEN, op_tok.span); BinOpKind::Eq }
                };
                let rhs = self.parse_expr(arena, 0);
                arena.alloc(ExprNode::CaseIs { op, rhs })
            } else {
                let start = self.current_span().start;
                let lo = self.parse_expr(arena, 0);
                if self.eat(TokenKind::Kw(Kw::To)) {
                    let hi = self.parse_expr(arena, 0);
                    let range = arena.alloc(ExprNode::RangeTo { lo, hi });
                    let len = self.last_token_end.saturating_sub(start);
                    self.set_span(range, Span { start, len });
                    range
                } else {
                    lo
                }
            };
            args.push(t);
            if !self.eat(TokenKind::Kw(Kw::Comma)) { break; }
        }
        arena.alloc(ExprNode::ArgList { args })
    }

    fn parse_call_stmt(&mut self, arena: &mut ExprArena) -> NodeId {
        self.advance(); // consume Call
        if self.peek().kind == TokenKind::Kw(Kw::Rem) {
            let span = self.current_span();
            self.diagnostics.push(ERR_SYNTAX_ERROR, span);
            self.skip_line();
            return arena.alloc(ExprNode::Block { stmts: Vec::new() });
        }
        if self.at_reserved_word() && self.peek().kind != TokenKind::Kw(Kw::Me) {
            self.reject_reserved_word();
            self.skip_line();
            return arena.alloc(ExprNode::Block { stmts: Vec::new() });
        }
        let callee = self.parse_target_expr(arena);
        // `Call Foo(1, 2)`: the brackets were read with the callee, so nothing may follow them.
        let args = if matches!(arena.get(callee), ExprNode::Call { .. }) {
            arena.alloc(ExprNode::ArgList { args: Vec::new() })
        } else if self.eat(TokenKind::Kw(Kw::LParen)) {
            let a = self.parse_arg_list(arena);
            self.expect(TokenKind::Kw(Kw::RParen), ERR_EXPECTED_RPAREN);
            a
        } else {
            // `Call Foo` takes its arguments in brackets or not at all.
            arena.alloc(ExprNode::ArgList { args: Vec::new() })
        };
        self.end_stmt();
        arena.alloc(ExprNode::CallStmt { callee, args })
    }

    fn parse_exit_stmt(&mut self, arena: &mut ExprArena) -> NodeId {
        let exit_span = self.current_span();
        self.advance(); // consume Exit
        let kind = match self.peek().kind {
            TokenKind::Kw(Kw::Sub)      => { self.advance(); ExitKind::Sub      }
            TokenKind::Kw(Kw::Function) => { self.advance(); ExitKind::Function }
            TokenKind::Kw(Kw::For)      => { self.advance(); ExitKind::For      }
            TokenKind::Kw(Kw::Do)       => { self.advance(); ExitKind::Do       }
            TokenKind::Kw(Kw::Property) => { self.advance(); ExitKind::Property }
            _ => {
                let span = self.current_span();
                self.diagnostics.push(ERR_UNEXPECTED_TOKEN, span);
                ExitKind::Sub
            }
        };
        self.check_exit_context(kind, exit_span);
        self.end_stmt();
        arena.alloc(ExprNode::ExitStmt { kind })
    }

    /// `Exit For`/`Exit Do` need an enclosing loop of that kind anywhere on the
    /// block stack; `Exit Sub`/`Function`/`Property` must match the procedure.
    fn check_exit_context(&mut self, kind: ExitKind, span: Span) {
        let (ok, code) = match kind {
            ExitKind::For => (
                self.blocks.iter().any(|b| matches!(b, Block::For(_))),
                ERR_EXIT_FOR_OUTSIDE_FOR,
            ),
            ExitKind::Do => (
                self.blocks.iter().any(|b| matches!(b, Block::Do { .. })),
                ERR_EXIT_DO_OUTSIDE_DO,
            ),
            ExitKind::Sub => (
                self.proc_kind.is_none_or(|k| k == ProcKind::Sub),
                ERR_EXIT_SUB_NOT_ALLOWED,
            ),
            ExitKind::Function => (
                self.proc_kind.is_none_or(|k| k == ProcKind::Function),
                ERR_EXIT_FUNCTION_NOT_ALLOWED,
            ),
            ExitKind::Property => (
                self.proc_kind.is_none_or(|k| {
                    matches!(k, ProcKind::PropGet | ProcKind::PropLet | ProcKind::PropSet)
                }),
                ERR_EXIT_PROPERTY_NOT_ALLOWED,
            ),
        };
        if !ok {
            self.diagnostics.push(code, span);
        }
    }

    fn parse_goto_stmt(&mut self, arena: &mut ExprArena, is_gosub: bool) -> NodeId {
        self.advance(); // consume GoTo / GoSub / Go
        // Space-separated forms: a bare `Go` is followed by `To` or `Sub`,
        // which determines the jump kind. The contiguous `GoTo`/`GoSub` tokens
        // are already complete, so neither `eat` matches and `is_gosub` stands.
        let is_gosub = if self.eat(TokenKind::Kw(Kw::To)) {
            is_gosub
        } else if self.eat(TokenKind::Kw(Kw::Sub)) {
            true
        } else {
            is_gosub
        };
        let target_span = self.current_span();
        let target = if matches!(self.peek().kind, TokenKind::IntLit | TokenKind::LongLit) {
            let tok = self.advance();
            LabelRef::Line(line_label_value(&tok.lit))
        } else {
            LabelRef::Name(self.expect_decl_name().sym.map(|s| s as u32).unwrap_or(0))
        };
        self.label_refs.push((target, target_span));
        self.end_stmt();
        if is_gosub {
            arena.alloc(ExprNode::GoSub { target })
        } else {
            arena.alloc(ExprNode::GoTo { target })
        }
    }

    fn parse_resume_stmt(&mut self, arena: &mut ExprArena) -> NodeId {
        self.advance(); // consume Resume
        // VB6 has four forms: `Resume`, `Resume Next`, `Resume <label>`, and
        // `Resume <line#>`. The operand must be consumed and recorded — leaving
        // it in the stream causes it to be mis-parsed as a separate statement.
        let target = if self.eat(TokenKind::Kw(Kw::Next)) {
            ResumeTarget::Next
        } else if matches!(self.peek().kind, TokenKind::IntLit | TokenKind::LongLit) {
            let tok = self.advance();
            let line = line_label_value(&tok.lit);
            // `Resume 0` is `Resume`: it names no label.
            if line != 0 {
                self.label_refs.push((LabelRef::Line(line), tok.span));
            }
            ResumeTarget::At(LabelRef::Line(line))
        } else if !self.peek().is_stmt_end()
            && !matches!(self.peek().kind, TokenKind::Kw(Kw::Colon) | TokenKind::Kw(Kw::Else))
        {
            let lbl = self.expect_decl_name();
            let sym = lbl.sym.map(|s| s as u32).unwrap_or(0);
            self.label_refs.push((LabelRef::Name(sym), lbl.span));
            ResumeTarget::At(LabelRef::Name(sym))
        } else {
            ResumeTarget::Retry
        };
        self.end_stmt();
        arena.alloc(ExprNode::Resume { target })
    }

    fn parse_on_stmt(&mut self, arena: &mut ExprArena) -> NodeId {
        self.advance(); // consume On
        // `On Local Error …` is the explicit form of `On Error …`; `Local` is the
        // default error scope and a reserved word, so eating it here is
        // unambiguous and leaves the handler form unchanged.
        if self.eat(TokenKind::Kw(Kw::Local)) && self.peek().kind != TokenKind::Kw(Kw::Error) {
            // `Local` only qualifies `Error`.
            let span = self.current_span();
            self.report_expected(span, "Error");
            self.skip_line();
            return arena.alloc(ExprNode::OnError { kind: OnErrorKind::Disable });
        }
        if self.eat(TokenKind::Kw(Kw::Error)) {
            if self.eat(TokenKind::Kw(Kw::Resume)) {
                self.expect(TokenKind::Kw(Kw::Next), ERR_EXPECTED_NEXT);
                self.end_stmt();
                return arena.alloc(ExprNode::OnError { kind: OnErrorKind::ResumeNext });
            }
            if self.eat(TokenKind::Kw(Kw::GoTo)) {
                // `On Error GoTo -1` leaves the error state of the handler.
                if self.peek().kind == TokenKind::Kw(Kw::Minus)
                    && matches!(self.peek2().kind, TokenKind::IntLit | TokenKind::LongLit)
                    && line_label_value(&self.peek2().lit) == 1
                {
                    self.advance();
                    self.advance();
                    self.end_stmt();
                    return arena.alloc(ExprNode::OnError { kind: OnErrorKind::ClearError });
                }
                if matches!(self.peek().kind, TokenKind::IntLit | TokenKind::LongLit) {
                    let tok = self.advance();
                    let line = line_label_value(&tok.lit);
                    self.end_stmt();
                    // Only `On Error GoTo 0` disables the handler; any nonzero
                    // line number installs a handler at that numeric line label.
                    let kind = if line == 0 {
                        OnErrorKind::Disable
                    } else {
                        self.label_refs.push((LabelRef::Line(line), tok.span));
                        OnErrorKind::Goto(LabelRef::Line(line))
                    };
                    return arena.alloc(ExprNode::OnError { kind });
                }
                let lbl = self.expect_ident();
                let lid = lbl.sym.map(|s| s as u32).unwrap_or(0);
                self.label_refs.push((LabelRef::Name(lid), lbl.span));
                self.end_stmt();
                return arena.alloc(ExprNode::OnError { kind: OnErrorKind::Goto(LabelRef::Name(lid)) });
            }
        }
        // On <expr> GoTo / GoSub label-list
        let expr = self.parse_expr(arena, 0);
        let is_gosub = match self.peek().kind {
            TokenKind::Kw(Kw::GoTo)  => { self.advance(); false }
            TokenKind::Kw(Kw::GoSub) => { self.advance(); true  }
            _ => {
                let span = self.current_span();
                self.diagnostics.push(ERR_UNEXPECTED_TOKEN, span);
                false
            }
        };
        // `On <expr> GoTo/GoSub <target-list>` — each target is a named label
        // or a numeric line label; both forms appear in real VB6.
        let mut labels = Vec::new();
        loop {
            let lref = if matches!(self.peek().kind, TokenKind::IntLit | TokenKind::LongLit) {
                let tok = self.advance();
                self.label_refs.push((LabelRef::Line(line_label_value(&tok.lit)), tok.span));
                let n = arena.alloc(ExprNode::JumpTarget {
                    target: LabelRef::Line(line_label_value(&tok.lit)),
                });
                self.set_span(n, tok.span);
                n
            } else {
                let lbl = self.expect_ident();
                let lid = lbl.sym.map(|s| s as u32).unwrap_or(0);
                self.label_refs.push((LabelRef::Name(lid), lbl.span));
                let nref = arena.alloc(ExprNode::JumpTarget { target: LabelRef::Name(lid) });
                self.set_span(nref, lbl.span);
                nref
            };
            labels.push(lref);
            if !self.eat(TokenKind::Kw(Kw::Comma)) { break; }
        }
        self.end_stmt();
        arena.alloc(ExprNode::OnGo { is_gosub, expr, labels })
    }

    fn parse_erase_stmt(&mut self, arena: &mut ExprArena) -> NodeId {
        self.advance(); // consume Erase
        let mut vars = Vec::new();
        loop {
            vars.push(self.parse_target_expr(arena));
            if !self.eat(TokenKind::Kw(Kw::Comma)) { break; }
        }
        self.end_stmt();
        arena.alloc(ExprNode::Erase { vars })
    }

    fn parse_raise_event(&mut self, arena: &mut ExprArena) -> NodeId {
        self.advance(); // consume RaiseEvent
        let name = self.expect_ident();
        let nid = name.sym.map(|s| s as u32).unwrap_or(0);
        let args = if self.eat(TokenKind::Kw(Kw::LParen)) {
            let a = self.parse_arg_list(arena);
            self.expect(TokenKind::Kw(Kw::RParen), ERR_EXPECTED_RPAREN);
            a
        } else {
            arena.alloc(ExprNode::ArgList { args: vec![] })
        };
        self.end_stmt();
        let id = arena.alloc(ExprNode::RaiseEvent { name: nid, args });
        self.set_span(id, name.span);
        id
    }

    fn parse_debug_stmt(&mut self, arena: &mut ExprArena) -> NodeId {
        self.advance(); // consume Debug
        self.expect(TokenKind::Kw(Kw::Dot), ERR_UNEXPECTED_TOKEN);
        let is_member = self.peek().kind == TokenKind::Kw(Kw::Print)
            || self.peek().sym.is_some_and(|sym| self.scanner.sym_name(sym as u32).eq_ignore_ascii_case("Assert"));
        if !is_member {
            let span = self.current_span();
            self.report_expected(span, "Print or Assert");
            self.skip_line();
            return arena.alloc(ExprNode::DebugPrint { args: Vec::new() });
        }
        self.advance(); // consume Print / Assert
        let mut args = Vec::new();
        self.parse_print_method_items(arena, &mut args);
        self.end_stmt();
        arena.alloc(ExprNode::DebugPrint { args })
    }

    // ── Conditional compilation ───────────────────────────────────────────────
    //
    // `#Const` / `#If` / `#ElseIf` / `#Else` / `#End If` are resolved on the
    // token stream, before the grammar sees it: a directive line is consumed
    // and the text of every inactive branch is dropped, so the statements of
    // the active branch parse in whatever construct surrounds the directives
    // (a `Case` clause, an `If` block, a procedure, the module) exactly as if
    // the directives were not there.

    /// Next raw token for directive processing (honours the one-token pushback).
    fn cc_peek(&mut self) -> &Token {
        if self.cc_pending.is_none() {
            let t = self.next_physical();
            self.cc_pending = Some(t);
        }
        self.cc_pending.as_ref().unwrap()
    }

    fn cc_advance(&mut self) -> Token {
        match self.cc_pending.take() {
            Some(t) => t,
            None => self.next_physical(),
        }
    }

    /// Consume the rest of a directive line, including its end of line.
    fn cc_skip_line(&mut self) {
        loop {
            match self.cc_peek().kind {
                TokenKind::Eof => return,
                TokenKind::Eol => {
                    self.cc_advance();
                    return;
                }
                _ => {
                    self.cc_advance();
                }
            }
        }
    }

    fn cc_active(&self) -> bool {
        self.cc_stack.last().map_or(true, |f| f.active)
    }

    /// Define the project's conditional-compilation arguments, which are
    /// visible to every `#If` of the module (a `#Const` of the same name
    /// overrides them).
    pub fn set_conditional_constants(&mut self, defs: Vec<(String, CcValue)>) {
        for (name, value) in defs {
            self.cc_set(&name, value);
        }
    }

    /// The end of a directive line: a line end, a comment, or a `:` after
    /// which the line goes on (`#Const A = 1 : #If A Then`). Anything else
    /// is reported and the line skipped.
    fn cc_end_of_directive(&mut self) {
        match self.cc_peek().kind.clone() {
            TokenKind::Eof => {}
            TokenKind::Eol => {
                self.cc_advance();
            }
            TokenKind::Kw(Kw::Apos) | TokenKind::Kw(Kw::Rem) => {
                self.cc_advance();
                if self.cc_peek().kind == TokenKind::Eol {
                    self.cc_advance();
                }
            }
            TokenKind::Kw(Kw::Colon) => {
                self.cc_advance();
            }
            _ => {
                let span = self.cc_peek().span;
                self.diagnostics.push_labeled(ERR_UNEXPECTED_TOKEN, span, "end of statement");
                self.cc_skip_line();
            }
        }
    }

    /// `#If` / `#ElseIf` condition: evaluated only when the branch can still
    /// be taken; the expression is followed by `Then`.
    fn cc_condition(&mut self, evaluate: bool) -> bool {
        if !evaluate {
            self.cc_skip_line();
            return false;
        }
        let value = self.eval_cc_expr(0).is_true();
        if self.cc_peek().kind == TokenKind::Kw(Kw::Then) {
            self.cc_advance();
            self.cc_end_of_directive();
        } else {
            let span = self.cc_peek().span;
            self.diagnostics.push(ERR_EXPECTED_THEN, span);
            self.cc_skip_line();
        }
        value
    }

    /// Handle the directive token `kind` (already consumed). Returns after the
    /// directive's whole line has been consumed.
    fn cc_directive(&mut self, kind: Kw, span: Span, one_word: bool) {
        match kind {
            Kw::CcConst => {
                if self.cc_active() {
                    let name_tok = self.cc_peek().clone();
                    let name_ok = name_tok.kind == TokenKind::Ident
                        || (matches!(name_tok.kind, TokenKind::Kw(_)) && name_tok.sym.is_some());
                    if !name_ok {
                        self.diagnostics.push_labeled(ERR_EXPECTED_IDENT, name_tok.span, "Identifier");
                        self.cc_skip_line();
                        return;
                    }
                    self.cc_advance();
                    let name = name_tok
                        .sym
                        .map(|id| self.scanner.sym_name(id as u32).to_ascii_lowercase())
                        .unwrap_or_default();
                    if self.cc_peek().kind == TokenKind::Kw(Kw::Eq) {
                        self.cc_advance();
                    } else {
                        let span = self.cc_peek().span;
                        self.diagnostics.push(ERR_EXPECTED_EQ, span);
                    }
                    let value = self.eval_cc_expr(0);
                    self.cc_set(&name, value);
                    self.cc_end_of_directive();
                } else {
                    self.cc_skip_line();
                }
            }
            Kw::CcIf => {
                let parent_active = self.cc_active();
                let taken = self.cc_condition(parent_active);
                self.cc_stack.push(CcFrame { open: span, else_seen: false, parent_active, active: taken, taken });
            }
            Kw::CcElseIf => {
                let Some(frame) = self.cc_stack.last().copied() else {
                    self.diagnostics.push(ERR_CC_WITHOUT_IF, span);
                    self.cc_skip_line();
                    return;
                };
                if frame.else_seen {
                    self.diagnostics.push(ERR_CC_ELSEIF_AFTER_ELSE, span);
                    self.cc_skip_line();
                    return;
                }
                let evaluate = frame.parent_active && !frame.taken;
                let cond = self.cc_condition(evaluate);
                let top = self.cc_stack.last_mut().unwrap();
                top.active = cond;
                top.taken |= cond;
            }
            Kw::CcElse => {
                if self.cc_active() || self.cc_stack.last().is_some_and(|f| f.parent_active) {
                    self.cc_end_of_directive();
                } else {
                    self.cc_skip_line();
                }
                match self.cc_stack.last_mut() {
                    None => self.diagnostics.push(ERR_CC_WITHOUT_IF, span),
                    Some(top) if top.else_seen => self.diagnostics.push(ERR_CC_DUPLICATE_ELSE, span),
                    Some(top) => {
                        top.else_seen = true;
                        top.active = top.parent_active && !top.taken;
                        top.taken = true;
                    }
                }
            }
            _ => {
                // `#End If`, or the one-word `#EndIf`.
                if !one_word {
                    if self.cc_peek().kind == TokenKind::Kw(Kw::If) {
                        self.cc_advance();
                    } else if self.cc_stack.last().is_some_and(|f| f.parent_active) {
                        let span = self.cc_peek().span;
                        self.diagnostics.push_labeled(ERR_UNEXPECTED_TOKEN, span, "If");
                    }
                }
                if self.cc_stack.last().is_some_and(|f| f.parent_active) {
                    self.cc_end_of_directive();
                } else {
                    self.cc_skip_line();
                }
                if self.cc_stack.pop().is_none() {
                    self.diagnostics.push(ERR_CC_WITHOUT_IF, span);
                }
            }
        }
    }

    /// At the end of the module, report the innermost `#If` that was never
    /// closed (once).
    fn report_open_cc_blocks(&mut self) {
        if let Some(frame) = self.cc_stack.last().copied() {
            self.diagnostics.push(ERR_CC_IF_NOT_TERMINATED, frame.open);
            self.cc_stack.clear();
        }
    }

    // ── CC expression evaluator ───────────────────────────────────────────────

    /// Look up a CC constant by (lowercase) name; an unknown name is Empty.
    fn cc_lookup(&self, name: &str) -> CcValue {
        self.cc_defs.iter().rev()
            .find(|(n, _)| n == name)
            .map(|(_, v)| v.clone())
            .unwrap_or(CcValue::Empty)
    }

    /// Set (or insert) a CC constant.
    fn cc_set(&mut self, name: &str, value: CcValue) {
        if let Some(entry) = self.cc_defs.iter_mut().find(|(n, _)| n == name) {
            entry.1 = value;
        } else {
            self.cc_defs.push((name.to_string(), value));
        }
    }

    /// Evaluate a CC prefix (unary) operator, or fall through to a primary.
    /// `^` binds tighter than a sign, so `-2 ^ 2` is `-(2 ^ 2)`.
    fn eval_cc_prefix(&mut self) -> CcValue {
        match self.cc_peek().kind.clone() {
            TokenKind::Kw(Kw::Not) => {
                self.cc_advance();
                self.eval_cc_expr(6).not()
            }
            TokenKind::Kw(Kw::Minus) => {
                self.cc_advance();
                self.eval_cc_expr(20).negate()
            }
            TokenKind::Kw(Kw::Plus) => {
                self.cc_advance();
                self.eval_cc_expr(20)
            }
            _ => self.eval_cc_primary(),
        }
    }

    /// Binding powers + operator for a CC infix operator, or `None`.
    fn cc_infix_op(&mut self) -> Option<(u8, u8, CcOp)> {
        let entry = match self.cc_peek().kind {
            TokenKind::Kw(Kw::Caret)      => (22, 23, CcOp::Pow),
            TokenKind::Kw(Kw::Star)       => (18, 19, CcOp::Mul),
            TokenKind::Kw(Kw::Slash)      => (18, 19, CcOp::Div),
            TokenKind::Kw(Kw::Backslash)  => (16, 17, CcOp::IDiv),
            TokenKind::Kw(Kw::Mod)        => (14, 15, CcOp::Mod),
            TokenKind::Kw(Kw::Plus)       => (12, 13, CcOp::Add),
            TokenKind::Kw(Kw::Minus)      => (12, 13, CcOp::Sub),
            TokenKind::Kw(Kw::Amp)        => (10, 11, CcOp::Cat),
            TokenKind::Kw(Kw::Eq)         => (8,  9,  CcOp::Eq),
            TokenKind::Kw(Kw::Ne)         => (8,  9,  CcOp::Ne),
            TokenKind::Kw(Kw::Lt)         => (8,  9,  CcOp::Lt),
            TokenKind::Kw(Kw::Gt)         => (8,  9,  CcOp::Gt),
            TokenKind::Kw(Kw::Le)         => (8,  9,  CcOp::Le),
            TokenKind::Kw(Kw::Ge)         => (8,  9,  CcOp::Ge),
            TokenKind::Kw(Kw::And)        => (6,  7,  CcOp::And),
            TokenKind::Kw(Kw::Or)         => (4,  5,  CcOp::Or),
            TokenKind::Kw(Kw::Xor)        => (4,  5,  CcOp::Xor),
            TokenKind::Kw(Kw::Eqv)        => (2,  3,  CcOp::Eqv),
            TokenKind::Kw(Kw::Imp)        => (2,  3,  CcOp::Imp),
            _ => return None,
        };
        Some(entry)
    }

    /// Evaluate a CC expression.
    fn eval_cc_expr(&mut self, min_bp: u8) -> CcValue {
        let mut lhs = self.eval_cc_prefix();
        loop {
            let Some((l_bp, r_bp, op)) = self.cc_infix_op() else { break };
            if l_bp <= min_bp { break; }
            self.cc_advance();
            let rhs = self.eval_cc_expr(r_bp);
            lhs = CcValue::binary(op, &lhs, &rhs);
        }
        lhs
    }

    fn eval_cc_primary(&mut self) -> CcValue {
        use crate::frontend::token::Lit;
        match self.cc_peek().kind.clone() {
            TokenKind::Kw(Kw::LParen) => {
                self.cc_advance();
                let v = self.eval_cc_expr(0);
                if self.cc_peek().kind == TokenKind::Kw(Kw::RParen) {
                    self.cc_advance();
                } else {
                    let span = self.cc_peek().span;
                    self.diagnostics.push(ERR_EXPECTED_RPAREN, span);
                }
                v
            }
            TokenKind::IntLit | TokenKind::LongLit => {
                let tok = self.cc_advance();
                match tok.lit {
                    Some(Lit::Int(n)) | Some(Lit::Long(n)) => CcValue::Int(n),
                    _ => CcValue::Int(0),
                }
            }
            TokenKind::SngLit | TokenKind::DblLit | TokenKind::DateLit => {
                let tok = self.cc_advance();
                match tok.lit {
                    Some(Lit::Single(f)) => CcValue::Double(f as f64),
                    Some(Lit::Double(d)) | Some(Lit::Date(d)) => CcValue::Double(d),
                    _ => CcValue::Int(0),
                }
            }
            TokenKind::CurLit => {
                let tok = self.cc_advance();
                match tok.lit {
                    Some(Lit::Currency(c)) => CcValue::Double(c as f64 / 10_000.0),
                    _ => CcValue::Int(0),
                }
            }
            TokenKind::StrLit => {
                let tok = self.cc_advance();
                match tok.lit {
                    Some(Lit::Str(s)) => CcValue::Str(s.to_string()),
                    _ => CcValue::Str(String::new()),
                }
            }
            TokenKind::Kw(Kw::True)  => { self.cc_advance(); CcValue::TRUE }
            TokenKind::Kw(Kw::False) => { self.cc_advance(); CcValue::FALSE }
            TokenKind::Kw(Kw::Empty) | TokenKind::Kw(Kw::Null) => {
                self.cc_advance();
                CcValue::Empty
            }
            TokenKind::Eol
            | TokenKind::Eof
            | TokenKind::Kw(Kw::Then)
            | TokenKind::Kw(Kw::Colon)
            | TokenKind::Kw(Kw::Apos)
            | TokenKind::Kw(Kw::Rem)
            | TokenKind::Kw(Kw::RParen) => {
                // A missing operand: reported without consuming the token that
                // ends the directive.
                let span = self.cc_peek().span;
                self.diagnostics.push(ERR_UNEXPECTED_TOKEN, span);
                CcValue::Empty
            }
            _ => {
                let tok = self.cc_advance();
                if tok.kind == TokenKind::Ident || matches!(tok.kind, TokenKind::Kw(_)) {
                    if let Some(id) = tok.sym {
                        let name = self.scanner.sym_name(id as u32).to_ascii_lowercase();
                        return self.cc_lookup(&name);
                    }
                }
                self.diagnostics.push(ERR_UNEXPECTED_TOKEN, tok.span);
                CcValue::Empty
            }
        }
    }

    // ── Option statement parser ───────────────────────────────────────────────

    fn parse_option_stmt(&mut self, arena: &mut ExprArena) -> NodeId {
        let option_span = self.current_span();
        self.advance(); // consume Option
        match self.peek().kind.clone() {
            TokenKind::Kw(Kw::Explicit) => {
                if std::mem::replace(&mut self.seen_option_explicit, true) {
                    self.diagnostics.push(ERR_DUPLICATE_OPTION, option_span);
                }
                self.advance();
                self.end_stmt();
                arena.alloc(ExprNode::OptionExplicit)
            }
            TokenKind::Kw(Kw::Base) => {
                if std::mem::replace(&mut self.seen_option_base, true) {
                    self.diagnostics.push(ERR_DUPLICATE_OPTION, option_span);
                } else if self.module_array_seen {
                    self.diagnostics.push(ERR_ARRAY_ALREADY_DIMENSIONED, option_span);
                }
                self.advance();
                // The base is the literal 0 or 1.
                let value = match (self.peek().kind.clone(), self.peek().lit.clone()) {
                    (TokenKind::IntLit, Some(crate::frontend::token::Lit::Int(n))) if n == 0 || n == 1 => {
                        self.advance();
                        n as u8
                    }
                    _ => {
                        let span = self.current_span();
                        self.diagnostics.push(ERR_UNEXPECTED_TOKEN, span);
                        self.skip_line();
                        return arena.alloc(ExprNode::Block { stmts: vec![] });
                    }
                };
                self.end_stmt();
                arena.alloc(ExprNode::OptionBase { value })
            }
            TokenKind::Kw(Kw::Compare) => {
                if std::mem::replace(&mut self.seen_option_compare, true) {
                    self.diagnostics.push(ERR_DUPLICATE_OPTION, option_span);
                }
                self.advance();
                let mode = match self.peek().kind.clone() {
                    TokenKind::Kw(Kw::Binary) => { self.advance(); 0u8 }
                    TokenKind::Kw(Kw::Text)   => { self.advance(); 1u8 }
                    TokenKind::Kw(Kw::Database) => { self.advance(); 2u8 }
                    _ => {
                        let span = self.current_span();
                        self.diagnostics.push(ERR_UNEXPECTED_TOKEN, span);
                        self.skip_line();
                        return arena.alloc(ExprNode::Block { stmts: vec![] });
                    }
                };
                self.end_stmt();
                arena.alloc(ExprNode::OptionCompare { mode })
            }
            TokenKind::Kw(Kw::Private) => {
                if std::mem::replace(&mut self.seen_option_private, true) {
                    self.diagnostics.push(ERR_DUPLICATE_OPTION, option_span);
                }
                if self.module_kind != ModuleKind::Standard {
                    self.diagnostics.push(ERR_OPTION_PRIVATE_IN_OBJECT_MODULE, option_span);
                }
                self.advance(); // consume Private
                if self.peek().kind == TokenKind::Kw(Kw::Module) {
                    self.advance();
                    self.end_stmt();
                } else {
                    let span = self.current_span();
                    self.diagnostics.push(ERR_UNEXPECTED_TOKEN, span);
                    self.skip_line();
                }
                arena.alloc(ExprNode::Block { stmts: vec![] })
            }
            _ => {
                let span = self.current_span();
                self.diagnostics.push(ERR_UNEXPECTED_TOKEN, span);
                self.skip_line();
                arena.alloc(ExprNode::Block { stmts: vec![] })
            }
        }
    }

    // ── Implements ───────────────────────────────────────────────────────────

    fn parse_implements_stmt(&mut self, arena: &mut ExprArena) -> NodeId {
        self.reject_in_standard_module();
        self.advance(); // consume Implements
        if self.peek().kind == TokenKind::Kw(Kw::Object) {
            // `Object` is a type, not an interface that can be implemented.
            let span = self.current_span();
            self.diagnostics.push_labeled(ERR_EXPECTED_IDENT, span, "identifier");
        }
        let first = self.expect_decl_name();
        let first_id = first.sym.map(|s| s as u32).unwrap_or(0);
        // The node spans the interface name (after any `Library.` qualifier).
        let (library, nid, name_span) = if self.eat(TokenKind::Kw(Kw::Dot)) {
            let name = self.expect_ident();
            (Some(first_id), name.sym.map(|s| s as u32).unwrap_or(0), name.span)
        } else {
            (None, first_id, first.span)
        };
        self.end_stmt();
        let id = arena.alloc(ExprNode::Implements { library, name: nid });
        self.set_span(id, name_span);
        id
    }

    /// `Def<Type> letter[-letter][, letter[-letter]]…`
    fn parse_def_stmt(&mut self, arena: &mut ExprArena) -> NodeId {
        let type_kw = match self.peek().kind {
            TokenKind::Kw(k) => k as u16,
            _ => 0,
        };
        let def_span = self.current_span();
        self.advance(); // consume the Def<Type> keyword
        let mut ranges = Vec::new();
        let mut letters = 0u32;
        loop {
            let lo = self.def_letter();
            let hi = if self.eat(TokenKind::Kw(Kw::Minus)) {
                self.def_letter()
            } else {
                0
            };
            letters |= self.letter_range_mask(lo, if hi == 0 { lo } else { hi });
            ranges.push((lo, hi));
            if !self.eat(TokenKind::Kw(Kw::Comma)) {
                break;
            }
            // A comma may end the list.
            if self.peek().is_stmt_end() {
                break;
            }
        }
        if self.module_decl_seen {
            self.diagnostics.push(ERR_DEFTYPE_AFTER_DECLARATION, def_span);
        } else if letters & self.deftype_letters != 0 {
            self.diagnostics.push(ERR_DUPLICATE_DEFTYPE, def_span);
        }
        self.deftype_letters |= letters;
        self.end_stmt();
        arena.alloc(ExprNode::DefType { type_kw, ranges })
    }

    /// A `Def<Type>` letter: an identifier of exactly one character.
    fn def_letter(&mut self) -> u32 {
        let tok = self.expect_ident();
        let sym = tok.sym.map(|s| s as u32).unwrap_or(0);
        if tok.sym.is_some() && self.scanner.sym_name(sym).len() != 1 {
            self.diagnostics.push(ERR_UNEXPECTED_TOKEN, tok.span);
        }
        sym
    }

    /// Bit mask (bit 0 = `A`) of the letters from the first letter of the
    /// name `lo` to that of `hi`; empty when either is not a letter or the
    /// range runs backwards.
    fn letter_range_mask(&self, lo: u32, hi: u32) -> u32 {
        let letter = |sym: u32| {
            self.scanner
                .sym_name(sym)
                .bytes()
                .next()
                .filter(u8::is_ascii_alphabetic)
                .map(|b| b.to_ascii_lowercase() - b'a')
        };
        match (letter(lo), letter(hi)) {
            (Some(a), Some(b)) if a <= b => ((1u32 << (b + 1)) - 1) & !((1u32 << a) - 1),
            _ => 0,
        }
    }

    /// `Attribute name = value[, value…]` metadata line.
    fn parse_attribute_stmt(&mut self, arena: &mut ExprArena) -> NodeId {
        self.advance(); // consume Attribute
        // `Attribute VB_Name = …` or `Attribute Member.VB_UserMemId = …`: the
        // attribute is the last name; a preceding name is the member it applies to.
        // A line that does not begin with a name is not read at all.
        if self.peek().kind == TokenKind::Kw(Kw::As) {
            // `As` here is read as the start of a type clause.
            let span = self.current_span();
            self.diagnostics.push_labeled(0x9c6f, span, "New or type name");
            return self.ignore_attribute_line(arena);
        }
        if !self.at_attribute_word() || self.peek().kind == TokenKind::Kw(Kw::Attribute) {
            return self.ignore_attribute_line(arena);
        }
        let mut key = self.advance();
        let mut target: Option<u32> = None;
        // The dot joins names only when it follows the name directly.
        while self.peek().kind == TokenKind::Kw(Kw::Dot) && !self.blank_before_next() {
            self.advance();
            if !self.at_attribute_word() {
                return self.ignore_attribute_line(arena);
            }
            target = Some(key.sym.map(|s| s as u32).unwrap_or(0));
            key = self.advance();
        }
        let name = key.sym.map(|s| s as u32).unwrap_or(0);
        if !self.eat(TokenKind::Kw(Kw::Eq)) {
            let span = self.current_span();
            self.diagnostics.push(ERR_SYNTAX_ERROR, span);
            self.skip_line();
            return arena.alloc(ExprNode::Block { stmts: Vec::new() });
        }
        // Without a member before it, the name must be one of the header words.
        if target.is_none() && !is_header_word(self.scanner.sym_name(name)) {
            self.diagnostics.push(ERR_SYNTAX_ERROR, key.span);
            self.skip_line();
            return arena.alloc(ExprNode::Block { stmts: Vec::new() });
        }
        // The node spans its first value: for `VB_Name` that is where the
        // module name is written. Whatever follows the `=` is only read for
        // its values: a value that does not parse is not an error.
        let value_span = self.current_span();
        let mark = self.diagnostics.items().len();
        let mut values = Vec::new();
        if !self.peek().is_stmt_end() && self.peek().kind != TokenKind::Kw(Kw::Colon) {
            loop {
                values.push(self.parse_expr(arena, 0));
                if !self.eat(TokenKind::Kw(Kw::Comma)) {
                    break;
                }
            }
        }
        self.diagnostics.truncate(mark);
        while !self.peek().is_stmt_end() {
            self.advance();
        }
        self.end_stmt();
        let id = arena.alloc(ExprNode::Attribute { target, name, values });
        self.set_span(id, value_span);
        id
    }

    /// True when the next token can be a name in an `Attribute` line: a plain
    /// name, a keyword that is not reserved, or one of the header words.
    fn at_attribute_word(&mut self) -> bool {
        let t = self.peek();
        match (&t.kind, t.sym) {
            (TokenKind::Ident, _) => true,
            (TokenKind::Kw(_), Some(sym)) => {
                let sym = sym as u32;
                !self.scanner.is_reserved(sym) || is_header_word(self.scanner.sym_name(sym))
            }
            _ => false,
        }
    }

    /// True at `Else` or a line number that is followed, directly or after a colon, by an
    /// attribute line proper (`Attribute`, a name, any `.name` parts and an `=`). Such a line
    /// is an attribute line whatever precedes it; a bare `Attribute` after the prefix is not.
    fn at_prefixed_attribute_line(&mut self) -> bool {
        let prefix = matches!(
            self.peek().kind,
            TokenKind::Kw(Kw::Else) | TokenKind::IntLit | TokenKind::LongLit
        );
        if !prefix {
            return false;
        }
        let end = (self.peek().span.start + self.peek().span.len) as usize;
        let rest = line_rest(self.src, end);
        let mut i = skip_blanks(rest, 0);
        if rest.get(i) == Some(&b':') {
            i = skip_blanks(rest, i + 1);
        }
        attribute_line(&rest[i..])
    }

    /// True when the rest of the line holds a colon followed by an attribute line proper.
    fn line_has_trailing_attribute_line(&mut self) -> bool {
        let start = self.peek().span.start as usize;
        let rest = line_rest(self.src, start);
        let mut from = 0;
        while let Some(colon) = rest[from..].iter().position(|b| *b == b':') {
            let at = skip_blanks(rest, from + colon + 1);
            if attribute_line(&rest[at..]) {
                return true;
            }
            from += colon + 1;
        }
        false
    }

    /// The rest of an `Attribute` line that the compiler does not read.
    fn ignore_attribute_line(&mut self, arena: &mut ExprArena) -> NodeId {
        while !self.peek().is_stmt_end() {
            self.advance();
        }
        self.end_stmt();
        arena.alloc(ExprNode::Block { stmts: Vec::new() })
    }

    // ── File I/O statement parsers ──────────────────────────────────────────

    /// `Open path [For mode] [Access access] [Lock lock | Shared] As [#]filenum [Len=n]`
    ///
    /// `mode` is packed into `open_mode` (Input=1, Output=2, Random=4,
    /// Append=8, Binary=0x20; the runtime's default when `For` is omitted
    /// is Random). `open_flags` packs `(lock_code << 4) | access_code`
    /// (access: none=0, Read=1, Write=2, ReadWrite=3; lock: none=0,
    /// `Lock Read Write`=1, `Lock Write`=2, `Lock Read`=3, `Shared`=4).
    fn parse_file_open(&mut self, arena: &mut ExprArena) -> NodeId {
        self.advance(); // Open
        let path = self.parse_expr(arena, 0);

        let mut open_mode: u8 = 4; // default: Random
        if self.eat(TokenKind::Kw(Kw::For)) {
            open_mode = match self.peek().kind {
                TokenKind::Kw(Kw::Input) => { self.advance(); 1 }
                TokenKind::Kw(Kw::Output) => { self.advance(); 2 }
                TokenKind::Kw(Kw::Random) => { self.advance(); 4 }
                TokenKind::Kw(Kw::Append) => { self.advance(); 8 }
                TokenKind::Kw(Kw::Binary) => { self.advance(); 0x20 }
                _ => open_mode,
            };
        }

        let mut access_code: u8 = 0;
        let access_span = self.current_span();
        if self.eat(TokenKind::Kw(Kw::Access)) {
            let has_read = self.eat(TokenKind::Kw(Kw::Read));
            let has_write = self.eat(TokenKind::Kw(Kw::Write));
            // The access must be possible in the open mode: Read is not
            // available to Output, Write not to Input, and Read Write to
            // neither Input nor Output.
            let conflict = (has_read && open_mode == 2)
                || (has_write && open_mode == 1)
                || (has_read && has_write && open_mode == 2);
            if conflict {
                self.diagnostics.push(ERR_BAD_FILE_MODE, access_span);
            }
            access_code = match (has_read, has_write) {
                (true, true) => 3,
                (true, false) => 1,
                (false, true) => 2,
                (false, false) => 0,
            };
        }

        let mut lock_code: u8 = 0;
        if self.eat(TokenKind::Kw(Kw::Shared)) {
            lock_code = 4;
        } else if self.eat(TokenKind::Kw(Kw::Lock)) {
            let has_read = self.eat(TokenKind::Kw(Kw::Read));
            let has_write = self.eat(TokenKind::Kw(Kw::Write));
            lock_code = match (has_read, has_write) {
                (true, true) => 1,
                (true, false) => 3,
                (false, true) => 2,
                // Bare `Lock` with neither `Read` nor `Write` is not a
                // grounded shape (never captured); 0xf marks it so codegen
                // rejects it loudly instead of guessing a byte.
                (false, false) => 0xf,
            };
        }
        let open_flags = (lock_code << 4) | access_code;

        self.eat(TokenKind::Kw(Kw::As));
        self.eat(TokenKind::Kw(Kw::Hash));
        let ch = self.parse_expr(arena, 0);
        let mut args = vec![path];
        if self.eat(TokenKind::Kw(Kw::Len)) {
            self.expect(TokenKind::Kw(Kw::Eq), ERR_EXPECTED_EQ);
            args.push(self.parse_expr(arena, 0)); // `Len = <expr>` is a real use site
        }
        self.end_stmt();
        arena.alloc(ExprNode::FileIoStmt {
            kind: FileIoKind::Open,
            channel: Some(ch),
            args,
            open_mode,
            open_flags,
        })
    }

    /// `Close [[#]filenum, ...]`
    fn parse_file_close(&mut self, arena: &mut ExprArena) -> NodeId {
        self.advance(); // Close
        let mut chs = Vec::new();
        while !self.peek().is_stmt_end()
            && !matches!(self.peek().kind, TokenKind::Kw(Kw::Colon) | TokenKind::Kw(Kw::Else))
        {
            self.eat(TokenKind::Kw(Kw::Hash));
            chs.push(self.parse_expr(arena, 0));
            if !self.eat(TokenKind::Kw(Kw::Comma)) { break; }
        }
        self.end_stmt();
        arena.alloc(ExprNode::FileIoStmt { kind: FileIoKind::Close, channel: None, args: chs, open_mode: 0, open_flags: 0 })
    }

    /// `Print #filenum, [args]` or `Write #filenum, [args]`.
    fn parse_file_print_after_keyword(&mut self, arena: &mut ExprArena, kind: FileIoKind) -> NodeId {
        self.eat(TokenKind::Kw(Kw::Hash));
        let ch = self.parse_expr(arena, 0);
        let mut args = Vec::new();
        if self.expect_comma() {
            self.parse_print_method_items(arena, &mut args);
        }
        self.end_stmt();
        arena.alloc(ExprNode::FileIoStmt { kind, channel: Some(ch), args, open_mode: 0, open_flags: 0 })
    }

    /// `Input #filenum, varlist`.
    fn parse_file_input_after_keyword(&mut self, arena: &mut ExprArena) -> NodeId {
        self.eat(TokenKind::Kw(Kw::Hash));
        let ch = self.parse_expr(arena, 0);
        let mut vars = Vec::new();
        if self.expect_comma() {
            loop {
                vars.push(self.parse_expr(arena, 0));
                if !self.eat(TokenKind::Kw(Kw::Comma)) { break; }
            }
        }
        self.end_stmt();
        arena.alloc(ExprNode::FileIoStmt { kind: FileIoKind::Input, channel: Some(ch), args: vars, open_mode: 0, open_flags: 0 })
    }

    /// `Get #filenum, [recnum], var` or `Put #filenum, [recnum], val`.
    fn parse_file_get_put_after_keyword(&mut self, arena: &mut ExprArena, kind: FileIoKind) -> NodeId {
        self.eat(TokenKind::Kw(Kw::Hash));
        let ch = self.parse_expr(arena, 0);
        let mut args = Vec::new();
        // `Get/Put [#]file, [record], variable`: both commas are required.
        if self.expect_comma() {
            if self.peek().kind != TokenKind::Kw(Kw::Comma) {
                args.push(self.parse_expr(arena, 0));
            }
            if self.expect_comma() {
                args.push(self.parse_expr(arena, 0));
            }
        }
        self.end_stmt();
        arena.alloc(ExprNode::FileIoStmt { kind, channel: Some(ch), args, open_mode: 0, open_flags: 0 })
    }

    /// Consume a `,`, or report that one is expected; true when it was there.
    fn expect_comma(&mut self) -> bool {
        if self.eat(TokenKind::Kw(Kw::Comma)) {
            return true;
        }
        let span = self.current_span();
        self.diagnostics.push(ERR_UNEXPECTED_TOKEN, span);
        false
    }

    /// `Lock #filenum [, from [To to]]` or `Unlock ...`.
    fn parse_file_lock_after_keyword(&mut self, arena: &mut ExprArena, kind: FileIoKind) -> NodeId {
        self.eat(TokenKind::Kw(Kw::Hash));
        let ch = self.parse_expr(arena, 0);
        let mut args = Vec::new();
        if self.eat(TokenKind::Kw(Kw::Comma)) {
            if self.peek().kind == TokenKind::Kw(Kw::To) {
                // `Lock #1, To 10`: the range starts at the first record.
                self.advance();
                args.push(arena.alloc(ExprNode::Literal { lit: AstLit::Int(0) }));
                args.push(self.parse_expr(arena, 0));
            } else {
                args.push(self.parse_expr(arena, 0));
                if self.eat(TokenKind::Kw(Kw::To)) {
                    args.push(self.parse_expr(arena, 0));
                }
            }
        }
        self.end_stmt();
        arena.alloc(ExprNode::FileIoStmt { kind, channel: Some(ch), args, open_mode: 0, open_flags: 0 })
    }

    /// Parse ident-statement starting from a token already consumed from the stream.
    fn parse_ident_stmt_with_tok(&mut self, arena: &mut ExprArena, first: Token) -> NodeId {
        let sym = first.sym.map(|s| s as u32).unwrap_or(0);
        let lhs_base = arena.alloc(ExprNode::NameRef { sym, suffix: first.type_suffix });
        self.set_span(lhs_base, first.span);
        // Apply any postfix operators (.member, (args), etc.)
        let lhs = self.parse_postfix(arena, lhs_base);
        self.finish_ident_stmt(arena, lhs)
    }

    // ── Procedure / declaration parsers ───────────────────────────────────────

    /// Parse an access-modifier prefix (Public/Private/Friend/Static) then dispatch.
    fn parse_access_decl(&mut self, arena: &mut ExprArena) -> NodeId {
        let first = self.advance();
        let mut access_tok = first.kind.clone();
        // `Static` and an access modifier combine in either order
        // (`Static Private Sub`, `Private Static Sub`): every local of the
        // procedure is then static.
        let mut static_proc = access_tok == TokenKind::Kw(Kw::Static);
        let mut access_span = first.span;
        if static_proc {
            if matches!(
                self.peek().kind,
                TokenKind::Kw(Kw::Public) | TokenKind::Kw(Kw::Private) | TokenKind::Kw(Kw::Friend)
            ) {
                access_span = self.current_span();
                access_tok = self.advance().kind.clone();
            }
        } else if self.eat(TokenKind::Kw(Kw::Static)) {
            static_proc = true;
        }
        // `Static` outside a procedure only qualifies a procedure.
        if access_tok == TokenKind::Kw(Kw::Static)
            && !matches!(
                self.peek().kind,
                TokenKind::Kw(Kw::Sub) | TokenKind::Kw(Kw::Function) | TokenKind::Kw(Kw::Property)
            )
        {
            self.diagnostics.push(ERR_INVALID_OUTSIDE_PROCEDURE, first.span);
            self.skip_to_stmt_end();
            return arena.alloc(ExprNode::Block { stmts: vec![] });
        }
        let vis = match access_tok {
            TokenKind::Kw(Kw::Public) | TokenKind::Kw(Kw::Global) | TokenKind::Kw(Kw::Friend) => {
                Some(true)
            }
            TokenKind::Kw(Kw::Private) => Some(false),
            _ => None,
        };
        let is_friend = access_tok == TokenKind::Kw(Kw::Friend);
        if is_friend {
            self.reject_in_standard_module_at(access_span);
        }
        // `Friend` and a combined `Static` only qualify a procedure.
        let procedure_only = is_friend || (static_proc && access_tok != TokenKind::Kw(Kw::Static));
        if procedure_only
            && !matches!(
                self.peek().kind,
                TokenKind::Kw(Kw::Sub) | TokenKind::Kw(Kw::Function) | TokenKind::Kw(Kw::Property)
            )
        {
            let span = self.current_span();
            self.diagnostics.push_labeled(ERR_UNEXPECTED_TOKEN, span, "Function or Sub or Property");
            self.skip_to_stmt_end();
            return arena.alloc(ExprNode::Block { stmts: vec![] });
        }
        // An event is always public.
        if self.peek().kind == TokenKind::Kw(Kw::Event)
            && matches!(access_tok, TokenKind::Kw(Kw::Private) | TokenKind::Kw(Kw::Global) | TokenKind::Kw(Kw::Friend))
        {
            let span = self.current_span();
            self.diagnostics.push(ERR_UNEXPECTED_TOKEN, span);
            self.skip_to_stmt_end();
            return arena.alloc(ExprNode::Block { stmts: vec![] });
        }
        let node = match self.peek().kind.clone() {
            TokenKind::Kw(Kw::Sub)      => self.parse_proc_decl(arena, ProcKind::Sub),
            TokenKind::Kw(Kw::Function) => self.parse_proc_decl(arena, ProcKind::Function),
            TokenKind::Kw(Kw::Property) => self.parse_property_decl(arena),
            TokenKind::Kw(Kw::Declare)  => self.parse_declare_stmt(arena),
            TokenKind::Kw(Kw::Type)     => self.parse_type_decl(arena),
            TokenKind::Kw(Kw::Enum)     => self.parse_enum_decl(arena),
            TokenKind::Kw(Kw::Event)    => self.parse_event_decl(arena),
            TokenKind::Kw(Kw::Const) => {
                self.advance(); self.parse_const_list(arena)
            }
            // Public/Private/Static variable declaration; a leading `WithEvents`
            // belongs to the first variable and is handled per item.
            TokenKind::Ident | TokenKind::Kw(Kw::Dim) | TokenKind::Kw(Kw::WithEvents) => {
                if self.peek().kind == TokenKind::Kw(Kw::Dim) { self.advance(); }
                self.parse_dim_list_static(arena, false, static_proc)
            }
            // A contextual keyword names the variable (`Public Name As String`).
            TokenKind::Kw(_) if self.at_contextual_keyword() => {
                self.parse_dim_list_static(arena, false, static_proc)
            }
            _ => {
                let span = self.current_span();
                self.diagnostics.push(ERR_UNEXPECTED_TOKEN, span);
                self.skip_to_stmt_end();
                arena.alloc(ExprNode::Block { stmts: vec![] })
            }
        };
        if let Some(is_public) = vis {
            self.mark_decl_visibility(arena, node, is_public);
        }
        if access_tok == TokenKind::Kw(Kw::Friend) {
            self.decl_friend.insert(node.0);
        }
        if static_proc && matches!(arena.get(node), ExprNode::ProcDecl { .. }) {
            self.decl_static.insert(node.0);
        }
        node
    }

    /// Record explicit visibility for a declaration node.
    fn mark_decl_visibility(&mut self, arena: &ExprArena, node: NodeId, is_public: bool) {
        if let ExprNode::Block { stmts } = arena.get(node) {
            let stmts = stmts.clone();
            for s in stmts {
                self.decl_public.insert(s.0, is_public);
            }
        } else {
            self.decl_public.insert(node.0, is_public);
        }
    }

    /// Parses Sub or Function body.
    fn parse_proc_decl(&mut self, arena: &mut ExprArena, kind: ProcKind) -> NodeId {
        self.advance(); // consume Sub / Function
        self.module_decl_seen = true;
        self.error_var_local = false;
        let name = self.expect_decl_name();
        let name_id = name.sym.map(|s| s as u32).unwrap_or(0);
        // Parameter list
        let params = if self.eat(TokenKind::Kw(Kw::LParen)) {
            let p = self.parse_param_list(arena);
            self.expect(TokenKind::Kw(Kw::RParen), ERR_EXPECTED_RPAREN);
            Some(p)
        } else { None };
        self.check_param_list(arena, params, ParamOwner::Procedure);
        // Return type.
        let ret_type = if kind == ProcKind::Function && self.eat(TokenKind::Kw(Kw::As)) {
            let t = self.parse_type_spec(arena, true);
            if name.type_suffix != TypeSuffix::None {
                self.diagnostics.push(ERR_UNEXPECTED_TOKEN, name.span);
            }
            Some(t)
        } else if name.type_suffix != TypeSuffix::None {
            if kind == ProcKind::Function {
                self.type_node_from_suffix(arena, name.type_suffix)
            } else {
                self.diagnostics.push(ERR_UNEXPECTED_TOKEN, name.span);
                None
            }
        } else { None };
        let trailing_static = self.eat(TokenKind::Kw(Kw::Static));
        self.end_stmt();
        let body = self.parse_procedure_body(arena, kind);
        let id = arena.alloc(ExprNode::ProcDecl { kind, name: name_id, params, ret_type, body });
        self.set_span(id, name.span);
        if trailing_static {
            self.decl_static.insert(id.0);
        }
        id
    }

    /// A procedure's statements and its closing `End Sub` / `End Function` /
    /// `End Property`. A missing `End`, an `End` of the wrong kind, and text
    /// after the `End` are reported; only a comment may follow it.
    fn parse_procedure_body(&mut self, arena: &mut ExprArena, kind: ProcKind) -> NodeId {
        let outer_kind = self.proc_kind.replace(kind);
        let outer_reported = std::mem::replace(&mut self.block_error_reported, false);
        self.labels_defined.clear();
        self.label_refs.clear();
        let body = self.parse_block_proc(arena);
        self.proc_kind = outer_kind;
        self.block_error_reported = outer_reported;
        self.report_undefined_labels();

        let (closer, expected_code) = match kind {
            ProcKind::Sub => (Kw::Sub, ERR_EXPECTED_END_SUB),
            ProcKind::Function => (Kw::Function, ERR_EXPECTED_END_FUNCTION),
            ProcKind::PropGet | ProcKind::PropLet | ProcKind::PropSet => {
                (Kw::Property, ERR_EXPECTED_END_PROPERTY)
            }
        };
        let end_span = self.current_span();
        if self.peek().kind == TokenKind::Eof {
            self.diagnostics.push(expected_code, end_span);
            return body;
        }
        self.advance(); // End
        if self.peek().kind != TokenKind::Kw(closer) {
            self.diagnostics.push(expected_code, end_span);
        }
        self.advance(); // Sub / Function / Property
        if self.peek().is_stmt_end() {
            self.eat_eol();
        } else {
            let span = self.current_span();
            self.diagnostics.push(ERR_AFTER_END_OF_PROCEDURE, span);
            self.skip_line();
        }
        body
    }

    /// Every label a jump names must be defined in the same procedure.
    fn report_undefined_labels(&mut self) {
        for (target, span) in std::mem::take(&mut self.label_refs) {
            if !self.labels_defined.contains(&target) {
                self.diagnostics.push(ERR_LABEL_NOT_DEFINED, span);
            }
        }
        self.labels_defined.clear();
    }

    /// Parses Property Get/Let/Set.
    fn parse_property_decl(&mut self, arena: &mut ExprArena) -> NodeId {
        self.advance(); // consume Property
        self.module_decl_seen = true;
        self.error_var_local = false;
        let kind = match self.peek().kind {
            TokenKind::Kw(Kw::Get) => { self.advance(); ProcKind::PropGet }
            TokenKind::Kw(Kw::Let) => { self.advance(); ProcKind::PropLet }
            TokenKind::Kw(Kw::Set) => { self.advance(); ProcKind::PropSet }
            _ => {
                let span = self.current_span();
                self.diagnostics.push(ERR_UNEXPECTED_TOKEN, span);
                ProcKind::PropGet
            }
        };
        let name = self.expect_decl_name();
        let name_id = name.sym.map(|s| s as u32).unwrap_or(0);
        let params = if self.eat(TokenKind::Kw(Kw::LParen)) {
            let p = self.parse_param_list(arena);
            self.expect(TokenKind::Kw(Kw::RParen), ERR_EXPECTED_RPAREN);
            Some(p)
        } else { None };
        let owner = if kind == ProcKind::PropGet { ParamOwner::Procedure } else { ParamOwner::PropertyValue };
        self.check_param_list(arena, params, owner);
        if owner == ParamOwner::PropertyValue {
            let has_params = params.is_some_and(|p| matches!(arena.get(p), ExprNode::ArgList { args } if !args.is_empty()));
            if !has_params {
                self.diagnostics.push(ERR_PROPERTY_ARGUMENT_REQUIRED, name.span);
            }
        }
        let ret_type = if kind == ProcKind::PropGet && self.eat(TokenKind::Kw(Kw::As)) {
            Some(self.parse_type_spec(arena, true))
        } else if name.type_suffix != TypeSuffix::None {
            if kind == ProcKind::PropGet {
                self.type_node_from_suffix(arena, name.type_suffix)
            } else {
                self.diagnostics.push(ERR_UNEXPECTED_TOKEN, name.span);
                None
            }
        } else { None };
        let trailing_static = self.eat(TokenKind::Kw(Kw::Static));
        self.end_stmt();
        let body = self.parse_procedure_body(arena, kind);
        let id = arena.alloc(ExprNode::ProcDecl { kind, name: name_id, params, ret_type, body });
        self.set_span(id, name.span);
        if trailing_static {
            self.decl_static.insert(id.0);
        }
        id
    }

    /// Parses an Event declaration.
    /// Report "Only valid in object module" at the current token when parsing
    /// a standard module. Called on the keyword of every class-only statement,
    /// whichever path (bare, access-modified, inside `#If`) reached it.
    fn reject_in_standard_module(&mut self) {
        let span = self.current_span();
        self.reject_in_standard_module_at(span);
    }

    fn reject_in_standard_module_at(&mut self, span: Span) {
        if self.module_kind == ModuleKind::Standard {
            self.diagnostics.push(ERR_OBJECT_MODULE_ONLY, span);
        }
    }

    fn parse_event_decl(&mut self, arena: &mut ExprArena) -> NodeId {
        self.reject_in_standard_module();
        self.advance(); // consume Event
        let name = self.expect_decl_name();
        self.reject_name_suffix(&name);
        let nid = name.sym.map(|s| s as u32).unwrap_or(0);
        let params = if self.eat(TokenKind::Kw(Kw::LParen)) {
            let p = self.parse_param_list(arena);
            self.expect(TokenKind::Kw(Kw::RParen), ERR_EXPECTED_RPAREN);
            p
        } else {
            arena.alloc(ExprNode::ArgList { args: vec![] })
        };
        if self.scanner.sym_name(nid).contains('_') {
            self.diagnostics.push(ERR_INVALID_EVENT_NAME, name.span);
        }
        self.check_param_list(arena, Some(params), ParamOwner::Event);
        self.end_stmt();
        let id = arena.alloc(ExprNode::EventDecl { name: nid, params });
        self.set_span(id, name.span);
        id
    }

    /// Parse `Declare [Function|Sub] name Lib "lib" [Alias "alias"] ([params]) [As type]`
    fn parse_declare_stmt(&mut self, arena: &mut ExprArena) -> NodeId {
        self.advance(); // Declare
        self.module_decl_seen = true;
        let kind = match self.peek().kind.clone() {
            TokenKind::Kw(Kw::Function) => { self.advance(); ProcKind::Function }
            TokenKind::Kw(Kw::Sub)      => { self.advance(); ProcKind::Sub }
            _ => {
                let span = self.current_span();
                self.diagnostics.push(ERR_UNEXPECTED_TOKEN, span);
                self.skip_to_stmt_end();
                self.end_stmt();
                return arena.alloc(ExprNode::Block { stmts: vec![] });
            }
        };
        let name_tok = self.expect_decl_name();
        let name = name_tok.sym.map(|s| s as u32).unwrap_or(0);
        // A `Sub` returns nothing, so its name takes no type character.
        if kind == ProcKind::Sub && name_tok.type_suffix != TypeSuffix::None {
            self.report_expected(name_tok.span, "Identifier without type suffix");
        }
        // The calling-convention word may precede `Lib`.
        self.eat(TokenKind::Kw(Kw::CDecl));
        // Lib "libname": the library is a string literal.
        let lib = if self.eat(TokenKind::Kw(Kw::Lib)) {
            match self.declare_string_operand(arena) {
                Some(s) => s,
                None => {
                    self.skip_line();
                    return arena.alloc(ExprNode::Block { stmts: vec![] });
                }
            }
        } else {
            let span = self.current_span();
            self.diagnostics.push(ERR_UNEXPECTED_TOKEN, span);
            arena.alloc(ExprNode::Literal { lit: crate::frontend::ast::AstLit::Str("".into()) })
        };
        // Optional Alias "aliasname"
        let alias = if self.eat(TokenKind::Kw(Kw::Alias)) {
            match self.declare_string_operand(arena) {
                Some(s) => Some(s),
                None => {
                    self.skip_line();
                    return arena.alloc(ExprNode::Block { stmts: vec![] });
                }
            }
        } else {
            None
        };
        // Optional parameter list
        let params = if self.eat(TokenKind::Kw(Kw::LParen)) {
            let p = self.parse_declare_param_list(arena);
            self.expect(TokenKind::Kw(Kw::RParen), ERR_EXPECTED_RPAREN);
            Some(p)
        } else {
            None
        };
        // Optional return type.
        if kind == ProcKind::Function {
            self.reject_suffix_with_as(&name_tok, "end of statement");
        }
        let ret_type = if kind == ProcKind::Function && self.eat(TokenKind::Kw(Kw::As)) {
            Some(self.parse_type_spec(arena, false))
        } else {
            None
        };
        self.end_stmt();
        let id = arena.alloc(ExprNode::DeclareDecl { kind, name, lib, alias, params, ret_type });
        self.set_span(id, name_tok.span);
        id
    }

    /// The string literal operand of `Lib` / `Alias`; anything else is reported.
    fn declare_string_operand(&mut self, arena: &mut ExprArena) -> Option<NodeId> {
        if self.peek().kind == TokenKind::StrLit {
            let tok = self.advance();
            let n = arena.alloc(ExprNode::Literal { lit: ast_lit_from_token(&tok) });
            self.set_span(n, tok.span);
            Some(n)
        } else {
            let span = self.current_span();
            self.diagnostics.push_labeled(ERR_UNEXPECTED_TOKEN, span, "string constant");
            None
        }
    }

    /// The order and shape rules for a parameter list: `Optional` parameters
    /// follow the required ones, `ParamArray` is last and an array, events
    /// take neither, and an array cannot be `ByVal` or `Optional`. Only the
    /// first violation of a list is reported, as the compiler stops at the
    /// first error of a line.
    fn check_param_list(&mut self, arena: &ExprArena, params: Option<NodeId>, owner: ParamOwner) {
        let Some(list) = params else { return };
        let ExprNode::ArgList { args } = arena.get(list) else { return };
        // 0 = none yet, 1 = an Optional seen, 2 = a ParamArray seen.
        let mut section = 0;
        for (i, &param) in args.iter().enumerate() {
            let ExprNode::ParamDef { flags, .. } = arena.get(param) else { continue };
            let flags = *flags;
            let span = self.node_spans.get(param);
            let is_optional = flags & 0x01 != 0;
            let is_param_array = flags & 0x20 != 0;
            let is_array = flags & 0x08 != 0;
            let value_param = owner == ParamOwner::PropertyValue && i + 1 == args.len();
            let violation: Option<(u32, Option<&'static str>)> = if owner == ParamOwner::Event
                && (is_optional || is_param_array)
            {
                Some((ERR_SYNTAX_ERROR, None))
            } else if is_param_array {
                if section == 2 {
                    Some((ERR_UNEXPECTED_TOKEN, Some("identifier")))
                } else if section == 1 {
                    Some((ERR_PARAM_ARRAY_WITH_OPTIONAL, None))
                } else if !is_array || !param_type_is_variant(arena, param) {
                    Some((ERR_PARAM_ARRAY_NOT_ARRAY, None))
                } else if i + 1 < args.len() && owner != ParamOwner::PropertyValue {
                    Some((ERR_UNEXPECTED_TOKEN, Some(")")))
                } else {
                    None
                }
            } else if is_optional {
                if section == 2 {
                    Some((ERR_UNEXPECTED_TOKEN, Some(")")))
                } else if is_array {
                    Some((ERR_OPTIONAL_ARRAY, None))
                } else {
                    None
                }
            } else if section != 0 && !value_param {
                Some((ERR_UNEXPECTED_TOKEN, Some(if section == 1 { "Optional" } else { ")" })))
            } else if is_array && flags & 0x02 != 0 {
                Some((ERR_ARRAY_ARGUMENT_MUST_BE_BYREF, None))
            } else {
                None
            };
            if let Some((code, label)) = violation {
                match label {
                    Some(label) => self.diagnostics.push_labeled(code, span, label),
                    None => self.diagnostics.push(code, span),
                }
                return;
            }
            if is_array && flags & 0x02 != 0 {
                self.diagnostics.push(ERR_ARRAY_ARGUMENT_MUST_BE_BYREF, span);
                return;
            }
            if is_param_array {
                section = 2;
            } else if is_optional {
                section = section.max(1);
            }
        }
    }

    /// Parses a procedure parameter list.
    fn parse_param_list(&mut self, arena: &mut ExprArena) -> NodeId {
        self.parse_param_list_of(arena, false)
    }

    /// Parses a `Declare` parameter list, where `As Any` is also legal.
    fn parse_declare_param_list(&mut self, arena: &mut ExprArena) -> NodeId {
        self.parse_param_list_of(arena, true)
    }

    fn parse_param_list_of(&mut self, arena: &mut ExprArena, in_declare: bool) -> NodeId {
        if self.peek().kind == TokenKind::Kw(Kw::RParen) {
            return arena.alloc(ExprNode::ArgList { args: vec![] });
        }
        let mut args = Vec::new();
        loop {
            args.push(self.parse_param(arena, in_declare));
            if !self.eat(TokenKind::Kw(Kw::Comma)) { break; }
            if self.peek().kind == TokenKind::Kw(Kw::RParen) {
                // A comma needs a parameter after it.
                let span = self.current_span();
                self.report_expected(span, "identifier");
                break;
            }
        }
        arena.alloc(ExprNode::ArgList { args })
    }

    fn parse_param(&mut self, arena: &mut ExprArena, in_declare: bool) -> NodeId {
        let mut flags: u16 = 0;
        // `ByVal` or `ByRef` may come before `Optional` as well as after it.
        let mut by_first = true;
        match self.peek().kind {
            TokenKind::Kw(Kw::ByVal) => { self.advance(); flags |= 0x02; }
            TokenKind::Kw(Kw::ByRef) => { self.advance(); flags |= 0x04; }
            _ => by_first = false,
        }
        if self.eat(TokenKind::Kw(Kw::Optional)) { flags |= 0x01; }
        // ParamArray must precede the parameter name; mutually exclusive with ByVal/ByRef.
        if by_first {
            // The pass-mode word was given already.
        } else if self.eat(TokenKind::Kw(Kw::ParamArray)) {
            flags |= 0x20;
        } else {
            match self.peek().kind {
                TokenKind::Kw(Kw::ByVal) => { self.advance(); flags |= 0x02; }
                TokenKind::Kw(Kw::ByRef) => { self.advance(); flags |= 0x04; }
                _ => {}
            }
        }
        let name = self.expect_decl_name();
        self.note_declared_name(&name, false);
        let nid = name.sym.map(|s| s as u32).unwrap_or(0);
        if self.eat(TokenKind::Kw(Kw::LParen)) {
            self.expect(TokenKind::Kw(Kw::RParen), ERR_EXPECTED_RPAREN);
            flags |= 0x08; // array parameter
        }
        self.reject_suffix_with_as(&name, "list separator or )");
        let type_node = if self.eat(TokenKind::Kw(Kw::As)) {
            if self.eat(TokenKind::Kw(Kw::New)) {
                flags |= 0x10;
                self.require_class_name_after_new();
            }
            if in_declare && self.eat(TokenKind::Kw(Kw::Any)) {
                Some(arena.alloc(ExprNode::AnyType))
            } else {
                Some(self.parse_type_spec(arena, false))
            }
        } else {
            // `Sub F(count%)` ≡ `count As Integer`
            self.type_node_from_suffix(arena, name.type_suffix)
        };
        // Default value for `Optional` — a real (constant) expression use site.
        let default = if flags & 0x01 != 0 && self.eat(TokenKind::Kw(Kw::Eq)) {
            Some(self.parse_expr(arena, 0))
        } else {
            None
        };
        let id = arena.alloc(ExprNode::ParamDef { flags, name: nid, type_node, default });
        self.set_span(id, name.span);
        id
    }

    /// Build a `UserType` from its already-consumed first name token, taking an
    /// optional `.Name` qualifier tail (`Project.Class`, `Module.Type`). The node
    /// spans the first name; the tail is a `NameRef` with its own span.
    fn finish_user_type(&mut self, arena: &mut ExprArena, name: Token) -> NodeId {
        let nid = name.sym.map(|s| s as u32).unwrap_or(0);
        let child = if self.eat(TokenKind::Kw(Kw::Dot)) {
            let sub = self.expect_ident();
            let sid = sub.sym.map(|s| s as u32).unwrap_or(0);
            let cref = arena.alloc(ExprNode::NameRef { sym: sid, suffix: TypeSuffix::None });
            self.set_span(cref, sub.span);
            Some(cref)
        } else {
            None
        };
        let id = arena.alloc(ExprNode::UserType { name: nid, child });
        self.set_span(id, name.span);
        id
    }

    /// Parse a type specifier.  `allow_array_suffix` controls whether a trailing
    /// `()` array marker is consumed.
    /// The length after `String *`: one whole-number literal or one constant
    /// name. Nothing larger is an expression here, so `String * 3 = "x"` in a
    /// `Const` leaves the `=` to the constant's value.
    fn parse_fixed_string_length(&mut self, arena: &mut ExprArena) -> NodeId {
        let len_span = self.current_span();
        let contextual = self.at_contextual_keyword();
        match self.peek().kind {
            TokenKind::IntLit | TokenKind::LongLit => {
                let tok = self.advance();
                let lit = ast_lit_from_token(&tok);
                // A literal length must be 1 up to (not including) 65527.
                if let AstLit::Int(n) | AstLit::Long(n) = &lit {
                    if *n <= 0 || *n >= FIXED_STRING_LIMIT {
                        self.diagnostics.push(ERR_INVALID_FIXED_STRING_LENGTH, len_span);
                    }
                }
                arena.alloc(ExprNode::Literal { lit })
            }
            TokenKind::Ident => {
                let tok = self.advance();
                let sym = tok.sym.map(|s| s as u32).unwrap_or(0);
                let name = arena.alloc(ExprNode::NameRef { sym, suffix: tok.type_suffix });
                self.set_span(name, tok.span);
                name
            }
            TokenKind::Kw(_) if contextual => {
                let tok = self.advance();
                let sym = tok.sym.map(|s| s as u32).unwrap_or(0);
                let name = arena.alloc(ExprNode::NameRef { sym, suffix: tok.type_suffix });
                self.set_span(name, tok.span);
                name
            }
            _ => {
                self.diagnostics.push_labeled(ERR_EXPECTED_IDENT, len_span, "Identifier");
                arena.alloc(ExprNode::Literal { lit: AstLit::Int(0) })
            }
        }
    }

    fn parse_type_spec(&mut self, arena: &mut ExprArena, allow_array_suffix: bool) -> NodeId {
        let node = match self.peek().kind.clone() {
            TokenKind::Kw(Kw::Integer)  => { self.advance(); arena.alloc(ExprNode::BuiltinType { kind: 2  }) }
            TokenKind::Kw(Kw::Long)     => { self.advance(); arena.alloc(ExprNode::BuiltinType { kind: 3  }) }
            TokenKind::Kw(Kw::Single)   => { self.advance(); arena.alloc(ExprNode::BuiltinType { kind: 4  }) }
            TokenKind::Kw(Kw::Double)   => { self.advance(); arena.alloc(ExprNode::BuiltinType { kind: 5  }) }
            TokenKind::Kw(Kw::Currency) => { self.advance(); arena.alloc(ExprNode::BuiltinType { kind: 6  }) }
            TokenKind::Kw(Kw::Date)     => { self.advance(); arena.alloc(ExprNode::BuiltinType { kind: 7  }) }
            TokenKind::Kw(Kw::Byte)     => { self.advance(); arena.alloc(ExprNode::BuiltinType { kind: 17 }) }
            TokenKind::Kw(Kw::Boolean)  => { self.advance(); arena.alloc(ExprNode::BuiltinType { kind: 11 }) }
            TokenKind::Kw(Kw::Variant)  => { self.advance(); arena.alloc(ExprNode::BuiltinType { kind: 12 }) }
            TokenKind::Kw(Kw::Object)   => { self.advance(); arena.alloc(ExprNode::BuiltinType { kind: 9  }) }
            TokenKind::Kw(Kw::Decimal)  => {
                // `Decimal` is the subtype of a Variant, never the type of a declaration.
                let span = self.current_span();
                self.report_expected(span, "type name");
                self.advance();
                arena.alloc(ExprNode::BuiltinType { kind: 0 })
            }
            TokenKind::Kw(Kw::String) => {
                self.advance();
                let fixed_len = if self.eat(TokenKind::Kw(Kw::Star)) {
                    Some(self.parse_fixed_string_length(arena))
                } else { None };
                arena.alloc(ExprNode::StringType { fixed_len })
            }
            TokenKind::Kw(Kw::New) => {
                self.advance();
                self.require_class_name_after_new();
                let name = self.expect_ident();
                self.finish_user_type(arena, name)
            }
            // A user type name; a contextual keyword (`Error`, `Line`) can be one.
            TokenKind::Ident | TokenKind::Kw(_) if self.at_contextual_keyword() || self.peek().kind == TokenKind::Ident => {
                let name = self.advance();
                self.finish_user_type(arena, name)
            }
            _ => {
                let span = self.current_span();
                self.diagnostics.push(ERR_UNEXPECTED_TOKEN, span);
                arena.alloc(ExprNode::BuiltinType { kind: 0 })
            }
        };
        // The optional array marker `()` makes the type an array of it.
        if allow_array_suffix && self.peek().kind == TokenKind::Kw(Kw::LParen) {
            self.advance();
            self.expect(TokenKind::Kw(Kw::RParen), ERR_EXPECTED_RPAREN);
            return arena.alloc(ExprNode::ArrayType { elem: node });
        }
        node
    }

    /// Parse a Type declaration block.
    fn parse_type_decl(&mut self, arena: &mut ExprArena) -> NodeId {
        self.advance(); // consume Type
        let name = self.expect_decl_name();
        self.reject_name_suffix(&name);
        let nid = name.sym.map(|s| s as u32).unwrap_or(0);
        self.end_stmt();
        let mut members = Vec::new();
        loop {
            while self.eat(TokenKind::Eol) {}
            match self.peek().kind.clone() {
                // `End` followed by `As` or `(` is a member named `End`.
                TokenKind::Kw(Kw::End)
                    if !matches!(
                        self.peek2().kind,
                        TokenKind::Kw(Kw::As) | TokenKind::Kw(Kw::LParen)
                    ) =>
                {
                    let end_span = self.current_span();
                    self.advance();
                    self.expect(TokenKind::Kw(Kw::Type), ERR_UNEXPECTED_TOKEN);
                    if members.is_empty() {
                        self.diagnostics.push(ERR_TYPE_WITHOUT_MEMBERS, end_span);
                    }
                    self.end_stmt();
                    break;
                }
                TokenKind::Eof => {
                    // The file ended inside the block.
                    let span = self.current_span();
                    self.diagnostics.push(ERR_STATEMENT_IN_TYPE, span);
                    break;
                }
                // A comment line between members is a blank line, and so is a lone colon.
                TokenKind::Kw(Kw::Apos) | TokenKind::Kw(Kw::Rem) | TokenKind::Kw(Kw::Colon) => { self.advance(); }
                // A statement keyword cannot start a member. A keyword followed by
                // `As` is a member that happens to be named like one, and so is one of
                // the few reserved words that also name callable things when bounds follow.
                TokenKind::Kw(_)
                    if !self.at_contextual_keyword()
                        && !matches!(self.peek2().kind, TokenKind::Kw(Kw::As))
                        && !(self.peek2().kind == TokenKind::Kw(Kw::LParen) && self.at_callable_reserved_word()) =>
                {
                    let span = self.current_span();
                    self.diagnostics.push(ERR_STATEMENT_IN_TYPE, span);
                    self.skip_line();
                }
                _ => {
                    let line_span = self.current_span();
                    // A keyword that starts a statement of its own cannot take bounds as a member.
                    if self.peek2().kind == TokenKind::Kw(Kw::LParen) && self.at_statement_word() {
                        self.report_expected(line_span, "end of statement");
                        self.skip_line();
                        continue;
                    }
                    self.reserved_names_allowed = true;
                    self.last_item_had_as = false;
                    let m = self.parse_dim_item(arena, false, false, false);
                    self.reserved_names_allowed = false;
                    if !self.last_item_had_as && self.peek().kind != TokenKind::Kw(Kw::Eq) {
                        self.diagnostics.push(ERR_STATEMENT_IN_TYPE, line_span);
                        self.skip_line();
                        continue;
                    }
                    // `name = value` is a statement, not a member.
                    if self.peek().kind == TokenKind::Kw(Kw::Eq) {
                        self.diagnostics.push(ERR_STATEMENT_IN_TYPE, line_span);
                        self.skip_line();
                    } else {
                        self.end_stmt();
                    }
                    members.push(m);
                }
            }
        }
        let id = arena.alloc(ExprNode::TypeDecl { name: nid, members });
        self.set_span(id, name.span);
        id
    }

    /// Parse an Enum declaration block.
    fn parse_enum_decl(&mut self, arena: &mut ExprArena) -> NodeId {
        self.advance(); // consume Enum
        let name = self.expect_decl_name();
        self.reject_name_suffix(&name);
        let nid = name.sym.map(|s| s as u32).unwrap_or(0);
        self.end_stmt();
        let mut members = Vec::new();
        loop {
            while self.eat(TokenKind::Eol) {}
            match self.peek().kind.clone() {
                TokenKind::Kw(Kw::End) => {
                    let end_span = self.current_span();
                    self.advance();
                    self.expect(TokenKind::Kw(Kw::Enum), ERR_UNEXPECTED_TOKEN);
                    if members.is_empty() {
                        self.diagnostics.push(ERR_ENUM_WITHOUT_MEMBERS, end_span);
                    }
                    self.end_stmt();
                    break;
                }
                TokenKind::Eof => {
                    // The file ended inside the block.
                    let span = self.current_span();
                    self.diagnostics.push(ERR_INVALID_IN_ENUM, span);
                    break;
                }
                // A comment line between members is a blank line, and `:`
                // separates members written on one line.
                TokenKind::Kw(Kw::Apos) | TokenKind::Kw(Kw::Rem) | TokenKind::Kw(Kw::Colon) => {
                    self.advance();
                }
                TokenKind::Ident => {
                    let m = self.parse_dim_item(arena, true, false, false);
                    self.end_stmt();
                    members.push(m);
                }
                // A non-reserved keyword (`Text`, `Key`) can name a member, unless it
                // starts a statement that needs operands (`Name`, `Width`, ...).
                // `Date`, `DoEvents` and `CDec` are reserved but name a member too.
                TokenKind::Kw(_) if self.names_enum_member() => {
                    self.reserved_names_allowed = true;
                    let m = self.parse_dim_item(arena, true, false, false);
                    self.reserved_names_allowed = false;
                    self.end_stmt();
                    members.push(m);
                }
                _ => {
                    // Not a member: report it and resume on the next line.
                    let span = self.current_span();
                    self.diagnostics.push(ERR_INVALID_IN_ENUM, span);
                    self.skip_line();
                }
            }
        }
        let id = arena.alloc(ExprNode::EnumDecl { name: nid, members });
        self.set_span(id, name.span);
        id
    }

    // ── Expression parser (Pratt) ─────────────────────────────────────────────

    /// Pratt precedence-climbing expression parser.
    /// `min_bp` = minimum binding power required to consume an infix operator.
    pub fn parse_expr(&mut self, arena: &mut ExprArena, min_bp: u8) -> NodeId {
        let lhs = self.parse_expr_prefix(arena);
        self.parse_infix(arena, lhs, min_bp)
    }

    /// The infix part of an expression whose left operand `lhs` is already
    /// parsed: consume operators binding tighter than `min_bp`.
    fn parse_infix(&mut self, arena: &mut ExprArena, mut lhs: NodeId, min_bp: u8) -> NodeId {
        loop {

            let kind = self.peek().kind.clone();
            let Some((l_bp, r_bp, bin_op)) = infix_bp(&kind) else { break };
            if l_bp <= min_bp { break; }
            self.advance(); // consume operator
            let rhs = self.parse_expr(arena, r_bp);
            lhs = arena.alloc(ExprNode::BinOp { op: bin_op, lhs, rhs });
        }
        lhs
    }

    /// Parse the prefix/primary portion of an expression: a unary operator,
    /// `TypeOf`, `New`, or a plain primary.
    fn parse_expr_prefix(&mut self, arena: &mut ExprArena) -> NodeId {
        if self.peek().kind == TokenKind::Kw(Kw::Plus) {
            // A unary plus builds no node. Its operand takes every operator that binds
            // tighter than `+` and `-` (`Mod` and above), so `+a * b` is `a * b`.
            self.advance();
            return self.parse_expr(arena, PLUS_OPERAND_MIN_BP);
        }
        if let Some((bp, uop)) = prefix_bp(&self.peek().kind.clone()) {
            self.advance();
            let operand = self.parse_expr(arena, bp);
            arena.alloc(ExprNode::UnOp { op: uop, operand })
        } else if self.peek().kind == TokenKind::Kw(Kw::TypeOf) {
            self.parse_typeof_expr(arena)
        } else if self.peek().kind == TokenKind::Kw(Kw::New) {
            self.parse_new_expr(arena)
        } else if matches!(self.peek().kind, TokenKind::Kw(k) if !keyword_starts_expression(k)) {
            // A reserved word that is not a value or a built-in function
            // cannot start an expression. It is consumed so that callers
            // looping over items always advance.
            let span = self.current_span();
            self.diagnostics.push(ERR_EXPECTED_EXPR, span);
            self.advance();
            arena.alloc(ExprNode::Literal { lit: AstLit::Int(0) })
        } else {
            self.parse_primary(arena)
        }
    }

    fn try_parse_expr(&mut self, arena: &mut ExprArena) -> Option<NodeId> {
        if self.peek().is_stmt_end() { return None; }
        match self.peek().kind {
            TokenKind::Kw(Kw::RParen)
            | TokenKind::Kw(Kw::Comma)
            | TokenKind::Kw(Kw::Colon)
            // The `Else` of a single-line `If` ends a call without arguments.
            | TokenKind::Kw(Kw::Else) => None,
            _ => Some(self.parse_expr(arena, 0)),
        }
    }

    /// `TypeOf <expr> Is <type>`
    fn parse_typeof_expr(&mut self, arena: &mut ExprArena) -> NodeId {
        self.advance(); // consume TypeOf
        let expr = self.parse_primary(arena);
        self.expect(TokenKind::Kw(Kw::Is), ERR_EXPECTED_IS);
        let type_spec = self.parse_type_spec(arena, false);
        arena.alloc(ExprNode::TypeOf { expr, type_spec })
    }

    /// `New <TypeName>`
    fn parse_new_expr(&mut self, arena: &mut ExprArena) -> NodeId {
        self.advance(); // consume New
        let type_spec = self.parse_type_spec(arena, false);
        arena.alloc(ExprNode::New { type_spec })
    }

    /// Parses primary expressions.
    fn parse_primary(&mut self, arena: &mut ExprArena) -> NodeId {
        let tok = self.peek().kind.clone();
        match tok {
            // Grouped expression
            TokenKind::Kw(Kw::LParen) => {
                self.advance();
                let inner = self.parse_expr(arena, 0);
                self.expect(TokenKind::Kw(Kw::RParen), ERR_EXPECTED_RPAREN);
                // Nothing joins to a group: `(a).b` and `(a)(1)` are not expressions.
                arena.alloc(ExprNode::Paren { inner })
            }
            // Numeric / string / date literals
            TokenKind::IntLit
            | TokenKind::LongLit
            | TokenKind::SngLit
            | TokenKind::DblLit
            | TokenKind::CurLit
            | TokenKind::StrLit
            | TokenKind::DateLit => {
                let tok = self.advance();
                arena.alloc(ExprNode::Literal { lit: ast_lit_from_token(&tok) })
            }
            // Boolean literals
            TokenKind::Kw(Kw::True) => {
                self.advance();
                arena.alloc(ExprNode::Literal { lit: AstLit::Bool(true) })
            }
            TokenKind::Kw(Kw::False) => {
                self.advance();
                arena.alloc(ExprNode::Literal { lit: AstLit::Bool(false) })
            }
            TokenKind::Kw(Kw::Nothing) => {
                self.advance();
                arena.alloc(ExprNode::Nothing)
            }
            // Me
            // `Me` takes member access and calls like any other object.
            TokenKind::Kw(Kw::Me) => {
                self.advance();
                let base = arena.alloc(ExprNode::Me);
                self.parse_postfix(arena, base)
            }
            // Empty / Null
            TokenKind::Kw(Kw::Empty) => {
                self.advance();
                arena.alloc(ExprNode::Literal { lit: AstLit::Empty })
            }
            TokenKind::Kw(Kw::Null) => {
                self.advance();
                arena.alloc(ExprNode::Literal { lit: AstLit::Null })
            }
            // Identifier or qualified name
            TokenKind::Ident => {
                let tok = self.advance();
                let sym_id = tok.sym.map(|s| s as u32).unwrap_or(0);
                let base = arena.alloc(ExprNode::NameRef { sym: sym_id, suffix: tok.type_suffix });
                self.set_span(base, tok.span);
                self.parse_postfix(arena, base)
            }
            // `AddressOf proc` — function-pointer operand.
            TokenKind::Kw(Kw::AddressOf) => {
                let at = self.advance();
                if !std::mem::take(&mut self.address_of_ok) {
                    self.report_expected(at.span, "expression");
                }
                let operand = self.parse_procedure_name(arena);
                arena.alloc(ExprNode::AddressOf { operand })
            }
            // Leading `.`/`!` — With-block member reference in expression
            // position (e.g. `y = .Field`), same as in a target position.
            TokenKind::Kw(Kw::Dot) | TokenKind::Kw(Kw::Bang) => {
                let access = self
                    .try_parse_leading_with_member(arena)
                    .expect("peek was Dot/Bang");
                self.parse_postfix(arena, access)
            }
            // A literal the scanner rejected (out of range, no digits).
            TokenKind::Error => {
                let span = self.current_span();
                self.advance();
                self.diagnostics.push(ERR_SYNTAX_ERROR, span);
                arena.alloc(ExprNode::Literal { lit: AstLit::Int(0) })
            }
            // Keyword that can start a primary (e.g. built-in functions)
            TokenKind::Kw(k) => {
                let tok = self.advance();
                let sym_id = tok.sym.map(|s| s as u32).unwrap_or(0);
                let base = arena.alloc(ExprNode::NameRef { sym: sym_id, suffix: tok.type_suffix });
                self.set_span(base, tok.span);
                let name = self.scanner.sym_name(sym_id).to_string();
                if let Some(shape) = INTRINSIC_SHAPES
                    .iter()
                    .find(|s| !self.in_target && name.eq_ignore_ascii_case(s.name))
                {
                    return self.parse_shaped_intrinsic_call(arena, base, tok.span, shape);
                }
                if name.eq_ignore_ascii_case("Array") && self.peek().kind != TokenKind::Kw(Kw::LParen) {
                    // `Array` is a function: brackets are not optional.
                    self.report_expected(tok.span, "(");
                    return base;
                }
                let is_input_fn = matches!(k, Kw::Input | Kw::InputS | Kw::InputB | Kw::InputBS);
                if is_input_fn && self.peek().kind == TokenKind::Kw(Kw::LParen) {
                    // Nothing joins to the result of `Input(n, file)`.
                    self.parse_input_fn_call(arena, base)
                } else if is_input_fn {
                    // `Input` is a function: its brackets are not optional.
                    self.report_expected(tok.span, "(");
                    base
                } else {
                    self.parse_postfix(arena, base)
                }
            }
            // Unexpected — emit error, return placeholder
            _ => {
                let span = self.current_span();
                self.diagnostics.push(ERR_EXPECTED_EXPR, span);
                arena.alloc(ExprNode::Literal { lit: AstLit::Int(0) })
            }
        }
    }

    /// `Input(count, [#]filenumber)` and its `InputB`/`$` forms: the file
    /// number may carry a `#`, which does not change its meaning, so the
    /// result is the same `Call` as without it.
    fn parse_input_fn_call(&mut self, arena: &mut ExprArena, func: NodeId) -> NodeId {
        self.advance(); // consume (
        let count = self.parse_expr(arena, 0);
        self.expect(TokenKind::Kw(Kw::Comma), ERR_UNEXPECTED_TOKEN);
        self.eat(TokenKind::Kw(Kw::Hash));
        let file = self.parse_expr(arena, 0);
        self.expect(TokenKind::Kw(Kw::RParen), ERR_EXPECTED_RPAREN);
        let args = arena.alloc(ExprNode::ArgList { args: vec![count, file] });
        arena.alloc(ExprNode::Call { func, args })
    }

    /// Postfix: member access and call/index.
    fn parse_postfix(&mut self, arena: &mut ExprArena, mut base: NodeId) -> NodeId {
        loop {
            // A `.` or `!` after blanks on the same line starts a With-member
            // argument of the statement, not a member of what precedes it.
            if matches!(self.peek().kind, TokenKind::Kw(Kw::Dot) | TokenKind::Kw(Kw::Bang))
                && self.blank_before_next()
            {
                break;
            }
            match self.peek().kind.clone() {
                TokenKind::Kw(Kw::Dot) => {
                    self.advance();
                    let member = self.expect_ident();
                    let mid = member.sym.map(|s| s as u32).unwrap_or(0);
                    base = arena.alloc(ExprNode::MemberAccess { base, member: mid, bang: false });
                    self.set_span(base, member.span);
                }
                TokenKind::Kw(Kw::Bang) => {
                    self.advance();
                    // Unlike the dot, the bang takes its key with no blanks between.
                    if self.blank_before_next() {
                        let span = self.current_span();
                        self.diagnostics.push(ERR_SYNTAX_ERROR, span);
                    }
                    let member = self.expect_ident();
                    let mid = member.sym.map(|s| s as u32).unwrap_or(0);
                    base = arena.alloc(ExprNode::MemberAccess { base, member: mid, bang: true });
                    self.set_span(base, member.span);
                }
                TokenKind::Kw(Kw::LParen) => {
                    self.advance();
                    let args = self.parse_arg_list(arena);
                    self.expect(TokenKind::Kw(Kw::RParen), ERR_EXPECTED_RPAREN);
                    base = arena.alloc(ExprNode::Call { func: base, args });
                }
                _ => break,
            }
        }
        base
    }

    /// The qualified name of a procedure after `AddressOf`: names joined by dots.
    fn parse_procedure_name(&mut self, arena: &mut ExprArena) -> NodeId {
        let first = self.expect_decl_name();
        let sym = first.sym.map(|s| s as u32).unwrap_or(0);
        let mut base = arena.alloc(ExprNode::NameRef { sym, suffix: first.type_suffix });
        self.set_span(base, first.span);
        while self.peek().kind == TokenKind::Kw(Kw::Dot) && !self.blank_before_next() {
            self.advance();
            let member = self.expect_ident();
            let mid = member.sym.map(|s| s as u32).unwrap_or(0);
            base = arena.alloc(ExprNode::MemberAccess { base, member: mid, bang: false });
            self.set_span(base, member.span);
        }
        base
    }

    /// Notes that an `AddressOf` may start the argument that follows.
    fn allow_address_of_here(&mut self) {
        self.address_of_ok = self.peek().kind == TokenKind::Kw(Kw::AddressOf);
    }

    /// Parses the LHS of an assignment.
    fn parse_target_expr(&mut self, arena: &mut ExprArena) -> NodeId {
        // A leading `.`/`!` is a With-block member reference: `.Member` / `!Key`.
        // The leading dot/bang is semantically significant (it binds the member
        // to the active With object), so build a `MemberAccess` on an implicit
        // `WithContext` base — never a bare `NameRef`, which would drop it.
        if let Some(access) = self.try_parse_leading_with_member(arena) {
            return self.parse_postfix(arena, access);
        }
        // A target names something: a reserved word does not (`Me` and the few reserved
        // words that also name callable things apart).
        let at_start = std::mem::take(&mut self.target_at_statement_start);
        if self.peek().kind == TokenKind::Kw(Kw::Rem) {
            // A comment cannot stand where a target belongs.
            let span = self.current_span();
            self.diagnostics.push(ERR_SYNTAX_ERROR, span);
            self.skip_line();
            return self.parse_primary_placeholder(arena);
        }
        if self.at_reserved_word()
            && !(at_start && self.starts_name_statement())
            && self.peek().kind != TokenKind::Kw(Kw::Me)
            && !self.at_array_call()
        {
            return self.reject_reserved_word_target(arena);
        }
        // A target names something: a bracket or an operator does not.
        if self.at_stray_punctuation() {
            let span = self.current_span();
            self.diagnostics.push(ERR_SYNTAX_ERROR, span);
            self.skip_line();
            return self.parse_primary_placeholder(arena);
        }
        let was_target = std::mem::replace(&mut self.in_target, true);
        let target = self.parse_target_primary(arena);
        self.in_target = was_target;
        target
    }

    /// The part of a target that is not a leading member reference.
    fn parse_target_primary(&mut self, arena: &mut ExprArena) -> NodeId {
        // A target names something: a literal does not.
        if matches!(
            self.peek().kind,
            TokenKind::IntLit
                | TokenKind::LongLit
                | TokenKind::SngLit
                | TokenKind::DblLit
                | TokenKind::CurLit
                | TokenKind::StrLit
                | TokenKind::DateLit
                | TokenKind::Kw(Kw::True)
                | TokenKind::Kw(Kw::False)
                | TokenKind::Kw(Kw::Nothing)
                | TokenKind::Kw(Kw::Empty)
                | TokenKind::Kw(Kw::Null)
        ) {
            let span = self.current_span();
            self.diagnostics.push_labeled(ERR_EXPECTED_IDENT, span, "Identifier");
            self.advance();
            return self.parse_primary_placeholder(arena);
        }
        self.parse_primary(arena)
    }

    fn parse_primary_placeholder(&mut self, arena: &mut ExprArena) -> NodeId {
        arena.alloc(ExprNode::Literal { lit: AstLit::Int(0) })
    }

    /// If the next token is a leading `.` or `!` (a With-block member reference
    /// with no explicit base), consume `.Member` / `!Key` and return a
    /// `MemberAccess` whose base is the implicit [`ExprNode::WithContext`].
    /// Returns `None` (consuming nothing) otherwise.
    fn try_parse_leading_with_member(&mut self, arena: &mut ExprArena) -> Option<NodeId> {
        let bang = match self.peek().kind {
            TokenKind::Kw(Kw::Dot) => false,
            TokenKind::Kw(Kw::Bang) => true,
            _ => return None,
        };
        let dot_span = self.current_span();
        self.advance(); // consume `.` / `!`
        let member = self.expect_ident();
        let mid = member.sym.map(|s| s as u32).unwrap_or(0);
        let base = arena.alloc(ExprNode::WithContext);
        self.set_span(base, dot_span);
        let access = arena.alloc(ExprNode::MemberAccess { base, member: mid, bang });
        self.set_span(access, member.span);
        Some(access)
    }

    // ── Argument lists ────────────────────────────────────────────────────────

    /// Parses a comma-separated list of arguments.
    fn parse_arg_list(&mut self, arena: &mut ExprArena) -> NodeId {
        if self.peek().kind == TokenKind::Kw(Kw::RParen) {
            return arena.alloc(ExprNode::ArgList { args: vec![] });
        }
        let mut args = Vec::new();
        loop {
            // Missing argument.
            if self.peek().kind == TokenKind::Kw(Kw::Comma) {
                args.push(arena.alloc(ExprNode::MissingArg));
            } else if self.peek().kind == TokenKind::Kw(Kw::RParen) {
                // A comma was the last thing before the bracket.
                let span = self.current_span();
                self.report_expected(span, "expression");
                break;
            } else if let Some(named) = self.try_parse_named_arg(arena) {
                args.push(named);
            } else {
                // A call-site `ByVal` qualifies the argument that follows.
                self.eat(TokenKind::Kw(Kw::ByVal));
                self.allow_address_of_here();
                args.push(self.parse_expr(arena, 0));
            }
            if !self.eat(TokenKind::Kw(Kw::Comma)) { break; }
        }
        arena.alloc(ExprNode::ArgList { args })
    }

    /// Argument list on a call statement line (no surrounding parens).
    /// An omitted argument (`MsgBox s, , title`) is a `MissingArg`.
    fn parse_stmt_arg_list(&mut self, arena: &mut ExprArena) -> Vec<NodeId> {
        let mut args = Vec::new();
        while !self.peek().is_stmt_end() {
            let a = if self.peek().kind == TokenKind::Kw(Kw::Comma) {
                arena.alloc(ExprNode::MissingArg)
            } else if let Some(named) = self.try_parse_named_arg(arena) {
                named
            } else {
                // A call-site `ByVal` qualifies the argument that follows.
                self.eat(TokenKind::Kw(Kw::ByVal));
                self.allow_address_of_here();
                if let Some(expr) = self.try_parse_expr(arena) {
                    expr
                } else {
                    break;
                }
            };
            args.push(a);
            if !self.eat(TokenKind::Kw(Kw::Comma)) { break; }
            // A comma needs another argument after it, empty or not, before the line ends.
            if self.peek().is_stmt_end()
                || matches!(self.peek().kind, TokenKind::Kw(Kw::Colon) | TokenKind::Kw(Kw::Else))
            {
                let span = self.current_span();
                self.report_expected(span, "expression");
                break;
            }
        }
        args
    }

    /// If the next tokens are `ident :=`, consume them and return a `NamedArg`
    /// node. The label names the *callee's* parameter, not a caller-scope
    /// variable, so it is recorded as a symbol (with its span, for future
    /// goto-parameter) — never as a `NameRef` use that would otherwise resolve in
    /// the caller's scope or be flagged undeclared.
    fn try_parse_named_arg(&mut self, arena: &mut ExprArena) -> Option<NodeId> {
        // Any word names a parameter before `:=` - including words that are
        // keywords elsewhere (`Text`, `Name`, `Key`, `Type`).
        let is_word = match self.peek().kind {
            TokenKind::Ident => true,
            TokenKind::Kw(_) => match self.peek().sym {
                Some(sym) => self
                    .scanner
                    .sym_name(sym as u32)
                    .chars()
                    .next()
                    .map_or(false, |c| c.is_alphabetic() || c == '_'),
                None => false,
            },
            _ => false,
        };
        if !is_word || self.peek2().kind != TokenKind::Kw(Kw::ColonEq) {
            return None;
        }
        let name_tok = self.advance(); // parameter name
        let name = name_tok.sym.map(|s| s as u32).unwrap_or(0);
        self.advance(); // `:=`
        let value = self.parse_expr(arena, 0);
        let id = arena.alloc(ExprNode::NamedArg { name, value });
        self.set_span(id, name_tok.span);
        Some(id)
    }

    /// Parses a Mid argument list.
    pub fn parse_mid_arg_list(&mut self, arena: &mut ExprArena) -> NodeId {
        self.parse_arg_list(arena)
    }

    /// Parses a subscript list.
    pub fn parse_subscript_list(&mut self, arena: &mut ExprArena) -> NodeId {
        self.parse_arg_list(arena)
    }
}

// ── Error codes ──────────────────────────────────────────────────────────────
//
// 0x9c6f = 40047 "Expected: <various>"
//
// 0x9c70 = 40048 "Syntax error"
//
// 0xdee1 = 57057 "Only valid in object module"
/// Longest identifier, in bytes.
const MAX_IDENTIFIER_LEN: usize = 255;
/// Most dimensions an array declaration may list.
const MAX_DIMENSIONS: usize = 60;
/// A fixed-length string holds 1 to this many characters, exclusive.
const FIXED_STRING_LIMIT: i32 = 0xfff7;

const ERR_EXPECTED_EXPR:        u32 = 0x9c6f;
const ERR_EXPECTED_IDENT:       u32 = 0x9c6f;
const ERR_OBJECT_MODULE_ONLY:   u32 = 0xdee1;
const ERR_UNEXPECTED_TOKEN:     u32 = 0x9c6f;
/// "Bad file mode" (54): an `Open` access the file mode cannot have.
const ERR_BAD_FILE_MODE:        u32 = 54;
/// "Invalid inside procedure".
const ERR_INVALID_IN_PROCEDURE: u32 = 0x9c74;
/// "Invalid inside Enum".
const ERR_INVALID_IN_ENUM:      u32 = 0xdeff;
const ERR_EXPECTED_EQ:          u32 = 0x9c6f;
const ERR_EXPECTED_RPAREN:      u32 = 0x9c6f;
const ERR_EXPECTED_LPAREN:      u32 = 0x9c6f;
const ERR_EXPECTED_THEN:        u32 = 0x9c6f;
const ERR_EXPECTED_TO:          u32 = 0x9c6f;
const ERR_EXPECTED_NEXT:        u32 = 0x9c6f;
const ERR_EXPECTED_CASE:        u32 = 0x9c6f;
const ERR_EXPECTED_IS:          u32 = 0x9c6f;
const ERR_EXPECTED_IN:          u32 = 0x9c6f;

/// Return a human-readable label for a token kind, used as the `label` of an
/// `0x9c6f` ("Expected: X") diagnostic.
fn token_kind_label(kind: &TokenKind) -> &'static str {
    match kind {
        TokenKind::Kw(Kw::Then)     => "Then",
        TokenKind::Kw(Kw::To)       => "To",
        TokenKind::Kw(Kw::Next)     => "Next",
        TokenKind::Kw(Kw::Loop)     => "Loop",
        TokenKind::Kw(Kw::Wend)     => "Wend",
        TokenKind::Kw(Kw::End)      => "End",
        TokenKind::Kw(Kw::Case)     => "Case",
        TokenKind::Kw(Kw::Is)       => "Is",
        TokenKind::Kw(Kw::In)       => "In",
        TokenKind::Kw(Kw::Eq)       => "=",
        TokenKind::Kw(Kw::RParen)   => ")",
        TokenKind::Kw(Kw::LParen)   => "(",
        TokenKind::Kw(Kw::Comma)    => ",",
        TokenKind::Kw(Kw::Dot)      => ".",
        TokenKind::Kw(Kw::Minus)    => "-",
        TokenKind::Kw(Kw::With)     => "With",
        TokenKind::Kw(Kw::If)       => "If",
        TokenKind::Kw(Kw::Select)   => "Select",
        TokenKind::Kw(Kw::Type)     => "Type",
        TokenKind::Kw(Kw::Enum)     => "Enum",
        TokenKind::Kw(Kw::Property) => "Property",
        TokenKind::Ident             => "Identifier",
        _                            => "token",
    }
}

/// True when keyword `k` may begin an expression: it is not reserved, or it
/// is a value or built-in function word (`Date`, `Len`, `Array`, ...).
fn keyword_starts_expression(k: Kw) -> bool {
    use super::keyword_table::KEYWORD_TABLE;
    let Some(entry) = KEYWORD_TABLE.get(k.token_id() as usize) else { return true };
    // Operators and punctuation never start an operand, except the opening
    // parenthesis and the member-access dot and bang.
    if !entry.name.starts_with(|c: char| c.is_ascii_alphabetic()) {
        return matches!(entry.name, "(" | "." | "!");
    }
    let reserved = entry.w1 & 0x10 != 0;
    let value_or_function = (entry.w1 >> 6) & 3 != 0;
    !reserved
        || value_or_function
        || matches!(entry.name, "Array" | "CDec" | "DoEvents" | "New" | "Not" | "TypeOf" | "AddressOf"
                | "True" | "False" | "Nothing" | "Empty" | "Null" | "Me" | "Seek"
                | "Decimal" | "DefDec")
}

/// True when the parameter `param` is declared without a type or as `Variant`.
fn param_type_is_variant(arena: &ExprArena, param: NodeId) -> bool {
    let ExprNode::ParamDef { type_node, .. } = arena.get(param) else { return true };
    match type_node {
        None => true,
        Some(t) => matches!(arena.get(*t), ExprNode::BuiltinType { kind: 12 }),
    }
}

/// The words that name module-header attributes. They are keywords only
/// at the start of an `Attribute` line and may be used as ordinary names.
fn is_header_word(name: &str) -> bool {
    name.eq_ignore_ascii_case("attribute") || name.len() > 3 && name[..3].eq_ignore_ascii_case("vb_")
}

/// Keywords that are not reserved but begin a statement of their own: used as a
/// name they need something after them that says so.
const STATEMENT_WORDS: [&str; 8] = ["Width", "Property", "Name", "Mid", "MidB", "Line", "Go", "Error"];

/// The argument grammar of a built-in function: the number of arguments it
/// takes, whether the first must be a variable, and whether the name may stand
/// without brackets.
struct IntrinsicShape {
    name: &'static str,
    min: usize,
    max: usize,
    name_first: bool,
    bare_ok: bool,
}

const fn shape(name: &'static str, min: usize, max: usize) -> IntrinsicShape {
    IntrinsicShape { name, min, max, name_first: false, bare_ok: false }
}

const INTRINSIC_SHAPES: &[IntrinsicShape] = &[
    shape("Abs", 1, 1),
    shape("CBool", 1, 1),
    shape("CByte", 1, 1),
    shape("CCur", 1, 1),
    shape("CDate", 1, 1),
    shape("CDbl", 1, 1),
    shape("CInt", 1, 1),
    shape("CLng", 1, 1),
    shape("CSng", 1, 1),
    shape("CStr", 1, 1),
    shape("CVErr", 1, 1),
    shape("CVar", 1, 1),
    shape("Fix", 1, 1),
    shape("Int", 1, 1),
    shape("Len", 1, 1),
    shape("LenB", 1, 1),
    shape("Sgn", 1, 1),
    shape("InStr", 2, 4),
    shape("InStrB", 2, 4),
    shape("StrComp", 2, 3),
    shape("String", 2, 2),
    IntrinsicShape { name: "UBound", min: 1, max: 2, name_first: true, bare_ok: false },
    IntrinsicShape { name: "LBound", min: 1, max: 2, name_first: true, bare_ok: false },
    IntrinsicShape { name: "Date", min: 0, max: 0, name_first: false, bare_ok: true },
];

/// The bytes of the line that starts at `from`, up to its end.
fn line_rest(src: &[u8], from: usize) -> &[u8] {
    let rest = src.get(from..).unwrap_or(&[]);
    let end = rest.iter().position(|b| *b == b'\n' || *b == b'\r').unwrap_or(rest.len());
    &rest[..end]
}

fn skip_blanks(line: &[u8], mut i: usize) -> usize {
    while i < line.len() && (line[i] == b' ' || line[i] == b'\t') {
        i += 1;
    }
    i
}

/// True when `line` begins with `Attribute`, a name, any `.name` parts and an `=`.
fn attribute_line(line: &[u8]) -> bool {
    let word_len = |from: usize| {
        line[from..].iter().take_while(|b| b.is_ascii_alphanumeric() || **b == b'_').count()
    };
    let i = skip_blanks(line, 0);
    let n = word_len(i);
    if n != 9 || !line[i..i + n].eq_ignore_ascii_case(b"Attribute") {
        return false;
    }
    let mut i = skip_blanks(line, i + n);
    let n = word_len(i);
    if n == 0 {
        return false;
    }
    i += n;
    while i < line.len() && line[i] == b'.' {
        let n = word_len(i + 1);
        if n == 0 {
            return false;
        }
        i += 1 + n;
    }
    let i = skip_blanks(line, i);
    i < line.len() && line[i] == b'='
}
