//! VB6 recursive-descent parser.
//!
//! Implements the source-to-AST stage for Visual Basic 6.
//!
//! **Representation choices:**
//! * Global error-flag + continue → `Diagnostics` collector.
//! * Process-global scanner state → owned `Scanner` threaded via `Parser`.
//! * List chain nodes → `Vec<NodeId>` in typed `ArgList`/`Block` variants.

use crate::frontend::ast::{
    AstLit, BinOpKind, DoKind, ExitKind, ExprArena, ExprNode, FileIoKind, LabelRef, NodeId,
    NodeSpans, OnErrorKind, ProcKind, ResumeTarget, UnOpKind,
};
use crate::frontend::diagnostics::Diagnostics;
use crate::frontend::scanner::{Scanner, ScannerContext};
use crate::frontend::token::{Kw, Lit, Span, Token, TokenKind, TypeSuffix};


// ── Binding-power table (Pratt) ───────────────────────────────────────────────

/// Infix binding powers for VB6 binary operators.
///
/// Returns `(left_bp, right_bp, BinOpKind)` or `None`.
///
/// Precedence levels are: `Imp`=1, `Eqv`=2, `Xor`=3, `Or`=4, `And`=5,
/// comparison/`Like`/`Is`=6, `&`=7, `+`/`-`=8, `Mod`=9, `\`=10, `*`/`/`=11,
/// `^`=12.  Member access (`.`/`!`) is tighter than any of these.
///
/// Each `(left_bp, right_bp)` pair below encodes level `N` as
/// `(2N, 2N+1)` (left-associative) so the relative ordering matches VB6
/// exactly; `^` is right-associative (`left > right`).
fn infix_bp(kind: &TokenKind) -> Option<(u8, u8, BinOpKind)> {
    use BinOpKind as B;
    let (l, r, o) = match kind {
        // Member access (highest infix) — also handled by parse_postfix for Ident bases
        TokenKind::Kw(Kw::Dot)       => (26, 27, B::Dot),
        TokenKind::Kw(Kw::Bang)      => (26, 27, B::Bang),
        // ^ exponentiation, level 12, right-associative
        TokenKind::Kw(Kw::Caret)     => (25, 24, B::Pow),
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
        TokenKind::Kw(Plus)  => Some((24, UnOpKind::Pos)),
        TokenKind::Kw(Not)   => Some((11, UnOpKind::Not)),
        _ => None,
    }
}


/// If `k` is a `Def<Type>` keyword (`DefBool`..`DefVar`, token ids 0x31–0x3c),
/// return its token id; otherwise `None`.  Used by module-level dispatch.
fn def_type_kw(k: Kw) -> Option<u16> {
    use Kw::*;
    matches!(
        k,
        DefBool | DefByte | DefCur | DefDate | DefDec | DefDbl
            | DefInt | DefLng | DefObj | DefSng | DefStr | DefVar
    )
    .then_some(k as u16)
}

/// Apply a conditional-compilation binary operator (tagged by [`eval_cc_expr`]'s
/// `cc_infix_op`) to two i32 operands. Boolean results use VB's -1/0 convention.
fn apply_cc_binop(lhs: i32, rhs: i32, op: u8) -> i32 {
    // VB's boolean result convention: True = -1, False = 0. Folding the six
    // comparison operators through this keeps each a single arm with no branch.
    let vb_bool = |b: bool| if b { -1 } else { 0 };
    // Integer divide / modulo with a zero divisor guarded to 0 (lenient CC eval).
    let checked = |a: i32| if rhs == 0 { 0 } else { a / rhs };
    match op {
        b'*'         => lhs.wrapping_mul(rhs),
        b'/' | b'\\' => checked(lhs),
        b'm'         => if rhs == 0 { 0 } else { lhs % rhs },
        b'+'         => lhs.wrapping_add(rhs),
        b'-'         => lhs.wrapping_sub(rhs),
        b'='         => vb_bool(lhs == rhs),
        b'n'         => vb_bool(lhs != rhs),
        b'<'         => vb_bool(lhs <  rhs),
        b'>'         => vb_bool(lhs >  rhs),
        b'l'         => vb_bool(lhs <= rhs),
        b'g'         => vb_bool(lhs >= rhs),
        b'&'         => lhs & rhs,
        b'|'         => lhs | rhs,
        b'^'         => lhs ^ rhs,
        b'e'         => !(lhs ^ rhs),
        b'i'         => !lhs | rhs,
        _            => 0,
    }
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
    peeked:      Option<Token>,
    /// Second lookahead slot (for `peek2`), filled lazily after `peeked`.
    peeked2:     Option<Token>,
    pub diagnostics: Diagnostics,
    /// Source-span side table for symbol-naming nodes (`NameRef` uses and
    /// declaration-name identifiers).  Consumed by the binder and LSP layer.
    pub node_spans: NodeSpans,
    /// Explicit visibility per declaration node (`NodeId.0` → is_public), set
    /// when an access modifier (`Public`/`Private`/`Friend`/`Global`) is present.
    /// Declarations with no modifier are absent; the binder applies VB6 defaults.
    pub decl_public: std::collections::HashMap<u32, bool>,
    /// Module kind — controls accept/reject for class-only syntax.
    pub module_kind: ModuleKind,
    /// Conditional-compilation constant table: `(lowercase_name, value)`.
    /// Pre-seeded with VB6 predefined constants.
    cc_defs: Vec<(String, i32)>,
}

impl<'src> Parser<'src> {
    /// Construct a parser over `src` using the provided scanner context.
    pub fn new(ctx: &'src mut ScannerContext, src: &'src [u8]) -> Self {
        Self {
            scanner:     Scanner::new(ctx, src),
            peeked:      None,
            peeked2:     None,
            diagnostics: Diagnostics::new(),
            node_spans:  NodeSpans::new(),
            decl_public: std::collections::HashMap::new(),
            module_kind: ModuleKind::Standard,
            cc_defs: vec![
                ("win32".into(), -1),
                ("win64".into(),  0),
                ("win16".into(),  0),
                ("vba6".into(),  -1),
                ("vba7".into(),   0),
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
                self.peeked2.take().unwrap_or_else(|| self.scanner.next_token()),
            );
        }
        self.peeked.as_ref().unwrap()
    }

    /// Two-token lookahead.  Used for label detection (`name :`) and other
    /// cases where one token of lookahead is insufficient.
    fn peek2(&mut self) -> &Token {
        self.peek(); // ensure `peeked` is filled first
        if self.peeked2.is_none() {
            self.peeked2 = Some(self.scanner.next_token());
        }
        self.peeked2.as_ref().unwrap()
    }

    fn advance(&mut self) -> Token {
        match self.peeked.take() {
            Some(t) => t,
            None    => self.peeked2.take().unwrap_or_else(|| self.scanner.next_token()),
        }
    }

    fn current_span(&mut self) -> Span {
        self.peek().span
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

    /// Consume a line-continuation sequence (`_` followed by an end-of-line).
    ///
    /// VB6 allows `_` at the end of a physical line to join it with the next.
    /// The scanner emits a `LineCont` token for the `_` and a separate `Eol`
    /// for the newline.  This helper eats both so callers can treat the logical
    /// line as unbroken.  Returns `true` if a `LineCont` was consumed.
    fn eat_line_cont(&mut self) -> bool {
        if self.peek().kind == TokenKind::Kw(Kw::LineCont) {
            self.advance();          // consume _
            self.eat(TokenKind::Eol); // consume the following newline if present
            true
        } else {
            false
        }
    }

    /// Skip all consecutive line-continuation sequences at the current position.
    fn skip_line_conts(&mut self) {
        while self.eat_line_cont() {}
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

    fn at_block_end(&mut self) -> bool {
        matches!(
            self.peek().kind,
            TokenKind::Eof
                | TokenKind::Kw(Kw::End)
                | TokenKind::Kw(Kw::EndIf)
                | TokenKind::Kw(Kw::Else)
                | TokenKind::Kw(Kw::ElseIf)
                | TokenKind::Kw(Kw::Loop)
                | TokenKind::Kw(Kw::Next)
                | TokenKind::Kw(Kw::Wend)
                | TokenKind::Kw(Kw::Case)
        )
    }

    /// Like `at_block_end` but only stops at the specific `End Sub/Function/Property`
    /// tokens that terminate a procedure body. This allows stranded `End If` (left by
    /// `#End If` splitting an If/Else block) to be consumed inside the proc body by
    /// `parse_stmt` rather than being mistaken for the proc terminator.
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

    /// Parse a procedure body — like `parse_block` but uses `at_proc_end` as the
    /// stop condition so that orphaned `End If` tokens (produced when `#End If`
    /// splits an If/Else block) are consumed as statements rather than halting
    /// block parsing early.
    fn parse_block_proc(&mut self, arena: &mut ExprArena) -> NodeId {
        let mut stmts: Vec<NodeId> = Vec::new();
        loop {
            while self.eat(TokenKind::Eol) {}
            if self.at_proc_end() || self.peek().kind == TokenKind::Eof { break; }
            if let Some(n) = self.parse_stmt(arena) {
                stmts.push(n);
            }
        }
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
                TokenKind::Kw(Kw::Public)
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
                    let span = self.current_span();
                    if self.module_kind == ModuleKind::Standard {
                        self.diagnostics.push(ERR_OBJECT_MODULE_ONLY, span);
                    }
                    let n = self.parse_event_decl(arena);
                    nodes.push(n);
                }
                // Implements — class-module-only
                TokenKind::Kw(Kw::Implements) => {
                    let span = self.current_span();
                    if self.module_kind == ModuleKind::Standard {
                        self.diagnostics.push(ERR_OBJECT_MODULE_ONLY, span);
                    }
                    let n = self.parse_implements_stmt(arena);
                    nodes.push(n);
                }
                // Conditional compilation directives
                TokenKind::Kw(Kw::CcIf)
                | TokenKind::Kw(Kw::CcElse)
                | TokenKind::Kw(Kw::CcElseIf)
                | TokenKind::Kw(Kw::CcEnd)
                | TokenKind::Kw(Kw::CcConst) => {
                    let n = self.parse_cc_directive(arena);
                    nodes.push(n);
                }
                // Option statements
                TokenKind::Kw(Kw::Option) => {
                    let n = self.parse_option_stmt(arena);
                    nodes.push(n);
                }
                // Module-level Dim / Const / ReDim
                TokenKind::Kw(Kw::Dim)
                | TokenKind::Kw(Kw::Const)
                | TokenKind::Kw(Kw::ReDim)
                | TokenKind::Kw(Kw::Global) => {
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
                        self.recover_module_unexpected();
                    }
                }
                _ => self.recover_module_unexpected(),
            }
        }
        nodes
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
            if self.eat(TokenKind::Eol) { continue; }
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
            self.parse_ident_stmt_with_tok(arena, t)
        }
    }

    /// `Input`/`Input$`/`InputB`/`InputB$`: `#n, vars` reads a file; otherwise
    /// an implicit call.
    fn parse_input_keyword_stmt(&mut self, arena: &mut ExprArena) -> NodeId {
        let t = self.advance();
        if self.peek().kind == TokenKind::Kw(Kw::Hash) {
            self.parse_file_input_after_keyword(arena)
        } else {
            self.parse_ident_stmt_with_tok(arena, t)
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
            self.eat_eol();
            arena.alloc(ExprNode::FileIoStmt {
                kind: FileIoKind::LineInput,
                channel: Some(ch),
                args: vec![var],
            })
        } else if self.peek().kind == TokenKind::Kw(Kw::LParen)
            || (self.peek().kind == TokenKind::Kw(Kw::Step)
                && self.peek2().kind == TokenKind::Kw(Kw::LParen))
        {
            // `Line [Step](x1,y1)-(x2,y2)[,color[,flags]]` — VB6 graphics method.
            self.parse_line_graphics_stmt(arena)
        } else {
            self.parse_ident_stmt_with_tok(arena, t)
        }
    }

    /// Parse `[Step](x, y)` — a VB6 graphics coordinate pair.
    /// Returns `(x_node, y_node)`.
    fn parse_coord_pair(&mut self, arena: &mut ExprArena) -> (NodeId, NodeId) {
        self.eat(TokenKind::Kw(Kw::Step)); // optional Step keyword
        self.expect(TokenKind::Kw(Kw::LParen), ERR_UNEXPECTED_TOKEN);
        let x = self.parse_expr(arena, 0);
        self.expect(TokenKind::Kw(Kw::Comma), ERR_UNEXPECTED_TOKEN);
        let y = self.parse_expr(arena, 0);
        self.expect(TokenKind::Kw(Kw::RParen), ERR_EXPECTED_RPAREN);
        (x, y)
    }

    /// Parse `[Step](x1,y1)-[Step](x2,y2)[,color[,flags]]`.
    /// Called after the `Line` keyword has already been consumed.
    fn parse_line_graphics_stmt(&mut self, arena: &mut ExprArena) -> NodeId {
        let (x1, y1) = self.parse_coord_pair(arena);
        self.expect(TokenKind::Kw(Kw::Minus), ERR_UNEXPECTED_TOKEN);
        let (x2, y2) = self.parse_coord_pair(arena);
        let mut args = vec![x1, y1, x2, y2];
        // Optional: `, color [, flags…]`
        if self.eat(TokenKind::Kw(Kw::Comma)) {
            if let Some(color) = self.try_parse_expr(arena) {
                args.push(color);
            }
            while self.eat(TokenKind::Kw(Kw::Comma)) {
                if let Some(flag) = self.try_parse_expr(arena) {
                    args.push(flag);
                }
            }
        }
        self.eat_eol();
        let arg_list = arena.alloc(ExprNode::ArgList { args });
        arena.alloc(ExprNode::Block { stmts: vec![arg_list] })
    }

    /// `Get`/`Put` random-access I/O: `#n, …`; otherwise an implicit call.
    fn parse_get_put_keyword_stmt(&mut self, arena: &mut ExprArena, kind: FileIoKind) -> NodeId {
        let t = self.advance();
        if self.peek().kind == TokenKind::Kw(Kw::Hash) {
            self.parse_file_get_put_after_keyword(arena, kind)
        } else {
            self.parse_ident_stmt_with_tok(arena, t)
        }
    }

    /// `Lock`/`Unlock`: `#n, …`; otherwise an implicit call.
    fn parse_lock_keyword_stmt(&mut self, arena: &mut ExprArena, kind: FileIoKind) -> NodeId {
        let t = self.advance();
        if self.peek().kind == TokenKind::Kw(Kw::Hash) {
            self.parse_file_lock_after_keyword(arena, kind)
        } else {
            self.parse_ident_stmt_with_tok(arena, t)
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
        if self.peek().kind == TokenKind::Kw(Kw::Hash) {
            self.eat(TokenKind::Kw(Kw::Hash));
            let ch = self.parse_expr(arena, 0);
            self.expect(TokenKind::Kw(Kw::Comma), ERR_UNEXPECTED_TOKEN);
            let val = self.parse_expr(arena, 0);
            self.eat_eol();
            arena.alloc(ExprNode::FileIoStmt {
                kind,
                channel: Some(ch),
                args: vec![val],
            })
        } else {
            self.parse_ident_stmt_with_tok(arena, t)
        }
    }

    /// `Name "old" As "new"`; or a `Name`-identifier assignment / call fallback.
    fn parse_name_keyword_stmt(&mut self, arena: &mut ExprArena) -> NodeId {
        let t = self.advance();
        if self.peek().is_stmt_end() {
            return self.parse_ident_stmt_with_tok(arena, t);
        }
        // Name "old" As "new"
        let old_path = self.parse_expr(arena, 0);
        if self.eat(TokenKind::Kw(Kw::As)) {
            let new_path = self.parse_expr(arena, 0);
            self.eat_eol();
            return arena.alloc(ExprNode::FileIoStmt {
                kind: FileIoKind::Name,
                channel: None,
                args: vec![old_path, new_path],
            });
        }
        // Could be assignment: Name = expr
        if self.eat(TokenKind::Kw(Kw::Eq)) {
            let value = self.parse_expr(arena, 0);
            self.eat_eol();
            let sym = t.sym.map(|s| s as u32).unwrap_or(0);
            let lhs = arena.alloc(ExprNode::NameRef { sym, suffix: t.type_suffix });
            self.set_span(lhs, t.span);
            return arena.alloc(ExprNode::Assign { target: lhs, value });
        }
        self.eat_eol();
        old_path
    }

    /// `End`: bare `End` terminates the program; `End If`/`End With`/… are
    /// structural block terminators consumed by the parent block parser.
    fn parse_end_stmt(&mut self, arena: &mut ExprArena) -> Option<NodeId> {
        // "End" alone terminates the program.
        // End Sub / End Function etc. are handled in their parent.
        self.advance();
        if self.peek().is_stmt_end() {
            self.eat_eol();
            return Some(arena.alloc(ExprNode::EndStmt));
        }
        // End If / End With / End Select / End Type / End Enum
        self.advance(); // consume the sub-keyword
        self.eat_eol();
        None // structural: consumed by the block parser
    }

    /// `LSet`/`RSet target = expr` — range-copy assignment preserving justification.
    fn parse_range_assign_stmt(&mut self, arena: &mut ExprArena) -> NodeId {
        let right_justify = matches!(self.peek().kind, TokenKind::Kw(Kw::RSet));
        self.advance();
        let target = self.parse_target_expr(arena);
        self.expect(TokenKind::Kw(Kw::Eq), ERR_EXPECTED_EQ);
        let value = self.parse_expr(arena, 0);
        self.eat_eol();
        arena.alloc(ExprNode::RangeAssign { right_justify, target, value })
    }

    // ── Statement dispatch ────────────────────────────────────────────────────

    /// Dispatches to the appropriate sub-parser based on the current token.
    /// Returns `None` on a pure blank/comment line.
    pub fn parse_stmt(&mut self, arena: &mut ExprArena) -> Option<NodeId> {
        // Skip statement ends
        while self.eat(TokenKind::Eol) {}
        if self.peek().kind == TokenKind::Eof { return None; }

        let kind = self.peek().kind.clone();
        let node = match &kind {
            // Variable / constant declarations
            TokenKind::Kw(Kw::Dim)
            | TokenKind::Kw(Kw::Static)
            | TokenKind::Kw(Kw::Public)
            | TokenKind::Kw(Kw::Private)
            | TokenKind::Kw(Kw::Global) => {
                self.advance();
                self.parse_dim_list(arena, false)
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
                self.eat_eol();
                arena.alloc(ExprNode::ReturnStmt)
            }
            TokenKind::Kw(Kw::GoTo)   => self.parse_goto_stmt(arena, false),
            TokenKind::Kw(Kw::GoSub)  => self.parse_goto_stmt(arena, true),
            // Space-separated `Go To` / `Go Sub` — the second keyword selects
            // the kind. VB6 accepts both spellings (oracle-confirmed).
            TokenKind::Kw(Kw::Go)     => self.parse_goto_stmt(arena, false),
            TokenKind::Kw(Kw::Resume) => self.parse_resume_stmt(arena),
            TokenKind::Kw(Kw::On)     => self.parse_on_stmt(arena),
            // Error
            TokenKind::Kw(Kw::Error) => {
                self.advance();
                let expr = self.parse_expr(arena, 0);
                self.eat_eol();
                arena.alloc(ExprNode::ErrorStmt { expr })
            }
            // Misc
            TokenKind::Kw(Kw::Stop) => {
                self.advance(); self.eat_eol();
                arena.alloc(ExprNode::Stop)
            }
            TokenKind::Kw(Kw::End) => return self.parse_end_stmt(arena),
            TokenKind::Kw(Kw::Erase)  => self.parse_erase_stmt(arena),
            TokenKind::Kw(Kw::RaiseEvent) => self.parse_raise_event(arena),
            TokenKind::Kw(Kw::Debug)  => self.parse_debug_stmt(arena),

            // Attribute metadata line
            TokenKind::Kw(Kw::Attribute) => self.parse_attribute_stmt(arena),
            // Conditional compilation directives (inside proc body)
            TokenKind::Kw(Kw::CcIf)
            | TokenKind::Kw(Kw::CcElseIf)
            | TokenKind::Kw(Kw::CcElse)
            | TokenKind::Kw(Kw::CcEnd)
            | TokenKind::Kw(Kw::CcConst) => self.parse_cc_directive(arena),
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
                self.parse_mid_assign(arena)
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
            // Sub / Function / Property — nested (error but recover)
            TokenKind::Kw(Kw::Sub) | TokenKind::Kw(Kw::Function) | TokenKind::Kw(Kw::Property) => {
                let span = self.current_span();
                self.diagnostics.push(ERR_UNEXPECTED_TOKEN, span);
                self.skip_to_stmt_end();
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
                let sym = name.sym.map(|s| s as u32).unwrap_or(0);
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
                let tok = self.advance(); // consume the line number
                let line = line_label_value(&tok.lit);
                let _ = self.eat(TokenKind::Kw(Kw::Colon)); // optional colon
                let id = arena.alloc(ExprNode::Label { target: LabelRef::Line(line) });
                self.set_span(id, tok.span);
                id
            }
            // Non-reserved keyword used as identifier.
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
        let lhs = self.parse_target_expr(arena);
        if self.eat(TokenKind::Kw(Kw::Eq)) {
            let value = self.parse_expr(arena, 0);
            self.eat_eol();
            arena.alloc(ExprNode::Assign { target: lhs, value })
        } else if self.eat(TokenKind::Kw(Kw::ColonEq)) {
            // Named argument in a call — parse rest as expression
            let value = self.parse_expr(arena, 0);
            self.eat_eol();
            arena.alloc(ExprNode::Assign { target: lhs, value })
        } else if self.peek().kind == TokenKind::Kw(Kw::Minus)
            && (self.peek2().kind == TokenKind::Kw(Kw::LParen)
                || self.peek2().kind == TokenKind::Kw(Kw::Step))
        {
            // `Obj.Line (x1,y1)-(x2,y2)` — postfix consumed `(x1,y1)` as a call
            // arg-list; the `-(x2,y2)` continuation signals a graphics coord-pair call.
            self.advance(); // consume `-`
            let (x2, y2) = self.parse_coord_pair(arena);
            let mut extra = vec![x2, y2];
            if self.eat(TokenKind::Kw(Kw::Comma)) {
                if let Some(color) = self.try_parse_expr(arena) {
                    extra.push(color);
                }
                while self.eat(TokenKind::Kw(Kw::Comma)) {
                    if let Some(flag) = self.try_parse_expr(arena) {
                        extra.push(flag);
                    }
                }
            }
            self.eat_eol();
            let args = arena.alloc(ExprNode::ArgList { args: extra });
            arena.alloc(ExprNode::CallStmt { callee: lhs, args })
        } else {
            // Implicit call with optional trailing arguments: `Foo a, b`
            let args = self.parse_stmt_arg_list(arena);
            self.eat_eol();
            if args.is_empty() {
                lhs
            } else {
                let args = arena.alloc(ExprNode::ArgList { args });
                arena.alloc(ExprNode::CallStmt { callee: lhs, args })
            }
        }
    }

    /// Parses a `Let` assignment statement.
    fn parse_let_stmt(&mut self, arena: &mut ExprArena) -> NodeId {
        self.advance(); // consume Let
        let target = self.parse_target_expr(arena);
        self.expect(TokenKind::Kw(Kw::Eq), ERR_EXPECTED_EQ);
        let value = self.parse_expr(arena, 0);
        self.eat_eol();
        arena.alloc(ExprNode::Assign { target, value })
    }

    /// Parses a `Set` assignment statement.
    fn parse_set_stmt(&mut self, arena: &mut ExprArena) -> NodeId {
        self.advance(); // consume Set
        let target = self.parse_target_expr(arena);
        self.expect(TokenKind::Kw(Kw::Eq), ERR_EXPECTED_EQ);
        let value = self.parse_expr(arena, 0);
        self.eat_eol();
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
        self.expect(TokenKind::Kw(Kw::RParen), ERR_EXPECTED_RPAREN);
        self.expect(TokenKind::Kw(Kw::Eq), ERR_EXPECTED_EQ);
        let value = self.parse_expr(arena, 0);
        self.eat_eol();
        arena.alloc(ExprNode::MidAssign { byte_oriented, dollar, args, value })
    }

    /// Parse a `Print`/`?` output list.
    fn parse_print_list(&mut self, arena: &mut ExprArena, channel: Option<NodeId>) -> NodeId {
        let mut args = Vec::new();
        while !self.peek().is_stmt_end() {
            if let Some(n) = self.try_parse_expr(arena) { args.push(n); }
            if !self.eat(TokenKind::Kw(Kw::Comma)) && !self.eat(TokenKind::Kw(Kw::Semi)) {
                break;
            }
        }
        self.eat_eol();
        arena.alloc(ExprNode::FileIoStmt { kind: FileIoKind::Print, channel, args })
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
        let mut stmts = Vec::new();
        loop {
            // Variable declarations accept `As Type()`; `Const` does not.
            let item = self.parse_dim_item(arena, is_const, !is_const);
            stmts.push(item);
            if !self.eat(TokenKind::Kw(Kw::Comma)) { break; }
        }
        self.eat_eol();
        arena.alloc(ExprNode::Block { stmts })
    }

    /// Parse a Const declaration list.
    fn parse_const_list(&mut self, arena: &mut ExprArena) -> NodeId {
        self.parse_dim_list(arena, true)
    }

    /// Single variable: `name [As type]` or `name(dims) [As type]` or `name = expr`.
    /// `array_suffix_ok` permits a trailing `As Type()` array marker.
    fn parse_dim_item(&mut self, arena: &mut ExprArena, is_const: bool, array_suffix_ok: bool) -> NodeId {
        let name = self.expect_ident();
        let name_id = name.sym.map(|s| s as u32).unwrap_or(0);
        let bounds = if self.eat(TokenKind::Kw(Kw::LParen)) {
            let d = self.parse_array_bounds(arena);
            self.expect(TokenKind::Kw(Kw::RParen), ERR_EXPECTED_RPAREN);
            d
        } else {
            None
        };
        let type_node = if self.eat(TokenKind::Kw(Kw::As)) {
            Some(self.parse_type_spec(arena, array_suffix_ok))
        } else {
            // `Dim x%` ≡ `Dim x As Integer`
            self.type_node_from_suffix(arena, name.type_suffix)
        };
        let init = if is_const || self.peek().kind == TokenKind::Kw(Kw::Eq) {
            if self.eat(TokenKind::Kw(Kw::Eq)) { Some(self.parse_expr(arena, 0)) } else { None }
        } else { None };
        let id = arena.alloc(ExprNode::DimItem { name: name_id, is_const, bounds, type_node, init });
        self.set_span(id, name.span);
        id
    }

    /// Parses array bounds. Returns the bounds-list `ArgList` node, or `None`
    /// for a dynamic array (empty `()`).
    fn parse_array_bounds(&mut self, arena: &mut ExprArena) -> Option<NodeId> {
        if self.peek().kind == TokenKind::Kw(Kw::RParen) {
            return None; // dynamic array: no bounds
        }
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
            if self.peek().kind == TokenKind::Kw(Kw::RParen) { break; }
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
            let name = self.expect_ident();
            let name_id = name.sym.map(|s| s as u32).unwrap_or(0);
            self.expect(TokenKind::Kw(Kw::LParen), ERR_EXPECTED_LPAREN);
            let bounds = self.parse_array_bounds(arena);
            self.expect(TokenKind::Kw(Kw::RParen), ERR_EXPECTED_RPAREN);
            let type_node = if self.eat(TokenKind::Kw(Kw::As)) {
                Some(self.parse_type_spec(arena, false))
            } else { None };
            stmts.push(arena.alloc(ExprNode::ReDimItem { preserve, name: name_id, bounds, type_node }));
            if !self.eat(TokenKind::Kw(Kw::Comma)) { break; }
        }
        self.eat_eol();
        arena.alloc(ExprNode::Block { stmts })
    }

    /// Parses a For loop.
    fn parse_for_stmt(&mut self, arena: &mut ExprArena) -> NodeId {
        self.advance(); // consume For
        if self.eat(TokenKind::Kw(Kw::Each)) {
            return self.parse_for_each_stmt(arena);
        }
        let var   = self.parse_target_expr(arena);
        self.expect(TokenKind::Kw(Kw::Eq), ERR_EXPECTED_EQ);
        let start = self.parse_expr(arena, 0);
        self.expect(TokenKind::Kw(Kw::To), ERR_EXPECTED_TO);
        let end = self.parse_expr(arena, 0);
        let step = if self.eat(TokenKind::Kw(Kw::Step)) {
            Some(self.parse_expr(arena, 0))
        } else { None };
        self.eat_eol();
        let body = self.parse_block(arena);
        // For is implicitly terminated by the procedure boundary if Next is missing.
        let at_term = matches!(self.peek().kind, TokenKind::Eof | TokenKind::Kw(Kw::End));
        if !at_term {
            self.expect(TokenKind::Kw(Kw::Next), ERR_EXPECTED_NEXT);
            // `Next [var [, var]…]`
            while !self.peek().is_stmt_end() {
                match self.peek().kind {
                    TokenKind::Ident | TokenKind::Kw(_) => { self.advance(); }
                    _ => break,
                }
                if !self.eat(TokenKind::Kw(Kw::Comma)) { break; }
            }
            self.eat_eol();
        }
        arena.alloc(ExprNode::For { var, start, end, step, body })
    }

    fn parse_for_each_stmt(&mut self, arena: &mut ExprArena) -> NodeId {
        let var = self.parse_target_expr(arena);
        self.expect(TokenKind::Kw(Kw::In), ERR_EXPECTED_IN);
        let collection = self.parse_expr(arena, 0);
        self.eat_eol();
        let body = self.parse_block(arena);
        self.expect(TokenKind::Kw(Kw::Next), ERR_EXPECTED_NEXT);
        while !self.peek().is_stmt_end() {
            match self.peek().kind {
                TokenKind::Ident | TokenKind::Kw(_) => { self.advance(); }
                _ => break,
            }
            if !self.eat(TokenKind::Kw(Kw::Comma)) { break; }
        }
        self.eat_eol();
        arena.alloc(ExprNode::ForEach { var, collection, body })
    }

    /// Parses a Do loop.
    fn parse_do_stmt(&mut self, arena: &mut ExprArena) -> NodeId {
        self.advance(); // consume Do
        let pre = match self.peek().kind.clone() {
            TokenKind::Kw(Kw::While) => { self.advance(); Some((DoKind::PreWhile, self.parse_expr(arena, 0))) }
            TokenKind::Kw(Kw::Until) => { self.advance(); Some((DoKind::PreUntil, self.parse_expr(arena, 0))) }
            _ => None,
        };
        self.eat_eol();
        let body = self.parse_block(arena);
        // Do is implicitly terminated by the procedure boundary if Loop is missing.
        let post = if matches!(self.peek().kind, TokenKind::Eof | TokenKind::Kw(Kw::End)) {
            None
        } else {
            self.expect(TokenKind::Kw(Kw::Loop), ERR_EXPECTED_LOOP);
            let p = match self.peek().kind.clone() {
                TokenKind::Kw(Kw::While) => { self.advance(); Some((DoKind::PostWhile, self.parse_expr(arena, 0))) }
                TokenKind::Kw(Kw::Until) => { self.advance(); Some((DoKind::PostUntil, self.parse_expr(arena, 0))) }
                _ => None,
            };
            self.eat_eol();
            p
        };
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
        self.eat_eol();
        let body = self.parse_block(arena);
        self.expect(TokenKind::Kw(Kw::Wend), ERR_EXPECTED_WEND);
        self.eat_eol();
        arena.alloc(ExprNode::While { cond, body })
    }

    /// Parses a With statement.
    fn parse_with_stmt(&mut self, arena: &mut ExprArena) -> NodeId {
        self.advance(); // consume With
        let obj = self.parse_expr(arena, 0);
        self.eat_eol();
        let body = self.parse_block(arena);
        self.expect(TokenKind::Kw(Kw::End), ERR_EXPECTED_END);
        self.expect(TokenKind::Kw(Kw::With), ERR_UNEXPECTED_TOKEN);
        self.eat_eol();
        arena.alloc(ExprNode::With { obj, body })
    }

    /// Parses an If statement.
    fn parse_if_stmt(&mut self, arena: &mut ExprArena) -> NodeId {
        self.advance(); // consume If
        let cond = self.parse_expr(arena, 0);

        // Legacy single-line form `If <cond> GoTo|GoSub <label>` (no `Then`).
        let legacy_goto = matches!(
            self.peek().kind,
            TokenKind::Kw(Kw::GoTo) | TokenKind::Kw(Kw::GoSub) | TokenKind::Kw(Kw::Go)
        );
        let has_then = if legacy_goto {
            false
        } else {
            self.expect(TokenKind::Kw(Kw::Then), ERR_EXPECTED_THEN);
            true
        };

        // Single-line If: a `Then` followed immediately by a statement (not Eol),
        // or the legacy `GoTo` form.
        if legacy_goto || (has_then && !self.peek().is_stmt_end()) {
            let then_n = if let Some(n) = self.parse_stmt(arena) { n }
                         else { arena.alloc(ExprNode::Block { stmts: vec![] }) };
            let then_body = arena.alloc(ExprNode::Block { stmts: vec![then_n] });
            let else_body = if self.eat(TokenKind::Kw(Kw::Else)) {
                let e = if let Some(n) = self.parse_stmt(arena) { n }
                         else { arena.alloc(ExprNode::Block { stmts: vec![] }) };
                Some(arena.alloc(ExprNode::Block { stmts: vec![e] }))
            } else { None };
            return arena.alloc(ExprNode::If { cond, then_body, else_body });
        }
        self.eat_eol();
        let then_body = self.parse_block(arena);
        let else_body = self.parse_else_clause(arena);
        self.end_if(); // `End If` or the one-word `EndIf`
        self.eat_eol();
        arena.alloc(ExprNode::If { cond, then_body, else_body })
    }

    /// Consume a block-`If` terminator.
    fn end_if(&mut self) {
        if self.eat(TokenKind::Kw(Kw::EndIf)) {
            return;
        }
        self.expect(TokenKind::Kw(Kw::End), ERR_EXPECTED_END);
        self.expect(TokenKind::Kw(Kw::If), ERR_UNEXPECTED_TOKEN);
    }

    fn parse_else_clause(&mut self, arena: &mut ExprArena) -> Option<NodeId> {
        match self.peek().kind.clone() {
            TokenKind::Kw(Kw::Else) => {
                self.advance();
                self.eat_eol();
                Some(self.parse_block(arena))
            }
            TokenKind::Kw(Kw::ElseIf) => {
                self.advance();
                let cond = self.parse_expr(arena, 0);
                self.expect(TokenKind::Kw(Kw::Then), ERR_EXPECTED_THEN);
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
        self.eat_eol();
        let mut pre = Vec::new();
        let mut cases = Vec::new();
        loop {
            while self.eat(TokenKind::Eol) {}
            match self.peek().kind.clone() {
                TokenKind::Kw(Kw::Case) => {
                    self.advance();
                    cases.push(self.parse_case_clause(arena));
                }
                TokenKind::Kw(Kw::End) => {
                    self.advance();
                    self.expect(TokenKind::Kw(Kw::Select), ERR_UNEXPECTED_TOKEN);
                    self.eat_eol();
                    break;
                }
                TokenKind::Eof => break,
                _ => {
                    // Statements before the first Case clause.
                    if self.peek().is_stmt_end() {
                        self.advance();
                    } else if let Some(stmt) = self.parse_stmt(arena) {
                        pre.push(stmt);
                    }
                }
            }
        }
        arena.alloc(ExprNode::SelectCase { subject, pre, cases })
    }

    fn parse_case_clause(&mut self, arena: &mut ExprArena) -> NodeId {
        if self.eat(TokenKind::Kw(Kw::Else)) {
            self.eat_eol();
            let body = self.parse_block(arena);
            arena.alloc(ExprNode::CaseElse { body })
        } else {
            let test = self.parse_case_test_expr(arena);
            self.eat_eol();
            let body = self.parse_block(arena);
            arena.alloc(ExprNode::CaseBlock { test, body })
        }
    }

    /// Parses Case test expressions like `Is op expr`, `lo To hi`, or bare `expr`.
    fn parse_case_test_expr(&mut self, arena: &mut ExprArena) -> NodeId {
        let mut args = Vec::new();
        loop {
            let t = if self.eat(TokenKind::Kw(Kw::Is)) {
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
                let lo = self.parse_expr(arena, 0);
                if self.eat(TokenKind::Kw(Kw::To)) {
                    let hi = self.parse_expr(arena, 0);
                    arena.alloc(ExprNode::RangeTo { lo, hi })
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
        let callee = self.parse_target_expr(arena);
        let args = if self.eat(TokenKind::Kw(Kw::LParen)) {
            let a = self.parse_arg_list(arena);
            self.expect(TokenKind::Kw(Kw::RParen), ERR_EXPECTED_RPAREN);
            a
        } else {
            let args = self.parse_stmt_arg_list(arena);
            arena.alloc(ExprNode::ArgList { args })
        };
        self.eat_eol();
        arena.alloc(ExprNode::CallStmt { callee, args })
    }

    fn parse_exit_stmt(&mut self, arena: &mut ExprArena) -> NodeId {
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
        self.eat_eol();
        arena.alloc(ExprNode::ExitStmt { kind })
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
        let target = if matches!(self.peek().kind, TokenKind::IntLit | TokenKind::LongLit) {
            let tok = self.advance();
            LabelRef::Line(line_label_value(&tok.lit))
        } else {
            LabelRef::Name(self.expect_ident().sym.map(|s| s as u32).unwrap_or(0))
        };
        self.eat_eol();
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
            ResumeTarget::At(LabelRef::Line(line_label_value(&tok.lit)))
        } else if !self.peek().is_stmt_end() {
            let lbl = self.expect_ident();
            ResumeTarget::At(LabelRef::Name(lbl.sym.map(|s| s as u32).unwrap_or(0)))
        } else {
            ResumeTarget::Retry
        };
        self.eat_eol();
        arena.alloc(ExprNode::Resume { target })
    }

    fn parse_on_stmt(&mut self, arena: &mut ExprArena) -> NodeId {
        self.advance(); // consume On
        // `On Local Error …` is the explicit form of `On Error …`; `Local` is the
        // default error scope and a reserved word, so eating it here is
        // unambiguous and leaves the handler form unchanged. (oracle-confirmed)
        self.eat(TokenKind::Kw(Kw::Local));
        if self.eat(TokenKind::Kw(Kw::Error)) {
            if self.eat(TokenKind::Kw(Kw::Resume)) {
                self.expect(TokenKind::Kw(Kw::Next), ERR_EXPECTED_NEXT);
                self.eat_eol();
                return arena.alloc(ExprNode::OnError { kind: OnErrorKind::ResumeNext });
            }
            if self.eat(TokenKind::Kw(Kw::GoTo)) {
                if matches!(self.peek().kind, TokenKind::IntLit | TokenKind::LongLit) {
                    let tok = self.advance();
                    let line = line_label_value(&tok.lit);
                    self.eat_eol();
                    // Only `On Error GoTo 0` disables the handler; any nonzero
                    // line number installs a handler at that numeric line label.
                    let kind = if line == 0 {
                        OnErrorKind::Disable
                    } else {
                        OnErrorKind::Goto(LabelRef::Line(line))
                    };
                    return arena.alloc(ExprNode::OnError { kind });
                }
                let lbl = self.expect_ident();
                let lid = lbl.sym.map(|s| s as u32).unwrap_or(0);
                self.eat_eol();
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
                let n = arena.alloc(ExprNode::Literal { lit: AstLit::Int(line_label_value(&tok.lit)) });
                self.set_span(n, tok.span);
                n
            } else {
                let lbl = self.expect_ident();
                let lid = lbl.sym.map(|s| s as u32).unwrap_or(0);
                let nref = arena.alloc(ExprNode::NameRef { sym: lid, suffix: TypeSuffix::None });
                self.set_span(nref, lbl.span);
                nref
            };
            labels.push(lref);
            if !self.eat(TokenKind::Kw(Kw::Comma)) { break; }
        }
        self.eat_eol();
        arena.alloc(ExprNode::OnGo { is_gosub, expr, labels })
    }

    fn parse_erase_stmt(&mut self, arena: &mut ExprArena) -> NodeId {
        self.advance(); // consume Erase
        let mut vars = Vec::new();
        loop {
            vars.push(self.parse_target_expr(arena));
            if !self.eat(TokenKind::Kw(Kw::Comma)) { break; }
        }
        self.eat_eol();
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
        self.eat_eol();
        arena.alloc(ExprNode::RaiseEvent { name: nid, args })
    }

    fn parse_debug_stmt(&mut self, arena: &mut ExprArena) -> NodeId {
        self.advance(); // consume Debug
        self.expect(TokenKind::Kw(Kw::Dot), ERR_UNEXPECTED_TOKEN);
        self.advance(); // consume Print / Assert etc.
        let mut args = Vec::new();
        while !self.peek().is_stmt_end() {
            if let Some(n) = self.try_parse_expr(arena) {
                args.push(n);
            }
            if !self.eat(TokenKind::Kw(Kw::Semi)) && !self.eat(TokenKind::Kw(Kw::Comma)) {
                break;
            }
        }
        self.eat_eol();
        arena.alloc(ExprNode::DebugPrint { args })
    }

    // ── Conditional-compilation directive parser ─────────────────────────────

    fn parse_cc_directive(&mut self, arena: &mut ExprArena) -> NodeId {
        match self.peek().kind.clone() {
            TokenKind::Kw(Kw::CcConst) => {
                self.advance(); // #Const
                let name_tok = self.expect_ident();
                let name = name_tok.sym
                    .map(|id| self.scanner.sym_name(id as u32).to_ascii_lowercase())
                    .unwrap_or_default();
                self.expect(TokenKind::Kw(Kw::Eq), ERR_EXPECTED_EQ);
                let value = self.eval_cc_expr(0);
                self.cc_set(&name, value);
                self.eat_eol();
                arena.alloc(ExprNode::Block { stmts: vec![] })
            }
            TokenKind::Kw(Kw::CcIf) => {
                self.advance(); // #If
                let cond = self.eval_cc_expr(0);
                self.eat(TokenKind::Kw(Kw::Then));
                self.eat_eol();
                self.parse_cc_chain(arena, cond != 0)
            }
            _ => {
                // Standalone #Else, #ElseIf, #End — consume and skip
                self.advance();
                self.skip_to_stmt_end();
                self.eat_eol();
                arena.alloc(ExprNode::Block { stmts: vec![] })
            }
        }
    }

    /// Parse the body of a CC if/elseif/else chain.
    /// `include`: whether this branch is currently active (should be parsed).
    /// Returns a Block node of the active branch's statements.
    fn parse_cc_chain(&mut self, arena: &mut ExprArena, include: bool) -> NodeId {
        let result = if include {
            self.parse_cc_active_branch(arena)
        } else {
            self.skip_cc_false_branch();
            arena.alloc(ExprNode::Block { stmts: vec![] })
        };
        self.parse_cc_continuation(arena, include, result)
    }

    /// Parse the statements of an active CC branch up to the next CC terminator.
    fn parse_cc_active_branch(&mut self, arena: &mut ExprArena) -> NodeId {
        let mut stmts = Vec::new();
        loop {
            while self.eat(TokenKind::Eol) {}
            if self.is_cc_terminator() { break; }
            match self.peek().kind.clone() {
                TokenKind::Eof => break,
                TokenKind::Kw(Kw::Sub)      => stmts.push(self.parse_proc_decl(arena, ProcKind::Sub)),
                TokenKind::Kw(Kw::Function) => stmts.push(self.parse_proc_decl(arena, ProcKind::Function)),
                TokenKind::Kw(Kw::Property) => stmts.push(self.parse_property_decl(arena)),
                TokenKind::Kw(Kw::Declare)  => stmts.push(self.parse_declare_stmt(arena)),
                TokenKind::Kw(Kw::Type)     => stmts.push(self.parse_type_decl(arena)),
                TokenKind::Kw(Kw::Enum)     => stmts.push(self.parse_enum_decl(arena)),
                TokenKind::Kw(Kw::Event)    => stmts.push(self.parse_event_decl(arena)),
                TokenKind::Kw(Kw::Implements) => stmts.push(self.parse_implements_stmt(arena)),
                TokenKind::Kw(Kw::Public)
                | TokenKind::Kw(Kw::Private)
                | TokenKind::Kw(Kw::Friend)
                | TokenKind::Kw(Kw::Static) => stmts.push(self.parse_access_decl(arena)),
                TokenKind::Kw(Kw::Option)   => stmts.push(self.parse_option_stmt(arena)),
                _ => { if let Some(n) = self.parse_stmt(arena) { stmts.push(n); } }
            }
        }
        arena.alloc(ExprNode::Block { stmts })
    }

    /// Handle the `#ElseIf` / `#Else` / `#End If` continuation of a CC chain.
    /// `result` is the node parsed for the branch that just ended.
    fn parse_cc_continuation(
        &mut self,
        arena: &mut ExprArena,
        include: bool,
        result: NodeId,
    ) -> NodeId {
        match self.peek().kind.clone() {
            TokenKind::Kw(Kw::CcElseIf) => {
                self.advance();
                let cond = self.eval_cc_expr(0);
                self.eat(TokenKind::Kw(Kw::Then));
                self.eat_eol();
                // Only the first matching branch is included
                let else_node = self.parse_cc_chain(arena, !include && cond != 0);
                if !include && cond != 0 { else_node } else { result }
            }
            TokenKind::Kw(Kw::CcElse) => {
                self.advance();
                self.eat_eol();
                let else_node = self.parse_cc_chain(arena, !include);
                if !include { else_node } else { result }
            }
            TokenKind::Kw(Kw::CcEnd) => {
                self.advance(); // #End
                self.eat(TokenKind::Kw(Kw::If));
                self.eat_eol();
                result
            }
            _ => result
        }
    }

    fn is_cc_terminator(&mut self) -> bool {
        matches!(self.peek().kind,
            TokenKind::Kw(Kw::CcElseIf) |
            TokenKind::Kw(Kw::CcElse) |
            TokenKind::Kw(Kw::CcEnd))
    }

    /// Skip statements in a false CC branch, tracking nested #If depth.
    fn skip_cc_false_branch(&mut self) {
        let mut cc_depth = 0i32;
        loop {
            while self.eat(TokenKind::Eol) {}
            match self.peek().kind.clone() {
                TokenKind::Eof => break,
                TokenKind::Kw(Kw::CcIf) => {
                    cc_depth += 1;
                    self.skip_to_stmt_end();
                    self.eat_eol();
                }
                TokenKind::Kw(Kw::CcEnd) if cc_depth > 0 => {
                    cc_depth -= 1;
                    self.skip_to_stmt_end();
                    self.eat_eol();
                }
                // Top-level CC terminator
                TokenKind::Kw(Kw::CcElseIf) |
                TokenKind::Kw(Kw::CcElse) |
                TokenKind::Kw(Kw::CcEnd) => break,
                _ => {
                    self.skip_to_stmt_end();
                    self.eat_eol();
                }
            }
        }
    }

    // ── CC expression evaluator ───────────────────────────────────────────────

    /// Look up a CC constant by (lowercase) name; returns 0 if not found.
    fn cc_lookup(&self, name: &str) -> i32 {
        self.cc_defs.iter().rev()
            .find(|(n, _)| n == name)
            .map(|(_, v)| *v)
            .unwrap_or(0)
    }

    /// Set (or insert) a CC constant.
    fn cc_set(&mut self, name: &str, value: i32) {
        if let Some(entry) = self.cc_defs.iter_mut().find(|(n, _)| n == name) {
            entry.1 = value;
        } else {
            self.cc_defs.push((name.to_string(), value));
        }
    }

    /// Evaluate a CC prefix (unary) operator, or fall through to a primary.
    fn eval_cc_prefix(&mut self) -> i32 {
        match self.peek().kind.clone() {
            TokenKind::Kw(Kw::Not) => {
                self.advance();
                let v = self.eval_cc_expr(20);
                if v != 0 { 0 } else { -1 }
            }
            TokenKind::Kw(Kw::Minus) => {
                self.advance();
                let v = self.eval_cc_expr(20);
                v.wrapping_neg()
            }
            TokenKind::Kw(Kw::Plus) => {
                self.advance();
                self.eval_cc_expr(20)
            }
            _ => self.eval_cc_primary(),
        }
    }

    /// Binding powers + operator tag for a CC infix operator, or `None`.
    fn cc_infix_op(&mut self) -> Option<(u8, u8, u8)> {
        let entry = match self.peek().kind {
            TokenKind::Kw(Kw::Star)       => (18, 19, b'*'),
            TokenKind::Kw(Kw::Slash)      => (18, 19, b'/'),
            TokenKind::Kw(Kw::Backslash)  => (16, 17, b'\\'),
            TokenKind::Kw(Kw::Mod)        => (14, 15, b'm'),
            TokenKind::Kw(Kw::Plus)       => (12, 13, b'+'),
            TokenKind::Kw(Kw::Minus)      => (12, 13, b'-'),
            TokenKind::Kw(Kw::Eq)         => (8,  9,  b'='),
            TokenKind::Kw(Kw::Ne)         => (8,  9,  b'n'),
            TokenKind::Kw(Kw::Lt)         => (8,  9,  b'<'),
            TokenKind::Kw(Kw::Gt)         => (8,  9,  b'>'),
            TokenKind::Kw(Kw::Le)         => (8,  9,  b'l'),
            TokenKind::Kw(Kw::Ge)         => (8,  9,  b'g'),
            TokenKind::Kw(Kw::And)        => (6,  7,  b'&'),
            TokenKind::Kw(Kw::Or)         => (4,  5,  b'|'),
            TokenKind::Kw(Kw::Xor)        => (4,  5,  b'^'),
            TokenKind::Kw(Kw::Eqv)        => (2,  3,  b'e'),
            TokenKind::Kw(Kw::Imp)        => (2,  3,  b'i'),
            _ => return None,
        };
        Some(entry)
    }

    /// Evaluate a CC expression to an i32 (True=-1, False=0).
    fn eval_cc_expr(&mut self, min_bp: u8) -> i32 {
        let mut lhs = self.eval_cc_prefix();
        loop {
            let Some((l_bp, r_bp, op)) = self.cc_infix_op() else { break };
            if l_bp <= min_bp { break; }
            self.advance();
            let rhs = self.eval_cc_expr(r_bp);
            lhs = apply_cc_binop(lhs, rhs, op);
        }
        lhs
    }

    fn eval_cc_primary(&mut self) -> i32 {
        match self.peek().kind.clone() {
            TokenKind::Kw(Kw::LParen) => {
                self.advance();
                let v = self.eval_cc_expr(0);
                self.eat(TokenKind::Kw(Kw::RParen));
                v
            }
            TokenKind::IntLit => {
                let tok = self.advance();
                if let Some(crate::frontend::token::Lit::Int(n)) = tok.lit { n } else { 0 }
            }
            TokenKind::LongLit => {
                let tok = self.advance();
                if let Some(crate::frontend::token::Lit::Long(n)) = tok.lit { n } else { 0 }
            }
            TokenKind::Kw(Kw::True)  => { self.advance(); -1 }
            TokenKind::Kw(Kw::False) => { self.advance();  0 }
            _ => {
                let tok = self.advance();
                if tok.kind == TokenKind::Ident || matches!(tok.kind, TokenKind::Kw(_)) {
                    if let Some(id) = tok.sym {
                        let name = self.scanner.sym_name(id as u32).to_ascii_lowercase();
                        return self.cc_lookup(&name);
                    }
                }
                0
            }
        }
    }

    // ── Option statement parser ───────────────────────────────────────────────

    fn parse_option_stmt(&mut self, arena: &mut ExprArena) -> NodeId {
        self.advance(); // consume Option
        match self.peek().kind.clone() {
            TokenKind::Kw(Kw::Explicit) => {
                self.advance();
                self.eat_eol();
                arena.alloc(ExprNode::OptionExplicit)
            }
            TokenKind::Kw(Kw::Base) => {
                self.advance();
                let value = match self.peek().kind {
                    TokenKind::IntLit => {
                        let tok = self.advance();
                        if let Some(crate::frontend::token::Lit::Int(n)) = tok.lit {
                            n as u8
                        } else { 0 }
                    }
                    _ => { self.advance(); 0 }
                };
                self.eat_eol();
                arena.alloc(ExprNode::OptionBase { value })
            }
            TokenKind::Kw(Kw::Compare) => {
                self.advance();
                let mode = match self.peek().kind.clone() {
                    TokenKind::Kw(Kw::Binary) => { self.advance(); 0u8 }
                    TokenKind::Kw(Kw::Text)   => { self.advance(); 1u8 }
                    TokenKind::Kw(Kw::Database) => { self.advance(); 2u8 }
                    TokenKind::Ident => {
                        self.advance();
                        2u8
                    }
                    _ => 0,
                };
                self.eat_eol();
                arena.alloc(ExprNode::OptionCompare { mode })
            }
            _ => {
                self.skip_to_stmt_end();
                self.eat_eol();
                arena.alloc(ExprNode::Block { stmts: vec![] })
            }
        }
    }

    // ── Implements ───────────────────────────────────────────────────────────

    fn parse_implements_stmt(&mut self, arena: &mut ExprArena) -> NodeId {
        self.advance(); // consume Implements
        let name = self.expect_ident();
        let nid = name.sym.map(|s| s as u32).unwrap_or(0);
        self.eat_eol();
        arena.alloc(ExprNode::Implements { name: nid })
    }

    /// `Def<Type> letter[-letter][, letter[-letter]]…`
    fn parse_def_stmt(&mut self, arena: &mut ExprArena) -> NodeId {
        let type_kw = match self.peek().kind {
            TokenKind::Kw(k) => k as u16,
            _ => 0,
        };
        self.advance(); // consume the Def<Type> keyword
        let mut ranges = Vec::new();
        loop {
            let lo = self.expect_ident().sym.map(|s| s as u32).unwrap_or(0);
            let hi = if self.eat(TokenKind::Kw(Kw::Minus)) {
                self.expect_ident().sym.map(|s| s as u32).unwrap_or(0)
            } else {
                0
            };
            ranges.push((lo, hi));
            if !self.eat(TokenKind::Kw(Kw::Comma)) {
                break;
            }
        }
        self.eat_eol();
        arena.alloc(ExprNode::DefType { type_kw, ranges })
    }

    /// `Attribute name = value[, value…]` metadata line.
    fn parse_attribute_stmt(&mut self, arena: &mut ExprArena) -> NodeId {
        self.advance(); // consume Attribute
        let key = self.expect_ident();
        let name = key.sym.map(|s| s as u32).unwrap_or(0);
        while self.eat(TokenKind::Kw(Kw::Dot)) {
            let _ = self.expect_ident();
        }
        self.expect(TokenKind::Kw(Kw::Eq), ERR_EXPECTED_EQ);
        let mut values = Vec::new();
        loop {
            values.push(self.parse_expr(arena, 0));
            if !self.eat(TokenKind::Kw(Kw::Comma)) {
                break;
            }
        }
        self.eat_eol();
        arena.alloc(ExprNode::Attribute { name, values })
    }

    // ── File I/O statement parsers ──────────────────────────────────────────

    /// `Open path For mode [Access access] [Lock lock] As [#]filenum [Len=n]`
    fn parse_file_open(&mut self, arena: &mut ExprArena) -> NodeId {
        self.advance(); // Open
        let path = self.parse_expr(arena, 0);
        loop {
            match self.peek().kind {
                TokenKind::Kw(Kw::As) | TokenKind::Eol | TokenKind::Eof => break,
                _ => { self.advance(); }
            }
        }
        self.eat(TokenKind::Kw(Kw::As));
        self.eat(TokenKind::Kw(Kw::Hash));
        let ch = self.parse_expr(arena, 0);
        let mut args = vec![path];
        if self.eat(TokenKind::Kw(Kw::Len)) {
            self.eat(TokenKind::Kw(Kw::Eq));
            args.push(self.parse_expr(arena, 0)); // `Len = <expr>` is a real use site
        }
        self.eat_eol();
        arena.alloc(ExprNode::FileIoStmt { kind: FileIoKind::Open, channel: Some(ch), args })
    }

    /// `Close [[#]filenum, ...]`
    fn parse_file_close(&mut self, arena: &mut ExprArena) -> NodeId {
        self.advance(); // Close
        let mut chs = Vec::new();
        while !self.peek().is_stmt_end() {
            self.eat(TokenKind::Kw(Kw::Hash));
            chs.push(self.parse_expr(arena, 0));
            if !self.eat(TokenKind::Kw(Kw::Comma)) { break; }
        }
        self.eat_eol();
        arena.alloc(ExprNode::FileIoStmt { kind: FileIoKind::Close, channel: None, args: chs })
    }

    /// `Print #filenum, [args]` or `Write #filenum, [args]`.
    fn parse_file_print_after_keyword(&mut self, arena: &mut ExprArena, kind: FileIoKind) -> NodeId {
        self.eat(TokenKind::Kw(Kw::Hash));
        let ch = self.parse_expr(arena, 0);
        let mut args = Vec::new();
        if self.eat(TokenKind::Kw(Kw::Comma)) {
            while !self.peek().is_stmt_end() {
                if let Some(n) = self.try_parse_expr(arena) { args.push(n); }
                if !self.eat(TokenKind::Kw(Kw::Comma)) && !self.eat(TokenKind::Kw(Kw::Semi)) { break; }
            }
        }
        self.eat_eol();
        arena.alloc(ExprNode::FileIoStmt { kind, channel: Some(ch), args })
    }

    /// `Input #filenum, varlist`.
    fn parse_file_input_after_keyword(&mut self, arena: &mut ExprArena) -> NodeId {
        self.eat(TokenKind::Kw(Kw::Hash));
        let ch = self.parse_expr(arena, 0);
        let mut vars = Vec::new();
        if self.eat(TokenKind::Kw(Kw::Comma)) {
            loop {
                vars.push(self.parse_expr(arena, 0));
                if !self.eat(TokenKind::Kw(Kw::Comma)) { break; }
            }
        }
        self.eat_eol();
        arena.alloc(ExprNode::FileIoStmt { kind: FileIoKind::Input, channel: Some(ch), args: vars })
    }

    /// `Get #filenum, [recnum], var` or `Put #filenum, [recnum], val`.
    fn parse_file_get_put_after_keyword(&mut self, arena: &mut ExprArena, kind: FileIoKind) -> NodeId {
        self.eat(TokenKind::Kw(Kw::Hash));
        let ch = self.parse_expr(arena, 0);
        let mut args = Vec::new();
        if self.eat(TokenKind::Kw(Kw::Comma)) {
            // Optional record number.
            if self.peek().kind != TokenKind::Kw(Kw::Comma) {
                args.push(self.parse_expr(arena, 0));
            }
            if self.eat(TokenKind::Kw(Kw::Comma)) {
                args.push(self.parse_expr(arena, 0));
            }
        }
        self.eat_eol();
        arena.alloc(ExprNode::FileIoStmt { kind, channel: Some(ch), args })
    }

    /// `Lock #filenum [, from [To to]]` or `Unlock ...`.
    fn parse_file_lock_after_keyword(&mut self, arena: &mut ExprArena, kind: FileIoKind) -> NodeId {
        self.eat(TokenKind::Kw(Kw::Hash));
        let ch = self.parse_expr(arena, 0);
        let mut args = Vec::new();
        if self.eat(TokenKind::Kw(Kw::Comma)) {
            args.push(self.parse_expr(arena, 0));
            if self.eat(TokenKind::Kw(Kw::To)) {
                args.push(self.parse_expr(arena, 0));
            }
        }
        self.eat_eol();
        arena.alloc(ExprNode::FileIoStmt { kind, channel: Some(ch), args })
    }

    /// Parse ident-statement starting from a token already consumed from the stream.
    fn parse_ident_stmt_with_tok(&mut self, arena: &mut ExprArena, first: Token) -> NodeId {
        let sym = first.sym.map(|s| s as u32).unwrap_or(0);
        let lhs_base = arena.alloc(ExprNode::NameRef { sym, suffix: first.type_suffix });
        self.set_span(lhs_base, first.span);
        // Apply any postfix operators (.member, (args), etc.)
        let lhs = self.parse_postfix(arena, lhs_base);
        if self.eat(TokenKind::Kw(Kw::Eq)) {
            let value = self.parse_expr(arena, 0);
            self.eat_eol();
            arena.alloc(ExprNode::Assign { target: lhs, value })
        } else if self.eat(TokenKind::Kw(Kw::ColonEq)) {
            let value = self.parse_expr(arena, 0);
            self.eat_eol();
            arena.alloc(ExprNode::Assign { target: lhs, value })
        } else if self.peek().kind == TokenKind::Kw(Kw::Minus)
            && (self.peek2().kind == TokenKind::Kw(Kw::LParen)
                || self.peek2().kind == TokenKind::Kw(Kw::Step))
        {
            // `Obj.Line (x1,y1)-(x2,y2)` — postfix consumed `(x1,y1)` as a call
            // arg-list; the `-(x2,y2)` continuation signals this is a graphics call.
            self.advance(); // consume `-`
            let (x2, y2) = self.parse_coord_pair(arena);
            let mut extra = vec![x2, y2];
            if self.eat(TokenKind::Kw(Kw::Comma)) {
                if let Some(color) = self.try_parse_expr(arena) {
                    extra.push(color);
                }
                while self.eat(TokenKind::Kw(Kw::Comma)) {
                    if let Some(flag) = self.try_parse_expr(arena) {
                        extra.push(flag);
                    }
                }
            }
            self.eat_eol();
            let args = arena.alloc(ExprNode::ArgList { args: extra });
            arena.alloc(ExprNode::CallStmt { callee: lhs, args })
        } else {
            // Implicit call with optional trailing arguments: `Foo a, b`
            let args = self.parse_stmt_arg_list(arena);
            self.eat_eol();
            if args.is_empty() {
                lhs
            } else {
                let args = arena.alloc(ExprNode::ArgList { args });
                arena.alloc(ExprNode::CallStmt { callee: lhs, args })
            }
        }
    }

    // ── Procedure / declaration parsers ───────────────────────────────────────

    /// Parse an access-modifier prefix (Public/Private/Friend/Static) then dispatch.
    fn parse_access_decl(&mut self, arena: &mut ExprArena) -> NodeId {
        let access_tok = self.advance().kind.clone();
        let vis = match access_tok {
            TokenKind::Kw(Kw::Public) | TokenKind::Kw(Kw::Global) | TokenKind::Kw(Kw::Friend) => {
                Some(true)
            }
            TokenKind::Kw(Kw::Private) => Some(false),
            _ => None,
        };
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
            // Public/Private variable declaration at module level
            TokenKind::Ident | TokenKind::Kw(Kw::Dim) => {
                if self.peek().kind == TokenKind::Kw(Kw::Dim) { self.advance(); }
                self.parse_dim_list(arena, false)
            }
            // WithEvents — class-module-only syntax
            TokenKind::Kw(Kw::WithEvents) => {
                let span = self.current_span();
                if self.module_kind == ModuleKind::Standard {
                    self.diagnostics.push(ERR_OBJECT_MODULE_ONLY, span);
                }
                self.advance();
                self.parse_dim_list(arena, false)
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
        let name = self.expect_ident();
        let name_id = name.sym.map(|s| s as u32).unwrap_or(0);
        // Parameter list
        let params = if self.eat(TokenKind::Kw(Kw::LParen)) {
            let p = self.parse_param_list(arena);
            self.expect(TokenKind::Kw(Kw::RParen), ERR_EXPECTED_RPAREN);
            Some(p)
        } else { None };
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
        self.eat_eol();
        let body = self.parse_block_proc(arena);
        if self.peek().kind != TokenKind::Eof {
            self.expect(TokenKind::Kw(Kw::End), ERR_EXPECTED_END);
            if matches!(self.peek().kind, TokenKind::Kw(Kw::Sub) | TokenKind::Kw(Kw::Function)) {
                self.advance();
            } else {
                let span = self.current_span();
                self.diagnostics.push(ERR_UNEXPECTED_TOKEN, span);
            }
            self.eat_eol();
        }
        let id = arena.alloc(ExprNode::ProcDecl { kind, name: name_id, params, ret_type, body });
        self.set_span(id, name.span);
        id
    }

    /// Parses Property Get/Let/Set.
    fn parse_property_decl(&mut self, arena: &mut ExprArena) -> NodeId {
        self.advance(); // consume Property
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
        let name = self.expect_ident();
        let name_id = name.sym.map(|s| s as u32).unwrap_or(0);
        let params = if self.eat(TokenKind::Kw(Kw::LParen)) {
            let p = self.parse_param_list(arena);
            self.expect(TokenKind::Kw(Kw::RParen), ERR_EXPECTED_RPAREN);
            Some(p)
        } else { None };
        let ret_type = if kind == ProcKind::PropGet && self.eat(TokenKind::Kw(Kw::As)) {
            let t = self.parse_type_spec(arena, true);
            if name.type_suffix != TypeSuffix::None {
                self.diagnostics.push(ERR_UNEXPECTED_TOKEN, name.span);
            }
            Some(t)
        } else if name.type_suffix != TypeSuffix::None {
            if kind == ProcKind::PropGet {
                self.type_node_from_suffix(arena, name.type_suffix)
            } else {
                self.diagnostics.push(ERR_UNEXPECTED_TOKEN, name.span);
                None
            }
        } else { None };
        self.eat_eol();
        let body = self.parse_block_proc(arena);
        self.expect(TokenKind::Kw(Kw::End), ERR_EXPECTED_END);
        self.expect(TokenKind::Kw(Kw::Property), ERR_UNEXPECTED_TOKEN);
        self.eat_eol();
        let id = arena.alloc(ExprNode::ProcDecl { kind, name: name_id, params, ret_type, body });
        self.set_span(id, name.span);
        id
    }

    /// Parses an Event declaration.
    fn parse_event_decl(&mut self, arena: &mut ExprArena) -> NodeId {
        self.advance(); // consume Event
        let name = self.expect_ident();
        let nid = name.sym.map(|s| s as u32).unwrap_or(0);
        let params = if self.eat(TokenKind::Kw(Kw::LParen)) {
            let p = self.parse_param_list(arena);
            self.expect(TokenKind::Kw(Kw::RParen), ERR_EXPECTED_RPAREN);
            p
        } else {
            arena.alloc(ExprNode::ArgList { args: vec![] })
        };
        self.eat_eol();
        let id = arena.alloc(ExprNode::EventDecl { name: nid, params });
        self.set_span(id, name.span);
        id
    }

    /// Parse `Declare [Function|Sub] name Lib "lib" [Alias "alias"] ([params]) [As type]`
    fn parse_declare_stmt(&mut self, arena: &mut ExprArena) -> NodeId {
        self.advance(); // Declare
        let kind = match self.peek().kind.clone() {
            TokenKind::Kw(Kw::Function) => { self.advance(); ProcKind::Function }
            TokenKind::Kw(Kw::Sub)      => { self.advance(); ProcKind::Sub }
            _ => {
                let span = self.current_span();
                self.diagnostics.push(ERR_UNEXPECTED_TOKEN, span);
                self.skip_to_stmt_end();
                self.eat_eol();
                return arena.alloc(ExprNode::Block { stmts: vec![] });
            }
        };
        let name_tok = self.expect_ident();
        let name = name_tok.sym.map(|s| s as u32).unwrap_or(0);
        // Lib "libname"
        self.skip_line_conts();
        let lib = if self.eat(TokenKind::Kw(Kw::Lib)) {
            let s = self.parse_expr(arena, 0);
            s
        } else {
            let span = self.current_span();
            self.diagnostics.push(ERR_UNEXPECTED_TOKEN, span);
            arena.alloc(ExprNode::Literal { lit: crate::frontend::ast::AstLit::Str("".into()) })
        };
        // Optional Alias "aliasname"
        self.skip_line_conts();
        let alias = if self.eat(TokenKind::Kw(Kw::Alias)) {
            Some(self.parse_expr(arena, 0))
        } else {
            None
        };
        // Optional parameter list
        self.skip_line_conts();
        let params = if self.eat(TokenKind::Kw(Kw::LParen)) {
            let p = self.parse_param_list(arena);
            self.expect(TokenKind::Kw(Kw::RParen), ERR_EXPECTED_RPAREN);
            Some(p)
        } else {
            None
        };
        // Optional return type.
        self.skip_line_conts();
        let ret_type = if kind == ProcKind::Function && self.eat(TokenKind::Kw(Kw::As)) {
            Some(self.parse_type_spec(arena, false))
        } else {
            None
        };
        self.eat_eol();
        let id = arena.alloc(ExprNode::DeclareDecl { kind, name, lib, alias, params, ret_type });
        self.set_span(id, name_tok.span);
        id
    }

    /// Parses a procedure parameter list.
    fn parse_param_list(&mut self, arena: &mut ExprArena) -> NodeId {
        self.skip_line_conts();
        if self.peek().kind == TokenKind::Kw(Kw::RParen) {
            return arena.alloc(ExprNode::ArgList { args: vec![] });
        }
        let mut args = Vec::new();
        loop {
            args.push(self.parse_param(arena));
            if !self.eat(TokenKind::Kw(Kw::Comma)) { break; }
            self.skip_line_conts();
            if self.peek().kind == TokenKind::Kw(Kw::RParen) { break; }
        }
        arena.alloc(ExprNode::ArgList { args })
    }

    fn parse_param(&mut self, arena: &mut ExprArena) -> NodeId {
        self.skip_line_conts();
        let mut flags: u16 = 0;
        if self.eat(TokenKind::Kw(Kw::Optional)) { flags |= 0x01; }
        // ParamArray must precede the parameter name; mutually exclusive with ByVal/ByRef.
        if self.eat(TokenKind::Kw(Kw::ParamArray)) {
            flags |= 0x20;
        } else {
            match self.peek().kind {
                TokenKind::Kw(Kw::ByVal) => { self.advance(); flags |= 0x02; }
                TokenKind::Kw(Kw::ByRef) => { self.advance(); flags |= 0x04; }
                _ => {}
            }
        }
        let name = self.expect_ident();
        let nid = name.sym.map(|s| s as u32).unwrap_or(0);
        if self.eat(TokenKind::Kw(Kw::LParen)) {
            self.expect(TokenKind::Kw(Kw::RParen), ERR_EXPECTED_RPAREN);
            flags |= 0x08; // array parameter
        }
        let type_node = if self.eat(TokenKind::Kw(Kw::As)) {
            if self.eat(TokenKind::Kw(Kw::New)) { flags |= 0x10; }
            Some(self.parse_type_spec(arena, false))
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

    /// Parse a type specifier.  `allow_array_suffix` controls whether a trailing
    /// `()` array marker is consumed.
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
            TokenKind::Kw(Kw::Decimal)  => { self.advance(); arena.alloc(ExprNode::BuiltinType { kind: 14 }) }
            TokenKind::Kw(Kw::String) | TokenKind::Kw(Kw::StringS) => {
                self.advance();
                let fixed_len = if self.eat(TokenKind::Kw(Kw::Star)) {
                    Some(self.parse_expr(arena, 0))
                } else { None };
                arena.alloc(ExprNode::StringType { fixed_len })
            }
            TokenKind::Kw(Kw::New) => {
                self.advance();
                let name = self.expect_ident();
                let nid = name.sym.map(|s| s as u32).unwrap_or(0);
                arena.alloc(ExprNode::UserType { name: nid, child: None })
            }
            TokenKind::Ident => {
                let name = self.advance();
                let nid = name.sym.map(|s| s as u32).unwrap_or(0);
                let child = if self.eat(TokenKind::Kw(Kw::Dot)) {
                    let sub = self.expect_ident();
                    let sid = sub.sym.map(|s| s as u32).unwrap_or(0);
                    let cref = arena.alloc(ExprNode::NameRef { sym: sid, suffix: TypeSuffix::None });
                    self.set_span(cref, sub.span);
                    Some(cref)
                } else { None };
                arena.alloc(ExprNode::UserType { name: nid, child })
            }
            _ => {
                let span = self.current_span();
                self.diagnostics.push(ERR_UNEXPECTED_TOKEN, span);
                arena.alloc(ExprNode::BuiltinType { kind: 0 })
            }
        };
        // Consume the optional array marker `()`.
        if allow_array_suffix && self.peek().kind == TokenKind::Kw(Kw::LParen) {
            self.advance();
            self.eat(TokenKind::Kw(Kw::RParen));
        }
        node
    }

    /// Parse a Type declaration block.
    fn parse_type_decl(&mut self, arena: &mut ExprArena) -> NodeId {
        self.advance(); // consume Type
        let name = self.expect_ident();
        let nid = name.sym.map(|s| s as u32).unwrap_or(0);
        self.eat_eol();
        let mut members = Vec::new();
        loop {
            while self.eat(TokenKind::Eol) {}
            match self.peek().kind {
                TokenKind::Kw(Kw::End) => {
                    self.advance();
                    self.expect(TokenKind::Kw(Kw::Type), ERR_UNEXPECTED_TOKEN);
                    self.eat_eol();
                    break;
                }
                TokenKind::Eof => break,
                _ => {
                    let m = self.parse_dim_item(arena, false, false);
                    self.eat_eol();
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
        let name = self.expect_ident();
        let nid = name.sym.map(|s| s as u32).unwrap_or(0);
        self.eat_eol();
        let mut members = Vec::new();
        loop {
            while self.eat(TokenKind::Eol) {}
            match self.peek().kind {
                TokenKind::Kw(Kw::End) => {
                    self.advance();
                    self.expect(TokenKind::Kw(Kw::Enum), ERR_UNEXPECTED_TOKEN);
                    self.eat_eol();
                    break;
                }
                TokenKind::Eof => break,
                TokenKind::Ident => {
                    let m = self.parse_dim_item(arena, true, false);
                    self.eat_eol();
                    members.push(m);
                }
                _ => { self.advance(); }
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
        let mut lhs = self.parse_expr_prefix(arena);

        // Infix loop
        loop {
            // A line-continuation sequence (`_ \n`) may appear before an infix
            // operator or before the RHS operand — skip it transparently.
            self.skip_line_conts();

            let kind = self.peek().kind.clone();
            // Check for Is / Is Not (two-token operator) separately from the bp table.
            if kind == TokenKind::Kw(Kw::Is) {
                match self.parse_is_operator(arena, min_bp, lhs) {
                    Some(node) => { lhs = node; continue; }
                    None => break,
                }
            }

            let Some((l_bp, r_bp, bin_op)) = infix_bp(&kind) else { break };
            if l_bp <= min_bp { break; }
            self.advance(); // consume operator
            self.skip_line_conts(); // continuation may follow the operator itself
            let rhs = self.parse_expr(arena, r_bp);
            lhs = arena.alloc(ExprNode::BinOp { op: bin_op, lhs, rhs });
        }
        lhs
    }

    /// Parse the prefix/primary portion of an expression: a unary operator,
    /// `TypeOf`, `New`, or a plain primary.
    fn parse_expr_prefix(&mut self, arena: &mut ExprArena) -> NodeId {
        self.skip_line_conts(); // continuation may precede the first operand
        if let Some((bp, uop)) = prefix_bp(&self.peek().kind.clone()) {
            self.advance();
            let operand = self.parse_expr(arena, bp);
            arena.alloc(ExprNode::UnOp { op: uop, operand })
        } else if self.peek().kind == TokenKind::Kw(Kw::TypeOf) {
            self.parse_typeof_expr(arena)
        } else if self.peek().kind == TokenKind::Kw(Kw::New) {
            self.parse_new_expr(arena)
        } else {
            self.parse_primary(arena)
        }
    }

    /// Handle an `Is` / `Is Not` operator at the current position. Returns
    /// `Some(new_lhs)` to continue the infix loop, or `None` to break it.
    fn parse_is_operator(
        &mut self,
        arena: &mut ExprArena,
        min_bp: u8,
        lhs: NodeId,
    ) -> Option<NodeId> {
        let is_tok = self.advance();
        if self.eat(TokenKind::Kw(Kw::Not)) {
            let rhs = self.parse_expr(arena, 9);
            return Some(arena.alloc(ExprNode::BinOp { op: BinOpKind::IsNot, lhs, rhs }));
        }
        if min_bp < 8 {
            let rhs = self.parse_expr(arena, 9);
            return Some(arena.alloc(ExprNode::BinOp { op: BinOpKind::Is, lhs, rhs }));
        }
        self.diagnostics.push(ERR_UNEXPECTED_TOKEN, is_tok.span);
        None
    }

    fn try_parse_expr(&mut self, arena: &mut ExprArena) -> Option<NodeId> {
        if self.peek().is_stmt_end() { return None; }
        match self.peek().kind {
            TokenKind::Kw(Kw::RParen)
            | TokenKind::Kw(Kw::Comma)
            | TokenKind::Kw(Kw::Colon) => None,
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
            TokenKind::Kw(Kw::Me) => {
                self.advance();
                arena.alloc(ExprNode::Me)
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
                self.advance();
                let operand = self.parse_primary(arena);
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
            // Keyword that can start a primary (e.g. built-in functions)
            TokenKind::Kw(_) => {
                let tok = self.advance();
                let sym_id = tok.sym.map(|s| s as u32).unwrap_or(0);
                let base = arena.alloc(ExprNode::NameRef { sym: sym_id, suffix: tok.type_suffix });
                self.set_span(base, tok.span);
                self.parse_postfix(arena, base)
            }
            // Unexpected — emit error, return placeholder
            _ => {
                let span = self.current_span();
                self.diagnostics.push(ERR_EXPECTED_EXPR, span);
                arena.alloc(ExprNode::Literal { lit: AstLit::Int(0) })
            }
        }
    }

    /// Postfix: member access and call/index.
    fn parse_postfix(&mut self, arena: &mut ExprArena, mut base: NodeId) -> NodeId {
        loop {
            match self.peek().kind.clone() {
                TokenKind::Kw(Kw::Dot) => {
                    self.advance();
                    let member = self.expect_ident();
                    let mid = member.sym.map(|s| s as u32).unwrap_or(0);
                    base = arena.alloc(ExprNode::MemberAccess { base, member: mid, bang: false });
                }
                TokenKind::Kw(Kw::Bang) => {
                    self.advance();
                    let member = self.expect_ident();
                    let mid = member.sym.map(|s| s as u32).unwrap_or(0);
                    base = arena.alloc(ExprNode::MemberAccess { base, member: mid, bang: true });
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

    /// Parses the LHS of an assignment.
    fn parse_target_expr(&mut self, arena: &mut ExprArena) -> NodeId {
        // A leading `.`/`!` is a With-block member reference: `.Member` / `!Key`.
        // The leading dot/bang is semantically significant (it binds the member
        // to the active With object), so build a `MemberAccess` on an implicit
        // `WithContext` base — never a bare `NameRef`, which would drop it.
        if let Some(access) = self.try_parse_leading_with_member(arena) {
            return self.parse_postfix(arena, access);
        }
        self.parse_primary(arena)
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
                break;
            } else if let Some(named) = self.try_parse_named_arg(arena) {
                args.push(named);
            } else {
                // Skip ByVal/ByRef at call-site qualifiers.
                let _ = self.eat(TokenKind::Kw(Kw::ByVal))
                    || self.eat(TokenKind::Kw(Kw::ByRef));
                args.push(self.parse_expr(arena, 0));
            }
            if !self.eat(TokenKind::Kw(Kw::Comma)) { break; }
        }
        arena.alloc(ExprNode::ArgList { args })
    }

    /// Argument list on a call statement line (no surrounding parens).
    fn parse_stmt_arg_list(&mut self, arena: &mut ExprArena) -> Vec<NodeId> {
        let mut args = Vec::new();
        while !self.peek().is_stmt_end() {
            let a = if let Some(named) = self.try_parse_named_arg(arena) {
                named
            } else if let Some(expr) = self.try_parse_expr(arena) {
                expr
            } else {
                break;
            };
            args.push(a);
            if !self.eat(TokenKind::Kw(Kw::Comma)) { break; }
        }
        args
    }

    /// If the next tokens are `ident :=`, consume them and return a `NamedArg`
    /// node. The label names the *callee's* parameter, not a caller-scope
    /// variable, so it is recorded as a symbol (with its span, for future
    /// goto-parameter) — never as a `NameRef` use that would otherwise resolve in
    /// the caller's scope or be flagged undeclared.
    fn try_parse_named_arg(&mut self, arena: &mut ExprArena) -> Option<NodeId> {
        if self.peek().kind != TokenKind::Ident
            || self.peek2().kind != TokenKind::Kw(Kw::ColonEq)
        {
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
const ERR_EXPECTED_EXPR:        u32 = 0x9c6f;
const ERR_EXPECTED_IDENT:       u32 = 0x9c6f;
const ERR_OBJECT_MODULE_ONLY:   u32 = 0xdee1;
const ERR_UNEXPECTED_TOKEN:     u32 = 0x9c6f;
const ERR_EXPECTED_EQ:          u32 = 0x9c6f;
const ERR_EXPECTED_RPAREN:      u32 = 0x9c6f;
const ERR_EXPECTED_LPAREN:      u32 = 0x9c6f;
const ERR_EXPECTED_THEN:        u32 = 0x9c6f;
const ERR_EXPECTED_TO:          u32 = 0x9c6f;
const ERR_EXPECTED_NEXT:        u32 = 0x9c6f;
const ERR_EXPECTED_LOOP:        u32 = 0x9c6f;
const ERR_EXPECTED_WEND:        u32 = 0x9c6f;
const ERR_EXPECTED_END:         u32 = 0x9c6f;
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
