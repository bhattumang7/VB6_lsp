//! VB6 recursive-descent parser.
//!
//! Implements the source-to-AST stage for Visual Basic 6.
//!
//! **Representation choices:**
//! * Global error-flag + continue → `Diagnostics` collector.
//! * Process-global scanner state → owned `Scanner` threaded via `Parser`.
//! * List chain nodes → `Vec<NodeId>` in typed `ArgList`/`Block` variants.

mod attribute;

use crate::frontend::{declarations, declared_names};
use crate::frontend::line_size::{self, cost, MAX_LINE_CODE_SIZE};
use crate::frontend::cond_compile::{CcOp, CcResult, CcValue, ERR_VARIABLE_NOT_DEFINED as CC_ERR_VARIABLE_NOT_DEFINED};

/// Binding power of the operand of `Not` and of a sign in a directive expression.
const CC_NOT_BP: u8 = 10;
const CC_SIGN_BP: u8 = 26;
use crate::frontend::ast::{
    AstLit, BinOpKind, DoKind, ExitKind, ExprArena, ExprNode, FileIoKind, LabelRef, NodeId,
    NodeSpans, OnErrorKind, ProcKind, Radix, ResumeTarget, UnOpKind,
};
use crate::frontend::diagnostics::*;
use crate::frontend::ordinal::converts_to_integer;
use crate::frontend::scanner::{fold_name, OrdinalStart, Scanner, ScannerContext};
use crate::frontend::token::{Kw, Lit, NumberFailure, Span, Token, TokenKind, TypeSuffix};

mod graphics;
mod param_rules;
use graphics::{ChainStart, GraphicsWord};

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
        // Not, level 5 like And: it is also a prefix operator, but in the operator table it has a
        // precedence, so between two operands it is taken as infix.
        TokenKind::Kw(Kw::Not)       => (10, 11, B::Not),
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
        DefBool | DefByte | DefCur | DefDate | DefDbl
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
    /// Tokens read ahead of the grammar, handed out before the scanner is asked again.
    pushback:    std::collections::VecDeque<Token>,
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
    /// Literal nodes (`NodeId.0`) written as octal or hexadecimal whole numbers
    /// (`&O17`, `&17`, `&H1F`). A decimal literal is absent.
    pub literal_radix: std::collections::HashMap<u32, Radix>,
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
    /// The statement before an `Attribute` line on the same line is being read: its errors are
    /// kept and the statement is dropped.
    in_attribute_head: bool,
    /// Single-line `If` statements whose statement part is being read and which will take an
    /// `Else` that follows a statement directly: only then does `Else` end a statement.
    inline_else_pending: u32,
    /// A line after the end of a procedure is being read as a statement: block statements
    /// are read as their header line alone and block terminators stand for themselves.
    line_scan: bool,
    /// Where the first line end read since `eol_watch` was set starts.
    first_eol_start: u32,
    /// The start of the next line end read is to be remembered in `first_eol_start`.
    eol_watch: bool,
    /// A statement's target is being parsed, where built-in functions are plain names.
    in_target: bool,
    /// The name being read may be written in brackets: the target of a jump, a name that
    /// is used rather than declared, a member of a type.
    bracketed_name_ok: bool,
    /// The type being read follows `New` in a declaration.
    type_after_new: bool,
    /// An enumeration member is being parsed.
    in_enum_member: bool,
    /// The declarator being read follows a comma in a declaration list: only the first
    /// declarator of a list refuses a name written in brackets.
    later_declarator: bool,
    /// The target about to be parsed is at the start of a statement, where a few reserved
    /// words (`Date`, `DoEvents`, `CDec`) may name what is assigned or called.
    target_at_statement_start: bool,
    /// Open block statements of the procedure being parsed, innermost last.
    blocks: Vec<Block>,
    /// The last token consumed was the end of a line.
    last_was_eol: bool,
    /// A `:` statement separator has been consumed on the current line.
    after_colon: bool,
    /// A token other than a line end has been read on the current logical line: a
    /// `Name:` is a line label only as the first thing on its line.
    line_has_content: bool,
    /// A token other than a line end has been consumed on the current logical line. Unlike
    /// `line_has_content` nothing resets it but the line's end: a line number is read only
    /// where nothing precedes it on its line.
    line_started: bool,
    /// Kind of the token consumed last.
    last_consumed: TokenKind,
    /// Line ends read so far: a block opener inside a single-line `If` must be closed
    /// before the line ends.
    eols_read: u32,
    /// A statement-separating colon has been read on the current logical line (the colon
    /// of a line label or of a line number does not count): a `Rem` then ends a statement.
    colon_on_line: bool,
    /// A module-level variable, procedure or `Declare` has been parsed.
    module_decl_seen: bool,
    /// True while the header of a procedure written with a leading `Static` is read.
    static_prefix: bool,
    /// The property procedures read so far, by property name.
    property_sigs: Vec<(u32, PropertySig)>,
    /// True once a procedure has been read: no declaration may follow it.
    procedure_seen: bool,
    /// True while the parameter list of a `Property Let` or `Property Set` is being read.
    value_owner_params: bool,
    /// True once the header marks the class as exposed outside its project.
    class_exposed: bool,
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
    next_carry: Vec<(Option<u32>, Span, NodeId)>,
    /// A variable named `Error` is declared at module level / in the current
    /// procedure (parameters included). It shadows the `Error` statement, so a
    /// line that starts with `Error` is then an ordinary target expression.
    error_var_module: bool,
    error_var_local: bool,
    /// Conditional-compilation constant table: `(lowercase_name, value)`.
    /// Pre-seeded with VB6 predefined constants.
    cc_defs: Vec<(String, CcValue)>,
    /// Names given by a `#Const` of this module: giving one again is an error.
    cc_module_consts: Vec<String>,
    /// The last directive token read ended the line.
    cc_line_done: bool,
    /// Where the logical line being read begins: an error inside a continued line is reported
    /// on its first physical line.
    logical_line_start: u32,
    /// How many constant names the directive expressions have read so far; an operand that
    /// read none is made of literals alone.
    cc_names_read: u64,
    /// A token has been read on the current physical line.
    phys_line_has_token: bool,
    /// The token read last was the first on its physical line.
    tok_first_on_line: bool,
    /// Depth of the argument lists being read, and the first `ByVal` and first named
    /// argument met at depth one since the statement began.
    group_depth: u32,
    /// Continuation marks joined into the current logical line.
    continuations: u32,
    /// Continuation marks the logical line that ended last had.
    line_continuations: u32,
    /// Compiled bytes of the assignments already read on the current logical line, and
    /// whether every statement before the current one was such an assignment.
    line_cost: usize,
    line_cost_valid: bool,
    stmt_measured: bool,
    /// Bytes of the items of the `Const` statement read so far, while they are all measured.
    const_items_cost: Option<usize>,

    /// The argument values written with a call-site `ByVal`.
    pub by_val_args: Vec<NodeId>,
    /// The argument list about to be read follows another one directly (`f(1)(2)`).
    later_group: bool,
    stmt_byval: Option<Span>,
    stmt_named: Option<Span>,
    /// What the member chain being read at the start of a statement is the target of, while a
    /// graphics word after a dot may still end it.
    graphics_chain: Option<ChainStart>,
    /// The graphics word the member chain just read ended in.
    graphics_stop: Option<GraphicsWord>,
    /// The statements read as graphics method statements, by node: the word, and whether it was
    /// written as a member of an object (`true`) rather than bare.
    graphics_stmts: std::collections::HashMap<u32, (GraphicsWord, bool)>,
    /// Open `#If` blocks, innermost last.
    cc_stack: Vec<CcFrame>,
    /// One-token pushback for directive processing.
    cc_pending: Option<Token>,
    /// How names the module does not declare are resolved in constant expressions.
    external_names: Option<Box<declarations::ExternalResolver>>,
    /// The module-level names found by the declaration checks once the module is read.
    pub module_names: declared_names::ModuleNames,
}

/// Which part of an `If` line a list of statements is: after `Then` or after `Else` of a
/// single-line `If`, or after the `Then` of an `ElseIf` line. `colon_led` is set when the
/// `Then` is followed directly by a colon.
#[derive(Clone, Copy)]
enum InlineList {
    Then { colon_led: bool },
    Else { colon_led: bool },
    ElseIf { colon_led: bool },
}

impl InlineList {
    /// A line number that is the whole first statement is a jump, except in a colon-led list.
    fn jump_allowed(self) -> bool {
        match self {
            InlineList::Then { colon_led } | InlineList::Else { colon_led } | InlineList::ElseIf { colon_led } => !colon_led,
        }
    }
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

/// How a statement parsed between `Select Case` and the first `Case` is treated.
enum PreambleRule {
    /// Nothing is reported for it.
    Allowed,
    /// Reported as a statement before the first `Case`, after any error of its own.
    Reported,
    /// Reported as a statement before the first `Case` in place of any error of its own.
    ReportedInstead,
}

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
    /// `Property Let`: the last parameter is the value.
    PropertyValue,
    /// `Property Set`: the last parameter is the object value and has an object type.
    PropertySet,
    /// `Declare`: as a procedure, but parameter names may repeat.
    External,
    Event,
}

/// The type of a property's value or index, as far as the procedures of one property
/// have to agree on it.
#[derive(Clone, PartialEq)]
enum PropType {
    Variant,
    Builtin(u8),
    String,
    Object,
    User(u32),
    /// An array of the element type.
    Array(Box<PropType>),
}

/// What the procedures of one property have to agree on.
#[derive(Clone)]
struct PropertySig {
    kind: ProcKind,
    /// Name, pass-by-value, array and type of each index parameter.
    index: Vec<(u32, bool, bool, PropType)>,
    value: PropType,
}

impl PropertySig {
    /// True when `self` and `other`, two procedures of one property, agree.
    fn agrees_with(&self, other: &PropertySig) -> bool {
        if self.index != other.index {
            return false;
        }
        let (get, let_) = match (self.kind, other.kind) {
            (ProcKind::PropGet, ProcKind::PropLet) => (self, other),
            (ProcKind::PropLet, ProcKind::PropGet) => (other, self),
            // A `Set` value is an object whatever the getter returns.
            _ => return true,
        };
        get.value == let_.value
            || (matches!(get.value, PropType::Object | PropType::User(_)) && let_.value == PropType::Variant)
    }
}

fn prop_type(arena: &ExprArena, node: Option<NodeId>) -> PropType {
    let Some(node) = node else { return PropType::Variant };
    match arena.get(node) {
        ExprNode::BuiltinType { kind: 12 } => PropType::Variant,
        ExprNode::BuiltinType { kind: 9 } => PropType::Object,
        ExprNode::BuiltinType { kind } => PropType::Builtin(*kind as u8),
        ExprNode::StringType { .. } => PropType::String,
        ExprNode::UserType { name, .. } => PropType::User(*name),
        ExprNode::ArrayType { elem } => PropType::Array(Box::new(prop_type(arena, Some(*elem)))),
        _ => PropType::Variant,
    }
}

/// State of one open `#If … #End If` block.
#[derive(Clone, Copy)]
struct CcFrame {
    /// An `#Else` has been seen: only `#End If` may follow.
    else_seen: bool,
    /// The enclosing region is active.
    parent_active: bool,
    /// The current branch is active.
    active: bool,
    /// Some branch of this block has already been taken.
    taken: bool,
}

/// Which `Option` statement has been read before.
#[derive(Clone, Copy)]
enum OptionSeen {
    Explicit,
    Compare,
    Private,
}

impl<'src> Parser<'src> {
    /// Construct a parser over `src` using the provided scanner context.
    pub fn new(ctx: &'src mut ScannerContext, src: &'src [u8]) -> Self {
        let mut parser = Self {
            scanner:     Scanner::new(ctx, src),
            src,
            peeked:      None,
            peeked2:     None,
            pushback:    std::collections::VecDeque::new(),
            after_line_cont: None,
            diagnostics: Diagnostics::new(),
            node_spans:  NodeSpans::new(),
            decl_public: std::collections::HashMap::new(),
            decl_friend: std::collections::HashSet::new(),
            literal_radix: std::collections::HashMap::new(),
            decl_static: std::collections::HashSet::new(),
            in_procedure: false,
            reserved_names_allowed: false,
            last_item_had_as: false,
            address_of_ok: false,
            in_inline_if: false,
            in_attribute_head: false,
            inline_else_pending: 0,
            line_scan: false,
            first_eol_start: u32::MAX,
            eol_watch: false,
            in_target: false,
            bracketed_name_ok: false,
            later_declarator: false,
            in_enum_member: false,
            type_after_new: false,
            target_at_statement_start: false,
            blocks: Vec::new(),
            last_was_eol: false,
            after_colon: false,
            line_has_content: false,
            line_started: false,
            last_consumed: TokenKind::Eol,
            eols_read: 0,
            colon_on_line: false,
            module_decl_seen: false,
            static_prefix: false,
            property_sigs: Vec::new(),
            procedure_seen: false,
            class_exposed: false,
            value_owner_params: false,
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
            external_names: None,
            module_names: declared_names::ModuleNames::default(),
            cc_module_consts: Vec::new(),
            cc_line_done: false,
            logical_line_start: 0,
            cc_names_read: 0,
            phys_line_has_token: false,
            tok_first_on_line: true,
            group_depth: 0,
            continuations: 0,
            line_continuations: 0,
            line_cost: 0,
            line_cost_valid: true,
            stmt_measured: true,
            const_items_cost: None,
            by_val_args: Vec::new(),
            later_group: false,
            stmt_byval: None,
            stmt_named: None,
            graphics_chain: None,
            graphics_stop: None,
            graphics_stmts: std::collections::HashMap::new(),
            cc_defs: vec![
                ("win16".into(), CcValue::Integer(0)),
                ("win32".into(), CcValue::Integer(1)),
                ("mac".into(), CcValue::Integer(0)),
            ],
        };
        parser
    }

    /// The first limit the text breaks while it is read, if any; its diagnostic is the first one
    /// that [`Parser::parse_module`] reports.
    pub fn read_problem(&self) -> Option<crate::frontend::line_reader::ReadProblem> {
        self.scanner.read_problem()
    }

    /// Declares the first `lines` physical lines of the text to be header lines that are not
    /// code, such as a designer section that was blanked: they do not count toward the limit on
    /// the lines of a module. Call it before [`Parser::parse_module`].
    pub fn set_header_lines(&mut self, lines: usize) {
        let starts = crate::frontend::line_reader::physical_line_starts(self.src);
        let end = starts.get(lines).copied().unwrap_or(self.src.len() as u32);
        self.scanner.set_header_end(end);
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
        self.line_has_content = tok.kind != TokenKind::Eol;
        self.line_started = tok.kind != TokenKind::Eol;
        self.last_consumed = tok.kind.clone();
        self.last_token_end = tok.span.start + tok.span.len;
        match tok.kind {
            TokenKind::Kw(Kw::Colon) => {
                self.after_colon = true;
                self.colon_on_line = true;
            }
            TokenKind::Eol => {
                self.after_colon = false;
                self.colon_on_line = false;
                self.eols_read += 1;
                if std::mem::replace(&mut self.eol_watch, false) {
                    self.first_eol_start = tok.span.start;
                }
            }
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
    /// The text after the `#` of a `Declare` alias must be an ordinal: a whole number at its
    /// start, and then a text that converts to an Integer. A number that cannot be read is
    /// reported for what is wrong with it.
    fn check_alias_ordinal(&mut self, ordinal: &str, span: Span) {
        let bytes: Vec<u8> = ordinal.chars().map(|c| u8::try_from(u32::from(c)).unwrap_or(b'?')).collect();
        match self.scanner.ordinal_start(&bytes) {
            OrdinalStart::WholeNumber if converts_to_integer(ordinal) => {}
            OrdinalStart::Failed(NumberFailure::Overflow) => self.diagnostics.push(ERR_OVERFLOW, span),
            OrdinalStart::Failed(NumberFailure::Malformed) => self.diagnostics.push(ERR_SYNTAX_ERROR, span),
            _ => self.diagnostics.push(ERR_INVALID_ORDINAL, span),
        }
    }

    /// A literal node for the number, string or date token `tok`, remembering how a
    /// whole number was written when that was in octal or hexadecimal.
    fn literal_node(&mut self, arena: &mut ExprArena, tok: &Token) -> NodeId {
        let id = arena.alloc(ExprNode::Literal { lit: ast_lit_from_token(tok) });
        if matches!(tok.kind, TokenKind::IntLit | TokenKind::LongLit) {
            let text = self.src.get(tok.span.start as usize..).unwrap_or(&[]);
            if let [b'&', marker, ..] = text {
                let radix = if marker.eq_ignore_ascii_case(&b'h') { Radix::Hexadecimal } else { Radix::Octal };
                self.literal_radix.insert(id.0, radix);
            }
        }
        id
    }

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
        match self.pushback.pop_front() {
            Some(token) => token,
            None => self.fetch_logical(),
        }
    }

    /// Skips the rest of the current line of the conditional compilation stream.
    fn skip_cc_line(&mut self) {
        while !matches!(self.cc_peek().kind, TokenKind::Eol | TokenKind::Eof) {
            self.cc_advance();
        }
    }

    /// A conditional compilation word that is read as a directive at the start of a line; elsewhere
    /// it is an ordinary word, and no word of that kind begins a statement.
    fn cc_word(&mut self, kw: Kw, tok: &Token, first_on_line: bool) {
        if first_on_line {
            self.cc_directive(kw, tok.span, tok.lit.is_some());
            return;
        }
        if !self.cc_active() {
            return;
        }
        // On a continued line the error is on the line the statement began.
        let began = self.logical_line_start as usize;
        let continued = self.src[began..tok.span.start as usize].contains(&b'\n');
        let span = if continued { Span { start: self.logical_line_start, len: tok.span.len } } else { tok.span };
        self.diagnostics.push(ERR_SYNTAX_ERROR, span);
        self.skip_cc_line();
    }

    /// At the end of the text, reports the `#If` blocks left open; reports a name that cannot be
    /// kept because the table of names is full.
    fn note_stream_end_or_full_table(&mut self, tok: &Token) {
        if tok.kind == TokenKind::Eof {
            self.report_open_cc_blocks();
        }
        if tok.kind == TokenKind::Error && self.scanner.symbol_table_full() && self.cc_active() {
            // The table of names is full: the name cannot be kept.
            self.diagnostics.push(ERR_OUT_OF_MEMORY, tok.span);
        }
    }

    /// Reports an identifier longer than the language allows.
    fn check_identifier_length(&mut self, tok: &Token) {
        if tok.kind != TokenKind::Ident {
            return;
        }
        let too_long =
            tok.sym.is_some_and(|id| self.scanner.sym_name(id as u32).chars().count() > MAX_IDENTIFIER_LEN);
        if too_long {
            self.diagnostics.push(ERR_IDENTIFIER_TOO_LONG, tok.span);
        }
    }

    fn fetch_logical(&mut self) -> Token {
        loop {
            let tok = self.next_physical();
            let first_on_line = self.tok_first_on_line;
            if let TokenKind::Kw(kw @ (Kw::CcIf | Kw::CcElseIf | Kw::CcElse | Kw::CcEnd | Kw::CcConst)) =
                tok.kind
            {
                self.cc_word(kw, &tok, first_on_line);
                continue;
            }
            if first_on_line && tok.kind == TokenKind::Kw(Kw::Hash) && self.cc_active() && !self.procedure_seen {
                // A line that opens with a number sign names a directive.
                self.report_expected(tok.span, "If or Else or ElseIf or End or EndIf or Const");
                self.skip_cc_line();
                continue;
            }
            self.note_stream_end_or_full_table(&tok);
            if !self.cc_active() {
                if tok.kind == TokenKind::Eof {
                    return tok;
                }
                continue;
            }
            self.check_identifier_length(&tok);
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
                    self.continuations += 1;
                    continue;
                }
                self.after_line_cont = Some(next);
            }
            if tok.kind == TokenKind::Eol {
                self.line_continuations = self.continuations;
                self.continuations = 0;
            }
            self.tok_first_on_line = !self.phys_line_has_token;
            if self.tok_first_on_line {
                self.logical_line_start = tok.span.start;
            }
            self.phys_line_has_token = !matches!(tok.kind, TokenKind::Eol);
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
            Token { kind: expected, sym: None, lit: None, type_suffix: TypeSuffix::None, bracketed: false, span }
        }
    }

    /// The keyword a word token spells. A bracketed word keeps the identity of the keyword
    /// it spells even though it is an ordinary name; a bare keyword has it by its kind.
    /// Every place that asks which word a name is goes through here, so a word is the same
    /// word wherever the grammar looks at it. A place that wants only the bare keyword
    /// compares the token kind instead.
    fn word_of(&self, t: &Token) -> Option<Kw> {
        let sym = t.sym? as u32;
        match t.kind {
            TokenKind::Kw(k) => Some(k),
            TokenKind::Ident => self.scanner.symbol_keyword(sym),
            _ => None,
        }
    }

    /// A name in brackets at the start of a statement is read as the keyword it spells when
    /// that keyword starts a statement of its own, unless what follows makes it a name: the
    /// target of an assignment, the object of a member, or a label.
    fn read_statement_word(&mut self) -> bool {
        let word = {
            let t = self.peek().clone();
            if t.kind != TokenKind::Ident || !t.bracketed {
                return false;
            }
            match self.word_of(&t) {
                Some(kw) if begins_own_statement(kw) => kw,
                _ => return false,
            }
        };
        if matches!(
            self.peek2().kind,
            TokenKind::Kw(Kw::Eq) | TokenKind::Kw(Kw::Dot) | TokenKind::Kw(Kw::Bang) | TokenKind::Kw(Kw::Colon)
        ) {
            return false;
        }
        // A remark word, bracketed or not, makes the rest of the line a comment.
        if word == Kw::Rem {
            while !matches!(self.peek().kind, TokenKind::Eol | TokenKind::Eof) {
                self.advance();
            }
            return true;
        }
        if let Some(tok) = self.peeked.as_mut() {
            tok.kind = TokenKind::Kw(word);
            tok.bracketed = false;
        }
        false
    }

    /// The keyword the next token spells, bare or in brackets.
    fn peek_word(&mut self) -> Option<Kw> {
        let t = self.peek().clone();
        self.word_of(&t)
    }

    /// True when the next token spells the keyword `kw`, bare or in brackets.
    fn at_word(&mut self, kw: Kw) -> bool {
        self.peek_word() == Some(kw)
    }

    /// Consumes the next token when it spells the keyword `kw`, bare or in brackets.
    fn eat_word(&mut self, kw: Kw) -> bool {
        if self.at_word(kw) {
            self.advance();
            true
        } else {
            false
        }
    }

    /// True when `tok` was written in square brackets (`[Print]`).
    fn is_bracketed(&self, tok: &Token) -> bool {
        self.src.get(tok.span.start as usize) == Some(&b'[')
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
        } else if self.at_underscore_word() {
            // A word cannot begin with an underscore: the whole word is the bad character.
            let span = self.current_span();
            self.report_invalid_character(span);
            let tok = self.advance();
            Token { kind: TokenKind::Ident, ..tok }
        } else {
            let span = self.current_span();
            let what = if self.later_declarator { "variable" } else { "identifier" };
            self.diagnostics.push_labeled(ERR_EXPECTED_IDENT, span, what);
            Token { kind: TokenKind::Ident, sym: None, lit: None, type_suffix: TypeSuffix::None, bracketed: false, span }
        }
    }

    /// True when the next token is a word the scanner refused because it begins with an
    /// underscore.
    fn at_underscore_word(&mut self) -> bool {
        let (kind, start) = { let t = self.peek(); (t.kind.clone(), t.span.start) };
        kind == TokenKind::Error && self.src.get(start as usize) == Some(&b'_')
    }

    /// Reports an invalid character: inside a procedure body the compiler gives only the
    /// general syntax error.
    fn report_invalid_character(&mut self, span: Span) {
        if self.in_procedure {
            self.diagnostics.push(ERR_SYNTAX_ERROR, span);
        } else {
            self.diagnostics.push(ERR_INVALID_CHARACTER, span);
        }
    }

    /// The name after the dot of a qualified type name. Any word names a type there, even a
    /// reserved one, except `Rem`, which starts a comment.
    fn expect_qualified_type_name(&mut self) -> Token {
        if self.peek().kind == TokenKind::Kw(Kw::Rem) {
            let span = self.current_span();
            self.report_expected(span, "identifier");
            return Token { kind: TokenKind::Ident, sym: None, lit: None, type_suffix: TypeSuffix::None, bracketed: false, span };
        }
        self.expect_ident()
    }

    /// A name being declared. A reserved word cannot name anything; the words that only
    /// mean something in a module header (`Attribute`, `VB_Name`, ...) are reserved words too.
    fn expect_decl_name(&mut self) -> Token {
        // A bracketed name is an external name, never a declared one.
        if !self.bracketed_name_ok && !self.later_declarator && self.peek().bracketed {
            let span = self.current_span();
            self.report_expected(span, "identifier");
            return self.advance();
        }
        // A `Type` member may take any reserved word but the two that begin an assignment.
        let me = self.peek().kind == TokenKind::Kw(Kw::Me) && self.peek2().kind != TokenKind::Kw(Kw::LParen);
        let begins_assignment = me || self.peek().kind == TokenKind::Kw(Kw::StringS);
        if (!self.reserved_names_allowed || begins_assignment) && self.at_reserved_word() {
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
                self.scanner.is_reserved(sym)
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

    /// The type-declaration character a declared name ends with. A keyword spelled with a
    /// trailing `$` that is not reserved (`Mid$`, `Dir$`, `Error$`, ...) is, as a name, the
    /// word without the `$` followed by the string character.
    fn declared_suffix(&self, name: &Token) -> TypeSuffix {
        if name.type_suffix != TypeSuffix::None {
            return name.type_suffix;
        }
        match (&name.kind, name.sym) {
            (TokenKind::Kw(_), Some(sym)) if !self.scanner.is_reserved(sym as u32) => self.dollar_word_suffix(name),
            _ => TypeSuffix::None,
        }
    }

    /// The string character a keyword spelled with a trailing `$` ends with, reserved or not.
    fn dollar_word_suffix(&self, name: &Token) -> TypeSuffix {
        match (&name.kind, name.sym) {
            (TokenKind::Kw(_), Some(sym)) if self.scanner.is_dollar_word(sym as u32) => TypeSuffix::String,
            _ => TypeSuffix::None,
        }
    }

    /// The name of a `Type`, `Enum` or `Event` carries no type-declaration character.
    fn reject_name_suffix(&mut self, name: &Token) {
        if name.type_suffix != TypeSuffix::None {
            self.report_expected(name.span, "identifier");
        }
    }

    /// As [`Self::reject_name_suffix`], naming what was expected instead.
    fn reject_name_suffix_as(&mut self, name: &Token, what: &'static str) {
        if self.declared_suffix(name) != TypeSuffix::None {
            self.report_expected(name.span, what);
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
        if ["CDec", "Date", "Date$", "DoEvents"].iter().any(|n| name.eq_ignore_ascii_case(n)) {
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
        ["Me", "Date", "Date$", "DoEvents", "CDec"].iter().any(|n| name.eq_ignore_ascii_case(n))
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
        if ["Date", "Date$", "DoEvents", "CDec"].iter().any(|n| name.eq_ignore_ascii_case(n)) {
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
        let (args, mut bad) = self.parse_intrinsic_args(arena);
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

    /// The comma-separated arguments of an intrinsic call, after its `(`, and whether any of them
    /// is left out or named.
    fn parse_intrinsic_args(&mut self, arena: &mut ExprArena) -> (Vec<NodeId>, bool) {
        let mut args = Vec::new();
        let mut bad = false;
        if self.peek().kind == TokenKind::Kw(Kw::RParen) {
            return (args, bad);
        }
        loop {
            if matches!(self.peek().kind, TokenKind::Kw(Kw::Comma) | TokenKind::Kw(Kw::RParen)) {
                bad = true; // an argument left out
                break;
            }
            if let Some(named) = self.try_parse_named_arg(arena) {
                bad = true; // a named argument
                args.push(named);
            } else {
                args.push(self.parse_expr(arena, 0));
            }
            if !self.eat(TokenKind::Kw(Kw::Comma)) {
                break;
            }
        }
        (args, bad)
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
    fn require_class_name_after_new(&mut self, what: &'static str) {
        if self.at_reserved_word() || self.at_word(Kw::Object) {
            let span = self.current_span();
            self.report_expected(span, what);
        }
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
            // In a `Type` the word is read as the target of an assignment, which wants `=`;
            // after a comma in a declaration list a variable is wanted.
            let what = if self.reserved_names_allowed {
                "="
            } else if self.later_declarator {
                "variable"
            } else {
                "identifier"
            };
            self.diagnostics.push_labeled(ERR_EXPECTED_IDENT, span, what);
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
    /// True at the end of a statement: the end of the line, a comment, or, inside a
    /// single-line `If`, a `Rem` comment.
    fn at_statement_end(&mut self) -> bool {
        self.peek().is_stmt_end() || (self.rem_ends_statement() && self.peek().kind == TokenKind::Kw(Kw::Rem))
    }

    /// Whether a `Rem` comment may follow a statement here: inside a single-line `If` or
    /// after a statement-separating colon.
    fn rem_ends_statement(&self) -> bool {
        self.in_inline_if || self.colon_on_line
    }

    /// As [`Self::end_stmt`] for a statement whose operand list does not stop at `Rem`: the
    /// word is a syntax error even inside a single-line `If`.
    fn end_stmt_before_rem_is_error(&mut self) {
        if self.rem_ends_statement() && self.peek().kind == TokenKind::Kw(Kw::Rem) {
            let span = self.current_span();
            self.diagnostics.push_labeled(0x9c6f, span, "end of statement");
            self.skip_line();
        } else {
            self.end_stmt();
        }
    }

    /// Reports a block statement's line whose expression takes the line's compiled code past
    /// its limit. Call with the end of the line next.
    fn check_condition_size(&mut self, arena: &ExprArena, condition: NodeId, at: Span) {
        if self.peek().kind != TokenKind::Eol || self.colon_on_line {
            return;
        }
        let lines = self.line_continuations as usize + 1;
        if line_size::expression_size(arena, condition, 0, lines).is_some_and(|size| size > MAX_LINE_CODE_SIZE) {
            self.diagnostics.push(ERR_STATEMENT_TOO_COMPLEX, at);
        }
    }

    /// Reports a line, alone and ending next, whose statement costs as much as an assignment's
    /// store plus `parts` and takes the line's compiled code past its limit.
    fn check_parts_size(&mut self, parts: Option<usize>, at: Span) {
        let Some(parts) = parts else { return };
        if self.peek().kind != TokenKind::Eol || self.colon_on_line {
            return;
        }
        let lines = self.line_continuations as usize + 1;
        if cost::ASSIGNMENT + cost::PER_PHYSICAL_LINE * lines + parts > MAX_LINE_CODE_SIZE {
            self.diagnostics.push(ERR_STATEMENT_TOO_COMPLEX, at);
        }
    }

    /// Whether an output list `a; b, c` that ends the line (with no separator after the last item)
    /// takes the line's compiled code past its limit: the first item as an assignment costs it,
    /// each further item its expression and a 2-byte separator, and `extra` is what the statement
    /// itself costs beyond an assignment. Lists with an item that is not a plain expression are not
    /// measured.
    fn output_list_too_complex(&mut self, arena: &ExprArena, items: &[NodeId], extra: usize) -> bool {
        if items.is_empty()
            || self.peek().kind != TokenKind::Eol
            || self.colon_on_line
            || matches!(self.last_consumed, TokenKind::Kw(Kw::Semi) | TokenKind::Kw(Kw::Comma))
        {
            return false;
        }
        let lines = self.line_continuations as usize + 1;
        let mut total = cost::ASSIGNMENT + cost::PER_PHYSICAL_LINE * lines + extra;
        for (i, item) in items.iter().enumerate() {
            let (expr, wrapper) = match arena.get(*item) {
                ExprNode::Call { func, args } if self.is_spc_or_tab(arena, *func) => match arena.get(*args) {
                    ExprNode::ArgList { args } if args.len() == 1 => (args[0], cost::SPC_OR_TAB),
                    _ => return false,
                },
                _ => (*item, 0),
            };
            let Some(size) = line_size::expression_part(arena, expr) else { return false };
            total += size + wrapper + if i > 0 { cost::ITEM_SEPARATOR } else { 0 };
        }
        total > MAX_LINE_CODE_SIZE
    }

    fn is_spc_or_tab(&self, arena: &ExprArena, func: NodeId) -> bool {
        match arena.get(func) {
            ExprNode::NameRef { sym, .. } => {
                let name = self.scanner.sym_name(*sym);
                name.eq_ignore_ascii_case("Spc") || name.eq_ignore_ascii_case("Tab")
            }
            _ => false,
        }
    }

    /// Reports a line that ends in the expression `value` and costs `extra` more than an
    /// assignment of it when that takes the line's compiled code past its limit.
    fn length_clause_too_complex(&mut self, arena: &ExprArena, value: NodeId, extra: usize, at: Span) {
        if self.peek().kind != TokenKind::Eol || self.colon_on_line {
            return;
        }
        let lines = self.line_continuations as usize + 1;
        if line_size::expression_size(arena, value, extra, lines).is_some_and(|size| size > MAX_LINE_CODE_SIZE) {
            self.diagnostics.push(ERR_STATEMENT_TOO_COMPLEX, at);
        }
    }

    fn end_stmt(&mut self) {
        if self.peek().is_stmt_end() || (self.rem_ends_statement() && self.peek().kind == TokenKind::Kw(Kw::Rem)) {
            self.advance();
        } else if !(self.peek().kind == TokenKind::Kw(Kw::Colon)
            || (self.in_inline_if && self.inline_else_pending > 0 && self.peek().kind == TokenKind::Kw(Kw::Else)))
        {
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
    fn stray_terminator(&mut self, arena: &mut ExprArena) -> Option<NodeId> {
        if self.line_scan {
            return self.scan_terminator(arena);
        }
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
        // A limit the text breaks is found while the text is read, before any line is parsed.
        if let Some(problem) = self.scanner.read_problem() {
            self.diagnostics.push(problem.code, Span { start: problem.offset, len: 0 });
        }
        let mut nodes = Vec::new();
        loop {
            self.skip_module_blank_lines();
            if self.peek().kind == TokenKind::Eof {
                break;
            }
            self.parse_module_line(arena, &mut nodes);
        }
        self.check_declarations(arena, &nodes);
        self.diagnostics.merge_directive_pass();
        nodes
    }

    /// Reads the declaration or statement line at the start of a line of the module into `nodes`.
    fn parse_module_line(&mut self, arena: &mut ExprArena, nodes: &mut Vec<NodeId>) {
        // `Else` or a line number in front of an attribute line does not change it.
        if self.at_prefixed_attribute_line() {
            self.advance();
            self.eat(TokenKind::Kw(Kw::Colon));
            let n = self.parse_attribute_stmt(arena);
            nodes.push(n);
            return;
        }

        // A colon that starts the line cannot precede the declaration of a procedure,
        // a `Declare` or an `Event`.
        if self.after_colon && self.starts_procedure_like_declaration() {
            let span = self.current_span();
            self.diagnostics.push_labeled(0x9c6f, span, "statement or end of statement");
        }

        // An `Attribute` line after a colon takes the whole line: what stands before it is
        // read for its errors only. A procedure, `Type` or `Enum` header before it opens
        // nothing.
        let trailing_attribute = self.trailing_attribute_line_at();
        if trailing_attribute.is_some() && self.starts_module_block() {
            for _ in 0..trailing_attribute.unwrap_or(0) {
                self.advance();
            }
            return;
        }
        let head_mark = nodes.len();
        if self.reject_line_after_procedure(arena, trailing_attribute.is_some()) {
            return;
        }

        // After a procedure only procedures (and attribute lines, conditional compilation)
        // may stand; any other declaration is reported alone.
        let late_declaration = if self.procedure_seen && self.starts_module_declaration() && trailing_attribute.is_none() {
            self.eol_watch = true;
            self.first_eol_start = u32::MAX;
            Some((self.diagnostics.len(), self.current_span()))
        } else {
            None
        };

        let line_taken = self.parse_module_member(arena, nodes);
        if let Some((mark, span)) = late_declaration {
            self.finish_late_declaration(mark, span);
        }
        if trailing_attribute.is_some() && !line_taken {
            nodes.truncate(head_mark);
        }
    }

    /// After a procedure, a line that a module does not take there is reported and read away.
    /// Returns true when the line has been dealt with.
    fn reject_line_after_procedure(&mut self, arena: &mut ExprArena, has_trailing_attribute: bool) -> bool {
        if !self.procedure_seen || has_trailing_attribute {
            return false;
        }
        // After a procedure a standard module takes no `Event` or `Implements`. A class module
        // takes `Implements`, and an `Event` there is text after the procedure's end.
        if self.module_kind == ModuleKind::Standard && self.starts_event_or_implements() {
            self.error_and_skip_line(ERR_SYNTAX_ERROR);
            return true;
        }

        // After a procedure a line that is neither a procedure, an attribute line nor
        // a declaration is read as a statement; it is rejected whether or not it reads.
        if self.is_statement_line_after_procedure() {
            self.scan_line_after_procedure(arena);
            return true;
        }
        false
    }

    /// Settles the diagnostics of a declaration that stands after the end of a procedure.
    fn finish_late_declaration(&mut self, mark: usize, span: Span) {
        self.eol_watch = false;
        // A declaration line that does not read is reported as it is; one that reads
        // is text after the procedure's end (errors inside a block it opens count
        // for nothing).
        let first_line_end = self.first_eol_start;
        let failed = self.diagnostics.items()[mark..]
            .iter()
            .position(|d| d.span.start <= first_line_end && matches!(d.code, 0x9c6f | 0x9c70));
        match failed {
            Some(i) => self.diagnostics.keep_one_from(mark, mark + i),
            None => {
                self.diagnostics.truncate(mark);
                self.diagnostics.push(ERR_AFTER_END_OF_PROCEDURE, span);
            }
        }
    }

    /// A word that starts a module line as a plain name: the `BEGIN ... END` designer block and
    /// the `VERSION` header line embedded in .cls/.frm files are skipped without diagnostics
    /// (neither is a keyword; both appear as names at module level); anything else is read as
    /// a statement. Returns true when the line was read as a statement.
    fn parse_module_name_line(&mut self, arena: &mut ExprArena, nodes: &mut Vec<NodeId>) -> bool {
        let name = self.peek().sym
            .map(|id| self.scanner.sym_name(id as u32).to_ascii_lowercase())
            .unwrap_or_default();
        if name == "begin" {
            self.skip_begin_end_block();
            return false;
        }
        if name == "version" {
            // VERSION x.y CLASS / VERSION 5.00 — file-format header; skip entire line.
            self.skip_line();
            return false;
        }
        nodes.extend(self.skip_executable_statement(arena));
        true
    }

    /// Reads the member that starts at the current token of a module line. Returns true when the
    /// line was read as an executable statement.
    fn parse_module_member(&mut self, arena: &mut ExprArena, nodes: &mut Vec<NodeId>) -> bool {
        match self.peek().kind.clone() {
            // Procedure declarations
            TokenKind::Kw(Kw::Sub) => nodes.push(self.parse_proc_decl(arena, ProcKind::Sub)),
            TokenKind::Kw(Kw::Function) => nodes.push(self.parse_proc_decl(arena, ProcKind::Function)),
            TokenKind::Kw(Kw::Property) => nodes.push(self.parse_property_decl(arena)),
            // Declaration modifiers that precede Sub/Function/Property/Declare
            // or a variable/constant list. `Global` is `Public`.
            TokenKind::Kw(Kw::Public)
            | TokenKind::Kw(Kw::Global)
            | TokenKind::Kw(Kw::Private)
            | TokenKind::Kw(Kw::Friend)
            | TokenKind::Kw(Kw::Static) => nodes.push(self.parse_access_decl(arena)),
            // Type / Enum declarations
            TokenKind::Kw(Kw::Type) => {
                let at = self.current_span();
                let n = self.parse_type_decl(arena);
                self.check_public_in_object_module(arena, n, at, true);
                nodes.push(n);
            }
            TokenKind::Kw(Kw::Enum) => nodes.push(self.parse_enum_decl(arena)),
            // External declarations
            TokenKind::Kw(Kw::Declare) => {
                let at = self.current_span();
                let n = self.parse_declare_stmt(arena);
                self.check_public_in_object_module(arena, n, at, true);
                nodes.push(n);
            }
            // Event declarations — class-module-only
            TokenKind::Kw(Kw::Event) => nodes.push(self.parse_event_decl(arena)),
            // Implements — class-module-only
            TokenKind::Kw(Kw::Implements) => nodes.push(self.parse_implements_stmt(arena)),
            // Option statements
            TokenKind::Kw(Kw::Option) => nodes.push(self.parse_option_stmt(arena)),
            // Module-level Dim / Const
            TokenKind::Kw(Kw::Dim) | TokenKind::Kw(Kw::Const) => nodes.extend(self.parse_stmt(arena)),
            // Def<Type> default-type declarations (DefInt A-Z, …).
            TokenKind::Kw(k) if def_type_kw(k).is_some() => nodes.push(self.parse_def_stmt(arena)),
            // Attribute metadata lines (hosts embed these in module text).
            TokenKind::Kw(Kw::Attribute) => nodes.push(self.parse_attribute_stmt(arena)),
            TokenKind::Ident => return self.parse_module_name_line(arena, nodes),
            TokenKind::Kw(k) if starts_executable_statement(k) => {
                nodes.extend(self.skip_executable_statement(arena));
                return true;
            }
            TokenKind::IntLit | TokenKind::LongLit => {
                nodes.extend(self.skip_executable_statement(arena));
                return true;
            }
            // `End Type` with no `Type` block open.
            TokenKind::Kw(Kw::End) if self.peek2().kind == TokenKind::Kw(Kw::Type) => {
                self.error_and_skip_line(ERR_END_TYPE_OUTSIDE_TYPE);
            }
            _ => self.recover_module_unexpected(),
        }
        false
    }

    /// True when the next tokens begin the header of a procedure, a `Type` or an `Enum`, with or
    /// without an access word before it.
    fn starts_module_block(&mut self) -> bool {
        let block_word = |k: &TokenKind| {
            matches!(k, TokenKind::Kw(Kw::Sub | Kw::Function | Kw::Property | Kw::Type | Kw::Enum))
        };
        let first = self.peek().kind.clone();
        if block_word(&first) {
            return true;
        }
        matches!(first, TokenKind::Kw(Kw::Public | Kw::Private | Kw::Friend | Kw::Global | Kw::Static))
            && block_word(&self.peek2().kind.clone())
    }

    /// True when the line begins with `Event` or `Implements`, with or without an access word.
    fn starts_event_or_implements(&mut self) -> bool {
        let is_word = |k: &TokenKind| matches!(k, TokenKind::Kw(Kw::Event | Kw::Implements));
        let first = self.peek().kind.clone();
        if is_word(&first) {
            return true;
        }
        matches!(
            first,
            TokenKind::Kw(Kw::Public | Kw::Private | Kw::Global | Kw::Friend | Kw::Static)
        ) && is_word(&self.peek2().kind.clone())
    }

    /// True when the line after a procedure is read as a statement: it does not begin a
    /// procedure, an attribute line or a declaration that has its own report.
    fn is_statement_line_after_procedure(&mut self) -> bool {
        if self.peek().kind == TokenKind::Eof
            || self.starts_module_declaration()
            || self.starts_procedure_like_declaration()
        {
            return false;
        }
        match self.peek().kind.clone() {
            TokenKind::Kw(Kw::Attribute | Kw::Rem | Kw::Apos | Kw::Implements) => false,
            // Access words are read by the module-level rules.
            TokenKind::Kw(Kw::Public | Kw::Private | Kw::Global | Kw::Friend) => false,
            TokenKind::Ident => {
                let name = self.peek().sym.map(|id| self.scanner.sym_name(id as u32).to_ascii_lowercase());
                !matches!(name.as_deref(), Some("begin" | "version"))
            }
            _ => true,
        }
    }

    /// Read the line after a procedure as statements. A line that does not read as a statement
    /// is reported by the statement reader (first error only); one that does is reported as
    /// text after the procedure's end.
    fn scan_line_after_procedure(&mut self, arena: &mut ExprArena) {
        let start = self.current_span();
        let mark = self.diagnostics.len();
        self.labels_defined.clear();
        self.label_refs.clear();
        let was_in_procedure = std::mem::replace(&mut self.in_procedure, true);
        let was_scanning = std::mem::replace(&mut self.line_scan, true);
        loop {
            if self.peek().kind == TokenKind::Kw(Kw::End)
                && matches!(
                    self.peek2().kind,
                    TokenKind::Kw(Kw::Sub | Kw::Function | Kw::Property | Kw::Type | Kw::Enum)
                )
            {
                // A closing line with nothing open to close.
                self.advance();
                self.advance();
                self.end_stmt();
            } else {
                let _ = self.parse_stmt(arena);
            }
            if self.last_was_eol || self.peek().kind == TokenKind::Eof {
                break;
            }
            if self.peek().is_stmt_end() {
                self.skip_line();
                break;
            }
        }
        self.line_scan = was_scanning;
        self.in_procedure = was_in_procedure;
        self.labels_defined.clear();
        self.label_refs.clear();
        // A line that reads is reported as text after the end whatever else it says about the
        // statement it spells (a graphics method with no object is such a statement).
        let failed = self.diagnostics.items()[mark..]
            .iter()
            .position(|d| d.code != ERR_METHOD_NOT_VALID_WITHOUT_OBJECT);
        match failed {
            Some(i) => self.diagnostics.keep_one_from(mark, mark + i),
            None => {
                self.diagnostics.truncate(mark);
                self.diagnostics.push(ERR_AFTER_END_OF_PROCEDURE, start);
            }
        }
    }

    /// A block terminator read on its own line (see [`Self::scan_line_after_procedure`]): the
    /// terminator and what it takes are read, nothing is matched against an open block.
    fn scan_terminator(&mut self, arena: &mut ExprArena) -> Option<NodeId> {
        match self.terminator() {
            Some(Term::Else) => {
                self.advance();
            }
            Some(Term::ElseIf) => {
                self.advance();
                self.parse_expr(arena, 0);
                self.expect(TokenKind::Kw(Kw::Then), ERR_EXPECTED_THEN);
            }
            Some(Term::EndIf | Term::EndWith | Term::EndSelect) => {
                if self.peek().kind == TokenKind::Kw(Kw::End) {
                    self.advance();
                }
                self.advance();
                self.end_stmt();
            }
            Some(Term::Next) => {
                self.advance();
                while !self.at_statement_end()
                    && !matches!(self.peek().kind, TokenKind::Kw(Kw::Colon) | TokenKind::Kw(Kw::Else))
                {
                    self.parse_target_expr(arena);
                    if !self.eat(TokenKind::Kw(Kw::Comma)) {
                        break;
                    }
                }
                self.end_stmt();
            }
            Some(Term::Loop) => {
                self.advance();
                if matches!(self.peek().kind, TokenKind::Kw(Kw::While | Kw::Until)) {
                    self.advance();
                    self.parse_expr(arena, 0);
                }
                self.end_stmt();
            }
            Some(Term::Wend) => {
                self.advance();
                self.end_stmt();
            }
            Some(Term::Case) => {
                self.advance();
                self.parse_case_clause(arena);
            }
            None => {
                self.error_and_skip_line(ERR_UNEXPECTED_TOKEN);
            }
        }
        None
    }

    /// Hands the module's declarations to the checks of constants, bounds, sizes, defaults and
    /// names, and records what they report and the names they found.
    fn check_declarations(&mut self, arena: &ExprArena, top: &[NodeId]) {
        let scanner = &self.scanner;
        let spelling = |sym: u32| scanner.sym_name(sym).to_string();
        let checked = declarations::check_declarations(&declarations::CheckInput {
            arena,
            spans: &self.node_spans,
            src: self.src,
            top,
            spelling: &spelling,
            external: self.external_names.as_deref(),
        });
        // A declaration after the end of a procedure is refused where it is read, which is after
        // the names and the declarations are checked and the headers' defaults are worked out,
        // and among the bodies as they come.
        for d in checked.diagnostics {
            self.diagnostics.push_diagnostic_before(ERR_AFTER_END_OF_PROCEDURE, d);
        }
        for d in checked.body_diagnostics {
            let at = d.span.start;
            self.diagnostics.push_diagnostic_before_position(ERR_AFTER_END_OF_PROCEDURE, at, d);
        }
        self.module_names = checked.names;
    }

    /// Sets how names the module does not declare are resolved in constant expressions: the
    /// constants of the project and of the libraries it uses. Without it such a name leaves the
    /// value of the expression open.
    pub fn set_external_names(&mut self, resolver: Box<declarations::ExternalResolver>) {
        self.external_names = Some(resolver);
    }

    /// True when the next tokens begin a declaration that belongs before the procedures of a
    /// module: `Option`, `Def…`, `Dim`, `Const`, `Type`, `Enum`, `Declare`, `Event`, or a
    /// variable, constant, type or external declaration behind an access word.
    fn starts_module_declaration(&mut self) -> bool {
        let first = self.peek().kind.clone();
        let is_procedure_word =
            |k: &TokenKind| matches!(k, TokenKind::Kw(Kw::Sub | Kw::Function | Kw::Property));
        match first {
            TokenKind::Kw(Kw::Option | Kw::Dim | Kw::Const | Kw::Type | Kw::Enum | Kw::Declare | Kw::Event) => true,
            TokenKind::Kw(k) if def_type_kw(k).is_some() => true,
            TokenKind::Kw(Kw::Public | Kw::Private | Kw::Global | Kw::Friend | Kw::Static) => {
                let second = self.peek2().kind.clone();
                !is_procedure_word(&second)
                    && second != TokenKind::Kw(Kw::Static)
                    && !(matches!(first, TokenKind::Kw(Kw::Friend | Kw::Static)))
                    && !matches!(second, TokenKind::Eol | TokenKind::Eof)
            }
            _ => false,
        }
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
    fn skip_executable_statement(&mut self, arena: &mut ExprArena) -> Option<NodeId> {
        if let Some(at) = self.trailing_attribute_line_at() {
            // The whole line is an attribute line: what comes before it is not read.
            for _ in 0..at {
                self.advance();
            }
            return Some(self.parse_attribute_stmt(arena));
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
        None
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
            // After a procedure a line holding only a colon is a statement line, not a blank one.
            if self.eat(TokenKind::Eol) || (!self.procedure_seen && self.eat(TokenKind::Kw(Kw::Colon))) { continue; }
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
        if self.line_scan {
            // Only the line of the block statement's header is read.
            return arena.alloc(ExprNode::Block { stmts });
        }
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
        let word = self.advance();
        if self.peek().kind == TokenKind::Kw(Kw::Hash) {
            self.parse_file_print_after_keyword(arena, FileIoKind::Print)
        } else {
            let mark = self.diagnostics.len();
            let node = self.parse_print_list(arena);
            // A standard module has no surface to print to.
            if self.module_kind == ModuleKind::Standard && self.diagnostics.len() == mark {
                self.diagnostics.push(ERR_METHOD_NOT_VALID_WITHOUT_OBJECT, word.span);
            }
            node
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
            // The `#` before the channel is required here (unlike `Input`, where it
            // is what tells the statement from a call).
            let marked = self.eat(TokenKind::Kw(Kw::Hash));
            if !marked {
                let span = self.current_span();
                self.report_expected(span, "#");
            }
            let ch = self.parse_expr(arena, 0);
            self.expect(TokenKind::Kw(Kw::Comma), ERR_UNEXPECTED_TOKEN);
            let var = self.parse_target_expr(arena);
            self.end_stmt();
            arena.alloc(ExprNode::FileIoStmt {
                kind: FileIoKind::LineInput,
                channel: Some(ch),
                channel_marks: vec![marked],
                args: vec![var],
                open_mode: 0,
                open_flags: 0,
            })
        } else {
            // `Line [Step](x1,y1)-(x2,y2)[,color[,B|BF]]` is the graphics method
            // of the enclosing form; otherwise an ordinary name.
            if self.graphics_word_is_name() {
                self.keyword_as_name_stmt(arena, t)
            } else {
                self.parse_bare_graphics_stmt(arena, t)
            }
        }
    }

    /// The items of a `Print` method statement, separated by `;` or `,`
    /// (either may also end the statement).
    fn parse_print_method_items(&mut self, arena: &mut ExprArena, args: &mut Vec<NodeId>) {
        while !self.at_statement_end() {
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
            // `Spc` takes exactly one bracketed expression, `Tab` that or nothing; the item
            // ends at the closing bracket, so nothing is chained onto it.
            if self.eat(TokenKind::Kw(Kw::LParen)) {
                let width = self.parse_expr(arena, 0);
                let args = arena.alloc(ExprNode::ArgList { args: vec![width] });
                self.expect(TokenKind::Kw(Kw::RParen), ERR_EXPECTED_RPAREN);
                return Some(arena.alloc(ExprNode::Call { func: base, args }));
            }
            if tok.kind == TokenKind::Kw(Kw::Spc) {
                self.expect(TokenKind::Kw(Kw::LParen), ERR_UNEXPECTED_TOKEN);
            }
            return Some(base);
        }
        self.try_parse_expr(arena)
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
            let marked = self.eat(TokenKind::Kw(Kw::Hash));
            let ch = self.parse_expr(arena, 0);
            self.expect(TokenKind::Kw(Kw::Comma), ERR_UNEXPECTED_TOKEN);
            let val = self.parse_expr(arena, 0);
            // `Seek #n, expr` costs the channel's expression over an assignment of `expr`; `Width`
            // costs 4 more.
            if matches!(kind, FileIoKind::Seek | FileIoKind::Width) {
                let extra = if kind == FileIoKind::Width { cost::WIDTH } else { 0 };
                let parts = line_size::expression_part(arena, ch).zip(line_size::expression_part(arena, val)).map(|(a, b)| a + b + extra);
                self.check_parts_size(parts, t.span);
            }
            self.end_stmt();
            arena.alloc(ExprNode::FileIoStmt {
                kind,
                channel: Some(ch),
                channel_marks: vec![marked],
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
                channel_marks: Vec::new(),
                args: vec![old_path, new_path],
                open_mode: 0,
                open_flags: 0,
            });
        }
        // `Name <path>` must go on to `As`.
        self.error_and_skip_line(ERR_SYNTAX_ERROR);
        old_path
    }

    /// `End`: bare `End` terminates the program; `End If`/`End With`/… are
    /// structural block terminators consumed by the parent block parser.
    fn parse_end_stmt(&mut self, arena: &mut ExprArena) -> Option<NodeId> {
        // "End" alone terminates the program.
        // End Sub / End Function etc. are handled in their parent.
        if self.terminator().is_some() {
            return self.stray_terminator(arena);
        }
        self.advance();
        if self.at_statement_end()
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
        let set_span = self.current_span();
        self.advance();
        let target = self.parse_target_expr(arena);
        self.expect(TokenKind::Kw(Kw::Eq), ERR_EXPECTED_EQ);
        let value = self.parse_expr(arena, 0);
        // `LSet`/`RSet target = v` costs its target and value, and nothing more than a store.
        let parts = line_size::expression_part(arena, target).zip(line_size::expression_part(arena, value)).map(|(a, b)| a + b);
        self.check_parts_size(parts, set_span);
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
        if self.read_statement_word() {
            return None;
        }
        self.begin_statement_cost();

        // `Else` or a line number in front of an attribute line does not change it.
        if self.at_prefixed_attribute_line() {
            self.advance();
            self.eat(TokenKind::Kw(Kw::Colon));
            return Some(self.parse_attribute_stmt(arena));
        }
        if let Some(result) = self.parse_trailing_attribute_line(arena) {
            return result;
        }

        self.promote_bracketed_graphics_word();
        let kind = self.peek().kind.clone();
        match kind {
            TokenKind::Kw(kw) => self.parse_keyword_stmt(arena, kw),
            TokenKind::Ident => self.parse_name_start_stmt(arena),
            TokenKind::IntLit | TokenKind::LongLit => self.parse_line_number_stmt(arena),
            // Anything else is unexpected
            _ => {
                self.error_and_skip_to_stmt_end(ERR_UNEXPECTED_TOKEN);
                None
            }
        }
    }

    /// A statement that follows another on a line adds to the line's compiled size only when
    /// the one before it was an assignment whose size is known.
    fn begin_statement_cost(&mut self) {
        if self.peek().kind == TokenKind::Kw(Kw::Colon) {
            return;
        }
        if !self.colon_on_line {
            self.line_cost = 0;
            self.line_cost_valid = true;
        } else if !self.stmt_measured {
            self.line_cost_valid = false;
        }
        self.stmt_measured = false;
    }

    /// An `Attribute` line after a colon takes the whole line: what stands before it is read
    /// for its errors only, and a block it opens is not opened. `None` when there is no such
    /// line; otherwise the result of the statement.
    fn parse_trailing_attribute_line(&mut self, arena: &mut ExprArena) -> Option<Option<NodeId>> {
        if self.in_attribute_head {
            return None;
        }
        let at = self.trailing_attribute_line_at()?;
        if matches!(self.peek().kind, TokenKind::Kw(Kw::Do | Kw::For | Kw::While | Kw::With | Kw::Select)) {
            for _ in 0..at {
                self.advance();
            }
            return Some(Some(self.parse_attribute_stmt(arena)));
        }
        self.in_attribute_head = true;
        let _ = self.parse_stmt(arena);
        self.in_attribute_head = false;
        Some(None)
    }

    /// Reports `code` at the current token and skips the rest of the line.
    fn error_and_skip_line(&mut self, code: u32) {
        let span = self.current_span();
        self.diagnostics.push(code, span);
        self.skip_line();
    }

    /// Reports `code` at the current token and skips to the end of the statement.
    fn error_and_skip_to_stmt_end(&mut self, code: u32) {
        let span = self.current_span();
        self.diagnostics.push(code, span);
        self.skip_to_stmt_end();
    }

    /// A name and a colon that start a line: a line label. The name has been checked.
    fn parse_name_label(&mut self, arena: &mut ExprArena) -> NodeId {
        let name = self.advance();
        self.advance(); // colon
        self.after_colon = false; // the colon belongs to the label
        self.colon_on_line = false;
        let sym = name.sym.map(|s| s as u32).unwrap_or(0);
        self.define_label(LabelRef::Name(sym), name.span);
        let id = arena.alloc(ExprNode::Label { target: LabelRef::Name(sym) });
        self.set_span(id, name.span);
        id
    }

    /// A statement that starts with the keyword `kw`.
    fn parse_keyword_stmt(&mut self, arena: &mut ExprArena, kw: Kw) -> Option<NodeId> {
        // A keyword that is not reserved names a line label like any name;
        // a reserved word that starts a statement is that statement.
        if self.peek().sym.is_some()
            && self.peek2().kind == TokenKind::Kw(Kw::Colon)
            && !self.line_has_content
            && kw != Kw::Randomize
            && !self.at_reserved_word()
        {
            return Some(self.parse_name_label(arena));
        }
        self.parse_declaration_keyword_stmt(arena, kw)
    }

    /// Statements that declare or that cannot be in a procedure.
    fn parse_declaration_keyword_stmt(&mut self, arena: &mut ExprArena, kw: Kw) -> Option<NodeId> {
        match kw {
            // Variable / constant declarations
            // An access word has no place in a procedure body.
            Kw::Public | Kw::Private | Kw::Global => {
                // An access word used as a name (`Global.Print`), or alone, is no declaration.
                let code = if matches!(self.peek2().kind, TokenKind::Kw(Kw::Dot | Kw::Bang)) {
                    ERR_SYNTAX_ERROR
                } else {
                    ERR_INVALID_ATTRIBUTE_IN_PROCEDURE
                };
                self.error_and_skip_line(code);
                None
            }
            Kw::Dim | Kw::Static => {
                self.advance();
                Some(self.parse_dim_list_static(arena, false, kw == Kw::Static))
            }
            Kw::Const => {
                self.advance();
                Some(self.parse_const_list(arena))
            }
            Kw::ReDim => Some(self.parse_redim_stmt(arena)),
            Kw::Erase => Some(self.parse_erase_stmt(arena)),
            Kw::RaiseEvent => Some(self.parse_raise_event(arena)),
            Kw::Debug => Some(self.parse_debug_stmt(arena)),
            // Attribute metadata line
            Kw::Attribute => Some(self.parse_attribute_stmt(arena)),
            // `Implements` is read wherever it stands; whether it is allowed there is the binder's
            // question.
            Kw::Implements => Some(self.parse_implements_stmt(arena)),
            // `Property` used as a name: assigned, or with a member.
            Kw::Property if matches!(self.peek2().kind, TokenKind::Kw(Kw::Eq | Kw::Dot | Kw::Bang)) => {
                let t = self.advance();
                Some(self.parse_ident_stmt_with_tok(arena, t))
            }
            // Sub / Function / Property — nested (error but recover)
            Kw::Sub | Kw::Function | Kw::Property | Kw::Declare | Kw::Event | Kw::Option => {
                self.error_and_skip_line(ERR_INVALID_IN_PROCEDURE);
                None
            }
            // A `Type` or `Enum` block cannot be declared inside a procedure; the
            // whole block is consumed so its members are not read as statements.
            Kw::Type | Kw::Enum => {
                let span = self.current_span();
                let mark = self.diagnostics.len();
                if kw == Kw::Type {
                    self.parse_type_decl(arena);
                } else {
                    self.parse_enum_decl(arena);
                }
                self.diagnostics.truncate(mark);
                self.diagnostics.push(ERR_INVALID_IN_PROCEDURE, span);
                None
            }
            k if def_type_kw(k).is_some() => {
                self.error_and_skip_line(ERR_INVALID_IN_PROCEDURE);
                None
            }
            _ => self.parse_flow_keyword_stmt(arena, kw),
        }
    }

    /// Control flow, assignment keywords, jumps and block terminators.
    fn parse_flow_keyword_stmt(&mut self, arena: &mut ExprArena, kw: Kw) -> Option<NodeId> {
        match kw {
            Kw::If => Some(self.parse_if_stmt(arena)),
            Kw::For => Some(self.parse_for_stmt(arena)),
            Kw::Do => Some(self.parse_do_stmt(arena)),
            Kw::While => Some(self.parse_while_stmt(arena)),
            Kw::Select => Some(self.parse_select_stmt(arena)),
            Kw::With => Some(self.parse_with_stmt(arena)),
            // Assignment keywords
            Kw::Let => Some(self.parse_let_stmt(arena)),
            Kw::Set => Some(self.parse_set_stmt(arena)),
            // Call / exit / goto
            Kw::Call => Some(self.parse_call_stmt(arena)),
            Kw::Exit => Some(self.parse_exit_stmt(arena)),
            Kw::Return => {
                self.advance();
                self.end_stmt();
                Some(arena.alloc(ExprNode::ReturnStmt))
            }
            Kw::GoTo => Some(self.parse_goto_stmt(arena, false)),
            Kw::GoSub => Some(self.parse_goto_stmt(arena, true)),
            Kw::Go => self.parse_go_stmt(arena),
            Kw::Resume => Some(self.parse_resume_stmt(arena)),
            Kw::On => Some(self.parse_on_stmt(arena)),
            // Error
            Kw::Error if self.error_starts_target() => Some(self.parse_ident_stmt(arena)),
            Kw::Error => {
                self.advance();
                let expr = self.parse_expr(arena, 0);
                self.end_stmt();
                Some(arena.alloc(ExprNode::ErrorStmt { expr }))
            }
            Kw::Stop => {
                self.advance();
                self.end_stmt();
                Some(arena.alloc(ExprNode::Stop))
            }
            Kw::End => self.parse_end_stmt(arena),
            // A block terminator that does not close the innermost open block.
            Kw::Else | Kw::ElseIf | Kw::EndIf | Kw::Next | Kw::Loop | Kw::Wend | Kw::Case => {
                self.stray_terminator(arena)
            }
            _ => self.parse_file_keyword_stmt(arena, kw),
        }
    }

    /// `Go` as the start of a statement: a name, or a space-separated `Go To` / `Go Sub`, where
    /// the second keyword selects the kind. `Go` is a jump only as `Go To` or `Go Sub`.
    fn parse_go_stmt(&mut self, arena: &mut ExprArena) -> Option<NodeId> {
        let next = self.peek2().kind.clone();
        if matches!(next, TokenKind::Kw(Kw::Eq | Kw::Dot | Kw::Bang)) {
            return Some(self.parse_ident_stmt(arena));
        }
        if !matches!(next, TokenKind::Kw(Kw::To | Kw::Sub)) {
            self.advance();
            self.error_and_skip_line(ERR_SYNTAX_ERROR);
            return None;
        }
        Some(self.parse_goto_stmt(arena, false))
    }

    /// File I/O statements.
    fn parse_file_keyword_stmt(&mut self, arena: &mut ExprArena, kw: Kw) -> Option<NodeId> {
        let node = match kw {
            Kw::Open => self.parse_file_open(arena),
            Kw::Close => self.parse_file_close(arena),
            // `?` is the `Print` shortcut.
            // It parses identically to `Print`: `? #n, args` writes
            // to a file, bare `? args` writes to the immediate window.
            Kw::Print | Kw::Question => self.parse_print_keyword_stmt(arena),
            Kw::Write => self.parse_write_keyword_stmt(arena),
            // `Input$ #n, vars` / `InputB #n, vars` / `InputB$ #n, vars` —
            // same file-read semantics as `Input #n, vars`.
            Kw::Input | Kw::InputS | Kw::InputB | Kw::InputBS => self.parse_input_keyword_stmt(arena),
            Kw::Line => self.parse_line_keyword_stmt(arena),
            Kw::Get => self.parse_get_put_keyword_stmt(arena, FileIoKind::Get),
            Kw::Put => self.parse_get_put_keyword_stmt(arena, FileIoKind::Put),
            Kw::Seek => self.parse_channel_expr_keyword_stmt(arena, FileIoKind::Seek),
            Kw::Lock => self.parse_lock_keyword_stmt(arena, FileIoKind::Lock),
            Kw::Unlock => self.parse_lock_keyword_stmt(arena, FileIoKind::Unlock),
            Kw::Width => self.parse_channel_expr_keyword_stmt(arena, FileIoKind::Width),
            Kw::Name => self.parse_name_keyword_stmt(arena),
            _ => return self.parse_name_like_keyword_stmt(arena, kw),
        };
        Some(node)
    }

    /// Keywords that also act as names, punctuation and comments.
    fn parse_name_like_keyword_stmt(&mut self, arena: &mut ExprArena, kw: Kw) -> Option<NodeId> {
        match kw {
            // Mid / Mid$ / MidB / MidB$ statement: `Mid(s, start[, len]) = value`.
            // The spelling selects the variant flag.
            Kw::Mid | Kw::MidS | Kw::MidB | Kw::MidBS => {
                if self.peek2().kind == TokenKind::Kw(Kw::LParen) {
                    return Some(self.parse_mid_assign(arena));
                }
                let t = self.advance();
                Some(self.keyword_as_name_stmt(arena, t))
            }
            // `LSet target = expr` / `RSet target = expr` — range-copy assignment.
            // The justification side is preserved.
            Kw::LSet | Kw::RSet => Some(self.parse_range_assign_stmt(arena)),
            // `Scale` alone, or with a coordinate pair, is the graphics method.
            // `PSet`, `Circle` and `Scale` are reserved: they name no variable, only the statement.
            Kw::PSet | Kw::Circle | Kw::Scale => {
                let t = self.advance();
                Some(self.parse_bare_graphics_stmt(arena, t))
            }
            // Comment tokens — consume and return None
            Kw::Apos | Kw::Rem => {
                self.advance(); // skip to EOL is implicit (scanner stops at Apos)
                None
            }
            // Leading dot — With-block member access (.Member = expr or .Method args);
            // a leading `!` is a With-object key.
            Kw::Dot | Kw::Bang => Some(self.parse_ident_stmt(arena)),
            // Colon — empty statement separator
            Kw::Colon => {
                self.advance();
                None
            }
            // `Me` starts a statement only as the object of a member, an element or an assignment.
            Kw::Me if matches!(
                self.peek2().kind,
                TokenKind::Kw(Kw::Dot | Kw::Bang | Kw::LParen | Kw::Eq)
            ) => self.parse_me_stmt(arena),
            // Punctuation, operators and reserved words that name no statement cannot
            // start one; any other keyword is an ordinary name.
            _ if !self.starts_name_statement() => {
                self.error_and_skip_line(ERR_SYNTAX_ERROR);
                None
            }
            _ => Some(self.parse_ident_stmt(arena)),
        }
    }

    /// `Me` followed by a member, an element or an assignment.
    fn parse_me_stmt(&mut self, arena: &mut ExprArena) -> Option<NodeId> {
        // `Me` and its member access are written together.
        let me_end = self.peek().span.start + self.peek().span.len;
        let next = self.peek2().clone();
        if matches!(next.kind, TokenKind::Kw(Kw::Dot) | TokenKind::Kw(Kw::Bang))
            && next.span.start > me_end
            && !self.src[me_end as usize..next.span.start as usize].iter().any(|&b| b == 10 || b == 13)
        {
            self.diagnostics.push(ERR_SYNTAX_ERROR, next.span);
            self.skip_line();
            return None;
        }
        Some(self.parse_ident_stmt(arena))
    }

    /// A statement that starts with a name: a line label (`Foo:`), an assignment or an implicit
    /// call.
    fn parse_name_start_stmt(&mut self, arena: &mut ExprArena) -> Option<NodeId> {
        if self.peek2().kind == TokenKind::Kw(Kw::Colon) && !self.line_has_content {
            return Some(self.parse_name_label(arena));
        }
        Some(self.parse_ident_stmt(arena))
    }

    /// Numeric line label: `10  stmt` or `10: stmt`. The label is a jump target in its own
    /// right (e.g. `GoTo 10`), so it is emitted as a `Label` node carrying the line number;
    /// any statement that follows on the same line is parsed by the enclosing block loop.
    fn parse_line_number_stmt(&mut self, arena: &mut ExprArena) -> Option<NodeId> {
        if self.line_started {
            // A line number only starts a line.
            self.error_and_skip_line(ERR_SYNTAX_ERROR);
            return None;
        }
        let tok = self.advance(); // consume the line number
        if matches!(self.peek().kind, TokenKind::IntLit | TokenKind::LongLit) {
            self.error_and_skip_line(ERR_SYNTAX_ERROR);
            return None;
        }
        let line = line_label_value(&tok.lit);
        if self.eat(TokenKind::Kw(Kw::Colon)) { // optional colon
            self.after_colon = false;
            self.colon_on_line = false;
        } else {
            // A line number alone does not make a later `Name:` a mid-line label.
            self.line_has_content = false;
        }
        self.define_label(LabelRef::Line(line), tok.span);
        let id = arena.alloc(ExprNode::Label { target: LabelRef::Line(line) });
        self.set_span(id, tok.span);
        Some(id)
    }

    // ── Individual statement parsers ──────────────────────────────────────────

    /// Implicit assignment or call when the line starts with an identifier.
    ///
    /// Covers: `target = expr` (Let), `Call target(args)`, label definition,
    /// property call without `Set`.
    fn parse_ident_stmt(&mut self, arena: &mut ExprArena) -> NodeId {
        self.target_at_statement_start = true;
        self.stmt_byval = None;
        self.stmt_named = None;
        let start = self.current_span();
        self.graphics_chain = Some(ChainStart::Statement);
        self.graphics_stop = None;
        let lhs = self.parse_target_expr(arena);
        self.graphics_chain = None;
        self.finish_ident_stmt(arena, lhs, start)
    }

    /// The rest of a statement that starts with a target expression `lhs` (read from `start`):
    /// an assignment, a graphics method statement, or an implicit call.
    fn finish_ident_stmt(&mut self, arena: &mut ExprArena, lhs: NodeId, start: Span) -> NodeId {
        let graphics_word = self.graphics_stop.take();
        // Without `Call`, the brackets after the callee hold plain expressions: a `ByVal` or a
        // named argument there is a syntax error, except that a named argument may stand in
        // the brackets of something assigned to.
        let (byval, named) = (self.stmt_byval.take(), self.stmt_named.take());
        let assigned = self.peek().kind == TokenKind::Kw(Kw::Eq);
        if let Some(span) = byval.or(if assigned { None } else { named }) {
            self.diagnostics.push(ERR_SYNTAX_ERROR, span);
            self.skip_line();
            return lhs;
        }
        // A member chain that ended in a graphics word hands the rest of the line to the reader
        // of that word's operands.
        if let Some(word) = graphics_word {
            return self.finish_graphics_method(arena, lhs, word, ChainStart::Statement, start);
        }
        let eq_span = self.current_span();
        if self.eat(TokenKind::Kw(Kw::Eq)) {
            return self.finish_assignment(arena, lhs, eq_span);
        }
        if self.peek().kind == TokenKind::Kw(Kw::ColonEq) {
            // A named argument is only meaningful inside an argument list, never as a statement.
            self.error_and_skip_line(ERR_SYNTAX_ERROR);
            return lhs;
        }
        if let Some((callee, arg)) = self.paren_first_argument(arena, lhs) {
            return self.finish_paren_call_stmt(arena, callee, arg);
        }
        if self.is_call_that_needs_call_keyword(arena, lhs) {
            // `Foo()`, `Foo(AddressOf Bar)`, `Foo(a, b)` and `Foo(a)(b)` are not statements: `Call` is needed for them.
            self.report_expected_and_skip_line("end of statement");
            return lhs;
        }
        self.finish_implicit_call(arena, lhs)
    }

    /// Reports that `what` was expected at the current token and skips the rest of the line.
    fn report_expected_and_skip_line(&mut self, what: &'static str) {
        let span = self.current_span();
        self.report_expected(span, what);
        self.skip_line();
    }

    /// True when `lhs` is a bracketed call that is not a statement without `Call`.
    fn is_call_that_needs_call_keyword(&self, arena: &ExprArena, lhs: NodeId) -> bool {
        self.has_empty_call(arena, lhs)
            || self.call_with_address_of_argument(arena, lhs)
            || self.call_with_several_arguments(arena, lhs)
            || matches!(arena.get(lhs), ExprNode::Call { func, .. } if matches!(arena.get(*func), ExprNode::Call { .. }))
    }

    /// The value of an assignment `lhs = value`, after the `=`.
    fn finish_assignment(&mut self, arena: &mut ExprArena, lhs: NodeId, eq_span: Span) -> NodeId {
        let value = self.parse_expr(arena, 0);
        self.measure_assignment(arena, lhs, value, eq_span);
        self.end_stmt();
        arena.alloc(ExprNode::Assign { target: lhs, value })
    }

    /// Adds the assignment to the compiled size of its line. A line of compiled code holds
    /// 0x3c00 bytes: the assignments before this one on the line (each followed by its 2-byte
    /// separator), this one, and 4 for each physical line.
    fn measure_assignment(&mut self, arena: &ExprArena, lhs: NodeId, value: NodeId, eq_span: Span) {
        match self.peek().kind {
            TokenKind::Eol => self.check_line_code_size(arena, lhs, value, eq_span),
            TokenKind::Kw(Kw::Colon) => match line_size::assignment_size(arena, lhs, value, 0) {
                Some(size) => {
                    self.line_cost += size + cost::OPERATOR;
                    self.stmt_measured = true;
                }
                None => self.line_cost_valid = false,
            },
            _ => {}
        }
    }

    /// Reports an assignment that ends its line when the line passes the size limit.
    fn check_line_code_size(&mut self, arena: &ExprArena, lhs: NodeId, value: NodeId, eq_span: Span) {
        let (before, valid) = (self.line_cost, self.line_cost_valid);
        let lines = self.line_continuations as usize + 1;
        if !valid {
            return;
        }
        let Some(size) = line_size::assignment_size(arena, lhs, value, lines) else { return };
        if before + size > MAX_LINE_CODE_SIZE {
            self.diagnostics.push(ERR_STATEMENT_TOO_COMPLEX, eq_span);
        }
    }

    /// `Foo (a) - (b), c` / `Foo (a), b`: in a call statement the
    /// parenthesised group after the callee begins the first argument
    /// expression; it is not the argument list.
    fn finish_paren_call_stmt(&mut self, arena: &mut ExprArena, callee: NodeId, arg: NodeId) -> NodeId {
        let first = arena.alloc(ExprNode::Paren { inner: arg });
        let first = self.parse_infix(arena, first, 0);
        let mut args = vec![first];
        if !self.eat(TokenKind::Kw(Kw::Comma)) {
            self.end_stmt();
        } else if self.at_call_arguments_end() {
            // A comma needs another argument after it, empty or not, before the line ends.
            self.report_expected_and_skip_line("expression");
        } else {
            args.extend(self.parse_stmt_arg_list(arena));
            self.end_stmt();
        }
        let args = arena.alloc(ExprNode::ArgList { args });
        arena.alloc(ExprNode::CallStmt { callee, args })
    }

    /// Whether the line, after a bracketed call, goes on with arguments: it does after a comma.
    fn call_goes_on_with_arguments(&mut self) -> bool {
        let at_end = self.peek().is_stmt_end()
            || matches!(
                self.peek().kind,
                TokenKind::Kw(Kw::Colon) | TokenKind::Kw(Kw::Else) | TokenKind::Eof
            );
        at_end || self.eat(TokenKind::Kw(Kw::Comma))
    }

    /// Implicit call with optional trailing arguments: `Foo a, b`.
    fn finish_implicit_call(&mut self, arena: &mut ExprArena, lhs: NodeId) -> NodeId {
        // After a plain name `As` opens an As clause, so it cannot name the first argument.
        if matches!(arena.get(lhs), ExprNode::NameRef { .. })
            && self.peek().kind == TokenKind::Kw(Kw::As)
            && self.peek2().kind == TokenKind::Kw(Kw::ColonEq)
        {
            self.error_and_skip_line(ERR_SYNTAX_ERROR);
            return lhs;
        }
        // After a parenthesised coordinate pair (`PSet (x, y), color`) the comma
        // only separates the pair from the remaining arguments.
        if matches!(arena.get(lhs), ExprNode::Call { .. }) && !self.call_goes_on_with_arguments() {
            self.report_expected_and_skip_line("end of statement");
            return lhs;
        }
        // A statement cannot end in a bang access, with or without arguments.
        if self.ends_in_bang_member(arena, lhs) {
            let span = if self.peek().is_stmt_end() { self.chain_span(arena, lhs) } else { self.current_span() };
            self.diagnostics.push(ERR_SYNTAX_ERROR, span);
            self.skip_line();
            return lhs;
        }
        let call_span = self.current_span();
        let args = self.parse_stmt_arg_list(arena);
        self.measure_randomize(arena, lhs, &args, call_span);
        self.end_stmt();
        if args.is_empty() {
            return lhs;
        }
        let args = arena.alloc(ExprNode::ArgList { args });
        arena.alloc(ExprNode::CallStmt { callee: lhs, args })
    }

    /// `Randomize seed` costs 2 over an assignment of the seed.
    fn measure_randomize(&mut self, arena: &ExprArena, lhs: NodeId, args: &[NodeId], call_span: Span) {
        let (ExprNode::NameRef { sym, .. }, [seed]) = (arena.get(lhs), args) else { return };
        if self.scanner.sym_name(*sym).eq_ignore_ascii_case("Randomize") {
            let parts = line_size::expression_part(arena, *seed).map(|part| cost::RANDOMIZE + part);
            self.check_parts_size(parts, call_span);
        }
    }

    /// True when `lhs` is a call, or a call of a call, with empty brackets.
    fn has_empty_call(&self, arena: &ExprArena, lhs: NodeId) -> bool {
        let ExprNode::Call { func, args } = arena.get(lhs) else { return false };
        let empty = matches!(arena.get(*args), ExprNode::ArgList { args } if args.is_empty());
        empty || self.has_empty_call(arena, *func)
    }

    /// True when `lhs` is a bracketed call of two or more arguments, which needs `Call`.
    fn call_with_several_arguments(&self, arena: &ExprArena, lhs: NodeId) -> bool {
        let ExprNode::Call { args, .. } = arena.get(lhs) else { return false };
        matches!(arena.get(*args), ExprNode::ArgList { args } if args.len() >= 2)
    }

    /// True when `lhs` is a bracketed call that has an `AddressOf` argument.
    fn call_with_address_of_argument(&self, arena: &ExprArena, lhs: NodeId) -> bool {
        let ExprNode::Call { args, .. } = arena.get(lhs) else { return false };
        matches!(arena.get(*args), ExprNode::ArgList { args }
            if args.iter().any(|a| matches!(arena.get(*a), ExprNode::AddressOf { .. })))
    }

    /// Where the chain `target` ends in a member or a name: a call has no span of its own, so
    /// this is the span of the thing that is called.
    fn chain_span(&self, arena: &ExprArena, mut target: NodeId) -> Span {
        while let ExprNode::Call { func, .. } = arena.get(target) {
            target = *func;
        }
        self.node_spans.get(target)
    }

    /// True when `target` is a `!` member access, or a call of one.
    fn ends_in_bang_member(&self, arena: &ExprArena, target: NodeId) -> bool {
        match arena.get(target) {
            ExprNode::MemberAccess { bang: true, .. } => true,
            ExprNode::Call { func, .. } => self.ends_in_bang_member(arena, *func),
            _ => false,
        }
    }

    /// True when a `Call` target cannot be called: a `!` access anywhere in its chain, or, in a
    /// chain that starts at a plain name, a call whose callee is itself a call (`Call f(1)(2)`).
    /// A chain that starts at `Me` or at the object of a `With` may have bracket groups in a row.
    fn call_target_is_malformed(&self, arena: &ExprArena, target: NodeId) -> bool {
        let mut root = target;
        let mut call_of_call = false;
        loop {
            match arena.get(root) {
                ExprNode::MemberAccess { bang: true, .. } => return true,
                ExprNode::MemberAccess { base, .. } => root = *base,
                ExprNode::Call { func, .. } => {
                    call_of_call |= matches!(arena.get(*func), ExprNode::Call { .. });
                    root = *func;
                }
                _ => break,
            }
        }
        call_of_call && matches!(arena.get(root), ExprNode::NameRef { .. })
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
        let (func, arg) = (*func, *arg);
        // A call, or a `!` member, in the callee's place is never the callee of a call statement.
        if matches!(arena.get(func), ExprNode::Call { .. } | ExprNode::MemberAccess { bang: true, .. }) {
            return None;
        }
        let kind = self.peek().kind.clone();
        let continues = kind == TokenKind::Kw(Kw::Comma) || infix_bp(&kind).is_some();
        continues.then_some((func, arg))
    }

    /// The target of a `Let` or `Set` statement. `Err` carries the target of a statement that
    /// has been reported already: a graphics word cannot be assigned to.
    fn parse_assignment_target(&mut self, arena: &mut ExprArena) -> Result<NodeId, NodeId> {
        let start = self.current_span();
        self.graphics_chain = Some(ChainStart::Assignment);
        self.graphics_stop = None;
        let target = self.parse_target_expr(arena);
        self.graphics_chain = None;
        match self.graphics_stop.take() {
            Some(word) => Err(self.finish_graphics_method(arena, target, word, ChainStart::Assignment, start)),
            None => Ok(target),
        }
    }

    /// Parses a `Let` assignment statement.
    fn parse_let_stmt(&mut self, arena: &mut ExprArena) -> NodeId {
        self.parse_keyword_assignment(arena, |target, value| ExprNode::Assign { target, value })
    }

    /// Parses a `Set` assignment statement.
    fn parse_set_stmt(&mut self, arena: &mut ExprArena) -> NodeId {
        self.parse_keyword_assignment(arena, |target, value| ExprNode::SetAssign { target, value })
    }

    /// An assignment that starts with the keyword `Let` or `Set` (the current token), built by
    /// `build` from the target and the value.
    fn parse_keyword_assignment(
        &mut self,
        arena: &mut ExprArena,
        build: impl FnOnce(NodeId, NodeId) -> ExprNode,
    ) -> NodeId {
        self.advance(); // consume Let / Set
        let target = match self.parse_assignment_target(arena) {
            Ok(target) => target,
            Err(reported) => return reported,
        };
        self.expect(TokenKind::Kw(Kw::Eq), ERR_EXPECTED_EQ);
        let value = self.parse_expr(arena, 0);
        self.end_stmt();
        arena.alloc(build(target, value))
    }

    /// `Mid`/`Mid$` string-replacement statement: `Mid(s, start[, len]) = value`.
    ///
    /// The spelling selects the variant flag: `MidB`/`MidB$` are byte-oriented;
    /// `Mid$`/`MidB$` carry the `$` bit.
    fn parse_mid_assign(&mut self, arena: &mut ExprArena) -> NodeId {
        let mid_span = self.current_span();
        let (byte_oriented, dollar) = match self.peek().kind {
            TokenKind::Kw(Kw::Mid)   => (false, false),
            TokenKind::Kw(Kw::MidS)  => (false, true),
            TokenKind::Kw(Kw::MidB)  => (true,  false),
            TokenKind::Kw(Kw::MidBS) => (true,  true),
            _ => (false, false),
        };
        self.advance(); // consume Mid / Mid$ / MidB / MidB$
        self.expect(TokenKind::Kw(Kw::LParen), ERR_EXPECTED_LPAREN);
        // `Mid(target, start [, length])`: the first operand is a target, the others are
        // plain expressions; none can be named, left out or written with a call-site `ByVal`.
        let mut items = vec![self.parse_target_expr(arena)];
        let mut well_formed = self.peek().kind == TokenKind::Kw(Kw::Comma);
        if well_formed {
            self.advance();
            items.push(self.parse_expr(arena, 0));
            if self.eat(TokenKind::Kw(Kw::Comma)) {
                items.push(self.parse_expr(arena, 0));
            }
            well_formed = self.peek().kind == TokenKind::Kw(Kw::RParen);
        }
        let args = arena.alloc(ExprNode::ArgList { args: items });
        if !well_formed {
            // The statement is abandoned at the first operand error, as the compiler does.
            self.error_and_skip_to_stmt_end(ERR_UNEXPECTED_TOKEN);
            let value = self.parse_primary_placeholder(arena);
            return arena.alloc(ExprNode::MidAssign { byte_oriented, dollar, args, value });
        }
        self.advance(); // )
        self.expect(TokenKind::Kw(Kw::Eq), ERR_EXPECTED_EQ);
        let value = self.parse_expr(arena, 0);
        // `Mid(s, a [, n]) = v` costs its operands and nothing more than an assignment's store,
        // in each of its four spellings.
        {
            let parts = line_size::expression_part(arena, value).and_then(|value| {
                let ExprNode::ArgList { args } = arena.get(args) else { return None };
                args.iter().try_fold(value, |total, arg| Some(total + line_size::expression_part(arena, *arg)?))
            });
            self.check_parts_size(parts, mid_span);
        }
        self.end_stmt();
        arena.alloc(ExprNode::MidAssign { byte_oriented, dollar, args, value })
    }

    /// Parse a `Print`/`?` output list.
    fn parse_print_list(&mut self, arena: &mut ExprArena) -> NodeId {
        let mut args = Vec::new();
        self.parse_print_method_items(arena, &mut args);
        self.end_stmt();
        arena.alloc(ExprNode::FileIoStmt {
            kind: FileIoKind::Print,
            channel: None,
            channel_marks: Vec::new(),
            args,
            open_mode: 0,
            open_flags: 0,
        })
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
        // `Shared` directly after `Dim`, `Public`, `Private` or `Global` is skipped; after
        // `Static` or `Const` (and in `ReDim`) it is an error like any reserved word.
        if !is_static && !is_const {
            self.eat(TokenKind::Kw(Kw::Shared));
        }
        if is_static && self.in_procedure && self.peek().kind == TokenKind::Kw(Kw::Property) {
            // After `Static` a `Property` starts a property procedure, which a procedure body cannot hold.
            self.error_and_skip_line(ERR_SYNTAX_ERROR);
            return arena.alloc(ExprNode::Block { stmts: Vec::new() });
        }
        let mut stmts = Vec::new();
        loop {
            // Variable declarations accept `As Type()`; `Const` does not.
            self.later_declarator = !is_const && !stmts.is_empty();
            let item = self.parse_dim_item_with(arena, is_const, is_static, false, is_const);
            self.later_declarator = false;
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
        // The first item of a `Const` statement starts the sum of its items' bytes.
        if is_const && matches!(self.last_consumed, TokenKind::Kw(Kw::Const)) {
            self.const_items_cost = Some(0);
        }
        let item_span = self.current_span();
        let with_events = self.parse_with_events_prefix(is_const);
        let name = self.expect_decl_name();
        self.note_declared_name(&name, !self.in_procedure);
        let name_id = name.sym.map(|s| s as u32).unwrap_or(0);
        self.reject_const_brackets(is_const);
        if with_events && self.peek().kind == TokenKind::Kw(Kw::LParen) {
            // An event source is one object variable: it takes no array bounds.
            self.report_expected_and_skip_line("As");
            let id = arena.alloc(ExprNode::DimItem {
                name: name_id, is_const, is_static, is_new: false, with_events, bounds: None, type_node: None, init: None,
            });
            self.set_span(id, name.span);
            return id;
        }
        let bounds = if self.eat(TokenKind::Kw(Kw::LParen)) {
            let d = self.parse_array_bounds(arena);
            self.expect(TokenKind::Kw(Kw::RParen), ERR_EXPECTED_RPAREN);
            d
        } else {
            None
        };
        self.note_module_declaration(is_const, bounds.is_some());
        self.reject_suffix_with_as(&name, "end of statement");
        self.reject_dir_name_with_as(&name, is_const);
        self.last_item_had_as = self.peek().kind == TokenKind::Kw(Kw::As);
        if with_events && !self.last_item_had_as {
            let span = self.current_span();
            self.report_expected(span, "As");
        }
        let (is_new, type_node) = if self.eat(TokenKind::Kw(Kw::As)) {
            let (is_new, ty) = self.parse_dim_as_type(arena, is_const, with_events, array_suffix_ok);
            (is_new, Some(ty))
        } else {
            // `Dim x%` ≡ `Dim x As Integer`
            (false, self.type_node_from_suffix(arena, name.type_suffix))
        };
        // Only a constant has an initial value.
        let init = if is_const { self.parse_const_value(arena, value_required, item_span) } else { None };
        let id = arena.alloc(ExprNode::DimItem { name: name_id, is_const, is_static, is_new, with_events, bounds, type_node, init });
        self.set_span(id, name.span);
        id
    }

    /// A constant is not an array: brackets after its name are reported.
    fn reject_const_brackets(&mut self, is_const: bool) {
        if is_const && !self.in_enum_member && self.peek().kind == TokenKind::Kw(Kw::LParen) {
            let span = self.current_span();
            self.report_expected(span, "As or =");
        }
    }

    /// Notes a declaration at module level, and whether it declares an array.
    fn note_module_declaration(&mut self, is_const: bool, is_array: bool) {
        if self.in_procedure {
            return;
        }
        self.module_decl_seen = true;
        if !is_const {
            self.module_array_seen |= is_array;
        }
    }

    /// A member named `WithEvents` is followed by its type, not by a name. Reads the word when
    /// it starts a variable of events, and says whether it did.
    fn parse_with_events_prefix(&mut self, is_const: bool) -> bool {
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
        with_events
    }

    /// `Dir$` is the one built-in `$` name that a variable or member declaration still reads
    /// as `Dir` with a type-declaration character.
    fn reject_dir_name_with_as(&mut self, name: &Token, is_const: bool) {
        if !is_const
            && name.sym.is_some_and(|s| self.scanner.sym_name(s as u32).eq_ignore_ascii_case("Dir$"))
            && self.peek().kind == TokenKind::Kw(Kw::As)
        {
            let span = self.current_span();
            self.report_expected(span, "end of statement");
        }
    }

    /// The type after `As` in a variable or constant declaration (the `As` is read): whether it
    /// is created with `New`, and the type node.
    fn parse_dim_as_type(
        &mut self,
        arena: &mut ExprArena,
        is_const: bool,
        with_events: bool,
        array_suffix_ok: bool,
    ) -> (bool, NodeId) {
        if with_events {
            self.check_event_source_type();
        }
        let is_new = self.eat(TokenKind::Kw(Kw::New));
        // `Object` is not a creatable class.
        if is_new && (self.at_word(Kw::Object) || self.at_reserved_word()) {
            // `New` takes a class name only.
            let span = self.current_span();
            self.report_expected(span, "identifier");
        }
        let type_span = self.current_span();
        self.type_after_new = is_new;
        let ty = self.parse_type_spec(arena, array_suffix_ok);
        // A constant has a built-in type: a class or user type name does not fit.
        if is_const && (is_new || matches!(arena.get(ty), ExprNode::UserType { .. })) {
            self.report_expected(type_span, "type name");
        }
        (is_new, ty)
    }

    /// An event source is a named class: `New` and the built-in types are out.
    fn check_event_source_type(&mut self) {
        let span = self.current_span();
        if self.peek().kind == TokenKind::Kw(Kw::New) {
            self.diagnostics.push(ERR_INVALID_NEW_KEYWORD, span);
        } else if matches!(
            self.peek_word(),
            Some(
                Kw::Integer | Kw::Long | Kw::Single | Kw::Double | Kw::Currency | Kw::Date | Kw::Byte
                    | Kw::Boolean | Kw::Variant | Kw::Object | Kw::String | Kw::Decimal
            )
        ) {
            self.report_expected(span, "identifier");
        }
    }

    /// The value of a constant, and the account of its size. `value_required` makes `= value`
    /// mandatory (a `Const` declaration, as opposed to an `Enum` member).
    fn parse_const_value(&mut self, arena: &mut ExprArena, value_required: bool, item_span: Span) -> Option<NodeId> {
        let value = if value_required {
            self.expect(TokenKind::Kw(Kw::Eq), ERR_EXPECTED_EQ);
            Some(self.parse_expr(arena, 0))
        } else if self.eat(TokenKind::Kw(Kw::Eq)) {
            Some(self.parse_expr(arena, 0))
        } else {
            None
        };
        // Each item costs its name and its value; items are separated by 2 bytes.
        let item = value.and_then(|v| line_size::expression_part(arena, v)).map(|part| cost::CONST + part);
        self.const_items_cost = self.const_items_cost.zip(item).map(|(total, item)| total + item);
        if self.peek().kind == TokenKind::Kw(Kw::Comma) {
            self.const_items_cost = self.const_items_cost.map(|total| total + cost::ITEM_SEPARATOR);
        } else {
            let total = self.const_items_cost.take();
            self.check_parts_size(total, item_span);
        }
        value
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
        let redim_span = self.current_span();
        self.advance(); // consume ReDim
        let preserve = self.eat(TokenKind::Kw(Kw::Preserve));
        let mut stmts = Vec::new();
        loop {
            stmts.push(self.parse_redim_item(arena, preserve));
            if !self.eat(TokenKind::Kw(Kw::Comma)) { break; }
        }
        self.check_redim_size(arena, &stmts, redim_span);
        self.end_stmt();
        arena.alloc(ExprNode::Block { stmts })
    }

    /// One array of a ReDim statement.
    fn parse_redim_item(&mut self, arena: &mut ExprArena, preserve: bool) -> NodeId {
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
        // `ReDim` names an existing array, so a bracketed name is allowed.
        self.bracketed_name_ok = true;
        let mut name = self.expect_decl_name();
        self.bracketed_name_ok = false;
        let bounds = self.parse_redim_name_chain(arena, &mut base, &mut name);
        let name_id = name.sym.map(|s| s as u32).unwrap_or(0);
        self.reject_suffix_with_as(&name, "end of statement");
        let type_node = if self.eat(TokenKind::Kw(Kw::As)) {
            Some(self.parse_type_spec(arena, false))
        } else { None };
        let item = arena.alloc(ExprNode::ReDimItem { preserve, base, name: name_id, bounds, type_node });
        self.set_span(item, name.span);
        item
    }

    /// The name of a ReDim item is the last of a chain of members and elements; reads the chain
    /// into `base` and `name`, and returns the bounds list that follows the last name.
    fn parse_redim_name_chain(&mut self, arena: &mut ExprArena, base: &mut Option<NodeId>, name: &mut Token) -> Option<NodeId> {
        loop {
            if self.peek().kind == TokenKind::Kw(Kw::LParen) {
                self.advance();
                let list = self.parse_array_bounds(arena);
                self.check_redim_bounds_given(arena, list);
                self.expect(TokenKind::Kw(Kw::RParen), ERR_EXPECTED_RPAREN);
                if self.peek().kind != TokenKind::Kw(Kw::Dot) {
                    return list;
                }
                // An index into the element named so far.
                let element = self.redim_chain_node(arena, *base, name);
                let args = list.unwrap_or_else(|| arena.alloc(ExprNode::ArgList { args: vec![] }));
                *base = Some(arena.alloc(ExprNode::Call { func: element, args }));
                self.advance(); // `.`
                *name = self.expect_ident();
            } else if self.eat(TokenKind::Kw(Kw::Dot)) {
                *base = Some(self.redim_chain_node(arena, *base, name));
                *name = self.expect_ident();
            } else {
                self.expect(TokenKind::Kw(Kw::LParen), ERR_EXPECTED_LPAREN);
                return None;
            }
        }
    }

    /// The name of a link of a ReDim chain: a plain name, or a member of what stands before it.
    fn redim_chain_node(&mut self, arena: &mut ExprArena, base: Option<NodeId>, name: &Token) -> NodeId {
        let sym = name.sym.map(|s| s as u32).unwrap_or(0);
        let node = match base {
            None => arena.alloc(ExprNode::NameRef { sym, suffix: name.type_suffix }),
            Some(b) => arena.alloc(ExprNode::MemberAccess { base: b, member: sym, bang: false, bracketed: self.is_bracketed(name) }),
        };
        self.set_span(node, name.span);
        node
    }

    /// Reports empty brackets or a left out bound in the bounds list of a ReDim item.
    fn check_redim_bounds_given(&mut self, arena: &ExprArena, list: Option<NodeId>) {
        let bounds_given = list.is_some_and(|l| {
            matches!(arena.get(l), ExprNode::ArgList { args }
                if !args.is_empty() && args.iter().all(|a| !matches!(arena.get(*a), ExprNode::MissingArg)))
        });
        if !bounds_given && self.peek().kind == TokenKind::Kw(Kw::RParen) {
            let span = self.current_span();
            self.report_expected(span, "expression");
        }
    }

    /// `ReDim a(n)` of one array with one bound costs 8 over an assignment of the bound.
    fn check_redim_size(&mut self, arena: &ExprArena, stmts: &[NodeId], redim_span: Span) {
        let [only] = stmts else { return };
        let ExprNode::ReDimItem { preserve: false, base: None, bounds: Some(list), type_node: None, .. } = arena.get(*only)
        else { return };
        let ExprNode::ArgList { args } = arena.get(*list) else { return };
        let [bound] = args.as_slice() else { return };
        let parts = line_size::expression_part(arena, *bound).map(|part| cost::REDIM + part);
        self.check_parts_size(parts, redim_span);
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
        let step = if self.eat_word(Kw::Step) {
            Some(self.parse_expr(arena, 0))
        } else { None };
        // `For i = a To b [Step c]` costs the variable, the bounds and the step, and its own code.
        if matches!(arena.get(var), ExprNode::NameRef { .. }) {
            let parts = line_size::expression_part(arena, start).zip(line_size::expression_part(arena, end)).and_then(|(a, b)| {
                let step_part = match step {
                    Some(step) => line_size::expression_part(arena, step)?,
                    None => 0,
                };
                Some(cost::NAME_OR_INTEGER + cost::FOR + a + b + step_part)
            });
            self.check_parts_size(parts, var_span);
        }
        self.end_stmt();
        if self.line_scan {
            let body = self.parse_block(arena);
            return arena.alloc(ExprNode::For { var, start, end, step, body, next_var: None });
        }
        self.blocks.push(Block::For(var_sym));
        let body = self.parse_block(arena);
        let next_var = self.finish_next(arena);
        self.blocks.pop();
        arena.alloc(ExprNode::For { var, start, end, step, body, next_var })
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
        if self.line_scan {
            let body = self.parse_block(arena);
            return arena.alloc(ExprNode::ForEach { var, collection, body, next_var: None });
        }
        self.blocks.push(Block::For(var_sym));
        let body = self.parse_block(arena);
        let next_var = self.finish_next(arena);
        self.blocks.pop();
        arena.alloc(ExprNode::ForEach { var, collection, body, next_var })
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
    fn finish_next(&mut self, arena: &mut ExprArena) -> Option<NodeId> {
        if !self.next_carry.is_empty() {
            let (name, span, node) = self.next_carry.remove(0);
            self.check_next_variable(name, span);
            return Some(node);
        }
        if self.peek().kind != TokenKind::Kw(Kw::Next) {
            self.report_unclosed(ERR_FOR_WITHOUT_NEXT);
            return None;
        }
        self.advance(); // Next
        let names = self.parse_next_variables(arena);
        if names.is_empty() && self.peek().kind == TokenKind::Kw(Kw::Else) {
            // `Next` alone does not end a single-line `If` part; `Next i` does.
            let span = self.current_span();
            self.diagnostics.push(ERR_SYNTAX_ERROR, span);
        }
        self.end_stmt();
        self.close_next_names(names)
    }

    /// The variables after `Next`: each with its symbol (when it is a plain name), its span and
    /// its node.
    fn parse_next_variables(&mut self, arena: &mut ExprArena) -> Vec<(Option<u32>, Span, NodeId)> {
        let mut names: Vec<(Option<u32>, Span, NodeId)> = Vec::new();
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
            names.push((name, span, var));
            if !self.eat(TokenKind::Kw(Kw::Comma)) { break; }
            if self.peek().is_stmt_end() {
                // A comma needs another variable after it.
                let span = self.current_span();
                self.report_expected(span, "variable");
            }
        }
        names
    }

    /// The first name closes this loop; each further name closes the next enclosing `For`,
    /// provided the `For` blocks are directly nested. Returns the node of the first name.
    fn close_next_names(&mut self, names: Vec<(Option<u32>, Span, NodeId)>) -> Option<NodeId> {
        let directly_nested = self.blocks.iter().rev().take_while(|b| matches!(b, Block::For(_))).count();
        let mut names = names.into_iter();
        let mut mine = None;
        if let Some((name, span, node)) = names.next() {
            self.check_next_variable(name, span);
            mine = Some(node);
        }
        for (i, (name, span, node)) in names.enumerate() {
            if i + 1 < directly_nested {
                self.next_carry.push((name, span, node));
            } else {
                self.diagnostics.push(ERR_NEXT_WITHOUT_FOR, span);
            }
        }
        mine
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
        let do_span = self.current_span();
        self.advance(); // consume Do
        let pre = self.parse_do_condition(arena, DoKind::PreWhile, DoKind::PreUntil);
        if let Some((_, cond)) = pre {
            self.check_condition_size(arena, cond, do_span);
        }
        self.end_stmt();
        if self.line_scan {
            let body = self.parse_block(arena);
            return Self::alloc_do(arena, pre, body);
        }
        self.blocks.push(Block::Do { pre_test: pre.is_some() });
        let body = self.parse_block(arena);
        let loop_span = self.current_span();
        let post = if self.close_block(ERR_DO_WITHOUT_LOOP) {
            self.parse_do_loop_condition(arena, pre.is_some(), loop_span)
        } else {
            None
        };
        self.blocks.pop();
        Self::alloc_do(arena, post.or(pre), body)
    }

    /// The `While` or `Until` condition of a `Do` or a `Loop`, if there is one.
    fn parse_do_condition(&mut self, arena: &mut ExprArena, while_kind: DoKind, until_kind: DoKind) -> Option<(DoKind, NodeId)> {
        let kind = match self.peek().kind {
            TokenKind::Kw(Kw::While) => while_kind,
            TokenKind::Kw(Kw::Until) => until_kind,
            _ => return None,
        };
        self.advance();
        Some((kind, self.parse_expr(arena, 0)))
    }

    /// The condition after `Loop`, and the end of its statement.
    fn parse_do_loop_condition(&mut self, arena: &mut ExprArena, has_pre_test: bool, loop_span: Span) -> Option<(DoKind, NodeId)> {
        let post = self.parse_do_condition(arena, DoKind::PostWhile, DoKind::PostUntil);
        // A loop tests its condition at the top or at the bottom, not both.
        if post.is_some() && has_pre_test {
            self.diagnostics.push(ERR_LOOP_WITHOUT_DO, loop_span);
        }
        if let Some((_, cond)) = post {
            self.check_condition_size(arena, cond, loop_span);
        }
        self.end_stmt();
        post
    }

    /// The `Do` node of a loop with the given test (none: a loop without a condition).
    fn alloc_do(arena: &mut ExprArena, test: Option<(DoKind, NodeId)>, body: NodeId) -> NodeId {
        let (kind, cond) = match test {
            Some((k, c)) => (k, Some(c)),
            None => (DoKind::Inf, None),
        };
        arena.alloc(ExprNode::Do { kind, cond, body })
    }

    /// Parses a While/Wend loop.
    fn parse_while_stmt(&mut self, arena: &mut ExprArena) -> NodeId {
        let while_span = self.current_span();
        self.advance(); // consume While
        let cond = self.parse_expr(arena, 0);
        self.check_condition_size(arena, cond, while_span);
        self.end_stmt();
        if self.line_scan {
            let body = self.parse_block(arena);
            return arena.alloc(ExprNode::While { cond, body });
        }
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
        if self.line_scan {
            let body = self.parse_block(arena);
            return arena.alloc(ExprNode::With { obj, body });
        }
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

        // Legacy single-line form `If <cond> GoTo <label>` (no `Then`).
        let legacy_goto = self.peek().kind == TokenKind::Kw(Kw::GoTo);
        let has_then = !legacy_goto;
        if has_then {
            self.expect(TokenKind::Kw(Kw::Then), ERR_EXPECTED_THEN);
            self.check_statement_after_then();
        }
        // Single-line If: a `Then` followed immediately by a statement (not Eol),
        // or the legacy `GoTo` form.
        if legacy_goto || (has_then && !self.peek().is_stmt_end()) {
            return self.parse_single_line_if(arena, cond, if_span, not_first_on_line, has_then);
        }
        self.parse_block_if(arena, cond, if_span, not_first_on_line)
    }

    /// What follows `Then`: a `Dim` cannot be the statement after it, and a comment cannot
    /// stand in for the statement.
    fn check_statement_after_then(&mut self) {
        if self.peek().kind == TokenKind::Kw(Kw::Dim) {
            let span = self.current_span();
            self.report_expected(span, "statement");
        }
        if self.peek().kind == TokenKind::Kw(Kw::Rem) {
            let span = self.current_span();
            self.diagnostics.push(ERR_SYNTAX_ERROR, span);
        }
    }

    /// A single-line `If`: the statements after `Then` (or the legacy `GoTo`), and its `Else`.
    fn parse_single_line_if(
        &mut self,
        arena: &mut ExprArena,
        cond: NodeId,
        if_span: Span,
        not_first_on_line: bool,
        has_then: bool,
    ) -> NodeId {
        // `Then` straight followed by a colon opens the statements of the line as a
        // colon-led list: it may not follow other statements, a line number is no jump
        // in it, and an `Else` follows a statement only after a colon.
        let colon_led = has_then && self.peek().kind == TokenKind::Kw(Kw::Colon);
        if colon_led && not_first_on_line {
            self.diagnostics.push(ERR_MUST_BE_FIRST_ON_LINE, if_span);
        }
        let stmts = self.parse_inline_statements(arena, InlineList::Then { colon_led });
        let then_body = arena.alloc(ExprNode::Block { stmts });
        let else_body = if !self.last_was_eol && self.eat(TokenKind::Kw(Kw::Else)) {
            let stmts = self.parse_inline_statements(arena, InlineList::Else { colon_led });
            Some(arena.alloc(ExprNode::Block { stmts }))
        } else { None };
        arena.alloc(ExprNode::If { cond, then_body, else_body })
    }

    /// A block `If` (`Then` ends the line), which must start its line.
    fn parse_block_if(&mut self, arena: &mut ExprArena, cond: NodeId, if_span: Span, not_first_on_line: bool) -> NodeId {
        if not_first_on_line {
            self.diagnostics.push(ERR_MUST_BE_FIRST_ON_LINE, if_span);
        }
        if self.in_inline_if {
            // A block `If` cannot be opened from inside a single-line `If`.
            let span = self.current_span();
            self.diagnostics.push(ERR_SYNTAX_ERROR, span);
        }
        self.check_condition_size(arena, cond, if_span);
        self.end_stmt();
        if self.line_scan {
            let then_body = self.parse_block(arena);
            return arena.alloc(ExprNode::If { cond, then_body, else_body: None });
        }
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
    fn parse_inline_statements(&mut self, arena: &mut ExprArena, part: InlineList) -> Vec<NodeId> {
        let was_inline = std::mem::replace(&mut self.in_inline_if, true);
        let was_pending = self.inline_else_pending;
        match part {
            InlineList::Then { colon_led: false } => self.inline_else_pending += 1,
            InlineList::Then { colon_led: true } => self.inline_else_pending = 0,
            InlineList::Else { .. } | InlineList::ElseIf { .. } => {}
        }
        let stmts = self.parse_inline_statement_list(arena, part);
        self.in_inline_if = was_inline;
        self.inline_else_pending = was_pending;
        stmts
    }

    fn parse_inline_statement_list(&mut self, arena: &mut ExprArena, part: InlineList) -> Vec<NodeId> {
        let mut stmts = Vec::new();
        // `If c Then 10`: a line number straight after `Then` / `Else` is a jump, unless the
        // list is colon-led.
        if part.jump_allowed() && matches!(self.peek().kind, TokenKind::IntLit | TokenKind::LongLit) {
            stmts.push(self.parse_inline_jump(arena));
        }
        loop {
            while self.eat(TokenKind::Kw(Kw::Colon)) {}
            if self.peek().kind == TokenKind::Kw(Kw::Else)
                && !self.last_was_eol
                && !matches!(part, InlineList::Then { .. })
                && self.inline_else_pending == 0
            {
                // Only the statements of a `Then` (or of an enclosing one) are followed by an
                // `Else`.
                self.error_and_skip_line(ERR_SYNTAX_ERROR);
                break;
            }
            if self.inline_list_ends(!stmts.is_empty()) {
                break;
            }
            let opener = self.peek().kind.clone();
            let opener_span = self.current_span();
            let eols_before = self.eols_read;
            if let Some(n) = self.parse_stmt(arena) {
                stmts.push(n);
            }
            self.check_block_opened_inline(&opener, opener_span, eols_before);
            if self.last_was_eol {
                break;
            }
        }
        stmts
    }

    /// A line number straight after `Then` / `Else`, read as a jump to that line.
    fn parse_inline_jump(&mut self, arena: &mut ExprArena) -> NodeId {
        let tok = self.advance();
        let target = LabelRef::Line(line_label_value(&tok.lit));
        self.label_refs.push((target, tok.span));
        let jump = arena.alloc(ExprNode::GoTo { target });
        self.set_span(jump, tok.span);
        // The jump is a statement like any other: a separator must follow it.
        self.end_stmt();
        jump
    }

    /// Whether the statements of a part of a single-line `If` end here.
    fn inline_list_ends(&mut self, has_statements: bool) -> bool {
        self.last_was_eol
            || self.peek().is_stmt_end()
            || self.peek().kind == TokenKind::Kw(Kw::Else)
            || (has_statements && self.peek().kind == TokenKind::Kw(Kw::Rem))
    }

    /// A loop, `With` or `Select Case` opened here is closed when the line ends, so
    /// its terminator on a later line has nothing to close.
    fn check_block_opened_inline(&mut self, opener: &TokenKind, opener_span: Span, eols_before: u32) {
        if self.eols_read > eols_before + 1
            && matches!(
                opener,
                TokenKind::Kw(Kw::For) | TokenKind::Kw(Kw::Do) | TokenKind::Kw(Kw::While)
                    | TokenKind::Kw(Kw::With) | TokenKind::Kw(Kw::Select)
            )
        {
            self.diagnostics.push(ERR_END_IF_WITHOUT_BLOCK_IF, opener_span);
        }
    }

    /// Reports an `Else` or `ElseIf` that follows a colon.
    fn check_else_first_on_line(&mut self) {
        if self.after_colon {
            let span = self.current_span();
            self.diagnostics.push(ERR_MUST_BE_FIRST_ON_LINE, span);
        }
    }

    fn parse_else_clause(&mut self, arena: &mut ExprArena) -> Option<NodeId> {
        match self.peek().kind.clone() {
            TokenKind::Kw(Kw::Else) => {
                self.check_else_first_on_line();
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
            TokenKind::Kw(Kw::ElseIf) => Some(self.parse_elseif_clause(arena)),
            _ => None,
        }
    }

    /// An `ElseIf` branch of a block `If`; the rest of the `If` is its else part.
    fn parse_elseif_clause(&mut self, arena: &mut ExprArena) -> NodeId {
        self.check_else_first_on_line();
        let elseif_span = self.current_span();
        self.advance();
        let cond = self.parse_expr(arena, 0);
        self.expect(TokenKind::Kw(Kw::Then), ERR_EXPECTED_THEN);
        self.check_condition_size(arena, cond, elseif_span);
        if self.peek().kind == TokenKind::Kw(Kw::Rem) {
            let span = self.current_span();
            self.diagnostics.push(ERR_SYNTAX_ERROR, span);
        }
        // Statements written after `Then` on the `ElseIf` line are read as those of a
        // single-line `If`; the lines that follow belong to the same branch.
        let then_body = if self.peek().is_stmt_end() {
            self.eat_eol();
            self.parse_block(arena)
        } else {
            let colon_led = self.peek().kind == TokenKind::Kw(Kw::Colon);
            let mut stmts = self.parse_inline_statements(arena, InlineList::ElseIf { colon_led });
            let rest = self.parse_block(arena);
            if let ExprNode::Block { stmts: more } = arena.get(rest) {
                stmts.extend(more.iter().copied());
            }
            arena.alloc(ExprNode::Block { stmts })
        };
        let else_body = self.parse_else_clause(arena);
        arena.alloc(ExprNode::If { cond, then_body, else_body })
    }

    /// Parses a Select Case statement.
    fn parse_select_stmt(&mut self, arena: &mut ExprArena) -> NodeId {
        let select_span = self.current_span();
        self.advance(); // consume Select
        self.expect(TokenKind::Kw(Kw::Case), ERR_EXPECTED_CASE);
        let subject = self.parse_expr(arena, 0);
        self.check_condition_size(arena, subject, select_span);
        self.end_stmt();
        if self.line_scan {
            return arena.alloc(ExprNode::SelectCase { subject, pre: Vec::new(), cases: Vec::new() });
        }
        self.blocks.push(Block::Select);
        let mut pre = Vec::new();
        let mut cases = Vec::new();
        let mut seen_else = false;
        loop {
            while self.eat(TokenKind::Eol) {}
            if !self.parse_select_part(arena, &mut pre, &mut cases, &mut seen_else) {
                break;
            }
        }
        self.blocks.pop();
        arena.alloc(ExprNode::SelectCase { subject, pre, cases })
    }

    /// One part of the body of a `Select Case`: a `Case` clause, the end of the block, or
    /// a statement before the first `Case`. Returns false when the block has ended.
    fn parse_select_part(
        &mut self,
        arena: &mut ExprArena,
        pre: &mut Vec<NodeId>,
        cases: &mut Vec<NodeId>,
        seen_else: &mut bool,
    ) -> bool {
        match self.peek().kind.clone() {
            TokenKind::Kw(Kw::Case) => {
                let case_span = self.current_span();
                self.advance();
                if *seen_else {
                    // `Case Else` is the last clause.
                    self.diagnostics.push(ERR_CASE_WITHOUT_SELECT, case_span);
                }
                let clause = self.parse_case_clause(arena);
                *seen_else |= matches!(arena.get(clause), ExprNode::CaseElse { .. });
                cases.push(clause);
                true
            }
            TokenKind::Kw(Kw::End) if self.terminator() == Some(Term::EndSelect) => {
                self.close_block(ERR_SELECT_WITHOUT_END_SELECT);
                self.end_stmt();
                false
            }
            TokenKind::Eof => false,
            _ if self.at_proc_end() => {
                self.report_unclosed(ERR_SELECT_WITHOUT_END_SELECT);
                false
            }
            _ => {
                self.parse_select_preamble_part(arena, pre);
                true
            }
        }
    }

    /// Statements before the first Case clause.
    fn parse_select_preamble_part(&mut self, arena: &mut ExprArena, pre: &mut Vec<NodeId>) {
        if self.peek().is_stmt_end() {
            self.advance();
            return;
        }
        if self.peek().kind == TokenKind::Kw(Kw::Loop)
            && matches!(self.peek2().kind, TokenKind::Kw(Kw::While) | TokenKind::Kw(Kw::Until))
        {
            // A `Loop` that carries a condition is a statement like any other here, not
            // a stray block end.
            self.error_and_skip_line(ERR_STATEMENT_BEFORE_FIRST_CASE);
            return;
        }
        let span = self.current_span();
        // An attribute line is read like anywhere else and is no statement.
        let attribute_line = self.at_attribute_line();
        let first = self.peek().kind.clone();
        let mark = self.diagnostics.len();
        let Some(stmt) = self.parse_stmt(arena) else { return };
        if attribute_line {
            return;
        }
        match self.preamble_rule(arena, &first, stmt) {
            PreambleRule::Allowed => {}
            PreambleRule::Reported => {
                self.diagnostics.push(ERR_STATEMENT_BEFORE_FIRST_CASE, span);
            }
            PreambleRule::ReportedInstead => {
                self.diagnostics.truncate(mark);
                self.diagnostics.push(ERR_STATEMENT_BEFORE_FIRST_CASE, span);
            }
        }
        pre.push(stmt);
    }

    /// What a statement parsed between `Select Case` and the first `Case` may be. Declarations
    /// (`Dim`, `Static`, `Const`, `ReDim`), a `Close` with no channel, an `Exit For` and a
    /// `Circle`, `Line`, `PSet` or `Scale` method of an object are allowed. An `Exit` of any other
    /// kind is reported as a statement before the first `Case` in place of its own error; every
    /// other statement is reported on top of whatever it reports itself.
    fn preamble_rule(&self, arena: &ExprArena, first: &TokenKind, stmt: NodeId) -> PreambleRule {
        if matches!(
            first,
            TokenKind::Kw(Kw::Dim) | TokenKind::Kw(Kw::Static) | TokenKind::Kw(Kw::Const) | TokenKind::Kw(Kw::ReDim)
        ) {
            return PreambleRule::Allowed;
        }
        match arena.get(stmt) {
            ExprNode::FileIoStmt { kind: FileIoKind::Close, channel: None, args, .. } if args.is_empty() => {
                PreambleRule::Allowed
            }
            ExprNode::ExitStmt { kind: ExitKind::For } => PreambleRule::Allowed,
            ExprNode::ExitStmt { .. } => PreambleRule::ReportedInstead,
            // A drawing method of an object is not tested; any other statement is.
            _ => match self.graphics_stmts.get(&stmt.0) {
                Some((word, true)) if word.draws() => PreambleRule::Allowed,
                _ => PreambleRule::Reported,
            },
        }
    }

    fn parse_case_clause(&mut self, arena: &mut ExprArena) -> NodeId {
        if self.eat(TokenKind::Kw(Kw::Else)) {
            self.end_stmt();
            let body = self.parse_block(arena);
            arena.alloc(ExprNode::CaseElse { body })
        } else {
            let case_span = self.current_span();
            let test = self.parse_case_test_expr(arena);
            // A `Case` line costs what a condition does: each test is its expression, a range
            // adds its operator and the tests are separated.
            if let ExprNode::ArgList { args } = arena.get(test) {
                let parts: Option<usize> = args.iter().enumerate().try_fold(0, |total, (i, arg)| {
                    let part = match arena.get(*arg) {
                        ExprNode::CaseIs { rhs, .. } => line_size::expression_part(arena, *rhs)?,
                        ExprNode::RangeTo { lo, hi } => {
                            line_size::expression_part(arena, *lo)? + line_size::expression_part(arena, *hi)? + cost::OPERATOR
                        }
                        _ => line_size::expression_part(arena, *arg)?,
                    };
                    Some(total + part + if i > 0 { cost::ITEM_SEPARATOR } else { 0 })
                });
                self.check_parts_size(parts, case_span);
            }
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
            self.error_and_skip_line(ERR_SYNTAX_ERROR);
            return arena.alloc(ExprNode::Block { stmts: Vec::new() });
        }
        if self.at_reserved_word() && self.peek().kind != TokenKind::Kw(Kw::Me) {
            self.reject_reserved_word();
            self.skip_line();
            return arena.alloc(ExprNode::Block { stmts: Vec::new() });
        }
        let start = self.current_span();
        self.graphics_chain = Some(ChainStart::Call);
        self.graphics_stop = None;
        let callee = self.parse_target_expr(arena);
        self.graphics_chain = None;
        if let Some(word) = self.graphics_stop.take() {
            return self.finish_graphics_method(arena, callee, word, ChainStart::Call, start);
        }
        if self.call_target_is_malformed(arena, callee) {
            let span = self.chain_span(arena, callee);
            self.diagnostics.push(ERR_SYNTAX_ERROR, span);
            self.skip_line();
            return arena.alloc(ExprNode::Block { stmts: Vec::new() });
        }
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
        self.end_stmt_before_rem_is_error();
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
        if self.line_scan {
            return;
        }
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
                self.proc_kind.is_none_or(|k| matches!(k, ProcKind::Sub | ProcKind::PropGet)),
                ERR_EXIT_SUB_NOT_ALLOWED,
            ),
            ExitKind::Function => (
                self.proc_kind.is_none_or(|k| matches!(k, ProcKind::Function | ProcKind::PropGet)),
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
        let spaced = self.advance().kind == TokenKind::Kw(Kw::Go); // consume GoTo / GoSub / Go
        // Space-separated forms: a bare `Go` is followed by `To` or `Sub`, which
        // determines the jump kind (the caller has checked that one of them follows).
        let is_gosub = if !spaced {
            is_gosub
        } else {
            self.advance().kind == TokenKind::Kw(Kw::Sub)
        };
        let target_span = self.current_span();
        let target = if matches!(self.peek().kind, TokenKind::IntLit | TokenKind::LongLit) {
            let tok = self.advance();
            LabelRef::Line(line_label_value(&tok.lit))
        } else {
            self.bracketed_name_ok = true;
            let name = self.expect_decl_name();
            self.bracketed_name_ok = false;
            LabelRef::Name(name.sym.map(|s| s as u32).unwrap_or(0))
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
            self.bracketed_name_ok = true;
            let lbl = self.expect_decl_name();
            self.bracketed_name_ok = false;
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
        if self.eat(TokenKind::Kw(Kw::Local)) && !self.at_word(Kw::Error) {
            // `Local` only qualifies `Error`.
            self.report_expected_and_skip_line("Error");
            return arena.alloc(ExprNode::OnError { kind: OnErrorKind::Disable });
        }
        if self.eat_word(Kw::Error) {
            return self.parse_on_error_stmt(arena);
        }
        self.parse_on_go_stmt(arena)
    }

    /// `On Error Resume Next` or `On Error GoTo target`, after `On Error`.
    fn parse_on_error_stmt(&mut self, arena: &mut ExprArena) -> NodeId {
        if self.eat(TokenKind::Kw(Kw::Resume)) {
            self.expect(TokenKind::Kw(Kw::Next), ERR_EXPECTED_NEXT);
            self.end_stmt();
            return arena.alloc(ExprNode::OnError { kind: OnErrorKind::ResumeNext });
        }
        if self.eat(TokenKind::Kw(Kw::GoTo)) {
            let kind = self.parse_on_error_goto_kind();
            return arena.alloc(ExprNode::OnError { kind });
        }
        // After `Error` only `Resume Next` and `GoTo` follow.
        self.error_and_skip_line(ERR_UNEXPECTED_TOKEN);
        arena.alloc(ExprNode::OnError { kind: OnErrorKind::Disable })
    }

    /// The target of `On Error GoTo`.
    fn parse_on_error_goto_kind(&mut self) -> OnErrorKind {
        // `On Error GoTo -1` leaves the error state of the handler.
        if self.peek().kind == TokenKind::Kw(Kw::Minus)
            && matches!(self.peek2().kind, TokenKind::IntLit | TokenKind::LongLit)
            && line_label_value(&self.peek2().lit) == 1
        {
            self.advance();
            self.advance();
            self.end_stmt();
            return OnErrorKind::ClearError;
        }
        if matches!(self.peek().kind, TokenKind::IntLit | TokenKind::LongLit) {
            let tok = self.advance();
            let line = line_label_value(&tok.lit);
            self.end_stmt();
            // Only `On Error GoTo 0` disables the handler; any nonzero
            // line number installs a handler at that numeric line label.
            if line == 0 {
                return OnErrorKind::Disable;
            }
            self.label_refs.push((LabelRef::Line(line), tok.span));
            return OnErrorKind::Goto(LabelRef::Line(line));
        }
        let lbl = self.expect_ident();
        let lid = lbl.sym.map(|s| s as u32).unwrap_or(0);
        self.label_refs.push((LabelRef::Name(lid), lbl.span));
        self.end_stmt();
        OnErrorKind::Goto(LabelRef::Name(lid))
    }

    /// `On <expr> GoTo / GoSub label-list`
    fn parse_on_go_stmt(&mut self, arena: &mut ExprArena) -> NodeId {
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
            labels.push(self.parse_jump_target(arena));
            if !self.eat(TokenKind::Kw(Kw::Comma)) { break; }
        }
        self.end_stmt();
        arena.alloc(ExprNode::OnGo { is_gosub, expr, labels })
    }

    /// One target of an `On ... GoTo` list: a numeric line label or a named label.
    fn parse_jump_target(&mut self, arena: &mut ExprArena) -> NodeId {
        let (target, span) = if matches!(self.peek().kind, TokenKind::IntLit | TokenKind::LongLit) {
            let tok = self.advance();
            (LabelRef::Line(line_label_value(&tok.lit)), tok.span)
        } else {
            let lbl = self.expect_ident();
            (LabelRef::Name(lbl.sym.map(|s| s as u32).unwrap_or(0)), lbl.span)
        };
        self.label_refs.push((target, span));
        let node = arena.alloc(ExprNode::JumpTarget { target });
        self.set_span(node, span);
        node
    }

    fn parse_erase_stmt(&mut self, arena: &mut ExprArena) -> NodeId {
        let erase_span = self.current_span();
        self.advance(); // consume Erase
        let mut vars = Vec::new();
        loop {
            vars.push(self.parse_target_expr(arena));
            if !self.eat(TokenKind::Kw(Kw::Comma)) { break; }
        }
        // `Erase a, b` costs 2 for each array name, a separator between them and 2 over a store.
        if vars.iter().all(|v| matches!(arena.get(*v), ExprNode::NameRef { .. })) {
            let parts = cost::ERASE + vars.len() * cost::ARRAY_NAME + (vars.len() - 1) * cost::ITEM_SEPARATOR;
            self.check_parts_size(Some(parts), erase_span);
        }
        self.end_stmt_before_rem_is_error();
        arena.alloc(ExprNode::Erase { vars })
    }

    fn parse_raise_event(&mut self, arena: &mut ExprArena) -> NodeId {
        if self.module_kind == ModuleKind::Standard {
            // Events belong to object modules; elsewhere the word starts nothing.
            self.error_and_skip_line(ERR_SYNTAX_ERROR);
            return arena.alloc(ExprNode::Block { stmts: Vec::new() });
        }
        self.advance(); // consume RaiseEvent
        let name_span = self.current_span();
        // The event is a (possibly qualified) name with an optional bracketed argument list.
        let target = self.parse_target_expr(arena);
        let (callee, args) = match arena.get(target) {
            ExprNode::Call { func, args } => (*func, Some(*args)),
            _ => (target, None),
        };
        let (nid, qualifier) = match arena.get(callee) {
            ExprNode::NameRef { sym, .. } => (*sym, None),
            ExprNode::MemberAccess { base, member, bang: false, .. } => (*member, Some(*base)),
            _ => {
                self.report_expected_and_skip_line("end of statement");
                return arena.alloc(ExprNode::Block { stmts: Vec::new() });
            }
        };
        let args = args.unwrap_or_else(|| arena.alloc(ExprNode::ArgList { args: vec![] }));
        // The arguments of an event are positional: there is no `name:=value` form.
        let named = matches!(arena.get(args), ExprNode::ArgList { args }
            if args.iter().any(|a| matches!(arena.get(*a), ExprNode::NamedArg { .. })));
        if named {
            self.diagnostics.push(ERR_SYNTAX_ERROR, name_span);
            self.skip_line();
            return arena.alloc(ExprNode::Block { stmts: Vec::new() });
        }
        self.end_stmt();
        let id = arena.alloc(ExprNode::RaiseEvent { name: nid, qualifier, args });
        self.set_span(id, name_span);
        id
    }

    fn parse_debug_stmt(&mut self, arena: &mut ExprArena) -> NodeId {
        let debug_span = self.current_span();
        self.advance(); // consume Debug
        // The dot follows `Debug` directly: blanks between them make a different statement.
        if self.blank_before_next() {
            self.diagnostics.push(ERR_SYNTAX_ERROR, debug_span);
            self.skip_line();
            return arena.alloc(ExprNode::DebugPrint { channel: None, args: Vec::new() });
        }
        self.expect(TokenKind::Kw(Kw::Dot), ERR_UNEXPECTED_TOKEN);
        // `Debug.?` is `Debug.Print`.
        let is_print = matches!(self.peek().kind, TokenKind::Kw(Kw::Print) | TokenKind::Kw(Kw::Question));
        let is_member = is_print
            || self.peek().sym.is_some_and(|sym| self.scanner.sym_name(sym as u32).eq_ignore_ascii_case("Assert"));
        if !is_member {
            self.report_expected_and_skip_line("Print or Assert");
            return arena.alloc(ExprNode::DebugPrint { channel: None, args: Vec::new() });
        }
        self.advance(); // consume Print / ? / Assert
        if !is_print {
            // `Debug.Assert` takes exactly one expression.
            let cond = self.parse_expr(arena, 0);
            self.end_stmt();
            return arena.alloc(ExprNode::DebugAssert { cond });
        }
        // `Debug.Print [#channel,] items`: a channel is followed by its comma.
        let mut channel = None;
        let mut list_follows = true;
        if self.eat(TokenKind::Kw(Kw::Hash)) {
            channel = Some(self.parse_expr(arena, 0));
            list_follows = self.expect_comma();
        }
        let mut args = Vec::new();
        if list_follows {
            self.parse_print_method_items(arena, &mut args);
        }
        // `Debug.Print` of expressions costs 2 more than an assignment of the first.
        if channel.is_none() && self.output_list_too_complex(arena, &args, cost::DEBUG_PRINT) {
            self.diagnostics.push(ERR_STATEMENT_TOO_COMPLEX, debug_span);
        }
        self.end_stmt();
        arena.alloc(ExprNode::DebugPrint { channel, args })
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
        let t = match self.cc_pending.take() {
            Some(t) => t,
            None => self.next_physical(),
        };
        self.cc_line_done = matches!(t.kind, TokenKind::Eol | TokenKind::Eof);
        t
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
            self.cc_set(&fold_name(&name), value);
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
            // Only a comment mark may follow a directive: neither `Rem` nor a colon does.
            TokenKind::Kw(Kw::Apos) => {
                self.cc_advance();
                if self.cc_peek().kind == TokenKind::Eol {
                    self.cc_advance();
                }
            }
            _ => {
                let span = self.cc_peek().span;
                self.report_expected(span, "end of statement");
                self.cc_skip_line();
            }
        }
    }

    /// `#If` / `#ElseIf` condition: evaluated only when the branch can still
    /// be taken; the expression is followed by `Then`.
    fn cc_condition(&mut self, evaluate: bool, span: Span) -> bool {
        let truth = self.eval_cc_expr(0).and_then(|v| v.is_true());
        let value = match truth {
            Ok(b) => b && evaluate,
            Err(code) => {
                // The expression raised an error: it is reported on the directive.
                self.diagnostics.push(code, span);
                self.cc_skip_line();
                return false;
            }
        };
        if self.cc_peek().kind == TokenKind::Kw(Kw::Then) {
            self.cc_advance();
            self.cc_end_of_directive();
        } else {
            let span = self.cc_peek().span;
            self.report_expected(span, "Then");
            self.cc_skip_line();
        }
        value
    }

    /// Handle the directive token `kind` (already consumed). Returns after the
    /// directive's whole line has been consumed.
    fn cc_directive(&mut self, kind: Kw, span: Span, one_word: bool) {
        let was_active = self.cc_active();
        let diag_len = self.diagnostics.len();
        let saved = (self.cc_stack.clone(), self.cc_defs.clone(), self.cc_module_consts.clone());
        self.cc_directive_line(kind, span, one_word);
        let mut found = self.diagnostics.split_off(diag_len);
        // A directive that does not read completely is not a directive: nothing it did
        // counts, and it is an ordinary line, whose errors count where code is compiled
        // and which is ignored where it is not.
        let malformed = found.iter().any(|d| matches!(d.code, 0x9c6f | 0x9c70));
        if malformed {
            (self.cc_stack, self.cc_defs, self.cc_module_consts) = saved;
            if !self.cc_line_done {
                self.cc_skip_line();
            }
            // The line stops at its first error, and is not read at all where code is not compiled.
            found.truncate(if was_active { 1 } else { 0 });
            self.diagnostics.extend(found);
        } else {
            // What a directive line reports is found before any statement is read.
            for diagnostic in found {
                self.diagnostics.push_directive_pass(diagnostic);
            }
        }
    }

    fn cc_directive_line(&mut self, kind: Kw, span: Span, one_word: bool) {
        match kind {
            Kw::CcConst => self.cc_const_directive(span),
            Kw::CcIf => {
                let parent_active = self.cc_active();
                let taken = self.cc_condition(parent_active, span);
                // The compiler holds at most 4079 open blocks.
                if self.cc_stack.len() >= MAX_CC_NESTING {
                    self.diagnostics.push(ERR_OUT_OF_MEMORY, span);
                }
                self.cc_stack.push(CcFrame { else_seen: false, parent_active, active: taken, taken });
            }
            Kw::CcElseIf => self.cc_elseif_directive(span),
            Kw::CcElse => self.cc_else_directive(span),
            // `#End If`, or the one-word `#EndIf`.
            _ => self.cc_end_directive(span, one_word),
        }
    }

    /// Whether the token can be the name of a `#Const`.
    fn cc_const_name_ok(&self, name_tok: &Token) -> bool {
        name_tok.kind == TokenKind::Ident
            || (matches!(name_tok.kind, TokenKind::Kw(_))
                && name_tok.sym.is_some_and(|sym| {
                    let sym = sym as u32;
                    !self.scanner.is_reserved(sym) && !is_header_word(self.scanner.sym_name(sym))
                }))
    }

    /// `#Const name = value`, after the directive word.
    fn cc_const_directive(&mut self, span: Span) {
        let name_tok = self.cc_peek().clone();
        if !self.cc_const_name_ok(&name_tok) {
            self.report_expected(name_tok.span, "identifier");
            self.cc_skip_line();
            return;
        }
        self.cc_advance();
        let name = name_tok
            .sym
            .map(|id| fold_name(self.scanner.sym_name(id as u32)))
            .unwrap_or_default();
        if self.cc_peek().kind == TokenKind::Kw(Kw::Eq) {
            self.cc_advance();
        } else {
            let span = self.cc_peek().span;
            self.report_expected(span, "=");
        }
        let value = match self.eval_cc_expr(0) {
            Ok(v) => v,
            Err(code) => {
                self.diagnostics.push(code, span);
                self.cc_skip_line();
                return;
            }
        };
        // A `#Const` is defined whether or not its block is active.
        if self.cc_module_consts.contains(&name) {
            self.diagnostics.push(ERR_CC_DUPLICATE_DEFINITION, span);
        } else {
            self.cc_module_consts.push(name.clone());
        }
        self.cc_set(&name, value);
        self.cc_end_of_directive();
    }

    /// `#ElseIf condition Then`, after the directive word.
    fn cc_elseif_directive(&mut self, span: Span) {
        let Some(frame) = self.cc_stack.last().copied() else {
            // No `#If` is open: the lone `#ElseIf` is reported as one that follows an `#Else`.
            self.diagnostics.push(ERR_CC_ELSEIF_AFTER_ELSE, span);
            self.cc_skip_line();
            return;
        };
        if frame.else_seen {
            self.diagnostics.push(ERR_CC_ELSEIF_AFTER_ELSE, span);
            self.cc_skip_line();
            return;
        }
        let evaluate = frame.parent_active && !frame.taken;
        // A branch that cannot be taken is still read as a directive.
        let cond = self.cc_condition(evaluate, span);
        let top = self.cc_stack.last_mut().unwrap();
        top.active = cond;
        top.taken |= cond;
    }

    /// `#Else`, after the directive word.
    fn cc_else_directive(&mut self, span: Span) {
        if self.cc_active() || self.cc_stack.last().is_some_and(|f| f.parent_active) {
            self.cc_end_of_directive();
        } else {
            self.cc_skip_line();
        }
        match self.cc_stack.last_mut() {
            // No `#If` is open: the lone `#Else` is reported as a repeated one.
            None => self.diagnostics.push(ERR_CC_DUPLICATE_ELSE, span),
            Some(top) if top.else_seen => self.diagnostics.push(ERR_CC_DUPLICATE_ELSE, span),
            Some(top) => {
                top.else_seen = true;
                top.active = top.parent_active && !top.taken;
                top.taken = true;
            }
        }
    }

    /// `#End If`, or the one-word `#EndIf`, after the directive word.
    fn cc_end_directive(&mut self, span: Span, one_word: bool) {
        if !one_word {
            self.cc_read_if_after_end();
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

    /// The `If` of `#End If`; reported when missing where the block is read.
    fn cc_read_if_after_end(&mut self) {
        if self.cc_peek().kind == TokenKind::Kw(Kw::If) {
            self.cc_advance();
        } else if self.cc_stack.last().is_some_and(|f| f.parent_active) {
            let span = self.cc_peek().span;
            self.report_expected(span, "If");
        }
    }

    /// At the end of the module, report the innermost `#If` that was never
    /// closed (once).
    fn report_open_cc_blocks(&mut self) {
        if self.cc_stack.last().is_some() {
            // Reported once, on the last line of the module.
            let text = &self.src[..];
            let end = text.iter().rposition(|b| !matches!(b, b'\r' | b'\n')).map_or(0, |i| i + 1);
            let start = text[..end].iter().rposition(|b| *b == b'\n').map_or(0, |i| i + 1);
            let span = Span { start: start as u32, len: (end - start) as u32 };
            self.diagnostics.push_directive_pass(Diagnostic::new(ERR_CC_IF_NOT_TERMINATED, span));
            self.cc_stack.clear();
        }
    }

    // ── CC expression evaluator ───────────────────────────────────────────────

    /// Look up a CC constant by (lowercase) name; a name with no definition is the Integer 0.
    fn cc_lookup(&self, name: &str) -> CcValue {
        self.cc_defs.iter().rev()
            .find(|(n, _)| n == name)
            .map(|(_, v)| v.clone())
            .unwrap_or(CcValue::Integer(0))
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
    fn eval_cc_prefix(&mut self) -> CcResult {
        match self.cc_peek().kind.clone() {
            TokenKind::Kw(Kw::Not) => {
                self.cc_advance();
                self.eval_cc_expr(CC_NOT_BP)?.not()
            }
            TokenKind::Kw(Kw::Minus) => {
                self.cc_advance();
                self.eval_cc_expr(CC_SIGN_BP)?.negate()
            }
            TokenKind::Kw(Kw::Plus) => {
                self.cc_advance();
                self.eval_cc_expr(CC_SIGN_BP)?.plus()
            }
            _ => self.eval_cc_primary(),
        }
    }

    /// Binding powers + operator for a CC infix operator, or `None`. From the loosest:
    /// `Imp`, `Eqv`, `Xor`, `Or`, `And`, (`Not`), comparisons, `&`, `+ -`, `Mod`, `\`,
    /// `* /`, `^`.
    fn cc_infix_op(&mut self) -> Option<(u8, u8, CcOp)> {
        let entry = match self.cc_peek().kind {
            TokenKind::Kw(Kw::Imp)        => (2,  3,  CcOp::Imp),
            TokenKind::Kw(Kw::Eqv)        => (4,  5,  CcOp::Eqv),
            TokenKind::Kw(Kw::Xor)        => (6,  7,  CcOp::Xor),
            TokenKind::Kw(Kw::Or)         => (8,  9,  CcOp::Or),
            TokenKind::Kw(Kw::And)        => (10, 11, CcOp::And),
            TokenKind::Kw(Kw::Eq)         => (14, 15, CcOp::Eq),
            TokenKind::Kw(Kw::Ne)         => (14, 15, CcOp::Ne),
            TokenKind::Kw(Kw::Lt)         => (14, 15, CcOp::Lt),
            TokenKind::Kw(Kw::Gt)         => (14, 15, CcOp::Gt),
            TokenKind::Kw(Kw::Le)         => (14, 15, CcOp::Le),
            TokenKind::Kw(Kw::Ge)         => (14, 15, CcOp::Ge),
            TokenKind::Kw(Kw::Like)       => (14, 15, CcOp::Like),
            TokenKind::Kw(Kw::Is)         => (14, 15, CcOp::Is),
            TokenKind::Kw(Kw::Amp)        => (16, 17, CcOp::Cat),
            TokenKind::Kw(Kw::Plus)       => (18, 19, CcOp::Add),
            TokenKind::Kw(Kw::Minus)      => (18, 19, CcOp::Sub),
            TokenKind::Kw(Kw::Mod)        => (20, 21, CcOp::Mod),
            TokenKind::Kw(Kw::Backslash)  => (22, 23, CcOp::IDiv),
            TokenKind::Kw(Kw::Star)       => (24, 25, CcOp::Mul),
            TokenKind::Kw(Kw::Slash)      => (24, 25, CcOp::Div),
            TokenKind::Kw(Kw::Caret)      => (28, 29, CcOp::Pow),
            _ => return None,
        };
        Some(entry)
    }

    /// Evaluate a CC expression.
    fn eval_cc_expr(&mut self, min_bp: u8) -> CcResult {
        let named_start = self.cc_names_read;
        let mut lhs = self.eval_cc_prefix()?;
        let mut lhs_named = self.cc_names_read != named_start;
        loop {
            let Some((l_bp, r_bp, op)) = self.cc_infix_op() else { break };
            if l_bp <= min_bp { break; }
            self.cc_advance();
            let named_before = self.cc_names_read;
            let rhs = self.eval_cc_expr(r_bp)?;
            let by_name = lhs_named || self.cc_names_read != named_before;
            lhs = CcValue::binary_with(op, &lhs, &rhs, by_name)?;
            lhs_named = by_name;
        }
        Ok(lhs)
    }

    /// Whether the word can stand as a name in a directive: an identifier, or a keyword
    /// that is not reserved (`Me`, `Date` and the like are names too).
    fn cc_word_is_name(&mut self, tok: &Token) -> bool {
        match tok.kind {
            TokenKind::Ident => true,
            TokenKind::Kw(_) => tok.sym.is_some_and(|sym| {
                let sym = sym as u32;
                let name = self.scanner.sym_name(sym);
                !self.scanner.is_reserved(sym)
                    || ["Me", "Date", "Date$", "DoEvents", "CDec"].iter().any(|n| name.eq_ignore_ascii_case(n))
            }),
            _ => false,
        }
    }

    fn eval_cc_primary(&mut self) -> CcResult {
        match self.cc_peek().kind.clone() {
            TokenKind::Kw(Kw::LParen) => {
                self.cc_advance();
                let v = self.eval_cc_expr(0)?;
                self.cc_close_paren();
                Ok(v)
            }
            TokenKind::IntLit
            | TokenKind::LongLit
            | TokenKind::SngLit
            | TokenKind::DblLit
            | TokenKind::DateLit
            | TokenKind::CurLit
            | TokenKind::StrLit => {
                let tok = self.cc_advance();
                Ok(Self::cc_literal_value(tok))
            }
            TokenKind::Kw(Kw::True)  => { self.cc_advance(); Ok(CcValue::TRUE) }
            TokenKind::Kw(Kw::False) => { self.cc_advance(); Ok(CcValue::FALSE) }
            TokenKind::Kw(Kw::Empty) => { self.cc_advance(); Ok(CcValue::Empty) }
            TokenKind::Kw(Kw::Null) => { self.cc_advance(); Ok(CcValue::Null) }
            TokenKind::Kw(Kw::Nothing) => { self.cc_advance(); Ok(CcValue::Nothing) }
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
                self.report_expected(span, "expression");
                Ok(CcValue::Empty)
            }
            _ => self.eval_cc_name(),
        }
    }

    /// Consumes the `)` that closes a bracketed group of a directive, or reports it missing.
    fn cc_close_paren(&mut self) {
        if self.cc_peek().kind == TokenKind::Kw(Kw::RParen) {
            self.cc_advance();
        } else {
            let span = self.cc_peek().span;
            self.diagnostics.push(ERR_EXPECTED_RPAREN, span);
        }
    }

    /// The value of a literal token of a directive.
    fn cc_literal_value(tok: Token) -> CcValue {
        match tok.kind {
            TokenKind::IntLit | TokenKind::LongLit => Self::cc_integer_value(tok),
            TokenKind::SngLit | TokenKind::DblLit | TokenKind::DateLit => Self::cc_float_or_date_value(tok),
            TokenKind::CurLit => match tok.lit {
                Some(Lit::Currency(c)) => CcValue::Currency(c),
                _ => CcValue::Integer(0),
            },
            _ => match tok.lit {
                Some(Lit::Str(s)) => CcValue::Str(s.to_string()),
                _ => CcValue::Str(String::new()),
            },
        }
    }

    fn cc_integer_value(tok: Token) -> CcValue {
        match tok.lit {
            Some(Lit::Int(n)) | Some(Lit::Long(n)) => match i16::try_from(n) {
                Ok(i) if tok.kind == TokenKind::IntLit => CcValue::Integer(i),
                _ => CcValue::Long(n),
            },
            _ => CcValue::Integer(0),
        }
    }

    fn cc_float_or_date_value(tok: Token) -> CcValue {
        match tok.lit {
            Some(Lit::Single(f)) => CcValue::Single(f),
            Some(Lit::Double(d)) => CcValue::Double(d),
            Some(Lit::Date(d)) => CcValue::Date(d),
            _ => CcValue::Integer(0),
        }
    }

    /// A word in a directive expression: a name, or a call of a built-in function.
    fn eval_cc_name(&mut self) -> CcResult {
        let tok = self.cc_advance();
        let is_word = tok.kind == TokenKind::Ident || matches!(tok.kind, TokenKind::Kw(_));
        let Some(id) = tok.sym.filter(|_| is_word) else {
            self.diagnostics.push(ERR_UNEXPECTED_TOKEN, tok.span);
            return Ok(CcValue::Empty);
        };
        let name = self.scanner.sym_name(id as u32).to_string();
        if self.cc_peek().kind == TokenKind::Kw(Kw::LParen) {
            return self.eval_cc_call(&name);
        }
        if CcValue::is_function_name(&name) || !self.cc_word_is_name(&tok) {
            // A function name without its brackets, or a reserved word.
            self.diagnostics.push(ERR_UNEXPECTED_TOKEN, tok.span);
            return Ok(CcValue::Empty);
        }
        self.cc_names_read += 1;
        Ok(self.cc_lookup(&fold_name(&name)))
    }

    /// A call: only the built-in conversion and numeric functions exist.
    fn eval_cc_call(&mut self, name: &str) -> CcResult {
        self.cc_advance();
        let mut args = Vec::new();
        if self.cc_peek().kind != TokenKind::Kw(Kw::RParen) {
            loop {
                args.push(self.eval_cc_expr(0)?);
                if self.cc_peek().kind == TokenKind::Kw(Kw::Comma) {
                    self.cc_advance();
                } else {
                    break;
                }
            }
        }
        self.cc_close_paren();
        CcValue::call(name, &args).unwrap_or(Err(CC_ERR_VARIABLE_NOT_DEFINED))
    }

    // ── Option statement parser ───────────────────────────────────────────────

    fn parse_option_stmt(&mut self, arena: &mut ExprArena) -> NodeId {
        let option_span = self.current_span();
        self.advance(); // consume Option
        // `Private` is a reserved word, so only the bare token is the keyword; the other
        // option words are names that spell their keyword, in brackets or not.
        let private = self.peek().kind == TokenKind::Kw(Kw::Private);
        match if private { None } else { self.peek_word() } {
            Some(Kw::Explicit) => {
                self.note_option_seen(OptionSeen::Explicit, option_span);
                self.advance();
                self.end_stmt();
                arena.alloc(ExprNode::OptionExplicit)
            }
            Some(Kw::Base) => self.parse_option_base(arena, option_span),
            Some(Kw::Compare) => self.parse_option_compare(arena, option_span),
            None if private => self.parse_option_private(arena, option_span),
            _ => {
                self.error_and_skip_line(ERR_UNEXPECTED_TOKEN);
                arena.alloc(ExprNode::Block { stmts: vec![] })
            }
        }
    }

    /// Reports an `Option` statement of a kind that has been read before.
    fn note_option_seen(&mut self, which: OptionSeen, option_span: Span) {
        let seen = match which {
            OptionSeen::Explicit => &mut self.seen_option_explicit,
            OptionSeen::Compare => &mut self.seen_option_compare,
            OptionSeen::Private => &mut self.seen_option_private,
        };
        if std::mem::replace(seen, true) {
            self.diagnostics.push(ERR_DUPLICATE_OPTION, option_span);
        }
    }

    /// `Option Base 0` / `Option Base 1`, at the word `Base`.
    fn parse_option_base(&mut self, arena: &mut ExprArena, option_span: Span) -> NodeId {
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
                self.error_and_skip_line(ERR_UNEXPECTED_TOKEN);
                return arena.alloc(ExprNode::Block { stmts: vec![] });
            }
        };
        self.end_stmt();
        arena.alloc(ExprNode::OptionBase { value })
    }

    /// `Option Compare Binary` / `Option Compare Text`, at the word `Compare`.
    fn parse_option_compare(&mut self, arena: &mut ExprArena, option_span: Span) -> NodeId {
        self.note_option_seen(OptionSeen::Compare, option_span);
        self.advance();
        let mode = if self.eat_word(Kw::Binary) {
            0u8
        } else if self.eat_word(Kw::Text) {
            1u8
        } else {
            let span = self.current_span();
            self.diagnostics.push_labeled(ERR_UNEXPECTED_TOKEN, span, "Text or Binary");
            self.skip_line();
            return arena.alloc(ExprNode::Block { stmts: vec![] });
        };
        self.end_stmt();
        arena.alloc(ExprNode::OptionCompare { mode })
    }

    /// `Option Private Module`, at the word `Private`.
    fn parse_option_private(&mut self, arena: &mut ExprArena, option_span: Span) -> NodeId {
        self.note_option_seen(OptionSeen::Private, option_span);
        self.advance(); // consume Private
        if self.eat_word(Kw::Module) {
            // Only a standard module can be private; that is found once the whole
            // statement has been read.
            if self.module_kind != ModuleKind::Standard {
                self.diagnostics.push(ERR_OPTION_PRIVATE_IN_OBJECT_MODULE, option_span);
            }
            self.end_stmt();
        } else {
            let span = self.current_span();
            self.diagnostics.push_labeled(ERR_UNEXPECTED_TOKEN, span, "Module");
            self.skip_line();
        }
        arena.alloc(ExprNode::Block { stmts: vec![] })
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
            // The list goes on only while a name follows the comma; anything else
            // (a colon, a comment, the end of the line) ends it.
            if !self.at_name() {
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

    /// Whether the next token reads as a name: an identifier, or a word that is also
    /// accepted where a name stands.
    fn at_name(&mut self) -> bool {
        let t = self.peek();
        t.kind == TokenKind::Ident || (matches!(t.kind, TokenKind::Kw(_)) && t.sym.is_some())
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
    /// name `lo` to that of `hi`; empty when either is not a letter. A range
    /// written backwards (`Z-X`) covers the same letters as the forward one.
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
            (Some(a), Some(b)) => {
                let (a, b) = if a <= b { (a, b) } else { (b, a) };
                ((1u32 << (b + 1)) - 1) & !((1u32 << a) - 1)
            }
            _ => 0,
        }
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
        let open_span = self.current_span();
        self.advance(); // Open
        let path = self.parse_expr(arena, 0);

        let open_mode = self.parse_open_mode();
        let Some(access_code) = self.parse_open_access(open_mode) else {
            return arena.alloc(ExprNode::Block { stmts: Vec::new() });
        };
        let Some(lock_code) = self.parse_open_lock() else {
            return arena.alloc(ExprNode::Block { stmts: Vec::new() });
        };
        let open_flags = (lock_code << 4) | access_code;

        if !self.eat(TokenKind::Kw(Kw::As)) {
            // The channel follows `As`.
            let span = self.current_span();
            self.diagnostics.push(ERR_SYNTAX_ERROR, span);
        }
        let marked = self.eat(TokenKind::Kw(Kw::Hash));
        let ch = self.parse_expr(arena, 0);
        let mut args = vec![path];
        if self.eat(TokenKind::Kw(Kw::Len)) {
            self.expect(TokenKind::Kw(Kw::Eq), ERR_EXPECTED_EQ);
            args.push(self.parse_expr(arena, 0)); // `Len = <expr>` is a real use site
        }
        self.check_open_size(arena, ch, &args, open_span);
        self.end_stmt();
        arena.alloc(ExprNode::FileIoStmt {
            kind: FileIoKind::Open,
            channel: Some(ch),
            channel_marks: vec![marked],
            args,
            open_mode,
            open_flags,
        })
    }

    /// The `For mode` clause of an `Open`; the default mode is Random.
    fn parse_open_mode(&mut self) -> u8 {
        let default_mode: u8 = 4;
        if !self.eat(TokenKind::Kw(Kw::For)) {
            return default_mode;
        }
        let mode = match self.peek_word() {
            Some(Kw::Input) => 1,
            Some(Kw::Output) => 2,
            Some(Kw::Random) => 4,
            Some(Kw::Append) => 8,
            Some(Kw::Binary) => 0x20,
            _ => {
                // `For` needs a mode word.
                let span = self.current_span();
                self.diagnostics.push(ERR_SYNTAX_ERROR, span);
                return default_mode;
            }
        };
        self.advance();
        mode
    }

    /// The `Access` clause of an `Open`: its code (0 when absent), or `None` when the clause is
    /// wrong and the line has been skipped.
    fn parse_open_access(&mut self, open_mode: u8) -> Option<u8> {
        let access_span = self.current_span();
        if !self.eat_word(Kw::Access) {
            return Some(0);
        }
        let (has_read, has_write) = self.read_write_words();
        // `Access` names what the file is opened for.
        if !has_read && !has_write {
            self.error_and_skip_line(ERR_SYNTAX_ERROR);
            return None;
        }
        if Self::open_access_conflicts(has_read, has_write, open_mode) {
            self.diagnostics.push(ERR_SYNTAX_ERROR, access_span);
        }
        Some(match (has_read, has_write) {
            (true, true) => 3,
            (true, false) => 1,
            (false, true) => 2,
            (false, false) => 0,
        })
    }

    /// The access must be possible in the open mode: Read is not
    /// available to Output, Write not to Input, and Read Write to
    /// neither Input nor Output.
    fn open_access_conflicts(has_read: bool, has_write: bool, open_mode: u8) -> bool {
        (has_read && open_mode == 2)
            || (has_write && open_mode == 1)
            || (has_read && has_write && open_mode == 2)
            // Append writes, so read-only access cannot go with it.
            || (has_read && !has_write && open_mode == 8)
    }

    /// The `Shared` or `Lock` clause of an `Open`: its code (0 when absent), or `None` when the
    /// clause is wrong and the line has been skipped.
    fn parse_open_lock(&mut self) -> Option<u8> {
        if self.eat(TokenKind::Kw(Kw::Shared)) {
            return Some(4);
        }
        if !self.eat(TokenKind::Kw(Kw::Lock)) {
            return Some(0);
        }
        let (has_read, has_write) = self.read_write_words();
        match (has_read, has_write) {
            (true, true) => Some(1),
            (true, false) => Some(3),
            (false, true) => Some(2),
            (false, false) => {
                self.error_and_skip_line(ERR_SYNTAX_ERROR);
                None
            }
        }
    }

    /// The size accounting of an `Open`.
    fn check_open_size(&mut self, arena: &ExprArena, ch: NodeId, args: &[NodeId], open_span: Span) {
        // The clauses of an `Open` of a path expression cost 8 more than an assignment of it.
        let channel_is_simple =
            matches!(arena.get(ch), ExprNode::NameRef { .. } | ExprNode::Literal { lit: AstLit::Int(_) });
        if channel_is_simple && args.len() == 1 && self.output_list_too_complex(arena, args, cost::OPEN) {
            self.diagnostics.push(ERR_STATEMENT_TOO_COMPLEX, open_span);
        }
        // With `Len = n` the clauses and the path cost 10 more than assigning `n` when the
        // path is a name, and the path's own bytes differ from a name's by as much as its size.
        if channel_is_simple && args.len() == 2 && matches!(arena.get(args[0]), ExprNode::NameRef { .. }) {
            self.length_clause_too_complex(arena, args[1], cost::OPEN_WITH_LENGTH, open_span);
        }
    }

    /// The words after `Access` or `Lock` in an `Open`: `Read`, `Write`, or `Read Write`.
    /// The first is a name that spells its keyword, bare or in brackets; the `Write` that
    /// follows `Read` must be the bare keyword.
    fn read_write_words(&mut self) -> (bool, bool) {
        if self.eat_word(Kw::Read) {
            let has_write = self.eat(TokenKind::Kw(Kw::Write));
            (true, has_write)
        } else {
            (false, self.eat_word(Kw::Write))
        }
    }

    /// `Close [[#]filenum, ...]`
    fn parse_file_close(&mut self, arena: &mut ExprArena) -> NodeId {
        let close_span = self.current_span();
        self.advance(); // Close
        let mut chs = Vec::new();
        let mut marks = Vec::new();
        // An empty list ends at the end of the statement or a colon; an `Else` ends the list
        // only behind a channel.
        if self.peek().kind == TokenKind::Kw(Kw::Else) {
            self.error_and_skip_line(ERR_SYNTAX_ERROR);
            return arena.alloc(ExprNode::FileIoStmt { kind: FileIoKind::Close, channel: None, args: chs, channel_marks: marks, open_mode: 0, open_flags: 0 });
        }
        while !self.peek().is_stmt_end()
            && self.peek().kind != TokenKind::Kw(Kw::Colon)
            && !(self.peek().kind == TokenKind::Kw(Kw::Else) && !chs.is_empty())
        {
            marks.push(self.eat(TokenKind::Kw(Kw::Hash)));
            chs.push(self.parse_expr(arena, 0));
            if !self.eat(TokenKind::Kw(Kw::Comma)) { break; }
            if self.peek().is_stmt_end() || matches!(self.peek().kind, TokenKind::Kw(Kw::Colon) | TokenKind::Kw(Kw::Else)) {
                // A comma needs a channel after it.
                let span = self.current_span();
                self.diagnostics.push(ERR_SYNTAX_ERROR, span);
                break;
            }
        }
        // `Close #a, #b` costs its channels, a separator between them and 2 over a store.
        if !chs.is_empty() && !matches!(self.last_consumed, TokenKind::Kw(Kw::Comma)) {
            let parts = chs.iter().enumerate().try_fold(cost::CLOSE, |total, (i, ch)| {
                let sep = if i > 0 { cost::ITEM_SEPARATOR } else { 0 };
                Some(total + line_size::expression_part(arena, *ch)? + sep)
            });
            self.check_parts_size(parts, close_span);
        }
        self.end_stmt_before_rem_is_error();
        arena.alloc(ExprNode::FileIoStmt {
            kind: FileIoKind::Close,
            channel: None,
            channel_marks: marks,
            args: chs,
            open_mode: 0,
            open_flags: 0,
        })
    }

    /// `Print #filenum, [args]` or `Write #filenum, [args]`.
    fn parse_file_print_after_keyword(&mut self, arena: &mut ExprArena, kind: FileIoKind) -> NodeId {
        let marked = self.eat(TokenKind::Kw(Kw::Hash));
        let ch = self.parse_expr(arena, 0);
        let mut args = Vec::new();
        let first_span = self.current_span();
        if self.expect_comma() {
            self.parse_print_method_items(arena, &mut args);
        }
        // `Print #n, expr` and `Write #n, expr` cost 2 more than an assignment of the first
        // expression, and the channel's own expression.
        if let Some(channel_size) = line_size::expression_part(arena, ch) {
            if self.output_list_too_complex(arena, &args, cost::FILE_OUTPUT + channel_size) {
                self.diagnostics.push(ERR_STATEMENT_TOO_COMPLEX, first_span);
            }
        }
        self.end_stmt();
        arena.alloc(ExprNode::FileIoStmt { kind, channel: Some(ch), channel_marks: vec![marked], args, open_mode: 0, open_flags: 0 })
    }

    /// `Input #filenum, varlist`.
    fn parse_file_input_after_keyword(&mut self, arena: &mut ExprArena) -> NodeId {
        let marked = self.eat(TokenKind::Kw(Kw::Hash));
        let ch = self.parse_expr(arena, 0);
        let mut vars = Vec::new();
        if self.expect_comma() {
            loop {
                vars.push(self.parse_target_expr(arena));
                if !self.eat(TokenKind::Kw(Kw::Comma)) { break; }
            }
        }
        self.end_stmt_before_rem_is_error();
        arena.alloc(ExprNode::FileIoStmt { kind: FileIoKind::Input, channel: Some(ch), channel_marks: vec![marked], args: vars, open_mode: 0, open_flags: 0 })
    }

    /// `Get #filenum, [recnum], var` or `Put #filenum, [recnum], val`.
    fn parse_file_get_put_after_keyword(&mut self, arena: &mut ExprArena, kind: FileIoKind) -> NodeId {
        let put_span = self.current_span();
        let marked = self.eat(TokenKind::Kw(Kw::Hash));
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
        // `Put #n, , value`, the record left out, costs the channel's expression and 2 over an
        // assignment of the value.
        if kind == FileIoKind::Put {
            if let [value] = args.as_slice() {
                let parts = line_size::expression_part(arena, ch)
                    .zip(line_size::expression_part(arena, *value))
                    .map(|(channel, value)| cost::MISSING_RECORD + channel + value);
                self.check_parts_size(parts, put_span);
            }
        }
        // With the record given, `Put` and `Get` cost the channel, the record and the value or
        // variable, and nothing more than an assignment's store.
        if let [record, value] = args.as_slice() {
            let parts = [ch, *record, *value]
                .iter()
                .try_fold(0, |total, part| Some(total + line_size::expression_part(arena, *part)?));
            self.check_parts_size(parts, put_span);
        }
        self.end_stmt();
        arena.alloc(ExprNode::FileIoStmt { kind, channel: Some(ch), channel_marks: vec![marked], args, open_mode: 0, open_flags: 0 })
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
        let lock_span = self.current_span();
        let marked = self.eat(TokenKind::Kw(Kw::Hash));
        let ch = self.parse_expr(arena, 0);
        let mut args = Vec::new();
        if self.eat(TokenKind::Kw(Kw::Comma)) {
            if self.peek().kind == TokenKind::Kw(Kw::To) {
                // `Lock #1, To 10`: the first record of the range is left out.
                self.advance();
                args.push(arena.alloc(ExprNode::MissingArg));
                args.push(self.parse_expr(arena, 0));
            } else {
                args.push(self.parse_expr(arena, 0));
                if self.eat(TokenKind::Kw(Kw::To)) {
                    args.push(self.parse_expr(arena, 0));
                }
            }
        }
        // `Lock #n, from` and `Unlock #n, from` cost the channel's expression and 2 over an
        // assignment of `from`.
        if let [from] = args.as_slice() {
            let parts = line_size::expression_part(arena, ch)
                .zip(line_size::expression_part(arena, *from))
                .map(|(channel, from)| cost::LOCK + channel + from);
            self.check_parts_size(parts, lock_span);
        }
        self.end_stmt();
        arena.alloc(ExprNode::FileIoStmt { kind, channel: Some(ch), channel_marks: vec![marked], args, open_mode: 0, open_flags: 0 })
    }

    /// Parse ident-statement starting from a token already consumed from the stream.
    fn parse_ident_stmt_with_tok(&mut self, arena: &mut ExprArena, first: Token) -> NodeId {
        let sym = first.sym.map(|s| s as u32).unwrap_or(0);
        let lhs_base = arena.alloc(ExprNode::NameRef { sym, suffix: first.type_suffix });
        self.set_span(lhs_base, first.span);
        self.stmt_byval = None;
        self.stmt_named = None;
        self.graphics_chain = Some(ChainStart::Statement);
        self.graphics_stop = None;
        // Apply any postfix operators (.member, (args), etc.)
        let lhs = self.parse_postfix(arena, lhs_base);
        self.graphics_chain = None;
        self.finish_ident_stmt(arena, lhs, first.span)
    }

    // ── Procedure / declaration parsers ───────────────────────────────────────

    /// Parse an access-modifier prefix (Public/Private/Friend/Static) then dispatch.
    fn parse_access_decl(&mut self, arena: &mut ExprArena) -> NodeId {
        let first = self.advance();
        let (access_tok, static_proc, access_span) = self.read_access_words(&first);
        // `Static` outside a procedure only qualifies a procedure. What follows it is read as
        // the list of variables it would declare, so a name or declarator that is wrong is
        // reported first; a list that reads well is refused as a whole.
        if access_tok == TokenKind::Kw(Kw::Static) && !self.at_procedure_keyword() {
            let reported = self.diagnostics.len();
            self.parse_dim_list_static(arena, false, true);
            if self.diagnostics.len() == reported {
                self.diagnostics.push(ERR_INVALID_OUTSIDE_PROCEDURE, first.span);
            }
            return arena.alloc(ExprNode::Block { stmts: vec![] });
        }
        let is_friend = access_tok == TokenKind::Kw(Kw::Friend);
        if is_friend {
            self.reject_in_standard_module_at(access_span);
        }
        // `Friend` and a combined `Static` only qualify a procedure.
        let procedure_only = is_friend || (static_proc && access_tok != TokenKind::Kw(Kw::Static));
        if procedure_only && !self.at_procedure_keyword() {
            let span = self.current_span();
            self.diagnostics.push_labeled(ERR_UNEXPECTED_TOKEN, span, "Function or Sub or Property");
            self.skip_to_stmt_end();
            return arena.alloc(ExprNode::Block { stmts: vec![] });
        }
        // An event is always public.
        if self.peek().kind == TokenKind::Kw(Kw::Event)
            && matches!(access_tok, TokenKind::Kw(Kw::Private) | TokenKind::Kw(Kw::Global) | TokenKind::Kw(Kw::Friend))
        {
            self.error_and_skip_to_stmt_end(ERR_UNEXPECTED_TOKEN);
            return arena.alloc(ExprNode::Block { stmts: vec![] });
        }
        self.static_prefix = static_proc;
        let node = self.parse_access_member(arena, static_proc);
        self.static_prefix = false;
        self.record_access_decl(arena, node, &access_tok, access_span, static_proc);
        node
    }

    /// The access words before a declaration. `first` is the first, already read. Returns the
    /// access word that decides the visibility, whether the declaration is `Static`, and where
    /// the access word is.
    fn read_access_words(&mut self, first: &Token) -> (TokenKind, bool, Span) {
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
        (access_tok, static_proc, access_span)
    }

    /// Records the visibility and the qualifiers of a declaration that follows access words, and
    /// reports what an object module does not allow to be public.
    fn record_access_decl(
        &mut self,
        arena: &ExprArena,
        node: NodeId,
        access_tok: &TokenKind,
        access_span: Span,
        static_proc: bool,
    ) {
        let vis = match access_tok {
            TokenKind::Kw(Kw::Public) | TokenKind::Kw(Kw::Global) | TokenKind::Kw(Kw::Friend) => {
                Some(true)
            }
            TokenKind::Kw(Kw::Private) => Some(false),
            _ => None,
        };
        if let Some(is_public) = vis {
            self.mark_decl_visibility(arena, node, is_public);
        }
        if *access_tok == TokenKind::Kw(Kw::Global) && self.module_kind != ModuleKind::Standard && !self.in_procedure {
            self.diagnostics.push(ERR_PUBLIC_MEMBER_NOT_ALLOWED, access_span);
        } else {
            self.check_public_in_object_module(arena, node, access_span, vis == Some(true));
        }
        if *access_tok == TokenKind::Kw(Kw::Friend) {
            self.decl_friend.insert(node.0);
        }
        if static_proc && matches!(arena.get(node), ExprNode::ProcDecl { .. }) {
            self.decl_static.insert(node.0);
        }
    }

    /// Whether the next token is `Sub`, `Function` or `Property`.
    fn at_procedure_keyword(&mut self) -> bool {
        matches!(
            self.peek().kind,
            TokenKind::Kw(Kw::Sub) | TokenKind::Kw(Kw::Function) | TokenKind::Kw(Kw::Property)
        )
    }

    /// The declaration that follows the access words.
    fn parse_access_member(&mut self, arena: &mut ExprArena, static_proc: bool) -> NodeId {
        match self.peek().kind.clone() {
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
                self.parse_dim_list_static(arena, false, static_proc)
            }
            // `Shared` directly after the access word is skipped by the variable list.
            TokenKind::Kw(Kw::Shared) if !static_proc => self.parse_dim_list_static(arena, false, false),
            // A contextual keyword names the variable (`Public Name As String`).
            TokenKind::Kw(_) if self.at_contextual_keyword() => {
                self.parse_dim_list_static(arena, false, static_proc)
            }
            _ => {
                let span = self.current_span();
                if self.at_underscore_word() {
                    self.report_invalid_character(span);
                } else {
                    self.diagnostics.push(ERR_UNEXPECTED_TOKEN, span);
                }
                self.skip_to_stmt_end();
                arena.alloc(ExprNode::Block { stmts: vec![] })
            }
        }
    }

    /// In an object module, constants, arrays, fixed-length strings, `Declare` statements and
    /// (unless the class is exposed) user types cannot be public members.
    fn check_public_in_object_module(&mut self, arena: &ExprArena, node: NodeId, at: Span, is_public: bool) {
        if !is_public || self.module_kind == ModuleKind::Standard || self.in_procedure {
            return;
        }
        let items = match arena.get(node) {
            ExprNode::Block { stmts } => stmts.clone(),
            _ => vec![node],
        };
        for item in items {
            match arena.get(item) {
                ExprNode::DeclareDecl { .. } => self.diagnostics.push(ERR_PUBLIC_MEMBER_NOT_ALLOWED, at),
                ExprNode::TypeDecl { .. } if !self.class_exposed => {
                    self.diagnostics.push(ERR_PUBLIC_TYPE_IN_PRIVATE_MODULE, at)
                }
                ExprNode::DimItem { is_const, bounds, type_node, .. } => {
                    let fixed_string = type_node.is_some_and(|t| {
                        matches!(arena.get(t), ExprNode::StringType { fixed_len: Some(_) })
                    });
                    if *is_const || bounds.is_some() || fixed_string {
                        self.diagnostics.push(ERR_PUBLIC_MEMBER_NOT_ALLOWED, at);
                    }
                }
                _ => {}
            }
        }
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
        // Parameter list
        let params = if self.eat(TokenKind::Kw(Kw::LParen)) {
            let p = self.parse_param_list(arena);
            self.expect(TokenKind::Kw(Kw::RParen), ERR_EXPECTED_RPAREN);
            Some(p)
        } else { None };
        self.check_param_list(arena, params, ParamOwner::Procedure);
        // Return type.
        let suffix = self.declared_suffix(&name);
        let ret_type = if kind == ProcKind::Function && self.eat(TokenKind::Kw(Kw::As)) {
            let t = self.parse_return_type(arena);
            if suffix != TypeSuffix::None {
                self.report_expected(name.span, "end of statement");
            }
            Some(t)
        } else if suffix != TypeSuffix::None {
            if kind == ProcKind::Function {
                self.type_node_from_suffix(arena, suffix)
            } else {
                self.report_expected(name.span, "Identifier without type suffix");
                None
            }
        } else { None };
        self.finish_proc_decl(arena, kind, &name, params, ret_type)
    }

    /// The end of a procedure's header (a trailing `Static`) and its body.
    fn finish_proc_decl(
        &mut self,
        arena: &mut ExprArena,
        kind: ProcKind,
        name: &Token,
        params: Option<NodeId>,
        ret_type: Option<NodeId>,
    ) -> NodeId {
        let name_id = name.sym.map(|s| s as u32).unwrap_or(0);
        let static_span = self.current_span();
        let trailing_static = self.eat(TokenKind::Kw(Kw::Static));
        if trailing_static && self.static_prefix {
            // `Static` is given once.
            self.diagnostics.push(ERR_SYNTAX_ERROR, static_span);
        }
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
        self.procedure_seen = true;
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
            // The procedure ends with the last line of the text: that line is where its
            // missing `End` is reported, the line the final line break ends included.
            let end = self.src.len();
            let break_len = if self.src.ends_with(b"\r\n") {
                2
            } else {
                usize::from(self.src.ends_with(b"\n") || self.src.ends_with(b"\r"))
            };
            let at = Span { start: (end - break_len) as u32, len: 0 };
            self.diagnostics.push(expected_code, at);
            return body;
        }
        self.advance(); // End
        // A property getter is closed by any of the three procedure closers.
        let closes = self.peek().kind == TokenKind::Kw(closer)
            || (kind == ProcKind::PropGet
                && matches!(self.peek().kind, TokenKind::Kw(Kw::Sub | Kw::Function | Kw::Property)));
        if !closes {
            self.diagnostics.push(expected_code, end_span);
        }
        self.advance(); // Sub / Function / Property
        if self.peek().is_stmt_end() {
            self.eat_eol();
        } else {
            // The closing line is a statement of its own: anything but a comment after it is
            // a syntax error.
            self.error_and_skip_line(ERR_SYNTAX_ERROR);
        }
        body
    }

    /// Records a label of the procedure; a second definition of the same label is reported.
    fn define_label(&mut self, label: LabelRef, span: Span) {
        if self.labels_defined.contains(&label) {
            self.diagnostics.push(ERR_DUPLICATE_DECLARATION, span);
        }
        self.labels_defined.push(label);
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
            self.value_owner_params = kind != ProcKind::PropGet;
            let p = self.parse_param_list(arena);
            self.value_owner_params = false;
            self.expect(TokenKind::Kw(Kw::RParen), ERR_EXPECTED_RPAREN);
            Some(p)
        } else { None };
        let owner = match kind {
            ProcKind::PropGet => ParamOwner::Procedure,
            ProcKind::PropSet => ParamOwner::PropertySet,
            _ => ParamOwner::PropertyValue,
        };
        self.check_param_list(arena, params, owner);
        if owner != ParamOwner::Procedure {
            let has_params = params.is_some_and(|p| matches!(arena.get(p), ExprNode::ArgList { args } if !args.is_empty()));
            if !has_params {
                self.diagnostics.push(ERR_PROPERTY_ARGUMENT_REQUIRED, name.span);
            }
        }
        let ret_type = if kind == ProcKind::PropGet && self.eat(TokenKind::Kw(Kw::As)) {
            Some(self.parse_return_type(arena))
        } else if kind == ProcKind::PropGet {
            self.type_node_from_suffix(arena, name.type_suffix)
        } else {
            // A procedure that sets a property has no type to give its name a character for.
            self.reject_name_suffix_as(&name, "Identifier without type suffix");
            None
        };
        self.check_property_agreement(arena, kind, name_id, name.span, params, ret_type);
        let static_span = self.current_span();
        let trailing_static = self.eat(TokenKind::Kw(Kw::Static));
        if trailing_static && self.static_prefix {
            // `Static` is given once.
            self.diagnostics.push(ERR_SYNTAX_ERROR, static_span);
        }
        self.end_stmt();
        let body = self.parse_procedure_body(arena, kind);
        let id = arena.alloc(ExprNode::ProcDecl { kind, name: name_id, params, ret_type, body });
        self.set_span(id, name.span);
        if trailing_static {
            self.decl_static.insert(id.0);
        }
        id
    }

    /// The procedures of one property must agree on its index parameters and value type.
    fn check_property_agreement(
        &mut self,
        arena: &ExprArena,
        kind: ProcKind,
        name: u32,
        span: Span,
        params: Option<NodeId>,
        ret_type: Option<NodeId>,
    ) {
        let list: Vec<NodeId> = match params.map(|p| arena.get(p)) {
            Some(ExprNode::ArgList { args }) => args.clone(),
            _ => Vec::new(),
        };
        let describe = |param: NodeId| match arena.get(param) {
            ExprNode::ParamDef { flags, name, type_node, .. } => {
                Some((*name, flags & 0x02 != 0, flags & 0x28 != 0, prop_type(arena, *type_node)))
            }
            _ => None,
        };
        let (index_params, value) = if kind == ProcKind::PropGet {
            (&list[..], prop_type(arena, ret_type))
        } else {
            let Some((last, rest)) = list.split_last() else { return };
            let value = match arena.get(*last) {
                // A value parameter written `v() As T` is an array of `T`.
                ExprNode::ParamDef { flags, type_node, .. } if flags & 0x08 != 0 => {
                    PropType::Array(Box::new(prop_type(arena, *type_node)))
                }
                ExprNode::ParamDef { type_node, .. } => prop_type(arena, *type_node),
                _ => return,
            };
            (rest, value)
        };
        let sig = PropertySig { kind, index: index_params.iter().filter_map(|&p| describe(p)).collect(), value };
        let clash = self
            .property_sigs
            .iter()
            .any(|(other_name, other)| *other_name == name && other.kind != kind && !other.agrees_with(&sig));
        if clash {
            self.diagnostics.push(ERR_PROPERTY_PROCEDURES_INCONSISTENT, span);
        }
        self.property_sigs.push((name, sig));
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
        self.reject_name_suffix_as(&name, "Identifier without type suffix");
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
        let Some(kind) = self.parse_declare_kind() else {
            return arena.alloc(ExprNode::Block { stmts: vec![] });
        };
        let name_tok = self.expect_decl_name();
        let name = name_tok.sym.map(|s| s as u32).unwrap_or(0);
        // A `Sub` returns nothing, so its name takes no type character.
        if kind == ProcKind::Sub {
            self.reject_name_suffix_as(&name_tok, "Identifier without type suffix");
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
        let Ok(alias) = self.parse_declare_alias(arena) else {
            return arena.alloc(ExprNode::Block { stmts: vec![] });
        };
        // Optional parameter list
        let params = if self.eat(TokenKind::Kw(Kw::LParen)) {
            let p = self.parse_declare_param_list(arena);
            self.expect(TokenKind::Kw(Kw::RParen), ERR_EXPECTED_RPAREN);
            Some(p)
        } else {
            None
        };
        self.check_param_list(arena, params, ParamOwner::External);
        // Optional return type.
        if kind == ProcKind::Function {
            self.reject_suffix_with_as(&name_tok, "end of statement");
        }
        let ret_type = if kind == ProcKind::Function && self.eat(TokenKind::Kw(Kw::As)) {
            Some(self.parse_return_type(arena))
        } else {
            None
        };
        self.end_stmt();
        let id = arena.alloc(ExprNode::DeclareDecl { kind, name, lib, alias, params, ret_type });
        self.set_span(id, name_tok.span);
        id
    }

    /// `Function` or `Sub` after `Declare`; anything else is reported and the statement ended.
    fn parse_declare_kind(&mut self) -> Option<ProcKind> {
        match self.peek().kind {
            TokenKind::Kw(Kw::Function) => {
                self.advance();
                Some(ProcKind::Function)
            }
            TokenKind::Kw(Kw::Sub) => {
                self.advance();
                Some(ProcKind::Sub)
            }
            _ => {
                self.error_and_skip_to_stmt_end(ERR_UNEXPECTED_TOKEN);
                self.end_stmt();
                None
            }
        }
    }

    /// The optional `Alias "name"` of a `Declare`. `Err` when the operand is wrong and the line
    /// has been skipped.
    fn parse_declare_alias(&mut self, arena: &mut ExprArena) -> Result<Option<NodeId>, ()> {
        if !self.eat(TokenKind::Kw(Kw::Alias)) {
            return Ok(None);
        }
        let alias_span = self.current_span();
        let Some(s) = self.declare_string_operand(arena) else {
            self.skip_line();
            return Err(());
        };
        if let ExprNode::Literal { lit: AstLit::Str(text) } = arena.get(s) {
            if let Some(ordinal) = text.strip_prefix('#') {
                self.check_alias_ordinal(ordinal, alias_span);
            }
        }
        Ok(Some(s))
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
        for (index, &param) in args.iter().enumerate() {
            let ExprNode::ParamDef { flags, .. } = arena.get(param) else { continue };
            let slot = param_rules::ParamSlot { owner, flags: *flags, index, count: args.len() };
            if let Some((code, label)) = param_rules::first_violation(arena, args, param, &slot, section) {
                let span = self.node_spans.get(param);
                match label {
                    Some(label) => self.diagnostics.push_labeled(code, span, label),
                    None => self.diagnostics.push(code, span),
                }
                return;
            }
            section = slot.section_after(section);
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
        // In a property assignment the value is the parameter that follows every Optional or
        // ParamArray one: `section` is set once such a parameter is read, `required_after`
        // once a required one follows it.
        let mut section = false;
        let mut required_after = false;
        // A standard module takes at most 60 parameters, an object module 59.
        let limit = if self.module_kind == ModuleKind::Standard { 60 } else { 59 };
        let mut too_many = None;
        loop {
            if args.len() == limit {
                too_many = Some(self.current_span());
            }
            let param = self.parse_param(arena, in_declare);
            Self::note_value_parameter(arena, param, &mut section, &mut required_after);
            args.push(param);
            if !self.eat(TokenKind::Kw(Kw::Comma)) { break; }
            if self.peek().kind == TokenKind::Kw(Kw::RParen) {
                // A comma needs a parameter after it.
                let span = self.current_span();
                self.report_expected(span, "identifier");
                break;
            }
        }
        if let Some(span) = too_many {
            self.diagnostics.push(ERR_TOO_MANY_ARGUMENTS, span);
        }
        if self.value_owner_params && section && !required_after {
            // The list ends without a required value parameter.
            let span = self.current_span();
            self.diagnostics.push(ERR_ARGUMENT_NOT_OPTIONAL, span);
        }
        arena.alloc(ExprNode::ArgList { args })
    }

    /// Tracks whether a required parameter follows an `Optional` or `ParamArray` one.
    fn note_value_parameter(arena: &ExprArena, param: NodeId, section: &mut bool, required_after: &mut bool) {
        let ExprNode::ParamDef { flags, .. } = arena.get(param) else { return };
        if flags & 0x21 != 0 {
            *section = true;
        } else if *section {
            *required_after = true;
        }
    }

    fn parse_param(&mut self, arena: &mut ExprArena, in_declare: bool) -> NodeId {
        let mut flags = self.parse_param_modifiers();
        let name = self.expect_decl_name();
        self.note_declared_name(&name, false);
        let nid = name.sym.map(|s| s as u32).unwrap_or(0);
        if self.eat(TokenKind::Kw(Kw::LParen)) {
            self.expect(TokenKind::Kw(Kw::RParen), ERR_EXPECTED_RPAREN);
            flags |= 0x08; // array parameter
        }
        self.reject_suffix_with_as(&name, "list separator or )");
        let type_node = if self.eat(TokenKind::Kw(Kw::As)) {
            Some(self.parse_param_type(arena, in_declare, &mut flags))
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
        // In a property assignment an Optional parameter is an error once its type and default
        // are read; whatever else stands in the parameter is not read.
        if self.value_owner_params && flags & 0x01 != 0 {
            self.skip_rest_of_parameter();
        }
        let id = arena.alloc(ExprNode::ParamDef { flags, name: nid, type_node, default });
        self.set_span(id, name.span);
        id
    }

    /// The pass-mode and `Optional` / `ParamArray` words before a parameter name, as flags.
    fn parse_param_modifiers(&mut self) -> u16 {
        // `ByVal` or `ByRef` may come before `Optional` as well as after it.
        let mut flags = self.take_pass_mode_word();
        let by_first = flags != 0;
        if self.eat(TokenKind::Kw(Kw::Optional)) { flags |= 0x01; }
        // ParamArray must precede the parameter name; mutually exclusive with ByVal/ByRef and
        // with Optional.
        if by_first || flags & 0x01 != 0 && !self.at_pass_mode_word() {
            // The pass-mode word was given already.
            return flags;
        }
        if self.eat(TokenKind::Kw(Kw::ParamArray)) {
            return flags | 0x20;
        }
        flags | self.take_pass_mode_word()
    }

    fn at_pass_mode_word(&mut self) -> bool {
        matches!(self.peek().kind, TokenKind::Kw(Kw::ByVal) | TokenKind::Kw(Kw::ByRef))
    }

    /// Reads a `ByVal` or `ByRef` and returns its flag (0 when there is none).
    fn take_pass_mode_word(&mut self) -> u16 {
        let flag = match self.peek().kind {
            TokenKind::Kw(Kw::ByVal) => 0x02,
            TokenKind::Kw(Kw::ByRef) => 0x04,
            _ => return 0,
        };
        self.advance();
        flag
    }

    /// The type after `As` of a parameter; sets the flag of a parameter that is a new object.
    fn parse_param_type(&mut self, arena: &mut ExprArena, in_declare: bool, flags: &mut u16) -> NodeId {
        if self.peek().kind == TokenKind::Kw(Kw::New) {
            // A parameter cannot be a new object.
            let span = self.current_span();
            self.report_expected(span, "type name");
            self.advance();
            *flags |= 0x10;
            self.require_class_name_after_new("type name");
        }
        if in_declare && self.eat(TokenKind::Kw(Kw::Any)) {
            return arena.alloc(ExprNode::AnyType);
        }
        let type_span = self.current_span();
        let t = self.parse_type_spec(arena, false);
        // A fixed-length string can only be passed as an array.
        if *flags & 0x08 == 0 && matches!(arena.get(t), ExprNode::StringType { fixed_len: Some(_) }) {
            self.diagnostics.push(ERR_EXPECTED_ARRAY, type_span);
        }
        t
    }

    /// Skips what stands in a parameter up to the comma or bracket that ends it.
    fn skip_rest_of_parameter(&mut self) {
        let mut depth = 0;
        while !self.peek().is_stmt_end()
            && !(depth == 0 && matches!(self.peek().kind, TokenKind::Kw(Kw::RParen) | TokenKind::Kw(Kw::Comma)))
        {
            match self.peek().kind {
                TokenKind::Kw(Kw::LParen) => depth += 1,
                TokenKind::Kw(Kw::RParen) => depth -= 1,
                _ => {}
            }
            self.advance();
        }
    }

    /// Build a `UserType` from its already-consumed first name token, taking an
    /// optional `.Name` qualifier tail (`Project.Class`, `Module.Type`). The node
    /// spans the first name; the tail is a `NameRef` with its own span.
    fn finish_user_type(&mut self, arena: &mut ExprArena, name: Token) -> NodeId {
        let nid = name.sym.map(|s| s as u32).unwrap_or(0);
        // `A.B` is a name with a child; each further `.C` nests: the child of `A.B.C` is the
        // type name `B.C`.
        let mut tail: Vec<Token> = Vec::new();
        while self.eat(TokenKind::Kw(Kw::Dot)) {
            tail.push(self.expect_qualified_type_name());
        }
        let mut child: Option<NodeId> = None;
        for (i, sub) in tail.iter().enumerate().rev() {
            let sid = sub.sym.map(|s| s as u32).unwrap_or(0);
            let node = match child {
                None => arena.alloc(ExprNode::NameRef { sym: sid, suffix: TypeSuffix::None }),
                Some(_) => arena.alloc(ExprNode::UserType { name: sid, child }),
            };
            self.set_span(node, sub.span);
            let _ = i;
            child = Some(node);
        }
        let id = arena.alloc(ExprNode::UserType { name: nid, child });
        self.set_span(id, name.span);
        id
    }

    /// The type after `As` in a function or property getter header: never a new object, and
    /// never a fixed-length string.
    fn parse_return_type(&mut self, arena: &mut ExprArena) -> NodeId {
        if self.peek().kind == TokenKind::Kw(Kw::New) {
            let span = self.current_span();
            self.report_expected(span, "type name");
        }
        let type_span = self.current_span();
        let t = self.parse_type_spec(arena, true);
        if matches!(arena.get(t), ExprNode::StringType { fixed_len: Some(_) }) {
            self.report_expected(type_span, "end of statement");
        }
        t
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
                let node = self.literal_node(arena, &tok);
                self.set_span(node, len_span);
                // A literal length too large for 16 bits is refused as it is read: inside a
                // procedure as a syntax error, elsewhere as an invalid length. A length from 65527
                // up to 65535, or zero, is judged with the other declarations afterwards.
                if let ExprNode::Literal { lit: AstLit::Int(n) | AstLit::Long(n) } = arena.get(node) {
                    if *n > 0xffff {
                        let code = if self.in_procedure { ERR_SYNTAX_ERROR } else { ERR_INVALID_FIXED_STRING_LENGTH };
                        self.diagnostics.push(code, len_span);
                    }
                }
                node
            }
            TokenKind::Ident => self.parse_length_constant_name(arena),
            TokenKind::Kw(_) if contextual => self.parse_length_constant_name(arena),
            _ => {
                self.diagnostics.push_labeled(ERR_EXPECTED_IDENT, len_span, "Identifier");
                arena.alloc(ExprNode::Literal { lit: AstLit::Int(0) })
            }
        }
    }

    /// A constant name as the length of a fixed-length string.
    fn parse_length_constant_name(&mut self, arena: &mut ExprArena) -> NodeId {
        let tok = self.advance();
        let sym = tok.sym.map(|s| s as u32).unwrap_or(0);
        let name = arena.alloc(ExprNode::NameRef { sym, suffix: tok.type_suffix });
        self.set_span(name, tok.span);
        name
    }

    fn parse_type_spec(&mut self, arena: &mut ExprArena, allow_array_suffix: bool) -> NodeId {
        // A type word is the same word in brackets: `[String] * 3` is a fixed-length string
        // and `[Integer]` an Integer, so the type is chosen by the keyword the name spells.
        // After `New` only a class can follow, and a bracketed word is just its name.
        let after_new = std::mem::take(&mut self.type_after_new);
        let word = if after_new && self.peek().bracketed { None } else { self.peek_word() };
        let node = match word {
            Some(Kw::Integer)  => { self.advance(); arena.alloc(ExprNode::BuiltinType { kind: 2  }) }
            Some(Kw::Long)     => { self.advance(); arena.alloc(ExprNode::BuiltinType { kind: 3  }) }
            Some(Kw::Single)   => { self.advance(); arena.alloc(ExprNode::BuiltinType { kind: 4  }) }
            Some(Kw::Double)   => { self.advance(); arena.alloc(ExprNode::BuiltinType { kind: 5  }) }
            Some(Kw::Currency) => { self.advance(); arena.alloc(ExprNode::BuiltinType { kind: 6  }) }
            Some(Kw::Date)     => { self.advance(); arena.alloc(ExprNode::BuiltinType { kind: 7  }) }
            Some(Kw::Byte)     => { self.advance(); arena.alloc(ExprNode::BuiltinType { kind: 17 }) }
            Some(Kw::Boolean)  => { self.advance(); arena.alloc(ExprNode::BuiltinType { kind: 11 }) }
            Some(Kw::Variant)  => { self.advance(); arena.alloc(ExprNode::BuiltinType { kind: 12 }) }
            Some(Kw::Object)   => { self.advance(); arena.alloc(ExprNode::BuiltinType { kind: 9  }) }
            Some(Kw::Decimal)  => {
                // `Decimal` is the subtype of a Variant, never the type of a declaration.
                let span = self.current_span();
                self.report_expected(span, "type name");
                self.advance();
                arena.alloc(ExprNode::BuiltinType { kind: 0 })
            }
            Some(Kw::String) => {
                self.advance();
                let fixed_len = if self.eat(TokenKind::Kw(Kw::Star)) {
                    Some(self.parse_fixed_string_length(arena))
                } else { None };
                arena.alloc(ExprNode::StringType { fixed_len })
            }
            _ => match self.peek().kind.clone() {
                TokenKind::Kw(Kw::New) => {
                    self.advance();
                    self.require_class_name_after_new("type name");
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
            },
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
        let type_span = self.current_span();
        self.advance(); // consume Type
        self.module_decl_seen = true;
        let name = self.expect_decl_name();
        self.reject_name_suffix(&name);
        let nid = name.sym.map(|s| s as u32).unwrap_or(0);
        self.end_stmt();
        let mut members = Vec::new();
        loop {
            let Some(after_member) = self.start_block_member(arena) else { continue };
            if self.peek2().kind != TokenKind::Kw(Kw::As) && self.read_statement_word() {
                continue;
            }
            if !self.parse_type_part(arena, type_span, &mut members, after_member) {
                break;
            }
        }
        let id = arena.alloc(ExprNode::TypeDecl { name: nid, members });
        self.set_span(id, name.span);
        id
    }

    /// The start of a line in the body of a `Type` or `Enum` block: blank lines are passed, and
    /// an attribute line is read. `None` when an attribute line has taken the line; otherwise
    /// whether an attribute line follows the member on this line.
    fn start_block_member(&mut self, arena: &mut ExprArena) -> Option<bool> {
        while self.eat(TokenKind::Eol) {}
        match self.take_body_attribute_line(arena) {
            attribute::BodyAttribute::Taken => None,
            attribute::BodyAttribute::AfterMember => Some(true),
            attribute::BodyAttribute::None => Some(false),
        }
    }

    /// One part of the body of a `Type` block. Returns false when the block has ended.
    fn parse_type_part(
        &mut self,
        arena: &mut ExprArena,
        type_span: Span,
        members: &mut Vec<NodeId>,
        after_member: bool,
    ) -> bool {
        match self.peek().kind.clone() {
            // `End` followed by `As` or `(` is a member named `End`.
            TokenKind::Kw(Kw::End)
                if !matches!(
                    self.peek2().kind,
                    TokenKind::Kw(Kw::As) | TokenKind::Kw(Kw::LParen)
                ) =>
            {
                self.advance();
                self.expect(TokenKind::Kw(Kw::Type), ERR_UNEXPECTED_TOKEN);
                if members.is_empty() {
                    // A type without members is reported where the type starts.
                    self.diagnostics.push(ERR_TYPE_WITHOUT_MEMBERS, type_span);
                }
                self.end_stmt();
                false
            }
            TokenKind::Eof => {
                // The file ended inside the block.
                let span = self.current_span();
                self.diagnostics.push(ERR_STATEMENT_IN_TYPE, span);
                false
            }
            // A comment line between members is a blank line, and so is a lone colon.
            TokenKind::Kw(Kw::Apos) | TokenKind::Kw(Kw::Rem) | TokenKind::Kw(Kw::Colon) => {
                self.advance();
                true
            }
            // A statement keyword cannot start a member. A keyword followed by
            // `As` is a member that happens to be named like one, and so is one of
            // the few reserved words that also name callable things when bounds follow.
            TokenKind::Kw(_)
                if !self.at_contextual_keyword()
                    && !matches!(self.peek2().kind, TokenKind::Kw(Kw::As))
                    && !(self.peek2().kind == TokenKind::Kw(Kw::LParen) && self.at_callable_reserved_word()) =>
            {
                self.error_and_skip_line(ERR_STATEMENT_IN_TYPE);
                true
            }
            _ => {
                if let Some(member) = self.parse_type_member(arena) {
                    if !after_member {
                        members.push(member);
                    }
                }
                true
            }
        }
    }

    /// A member line of a `Type` block. `None` when the line is not a member and has been
    /// skipped.
    fn parse_type_member(&mut self, arena: &mut ExprArena) -> Option<NodeId> {
        let line_span = self.current_span();
        // A keyword that starts a statement of its own cannot take bounds as a member.
        if self.peek2().kind == TokenKind::Kw(Kw::LParen) && self.at_statement_word() {
            self.report_expected_and_skip_line("end of statement");
            return None;
        }
        // A bracketed name is a member only when bounds follow it.
        let bracketed = self.peek().bracketed;
        if bracketed && self.peek2().kind != TokenKind::Kw(Kw::LParen) {
            if self.peek2().kind == TokenKind::Kw(Kw::As) {
                self.report_expected(line_span, "expression");
            } else {
                self.diagnostics.push(ERR_STATEMENT_IN_TYPE, line_span);
            }
            self.skip_line();
            return None;
        }
        self.reserved_names_allowed = true;
        self.bracketed_name_ok = bracketed;
        self.last_item_had_as = false;
        let m = self.parse_dim_item(arena, false, false, false);
        self.bracketed_name_ok = false;
        self.reserved_names_allowed = false;
        // A member whose bounds give a range needs no `As`: it is an array of Variant.
        let has_bounds = matches!(arena.get(m), ExprNode::DimItem { bounds: Some(b), .. }
            if matches!(arena.get(*b), ExprNode::ArgList { args }
                if args.iter().any(|a| matches!(arena.get(*a), ExprNode::RangeTo { .. }))));
        if !self.last_item_had_as && !has_bounds && self.peek().kind != TokenKind::Kw(Kw::Eq) {
            self.diagnostics.push(ERR_STATEMENT_IN_TYPE, line_span);
            self.skip_line();
            return None;
        }
        // `name = value` is a statement, not a member.
        if self.peek().kind == TokenKind::Kw(Kw::Eq) {
            self.diagnostics.push(ERR_STATEMENT_IN_TYPE, line_span);
            self.skip_line();
        } else {
            self.end_stmt();
        }
        Some(m)
    }

    /// Parse an Enum declaration block.
    fn parse_enum_decl(&mut self, arena: &mut ExprArena) -> NodeId {
        self.advance(); // consume Enum
        self.module_decl_seen = true;
        // An enumeration may be named with a bracketed word.
        self.bracketed_name_ok = true;
        let name = self.expect_decl_name();
        self.bracketed_name_ok = false;
        self.reject_name_suffix(&name);
        let nid = name.sym.map(|s| s as u32).unwrap_or(0);
        // A character that is no part of the language (a type character after a bracketed
        // name is one) cannot follow the name.
        if self.peek().kind == TokenKind::Error {
            let span = self.current_span();
            self.report_invalid_character(span);
            self.skip_line();
        } else {
            self.end_stmt();
        }
        let mut members = Vec::new();
        loop {
            let Some(after_member) = self.start_block_member(arena) else { continue };
            if self.read_statement_word() {
                continue;
            }
            if !self.parse_enum_part(arena, &mut members, after_member) {
                break;
            }
        }
        let id = arena.alloc(ExprNode::EnumDecl { name: nid, members });
        self.set_span(id, name.span);
        id
    }

    /// One part of the body of an `Enum` block. Returns false when the block has ended.
    fn parse_enum_part(&mut self, arena: &mut ExprArena, members: &mut Vec<NodeId>, after_member: bool) -> bool {
        match self.peek().kind.clone() {
            TokenKind::Kw(Kw::End) => return !self.parse_enum_end(members.is_empty()),
            TokenKind::Eof => {
                // The file ended inside the block.
                let span = self.current_span();
                self.diagnostics.push(ERR_INVALID_IN_ENUM, span);
                return false;
            }
            // A comment line between members is a blank line, and so is a lone colon.
            TokenKind::Kw(Kw::Apos) | TokenKind::Kw(Kw::Rem) | TokenKind::Kw(Kw::Colon) => {
                self.advance();
            }
            TokenKind::Ident => self.parse_enum_member(arena, members, after_member, false),
            // A non-reserved keyword (`Text`, `Key`) can name a member, unless it
            // starts a statement that needs operands (`Name`, `Width`, ...).
            // `Date`, `DoEvents` and `CDec` are reserved but name a member too.
            TokenKind::Kw(_) if self.names_enum_member() => {
                self.parse_enum_member(arena, members, after_member, true);
            }
            TokenKind::Error if self.at_underscore_word() => {
                let span = self.current_span();
                self.report_invalid_character(span);
                self.skip_line();
            }
            // Not a member: report it and resume on the next line.
            _ => self.error_and_skip_line(ERR_INVALID_IN_ENUM),
        }
        true
    }

    /// The `End` of an `Enum` block. Returns true when the block has ended, and false when the
    /// closing words are not on one physical line: the line is then reported and skipped and
    /// the block goes on.
    fn parse_enum_end(&mut self, no_members: bool) -> bool {
        let end_span = self.current_span();
        self.advance();
        // Inside an Enum the closing words must share one physical line.
        let gap_start = (end_span.start + end_span.len) as usize;
        let gap_end = self.peek().span.start as usize;
        if self.peek().kind == TokenKind::Kw(Kw::Enum)
            && self.src.get(gap_start..gap_end).is_some_and(|gap| gap.iter().any(|&b| b == b'\n' || b == b'\r'))
        {
            self.error_and_skip_line(ERR_INVALID_IN_ENUM);
            return false;
        }
        self.expect(TokenKind::Kw(Kw::Enum), ERR_UNEXPECTED_TOKEN);
        if no_members {
            self.diagnostics.push(ERR_ENUM_WITHOUT_MEMBERS, end_span);
        }
        self.end_stmt();
        true
    }

    /// A member of an `Enum` block: a name, or (`reserved_word`) a keyword that names one.
    fn parse_enum_member(&mut self, arena: &mut ExprArena, members: &mut Vec<NodeId>, after_member: bool, reserved_word: bool) {
        let start = self.current_span();
        let name_tok = self.peek().clone();
        if reserved_word {
            self.reserved_names_allowed = true;
        } else {
            self.bracketed_name_ok = name_tok.bracketed;
        }
        self.in_enum_member = true;
        let m = self.parse_dim_item(arena, true, false, false);
        self.in_enum_member = false;
        if reserved_word {
            self.reserved_names_allowed = false;
        } else {
            self.bracketed_name_ok = false;
        }
        self.check_enum_member_suffix(arena, m, &name_tok);
        self.finish_enum_member(arena, m, start, after_member);
        if !after_member {
            members.push(m);
        }
    }

    /// A member of an enumeration is a `Long`: given a value, its name may carry no
    /// type-declaration character but `&`, and a keyword spelled with a trailing `$` carries
    /// the string character.
    fn check_enum_member_suffix(&mut self, arena: &ExprArena, member: NodeId, name: &Token) {
        let has_value = matches!(arena.get(member), ExprNode::DimItem { init: Some(_), .. });
        let suffix = if name.type_suffix != TypeSuffix::None { name.type_suffix } else { self.dollar_word_suffix(name) };
        if has_value && !matches!(suffix, TypeSuffix::None | TypeSuffix::Long) {
            self.diagnostics.push(ERR_TYPE_CHARACTER_MISMATCH, name.span);
        }
    }

    /// The end of an `Enum` member: the member takes no type, stands on one physical line, and
    /// a colon may follow it only when it has a value.
    fn finish_enum_member(&mut self, arena: &ExprArena, member: NodeId, start: Span, before_attribute_line: bool) {
        let has_value = matches!(arena.get(member), ExprNode::DimItem { init: Some(_), .. });
        // A name followed by brackets is a call or an array element, not a member: it can only
        // be the start of a statement, and an empty pair of brackets on its own is one that
        // lacks the `=` of an assignment.
        if let ExprNode::DimItem { bounds: Some(bounds), .. } = arena.get(member) {
            let empty = matches!(arena.get(*bounds), ExprNode::ArgList { args } if args.is_empty());
            if empty && !has_value {
                let span = self.current_span();
                self.diagnostics.push_labeled(ERR_UNEXPECTED_TOKEN, span, "=");
            } else {
                self.diagnostics.push(ERR_INVALID_IN_ENUM, start);
            }
            self.skip_line();
            return;
        }
        let end = self.peek().span.start as usize;
        let joined = self
            .src
            .get(start.start as usize..end)
            .is_some_and(|text| text.iter().any(|&b| b == 10 || b == 13));
        let colon_without_value = !has_value && !before_attribute_line && self.peek().kind == TokenKind::Kw(Kw::Colon);
        if self.last_item_had_as || joined || colon_without_value {
            self.diagnostics.push(ERR_INVALID_IN_ENUM, start);
            self.skip_line();
        } else {
            self.end_stmt();
        }
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
        // The operand names something: a bracket, a literal, `Nothing` or `Array(...)` does not.
        if self.at_array_call() {
            let span = self.current_span();
            self.report_expected(span, "expression");
        }
        let expr = self.parse_target_expr(arena);
        self.expect(TokenKind::Kw(Kw::Is), ERR_EXPECTED_IS);
        // The type is an object type: `Object` or a class name, never a built-in value type
        // and never `New`.
        if matches!(
            self.peek().kind,
            TokenKind::Kw(
                Kw::New | Kw::Integer | Kw::Long | Kw::Single | Kw::Double | Kw::Currency | Kw::Date
                    | Kw::Byte | Kw::Boolean | Kw::Variant | Kw::String
            )
        ) {
            let span = self.current_span();
            self.report_expected(span, "identifier");
        }
        let type_spec = self.parse_type_spec(arena, false);
        arena.alloc(ExprNode::TypeOf { expr, type_spec })
    }

    /// `New <TypeName>`
    fn parse_new_expr(&mut self, arena: &mut ExprArena) -> NodeId {
        self.advance(); // consume New
        // Only a class name follows: a built-in type, `Object` or a second `New` is an error.
        self.require_class_name_after_new("identifier");
        let name = self.expect_ident();
        let type_spec = self.finish_user_type(arena, name);
        arena.alloc(ExprNode::New { type_spec })
    }

    /// Parses primary expressions.
    fn parse_primary(&mut self, arena: &mut ExprArena) -> NodeId {
        // The chain that follows the primary is the one that may end in a graphics word; what is
        // read inside the primary (a group, the operands of a function) is expression text.
        let chain = self.graphics_chain.take();
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
            ref kind if is_number_or_string_literal(kind) => {
                let tok = self.advance();
                self.literal_node(arena, &tok)
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
                { self.graphics_chain = chain; self.parse_postfix(arena, base) }
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
            TokenKind::Ident => self.parse_name_primary(arena, chain),
            // `AddressOf proc` — function-pointer operand.
            TokenKind::Kw(Kw::AddressOf) => {
                let at = self.advance();
                if !std::mem::take(&mut self.address_of_ok) {
                    self.report_expected(at.span, "expression");
                }
                let name = self.parse_procedure_name(arena);
                // The operand is a whole expression that starts with the procedure's name:
                // `AddressOf Bar + 1` takes the operators that follow.
                let operand = self.parse_infix(arena, name, 0);
                arena.alloc(ExprNode::AddressOf { operand })
            }
            // Leading `.`/`!` — With-block member reference in expression
            // position (e.g. `y = .Field`), same as in a target position.
            TokenKind::Kw(Kw::Dot) | TokenKind::Kw(Kw::Bang) => {
                let access = self
                    .try_parse_leading_with_member(arena)
                    .expect("peek was Dot/Bang");
                { self.graphics_chain = chain; self.parse_postfix(arena, access) }
            }
            // A literal the scanner rejected (out of range, no digits). Inside a procedure
            // every such failure is a syntax error; outside one the kind of failure shows.
            TokenKind::Error => self.parse_failed_literal(arena),
            // Keyword that can start a primary (e.g. built-in functions)
            TokenKind::Kw(k) => self.parse_keyword_primary(arena, k, chain),
            // Unexpected — emit error, return placeholder
            _ => {
                let span = self.current_span();
                self.diagnostics.push(ERR_EXPECTED_EXPR, span);
                arena.alloc(ExprNode::Literal { lit: AstLit::Int(0) })
            }
        }
    }

    /// A name as a primary: a call of one of the built-in functions with their own argument
    /// grammar, or a name with its member accesses and calls.
    fn parse_name_primary(&mut self, arena: &mut ExprArena, chain: Option<ChainStart>) -> NodeId {
        let tok = self.advance();
        let sym_id = tok.sym.map(|s| s as u32).unwrap_or(0);
        let base = arena.alloc(ExprNode::NameRef { sym: sym_id, suffix: tok.type_suffix });
        self.set_span(base, tok.span);
        // A name that spells one of these four built-in functions keeps their argument
        // grammar, bracketed or not; every other bracketed word is an ordinary name.
        if !self.in_target
            && matches!(self.word_of(&tok), Some(Kw::CByte | Kw::InStr | Kw::InStrB | Kw::StrComp))
        {
            let name = self.scanner.sym_name(sym_id).to_string();
            if let Some(shape) = INTRINSIC_SHAPES.iter().find(|s| name.eq_ignore_ascii_case(s.name)) {
                return self.parse_shaped_intrinsic_call(arena, base, tok.span, shape);
            }
        }
        self.graphics_chain = chain;
        self.parse_postfix(arena, base)
    }

    /// A literal the scanner rejected (out of range, no digits). Inside a procedure
    /// every such failure is a syntax error; outside one the kind of failure shows.
    fn parse_failed_literal(&mut self, arena: &mut ExprArena) -> NodeId {
        let tok = self.advance();
        match tok.lit {
            Some(Lit::Failed(NumberFailure::Overflow)) if !self.in_procedure => {
                self.diagnostics.push(ERR_OVERFLOW, tok.span);
            }
            Some(Lit::Failed(NumberFailure::Unreadable)) => self.report_expected(tok.span, "expression"),
            _ => self.diagnostics.push(ERR_SYNTAX_ERROR, tok.span),
        }
        arena.alloc(ExprNode::Literal { lit: AstLit::Int(0) })
    }

    /// A keyword that can start a primary (e.g. a built-in function).
    fn parse_keyword_primary(&mut self, arena: &mut ExprArena, k: Kw, chain: Option<ChainStart>) -> NodeId {
        let tok = self.advance();
        let sym_id = tok.sym.map(|s| s as u32).unwrap_or(0);
        let base = arena.alloc(ExprNode::NameRef { sym: sym_id, suffix: tok.type_suffix });
        self.set_span(base, tok.span);
        let name = self.scanner.sym_name(sym_id).to_string();
        if let Some(shape) = INTRINSIC_SHAPES
            .iter()
            .find(|s| !self.in_target && name.trim_end_matches('$').eq_ignore_ascii_case(s.name))
        {
            return self.parse_shaped_intrinsic_call(arena, base, tok.span, shape);
        }
        if name.eq_ignore_ascii_case("Array") && self.peek().kind != TokenKind::Kw(Kw::LParen) {
            // `Array` is a function: brackets are not optional.
            self.report_expected(tok.span, "(");
            return base;
        }
        if !matches!(k, Kw::Input | Kw::InputS | Kw::InputB | Kw::InputBS) {
            self.graphics_chain = chain;
            return self.parse_postfix(arena, base);
        }
        if self.peek().kind != TokenKind::Kw(Kw::LParen) {
            // `Input` is a function: its brackets are not optional.
            self.report_expected(tok.span, "(");
            return base;
        }
        // Nothing joins to the result of `Input(n, file)`.
        self.parse_input_fn_call(arena, base)
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
        // Only the chain that is the target of the statement can end in a graphics word; what is
        // read inside its brackets is expression text.
        let chain = self.graphics_chain.take();
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
                    let graphics = match chain {
                        Some(_) => self.graphics_word_after_dot(arena, base),
                        None => None,
                    };
                    if let Some(word) = graphics {
                        // The word ends the chain: the rest of the statement is its operands.
                        let member = self.advance();
                        let mid = self.graphics_member_symbol(&member);
                        base = arena.alloc(ExprNode::MemberAccess { base, member: mid, bang: false, bracketed: false });
                        self.set_span(base, member.span);
                        self.graphics_stop = Some(word);
                        break;
                    }
                    base = self.parse_member_access(arena, base, false);
                }
                TokenKind::Kw(Kw::Bang) => {
                    self.advance();
                    // Unlike the dot, the bang takes its key with no blanks between.
                    if self.blank_before_next() {
                        let span = self.current_span();
                        self.diagnostics.push(ERR_SYNTAX_ERROR, span);
                    }
                    base = self.parse_member_access(arena, base, true);
                }
                TokenKind::Kw(Kw::LParen) => {
                    self.advance();
                    self.later_group = matches!(arena.get(base), ExprNode::Call { .. });
                    let args = self.parse_arg_list(arena);
                    self.expect(TokenKind::Kw(Kw::RParen), ERR_EXPECTED_RPAREN);
                    base = arena.alloc(ExprNode::Call { func: base, args });
                }
                _ => break,
            }
        }
        base
    }

    /// The member name that follows a `.` or `!` after `base`, as a member access of it.
    fn parse_member_access(&mut self, arena: &mut ExprArena, base: NodeId, bang: bool) -> NodeId {
        let member = self.expect_ident();
        let mid = member.sym.map(|s| s as u32).unwrap_or(0);
        let access = arena.alloc(ExprNode::MemberAccess { base, member: mid, bang, bracketed: self.is_bracketed(&member) });
        self.set_span(access, member.span);
        access
    }

    /// The qualified name of a procedure after `AddressOf`: names joined by dots.
    fn parse_procedure_name(&mut self, arena: &mut ExprArena) -> NodeId {
        // A bracketed name is read as written: it names a procedure declared elsewhere.
        self.bracketed_name_ok = self.peek().bracketed;
        let first = self.expect_decl_name();
        self.bracketed_name_ok = false;
        let sym = first.sym.map(|s| s as u32).unwrap_or(0);
        let mut base = arena.alloc(ExprNode::NameRef { sym, suffix: first.type_suffix });
        self.set_span(base, first.span);
        while self.peek().kind == TokenKind::Kw(Kw::Dot) && !self.blank_before_next() {
            self.advance();
            base = self.parse_member_access(arena, base, false);
        }
        base
    }

    /// Reads the call-site `ByVal` that may come before an argument's value; an `AddressOf`
    /// may start the value only when no `ByVal` came before it.
    fn argument_operand_prefix(&mut self) -> bool {
        if self.peek().kind == TokenKind::Kw(Kw::ByVal) && self.group_depth == 1 && self.stmt_byval.is_none() {
            self.stmt_byval = Some(self.current_span());
        }
        if self.eat(TokenKind::Kw(Kw::ByVal)) {
            self.address_of_ok = false;
            if self.at_argument_end() {
                let span = self.current_span();
                self.report_expected(span, "expression");
            }
            true
        } else {
            self.allow_address_of_here();
            false
        }
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
            if self.graphics_stop.is_some() {
                // `.?` ended the chain.
                self.graphics_chain = None;
                return access;
            }
            return self.parse_postfix(arena, access);
        }
        // A target names something: a reserved word does not (`Me` and the few reserved
        // words that also name callable things apart).
        let at_start = std::mem::take(&mut self.target_at_statement_start);
        if self.peek().kind == TokenKind::Kw(Kw::Rem) {
            // A comment cannot stand where a target belongs.
            self.error_and_skip_line(ERR_SYNTAX_ERROR);
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
            self.error_and_skip_line(ERR_SYNTAX_ERROR);
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
        if is_number_or_string_literal(&self.peek().kind)
            || matches!(
                self.peek().kind,
                TokenKind::Kw(Kw::True)
                    | TokenKind::Kw(Kw::False)
                    | TokenKind::Kw(Kw::Nothing)
                    | TokenKind::Kw(Kw::Empty)
                    | TokenKind::Kw(Kw::Null)
            )
        {
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
        // At the start of a statement a question mark after the dot is the `Print` method of the
        // object of the `With`, and ends the chain.
        let question = !bang
            && self.graphics_chain == Some(ChainStart::Statement)
            && self.peek().kind == TokenKind::Kw(Kw::Question);
        let member = if question { self.advance() } else { self.expect_ident() };
        let mid = if question { self.graphics_member_symbol(&member) } else { member.sym.map(|s| s as u32).unwrap_or(0) };
        if question {
            self.graphics_stop = Some(GraphicsWord::Print);
        }
        let base = arena.alloc(ExprNode::WithContext);
        self.set_span(base, dot_span);
        let access = arena.alloc(ExprNode::MemberAccess { base, member: mid, bang, bracketed: self.is_bracketed(&member) });
        self.set_span(access, member.span);
        Some(access)
    }

    // ── Argument lists ────────────────────────────────────────────────────────

    /// Parses a comma-separated list of arguments.
    /// True when no argument expression follows: the list, the bracket or the statement ends.
    fn at_argument_end(&mut self) -> bool {
        self.peek().is_stmt_end()
            || matches!(
                self.peek().kind,
                TokenKind::Kw(Kw::Comma) | TokenKind::Kw(Kw::RParen) | TokenKind::Kw(Kw::Colon) | TokenKind::Kw(Kw::Else)
            )
    }

    fn parse_arg_list(&mut self, arena: &mut ExprArena) -> NodeId {
        // The arguments of a call are expressions, even when the call is itself a target.
        let was_target = std::mem::replace(&mut self.in_target, false);
        self.group_depth += 1;
        let list = self.parse_arg_list_items(arena);
        self.group_depth -= 1;
        self.in_target = was_target;
        list
    }

    fn parse_arg_list_items(&mut self, arena: &mut ExprArena) -> NodeId {
        if self.peek().kind == TokenKind::Kw(Kw::RParen) {
            return arena.alloc(ExprNode::ArgList { args: vec![] });
        }
        let mut args = Vec::new();
        let mut seen_named = false;
        // A list that follows another takes neither an omitted argument nor a `ByVal`.
        let later = std::mem::take(&mut self.later_group);
        loop {
            let Some(arg) = self.parse_arg_list_item(arena, later, &mut seen_named) else { break };
            args.push(arg);
            if !self.eat(TokenKind::Kw(Kw::Comma)) { break; }
        }
        arena.alloc(ExprNode::ArgList { args })
    }

    /// One argument of a bracketed list. `None` when a comma was the last thing before the
    /// bracket, which is reported.
    fn parse_arg_list_item(&mut self, arena: &mut ExprArena, later: bool, seen_named: &mut bool) -> Option<NodeId> {
        // Missing argument.
        if self.peek().kind == TokenKind::Kw(Kw::Comma) {
            if later {
                let span = self.current_span();
                self.report_expected(span, "expression");
            }
            return Some(arena.alloc(ExprNode::MissingArg));
        }
        if self.peek().kind == TokenKind::Kw(Kw::RParen) {
            // A comma was the last thing before the bracket.
            let span = self.current_span();
            self.report_expected(span, "expression");
            return None;
        }
        if let Some(named) = self.try_parse_named_arg(arena) {
            *seen_named = true;
            return Some(named);
        }
        if *seen_named {
            // After a named argument every argument is named.
            let span = self.current_span();
            self.report_expected(span, "named argument");
        }
        if later && self.peek().kind == TokenKind::Kw(Kw::ByVal) {
            let span = self.current_span();
            self.report_expected(span, "expression");
        }
        Some(self.parse_argument_value(arena))
    }

    /// The value of an argument, after any call-site `ByVal`, which is noted.
    fn parse_argument_value(&mut self, arena: &mut ExprArena) -> NodeId {
        let by_val = self.argument_operand_prefix();
        let value = self.parse_expr(arena, 0);
        if by_val {
            self.by_val_args.push(value);
        }
        value
    }

    /// Argument list on a call statement line (no surrounding parens).
    /// An omitted argument (`MsgBox s, , title`) is a `MissingArg`.
    fn parse_stmt_arg_list(&mut self, arena: &mut ExprArena) -> Vec<NodeId> {
        let mut args = Vec::new();
        let mut seen_named = false;
        while !self.peek().is_stmt_end() {
            if self.peek().kind == TokenKind::Kw(Kw::Else) {
                break;
            }
            let Some(a) = self.parse_stmt_arg(arena, &mut seen_named) else { break };
            args.push(a);
            if !self.eat(TokenKind::Kw(Kw::Comma)) { break; }
            // A comma needs another argument after it, empty or not, before the line ends.
            if self.at_call_arguments_end() {
                let span = self.current_span();
                self.report_expected(span, "expression");
                break;
            }
        }
        args
    }

    /// Whether the arguments of a call statement end here: the statement ends, or a colon or
    /// an `Else` follows.
    fn at_call_arguments_end(&mut self) -> bool {
        self.peek().is_stmt_end()
            || matches!(self.peek().kind, TokenKind::Kw(Kw::Colon) | TokenKind::Kw(Kw::Else))
    }

    /// One argument of a call statement. `None` when no expression follows.
    fn parse_stmt_arg(&mut self, arena: &mut ExprArena, seen_named: &mut bool) -> Option<NodeId> {
        if self.peek().kind == TokenKind::Kw(Kw::Comma) {
            return Some(arena.alloc(ExprNode::MissingArg));
        }
        if let Some(named) = self.try_parse_named_arg(arena) {
            *seen_named = true;
            return Some(named);
        }
        if *seen_named {
            // After a named argument every argument is named.
            let span = self.current_span();
            self.report_expected(span, "named argument");
        }
        let by_val = self.argument_operand_prefix();
        let expr = self.try_parse_expr(arena)?;
        if by_val {
            self.by_val_args.push(expr);
        }
        Some(expr)
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
        if self.group_depth == 1 && self.stmt_named.is_none() {
            self.stmt_named = Some(name_tok.span);
        }
        let name = name_tok.sym.map(|s| s as u32).unwrap_or(0);
        self.advance(); // `:=`
        let value = self.parse_argument_value(arena);
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
/// Most `#If` blocks that may be open at once.
const MAX_CC_NESTING: usize = 4079;
/// Most dimensions an array declaration may list.
const MAX_DIMENSIONS: usize = 60;

const ERR_EXPECTED_EXPR:        u32 = 0x9c6f;
const ERR_EXPECTED_IDENT:       u32 = 0x9c6f;
const ERR_OBJECT_MODULE_ONLY:   u32 = 0xdee1;
const ERR_UNEXPECTED_TOKEN:     u32 = 0x9c6f;
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
/// Whether the token is a number, string or date literal.
fn is_number_or_string_literal(kind: &TokenKind) -> bool {
    matches!(
        kind,
        TokenKind::IntLit
            | TokenKind::LongLit
            | TokenKind::SngLit
            | TokenKind::DblLit
            | TokenKind::CurLit
            | TokenKind::StrLit
            | TokenKind::DateLit
    )
}

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

/// The keywords that start a statement of their own, the ones a statement is chosen by when a
/// line begins with one. Every other keyword is an ordinary name at the start of a statement.
fn begins_own_statement(kw: Kw) -> bool {
    use Kw::*;
    matches!(
        kw,
        Public | Private | Global | Dim | Static | Const | ReDim | If | For | Do | While | Select | With
            | Let | Set | Call | Exit | Return | GoTo | GoSub | Go | Resume | On | Error | Stop | End
            | Else | ElseIf | EndIf | Next | Loop | Wend | Case | Erase | RaiseEvent | Debug | Attribute
            | Open | Close | Print | Write | Input | InputS | InputBS | Line | LineInput | Friend | WithEvents | Get | Put | Seek
            | Lock | Unlock | Width | Name | Mid | MidS | MidB | MidBS | LSet | RSet | Question | Property
            | Scale | Implements | Sub | Function | Declare | Event | Option | Type | Enum | Rem | Apos
            | PSet | Circle | DefBool | DefByte | DefCur | DefDate | DefDbl | DefInt | DefLng | DefObj
            | DefSng | DefStr | DefVar
    )
}

/// The words that name module-header attributes. They are keywords only
/// at the start of an `Attribute` line and may be used as ordinary names.
fn is_header_word(name: &str) -> bool {
    name.eq_ignore_ascii_case("attribute") || name.len() > 3 && name[..3].eq_ignore_ascii_case("vb_")
}

/// Keywords that are not reserved but begin a statement of their own: used as a
/// name they need something after them that says so.
const STATEMENT_WORDS: [&str; 10] = ["Width", "Property", "Name", "Mid", "MidB", "Mid$", "MidB$", "Line", "Go", "Error"];

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

