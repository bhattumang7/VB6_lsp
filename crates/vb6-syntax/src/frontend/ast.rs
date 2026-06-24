//! Typed AST node types for the VB6 compiler frontend.
//!
//! VB6 uses two node families:
//!
//! * **Expression nodes** — fixed 40-byte, bump-allocated by `alloc_ast_node`.
//!   Layout: word[0] = opcode (low 16 bits), word[1] = flags, word[4] = left
//!   child, word[5] = right child.
//!
//! * **Scope-container nodes** — variable size, allocated via `ast_node_create`.
//!   Sizes: 0x00→0, 0x01-0x10→8, 0x11→28, 0x12→20, 0x13→64, 0x14→36, 0x15→24,
//!   0x16→12, 0x17→18, 0x18→16, 0x19→12, 0x1a→32.
//!
//! * **Declaration/symbol nodes** — variable size, allocated via
//!   `alloc_decl_node`.
//!
//! Here we replace the raw-offset layouts with typed Rust structs. The `NodeId`
//! index (from `support::arena`) replaces the raw `u32 *` node pointer.

pub use crate::support::arena::{Arena, NodeId};
pub use crate::frontend::token::{Span, TypeSuffix};
pub use crate::frontend::diagnostics::Diagnostics;
use vb6_ast_derive::Children;

// ── Typed node enums ──────────────────────────────────────────────────────────

/// Literal value kinds produced by the parser (analogous to scanner `Lit` but
/// extended with Bool/Empty/Null/Date that the scanner emits as keywords).
#[derive(Debug, Clone, PartialEq)]
pub enum AstLit {
    Int(i32),
    Long(i32),
    Single(f32),
    Double(f64),
    Currency(i64),
    Str(Box<str>),
    Date(f64),
    Bool(bool),
    Empty,
    Null,
}

/// Binary operator kinds for [`ExprNode::BinOp`].
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum BinOpKind {
    Add, Sub, Mul, Div, IDiv, Mod, Pow, Cat,
    Eq, Ne, Lt, Gt, Le, Ge, Like, Is, IsNot,
    And, Or, Xor, Eqv, Imp,
    /// `.member` access (infix form, used when base is a non-Ident primary like `Me`).
    Dot,
    /// `!member` access (infix form).
    Bang,
}

/// Unary operator kinds for [`ExprNode::UnOp`].
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum UnOpKind { Neg, Pos, Not }

/// Procedure kind for [`ExprNode::ProcDecl`].
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum ProcKind { Sub, Function, PropGet, PropLet, PropSet }

/// Do-loop variant for [`ExprNode::Do`].
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum DoKind {
    Inf,       // Do ... Loop
    PreWhile,  // Do While ... Loop
    PreUntil,  // Do Until ... Loop
    PostWhile, // Do ... Loop While
    PostUntil, // Do ... Loop Until
}

/// Exit target for [`ExprNode::ExitStmt`].
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum ExitKind { Sub, Function, For, Do, Property }

/// A jump destination for `GoTo`/`GoSub`/`On Error GoTo`/`Resume`, and the
/// definition site for a line label.
///
/// VB6 line labels come in two forms: a named label (`Done:`) and a numeric
/// line label (`100`). Both are valid jump targets, so the destination must be
/// able to carry either an interned name symbol or the line-number value — the
/// value of a numeric target is never folded away to a sentinel.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum LabelRef {
    /// Named label — interned symbol id of the label identifier.
    Name(u32),
    /// Numeric line label — the line-number value (e.g. `GoTo 100`).
    Line(i32),
}

/// On Error variant for [`ExprNode::OnError`].
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum OnErrorKind {
    Goto(LabelRef), // On Error GoTo label / line#
    ResumeNext,     // On Error Resume Next
    Disable,        // On Error GoTo 0
}

/// Target of a `Resume` statement — the four VB6 forms.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum ResumeTarget {
    /// `Resume` — retry the faulting statement.
    Retry,
    /// `Resume Next` — continue at the statement after the fault.
    Next,
    /// `Resume label` / `Resume line#` — jump to a label or numeric line.
    /// Note: `Resume 0` is parsed as [`LabelRef::Line(0)`] and is semantically
    /// equivalent to [`ResumeTarget::Retry`]; the syntactic form is preserved.
    At(LabelRef),
}

/// An expression-class AST node.
///
/// The opcode values cited are the VB6 constants stored in word[0] of the
/// 40-byte raw layout.
///
/// [`Children`] derives [`ExprNode::for_each_child`], the single authoritative
/// enumeration of a node's child `NodeId`s. Every tree-walk (binding, the
/// resolution-coverage check, …) goes through it, so a child can never be
/// silently missed. A field is a child edge iff its type is `NodeId`,
/// `Option<NodeId>`, or `Vec<NodeId>`; all other fields are scalar payload.
#[derive(Children)]
pub enum ExprNode {
    /// Generic expression node — the direct output of `alloc_ast_node`.
    ///
    /// Covers all opcodes not given a typed variant.  Nodes are stored as
    /// 40-byte raw records: `word[0]` = opcode (low 16 bits), `word[1]` lo =
    /// `flags`, `word[4]` = `lhs`, `word[5]` = `rhs`.  `lhs`/`rhs` are raw
    /// `u32` because they carry either a `NodeId` index (when pointing at a
    /// child node) or a raw scalar (token id, enum tag, bound-list ptr, etc.)
    /// depending on the opcode.
    Generic {
        /// Opcode in the low 16 bits of word[0], stored directly.
        opcode: u16,
        /// Flags in the low 16 bits of word[1].
        flags: u16,
        /// Left child / primary operand (word[4]).  Interpretation depends on
        /// opcode: child `NodeId` index, raw token id, or enum discriminant.
        lhs: u32,
        /// Right child / secondary operand (word[5]).
        rhs: u32,
    },

    /// `For` loop node — opcode **0x86** with an inline range sub-node.
    ///
    /// VB6 packs this into a standard 40-byte node plus word[6] for the step,
    /// with a nested 0x7b range sub-node at word[5].  Here we lift both into
    /// named fields for clarity.
    ///
    /// * `loop_var` (word[4]): the loop-control-variable expression node id.
    /// * `range` (word[5]): id of the 0x7b range node (start To end).
    /// * `step` (word[6]): optional step expression node id (0 = absent).
    ForRange {
        loop_var: u32,
        range: u32,
        step: u32,
    },

    /// Opcode **0xb3** — built-in type specifier.
    ///
    /// The UDT variant (`make_udt_type_node`) uses the same opcode but adds
    /// flag bit `0x8000` in `type_flags`.
    ///
    /// Layout (40-byte node):
    /// * byte[0] = 0xb3 (opcode)
    /// * word[1] (bytes 4–5) = `type_flags`
    /// * word[4] lo (bytes 16–17) = `parent_scope`
    /// * word[5] (bytes 20–23) = `type_kind`
    /// * word[6] (bytes 24–27) = child node ptr (0 = absent)
    TypeSpec {
        /// Modifier flags: `0x4000` = fixed-length `String * n`,
        /// `0x8000` = user-defined type (UDT) path.
        type_flags: u16,
        /// Parent scope identifier, threaded through from the parser.
        parent_scope: u16,
        /// 5-bit VBA built-in type kind (0 = none, 2 = Integer,
        /// 3 = Long, 8 = String, 0x10 = Object, …).
        type_kind: u32,
        /// Optional child node (e.g. the qualified `Module.TypeName` ref
        /// produced by `make_type_node` when the type is a qualified name).
        child: Option<NodeId>,
    },

    /// Opcode **0xb3**, flag bit `0x8000` set — user-defined type specifier.
    ///
    /// Shares opcode 0xb3 with [`ExprNode::TypeSpec`] but stores extra UDT
    /// metadata in the high half of word[4].
    ///
    /// Layout (40-byte node):
    /// * byte[0] = 0xb3 (opcode)
    /// * word[1] (bytes 4–5) = `flags | 0x8000`
    /// * word[4] lo (bytes 16–17) = `parent_scope`
    /// * word[4] hi (bytes 18–19) = `udt_count`
    /// * word[5] (bytes 20–23) = type-list node ptr
    UdtTypeSpec {
        /// Modifier flags OR'd with `0x8000` (UDT marker).
        flags: u16,
        /// word[4] lo (bytes 16–17): parent scope identifier.
        parent_scope: u16,
        /// word[4] hi (bytes 18–19): count of UDT member fields.
        udt_count: u16,
        /// word[5] (bytes 20–23): head of the UDT field-type list.
        type_list: Option<NodeId>,
    },

    // ── Parser-produced typed nodes (opcode equivalents in comments) ─────────

    /// Literal value — corresponds to opcode `LIT` (0x60).
    Literal { lit: AstLit },

    /// Name reference (variable/function identifier) — opcode `NAME` (0x61).
    /// `suffix` is the type-declaration character on this use (`count%`), which
    /// is carried into the name's type node.
    NameRef { sym: u32, suffix: TypeSuffix },

    /// `Me` keyword — current class instance; opcode `ME` (0x62).
    Me,

    /// `Nothing` keyword — null object reference; opcode `NOTHING` (0x65).
    Nothing,

    /// Implicit base of a leading-dot/bang member reference inside a `With`
    /// block (`.Member` / `!Key`). It denotes the innermost active `With`
    /// object, which VB6 resolves at compile/run time against the With-block
    /// stack. Modeled as an explicit base so the leading `.`/`!` becomes a real
    /// `MemberAccess` rather than collapsing into a bare `NameRef` (which would
    /// drop the dot).
    WithContext,

    /// Parenthesised expression — opcode `PAREN` (0x07).
    Paren { inner: NodeId },

    /// Binary operation — opcodes `ADD`/`SUB`/… (0x04–0x1d).
    BinOp { op: BinOpKind, lhs: NodeId, rhs: NodeId },

    /// Unary operation — opcodes `UNEG`/`UPOS`/`NOT` (0x03/0x06/0x0f).
    UnOp { op: UnOpKind, operand: NodeId },

    /// Member access: `base.member` or `base!member` — opcodes `DOT`/`BANG` (0x22/0x23).
    MemberAccess { base: NodeId, member: u32, bang: bool },

    /// `AddressOf proc` — function-pointer operand (token 3, parsed via the
    /// general keyword path).  `operand` is the (possibly qualified) procedure
    /// name.  The node opcode is O3.
    AddressOf { operand: NodeId },

    /// Call/index expression: `func(args)` — opcode `CALL` (0x25).
    Call { func: NodeId, args: NodeId },

    /// Implicit/Let assignment — opcode `ASSIGN` (0x8f).
    Assign { target: NodeId, value: NodeId },

    /// Object assignment: `Set target = value` — opcode `SET_ASSIGN` (0x90).
    SetAssign { target: NodeId, value: NodeId },

    /// `LSet`/`RSet` range-copy assignment: `LSet target = value`.
    /// Parsed as a `0x2c` node with a flag distinguishing the two forms:
    /// `0x4000` = `LSet` (left-justify), `0x2000` = `RSet`
    /// (right-justify).  This is NOT a plain `Assign` — the justification side is
    /// semantically significant (it pads/truncates the fixed-length string target
    /// differently), so it is preserved here as `right_justify`.
    RangeAssign { right_justify: bool, target: NodeId, value: NodeId },

    /// `Mid`/`Mid$`/`MidB`/`MidB$` string-replacement statement:
    /// `Mid(s, start[, len]) = value`.  All four spellings are parsed in one
    /// handler as a `0x56` node.  The `variant` flag encodes
    /// the spelling and is semantically significant: bit `0x4000` selects the
    /// **byte-oriented** `MidB` family (vs character-oriented `Mid`); bit `0x8000`
    /// marks the `$` (explicit-string) spelling.  Collapsing these would lose the
    /// char-vs-byte distinction, so we preserve `byte_oriented` and `dollar`.
    /// `args` is the `(s, start[, len])` tuple; the arg count distinguishes the
    /// 2-arg and 3-arg forms.  Distinct from an `=` comparison expression.
    MidAssign { byte_oriented: bool, dollar: bool, args: NodeId, value: NodeId },

    /// Procedure declaration (Sub/Function/Property).
    /// `params` is the parameter `ArgList`; `ret_type` the return type-spec.
    ProcDecl { kind: ProcKind, name: u32, params: Option<NodeId>, ret_type: Option<NodeId>, body: NodeId },

    // ── Structural nodes ──────────────────────────────────────────────────────

    /// Sequence of statements — replaces BLOCK + STMT_LINK Generic chains.
    Block { stmts: Vec<NodeId> },
    /// Argument / parameter list — replaces LIST_HEAD Generic chains.
    ArgList { args: Vec<NodeId> },

    // ── Control-flow statement nodes ──────────────────────────────────────────

    /// `If cond Then then_body [Else else_body] End If`
    If { cond: NodeId, then_body: NodeId, else_body: Option<NodeId> },
    /// `While cond ... Wend`
    While { cond: NodeId, body: NodeId },
    /// `Do [While/Until cond] ... Loop [While/Until cond]`
    Do { kind: DoKind, cond: Option<NodeId>, body: NodeId },
    /// `For var = start To end [Step step] ... Next`
    For { var: NodeId, start: NodeId, end: NodeId, step: Option<NodeId>, body: NodeId },
    /// `For Each var In collection ... Next`
    ForEach { var: NodeId, collection: NodeId, body: NodeId },
    /// `With obj ... End With`
    With { obj: NodeId, body: NodeId },
    /// `Select Case subject ... End Select`; `cases` = Vec of CaseBlock/CaseElse nodes.
    /// `pre` holds any statements (e.g. `Dim`) appearing between the subject and
    /// the first `Case`, which VB6 tolerates; preserved here rather than
    /// discarded.
    SelectCase { subject: NodeId, pre: Vec<NodeId>, cases: Vec<NodeId> },
    /// `Case test-list ... body`
    CaseBlock { test: NodeId, body: NodeId },
    /// `Case Else ... body`
    CaseElse { body: NodeId },

    // ── Jump / control-transfer statement nodes ───────────────────────────────

    /// `GoTo label` / `GoTo line#`
    GoTo { target: LabelRef },
    /// `GoSub label` / `GoSub line#`
    GoSub { target: LabelRef },
    /// `Return` (return from GoSub)
    ReturnStmt,
    /// `On Error GoTo / Resume Next / GoTo 0`
    OnError { kind: OnErrorKind },
    /// `Resume` / `Resume Next` / `Resume label` / `Resume line#`
    Resume { target: ResumeTarget },
    /// `Stop`
    Stop,
    /// `End` (terminate program)
    EndStmt,
    /// `On expr GoTo / GoSub label-list`
    OnGo { is_gosub: bool, expr: NodeId, labels: Vec<NodeId> },
    /// `Exit Sub/Function/For/Do/Property`
    ExitStmt { kind: ExitKind },

    // ── Other statement nodes ─────────────────────────────────────────────────

    /// `Call callee(args)` — explicit Call keyword.
    CallStmt { callee: NodeId, args: NodeId },
    /// `Erase var1[, var2...]`
    Erase { vars: Vec<NodeId> },
    /// `RaiseEvent name[(args)]`
    RaiseEvent { name: u32, args: NodeId },
    /// `Debug.Print [args...]`
    DebugPrint { args: Vec<NodeId> },
    /// `Error expr` — set current error number.
    ErrorStmt { expr: NodeId },

    // ── Declaration item nodes ────────────────────────────────────────────────

    /// Single variable/constant in a Dim/Const/Static statement.
    /// `bounds` = array-bounds ArgList (None = scalar); `type_node` = type-spec.
    DimItem { name: u32, is_const: bool, is_static: bool, bounds: Option<NodeId>, type_node: Option<NodeId>, init: Option<NodeId> },
    /// Single array variable in a ReDim statement.
    ReDimItem { preserve: bool, name: u32, bounds: Option<NodeId>, type_node: Option<NodeId> },
    /// Parameter definition in a procedure/event declaration.
    /// `default` is the `Optional` default-value expression, if any.
    ParamDef { flags: u16, name: u32, type_node: Option<NodeId>, default: Option<NodeId> },
    /// `Type name ... End Type` declaration.
    TypeDecl { name: u32, members: Vec<NodeId> },
    /// `Enum name ... End Enum` declaration.
    EnumDecl { name: u32, members: Vec<NodeId> },
    /// `Event name(params)` declaration.
    EventDecl { name: u32, params: NodeId },
    /// `Implements InterfaceName`
    Implements { name: u32 },

    // ── Type-specifier nodes (parser-level; distinct from sema TypeSpec/UdtTypeSpec) ──

    /// Built-in type keyword: Integer=2, Long=3, Single=4, Double=5, Byte=17,
    /// Boolean=11, Variant=12, Object=9, Decimal=14, Currency=6.
    BuiltinType { kind: u32 },
    /// `String` or `String * n` type specifier.
    StringType { fixed_len: Option<NodeId> },
    /// User-defined or qualified type name: `TypeName` or `Module.TypeName`.
    UserType { name: u32, child: Option<NodeId> },

    // ── Additional expression nodes ───────────────────────────────────────────

    /// `lo To hi` — For loop bounds and Select Case ranges.
    RangeTo { lo: NodeId, hi: NodeId },
    /// `TypeOf expr Is typeName`
    TypeOf { expr: NodeId, type_spec: NodeId },
    /// `New TypeName`
    New { type_spec: NodeId },
    /// `Is <op> <expr>` arm of a Case clause (e.g. `Case Is > 5`).
    CaseIs { op: BinOpKind, rhs: NodeId },
    /// Absent argument in a call arg list (consecutive commas).
    MissingArg,

    // ── New nodes for D-P-NamedArg, D-P-Option, D-P-Declare, D-P-FileIO ─────

    /// Named argument `name:=value` in a function or method call.
    NamedArg { name: u32, value: NodeId },

    /// `Option Explicit` module-level directive.
    OptionExplicit,
    /// `Option Base 0|1` — default array lower bound (0 or 1).
    OptionBase { value: u8 },
    /// `Option Compare Binary|Text|Database` — string comparison mode.
    /// mode: 0=Binary, 1=Text, 2=Database.
    OptionCompare { mode: u8 },

    /// `Declare [Function|Sub] name Lib "lib" [Alias "alias"] ([params]) [As type]`
    DeclareDecl {
        kind: ProcKind,
        name: u32,
        lib: NodeId,
        alias: Option<NodeId>,
        params: Option<NodeId>,
        ret_type: Option<NodeId>,
    },

    /// File I/O statement (Open, Close, Print #, Write #, Input #, etc.)
    FileIoStmt { kind: FileIoKind, channel: Option<NodeId>, args: Vec<NodeId> },

    /// Line label definition at the start of a statement: a named label
    /// (`Foo:`) or a numeric line label (`100`). `target` carries the interned
    /// label symbol or the line-number value, so a numeric label is preserved
    /// as a distinct jump target rather than collapsed to a sentinel.
    Label { target: LabelRef },

    /// `Def<Type> letter[-letter][, …]` module-level default-type declaration
    /// (`DefInt A-Z`, `DefBool B`, …).
    /// `type_kw` is the keyword token id (e.g. `Kw::DefInt as u16`); `ranges`
    /// holds the interned start/end letter symbols for each range (end == 0 for
    /// a single-letter range).  Full letter-range modelling is deferred to sema.
    DefType { type_kw: u16, ranges: Vec<(u32, u32)> },

    /// `Attribute name = value[, …]` metadata line.  Hosts embed these in
    /// module text (`Attribute VB_Name = "Module1"`).  `name` is the interned
    /// attribute-key symbol; `values` are the literal value expressions.
    Attribute { name: u32, values: Vec<NodeId> },
}

/// Kind of file I/O statement for [`ExprNode::FileIoStmt`].
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum FileIoKind {
    Open, Close, Print, Write, Input, LineInput,
    Get, Put, Seek, Lock, Unlock, Width, Name,
}

/// Bump arena for expression-class AST nodes.
///
/// These nodes are allocated from a bump arena hung off the compiler context.
/// Here we use an `Arena<ExprNode>` backed by a `Vec`.
pub type ExprArena = Arena<ExprNode>;

/// Source-span side table for AST nodes, keyed by `NodeId.0`.
///
/// VB6's AST nodes do not retain source positions (a compiler does not
/// need them for codegen). The LSP layer does, so the parser records the
/// source span of the nodes that name a symbol — `NameRef` use sites and the
/// declared-name identifier of each declaration — into this parallel table.
/// Keeping spans out of the [`ExprNode`] structs preserves their
/// compact layout; this table is host metadata, not part of the
/// modeled node.
///
/// Entries default to [`Span::DUMMY`] for any node whose span was not
/// recorded (operators, literals, calls, …), which the LSP layer never
/// queries by position.
#[derive(Debug, Default, Clone)]
pub struct NodeSpans {
    spans: Vec<Span>,
}

impl NodeSpans {
    pub fn new() -> Self {
        Self { spans: Vec::new() }
    }

    /// Record the source span of `id`, growing the table as needed.
    pub fn set(&mut self, id: NodeId, span: Span) {
        let i = id.0 as usize;
        if i >= self.spans.len() {
            self.spans.resize(i + 1, Span::DUMMY);
        }
        self.spans[i] = span;
    }

    /// Return the recorded span for `id`, or [`Span::DUMMY`] if none.
    pub fn get(&self, id: NodeId) -> Span {
        self.spans.get(id.0 as usize).copied().unwrap_or(Span::DUMMY)
    }
}

/// Allocates a type-reference AST node.
///
/// * `node_kind`: stored in the low 16 bits of word[0] (the opcode field).
/// * `flags`: stored in `word[1]` lo; if `is_qualified`, bit `0x8000` is OR'd in
///   to mark a qualified-name reference (e.g. `Module.TypeName`).
/// * `p_type_desc` (word[4]): raw pointer/id for the type descriptor.
/// * `type_aux` (word[5]): auxiliary type data.
///
/// When `p_type_desc` is [`VARIANT_TYPE_DESC_MARKER`], `type_aux` is normalized
/// to `5` (the canonical Variant type identifier).
pub const VARIANT_TYPE_DESC_MARKER: u32 = 0x3f;

pub fn make_type_node(
    arena: &mut ExprArena,
    node_kind: u16,
    flags: u16,
    p_type_desc: u32,
    type_aux: u32,
    is_qualified: bool,
) -> NodeId {
    let adjusted_flags = if is_qualified { flags | 0x8000 } else { flags };
    let normalized_type_aux = if p_type_desc == VARIANT_TYPE_DESC_MARKER { 5 } else { type_aux };
    alloc_ast_node(arena, node_kind, adjusted_flags, p_type_desc, normalized_type_aux)
}

/// Creates a built-in type-specifier node (opcode 0xb3) and returns its id.
///
/// `type_flags`: modifier flags (0 = bare type, 0x4000 = fixed-length String).
/// `parent_scope`: parent scope id threaded down from the declaration parser.
/// `type_kind`: 5-bit VBA built-in type enumeration.
/// `child`: optional child node (qualified module.type reference, if any).
pub fn make_type_spec_node(
    arena: &mut ExprArena,
    type_flags: u16,
    parent_scope: u16,
    type_kind: u32,
    child: Option<NodeId>,
) -> NodeId {
    arena.alloc(ExprNode::TypeSpec { type_flags, parent_scope, type_kind, child })
}

/// Creates a user-defined type specifier node (opcode 0xb3, flag `0x8000`).
///
/// `fixed_len_flags`: additional modifier flags (e.g. `0x4000` for fixed-length
///   `String * n`). The `0x8000` UDT bit is added automatically.
/// `parent_scope`: parent scope identifier threaded from the declaration parser.
/// `udt_count`: number of UDT member fields.
/// `type_list`: head node of the field-type list (`None` = no fields yet).
///
/// Appends a node to a list. In VB6, list nodes are built as raw chain nodes
/// using opcodes `0xbb` (append-with-new-head) or `0x7c` (extend-existing-chain).
/// Here we use `Vec<NodeId>` instead — the list is a plain vec and
/// `append_list_node` is just `Vec::push`.
pub fn append_list_node(list: &mut Vec<NodeId>, node: NodeId) {
    list.push(node);
}

pub fn make_udt_type_node(
    arena: &mut ExprArena,
    fixed_len_flags: u16,
    parent_scope: u16,
    udt_count: u16,
    type_list: Option<NodeId>,
) -> NodeId {
    arena.alloc(ExprNode::UdtTypeSpec {
        flags: fixed_len_flags | 0x8000,
        parent_scope,
        udt_count,
        type_list,
    })
}

/// Builds an abstract member node (opcode `0xbb` — list/chain node), adjusting
/// the flags field based on a qualifying discriminant.
///
/// Behaviour:
/// * Creates a 0xbb node with `flags` / `lhs` / `rhs`.
/// * `qual_flag == 0` → sets bit 2 (0x04) in the low byte of word[1] (`flags`).
/// * `qual_flag == 4` → sets bit 7 (0x80) in the high byte of word[1] (`flags`).
/// * VB6 records a syntax error of type `0x9c6f` when creating this node;
///   the diagnostic is pushed into `diag` at `span`.
pub fn build_abstract_member_node(
    arena: &mut ExprArena,
    flags: u16,
    lhs: u32,
    rhs: u32,
    qual_flag: u32,
    diag: &mut Diagnostics,
    span: Span,
) -> NodeId {
    diag.push(0x9c6f, span);
    let adjusted_flags = match qual_flag {
        0 => flags | 0x04,
        4 => flags | 0x8000,
        _ => flags,
    };
    alloc_ast_node(arena, 0xbb, adjusted_flags, lhs, rhs)
}

/// Creates a leaf (operand-less) expression node with the given opcode.
///
/// In VB6 this first consumes a scanner token and then allocates the node. The
/// consume is a scanner-context operation that is omitted here.
pub fn emit_simple_node(arena: &mut ExprArena, opcode: u16) -> NodeId {
    alloc_ast_node(arena, opcode, 0, 0, 0)
}

/// Builds a `For`-loop AST node (opcode `0x86`).
///
/// Layout (word indices of the 40-byte node):
/// * word[0] lo = opcode 0x86
/// * word[1] = 0 (no flags)
/// * word[4] = `loop_var` — the loop-variable target expression node
/// * word[5] = a `0x7b` range node built from `start`/`end`
/// * word[6] = `step` — optional step expression (0 = absent)
///
/// The range sub-node (`0x7b`) is a nested `Generic` entry in the same
/// `ExprArena`; its id is stored in `rhs`.  The extra word[6] for the step is
/// stored in `ExprNode::ForRange` rather than squeezing it into the generic
/// 40-byte layout.
pub fn build_for_node(arena: &mut ExprArena, loop_var: u32, start: u32, end: u32, step: u32) -> NodeId {
    let range_id = alloc_ast_node(arena, 0x7b, 0, start, end);
    arena.alloc(ExprNode::ForRange { loop_var, range: range_id.0, step })
}

/// Builds an `On Error`/`On … GoTo` AST node (opcode `0x82`).
///
/// Layout:
/// * word[0] lo = opcode 0x82
/// * word[1] = 0
/// * word[4] = `on_kind` (0 = `Resume Next`, 1 = `GoTo 0`, etc.)
/// * word[5] = 0
pub fn build_on_node(arena: &mut ExprArena, on_kind: u32) -> NodeId {
    alloc_ast_node(arena, 0x82, 0, on_kind, 0)
}

/// Builds an array-dimension bounds node (opcode `0x7f`).
///
/// Layout:
/// * word[0]: low byte = 0x7f, stored directly.
/// * word[1] lo = `node_flags`
/// * word[4] = `bounds_list` — head of the list of `0x37` (lo To hi) bound nodes
/// * word[5] = `saved_token` — token id saved before the dimension list was parsed
pub fn build_array_bounds_node(
    arena: &mut ExprArena,
    node_flags: u16,
    bounds_list: u32,
    saved_token: u32,
) -> NodeId {
    alloc_ast_node(arena, 0x7f, node_flags, bounds_list, saved_token)
}

/// Allocates a generic 40-byte expression node.
///
/// The opcode is stored directly in the low 16 bits of word[0].
///
/// `flags`: placed in `word[1]` low half.
/// `lhs` / `rhs`: placed in `word[4]` / `word[5]`; may be a `NodeId` index or
/// a raw scalar depending on opcode.
pub fn alloc_ast_node(
    arena: &mut ExprArena,
    opcode: u16,
    flags: u16,
    lhs: u32,
    rhs: u32,
) -> NodeId {
    arena.alloc(ExprNode::Generic { opcode, flags, lhs, rhs })
}

// ---------------------------------------------------------------------------
// Scope-container node family (ast_node_create)
// ---------------------------------------------------------------------------

/// Kind discriminant for scope-container nodes created by `ast_node_create`.
///
/// The kind indexes a per-entry size table; the sizes listed in comments are
/// informational only (the Rust arena replaces the raw heap).
///
/// Named variants carry semantic labels where their role is established;
/// others are `K<hex>`.
#[derive(Clone, Copy, Debug, PartialEq, Eq)]
#[repr(u32)]
pub enum ScopeNodeKind {
    K00 = 0x00,                       // size 0
    K01 = 0x01, K02 = 0x02, K03 = 0x03, K04 = 0x04,
    K05 = 0x05, K06 = 0x06, K07 = 0x07, K08 = 0x08,
    K09 = 0x09, K0a = 0x0a, K0b = 0x0b, K0c = 0x0c,
    K0d = 0x0d, K0e = 0x0e, K0f = 0x0f, K10 = 0x10, // each: size 8
    K11 = 0x11,                       // size 28
    K12 = 0x12,                       // size 20
    /// Module scope node (size 64).  Always paired with a `ModuleList` child.
    Module = 0x13,
    K14 = 0x14, K15 = 0x15,           // size 36, 24
    /// Scope list-tail node (size 12).  Created as a paired child of
    /// `Module`; flag bit 0 is set on creation.
    ModuleList = 0x16,
    K17 = 0x17, K18 = 0x18, K19 = 0x19, // size 18, 16, 12
    K1a = 0x1a,                       // size 32
}

/// A scope-container node created by `ast_node_create`.
///
/// Replaces the raw variable-size allocation.  The `child` field is present
/// only on [`ScopeNodeKind::Module`] nodes (byte offset 32, pointing to the
/// paired `ModuleList` child).
pub struct ScopeNode {
    pub kind: ScopeNodeKind,
    /// Flags at byte 6.  Used when building predefined nodes:
    /// bit 2 (0x04) for the first predefined node, bit 3 (0x08) for the second,
    /// bit 7 (0x80) set after the first is linked.
    pub flags_b6: u8,
    /// Flags at byte 7.  Set to 1 for `ModuleList` nodes on creation.
    pub flags_b7: u8,
    /// For `Module` nodes: the paired `ModuleList` child (byte offset 32).
    pub child: Option<NodeId>,
    /// DeclNode linked to this scope by `ast_node_link` (byte offset 8).
    pub linked_decl: Option<NodeId>,
    /// Words at offsets 0x0c–0x18.
    ///
    /// For K1a (field access) and K13/Module (proc-call expr): the four words of
    /// the `AccessSpec` struct are copied here so the type system has access to
    /// them without chasing a raw pointer.
    /// For K18 (member ref): `extra[0]` holds the resolved member type, wired by
    /// `build_member_ref_node` from its `type_result` argument.
    /// Other kinds: unused (zero-initialised).
    pub extra: [u32; 4],
    /// Word at offset 0x1c.
    ///
    /// For K1a (field access): the entry from the parent DeclNode's field-type
    /// table ([`DeclNode::field_type_table`]) at slot `0xd + slot`, where `slot`
    /// is the discriminant-derived index. `0` when the parent's table has no
    /// entry at that slot.
    pub slot_info: u32,
}

/// Bump arena for scope-container nodes.
pub type ScopeArena = Arena<ScopeNode>;

/// Allocates a scope-container node of the given kind.
///
/// For [`ScopeNodeKind::Module`] (0x13): also allocates a companion
/// [`ScopeNodeKind::ModuleList`] (0x16) child and cross-links them.
/// The parent holds `child = Some(child_id)`; the back-pointer (child to
/// parent) is not stored here — it will be added when `ast_node_link`
/// requires it.
///
/// For [`ScopeNodeKind::ModuleList`] (0x16): flag bit 0 is set to 1.
pub fn ast_node_create(arena: &mut ScopeArena, kind: ScopeNodeKind) -> NodeId {
    let fresh = |kind, flags_b7, child| ScopeNode {
        kind, flags_b6: 0, flags_b7, child, linked_decl: None,
        extra: [0; 4], slot_info: 0,
    };
    match kind {
        ScopeNodeKind::Module => {
            let child_id = arena.alloc(fresh(ScopeNodeKind::ModuleList, 1, None));
            arena.alloc(fresh(ScopeNodeKind::Module, 0, Some(child_id)))
        }
        ScopeNodeKind::ModuleList => arena.alloc(fresh(ScopeNodeKind::ModuleList, 1, None)),
        _ => arena.alloc(fresh(kind, 0, None)),
    }
}

/// Creates a [`DeclNode`] (via [`alloc_decl_node`]) whose kind and flags depend
/// on the kind of `scope_node_id`:
/// * [`ScopeNodeKind::Module`] (0x13) → DeclNode kind `K08`, flags 3.
/// * Any other kind              → DeclNode kind `K07`, flags 1.
///
/// After allocation:
/// * Stores the new DeclNode id in `scope_node.linked_decl`.
/// * Sets `decl_node.link_param` to `link_kind`.
/// * Sets `decl_node.scope_parent` to `scope_node_id`.
/// * Based on `link_kind`:
///   - 1 → `decl_node.ext_flags |= 0x0d`
///   - 2 → `decl_node.ext_flags |= 0x01`
///   - 0 or other → no ext_flags change.
///
/// Returns the new DeclNode id.
pub fn ast_node_link(
    scope_nodes: &mut ScopeArena,
    decls: &mut DeclArena,
    scopes: &mut ScopeBlockArena,
    scope_node_id: NodeId,
    p_scope: Option<NodeId>,
    parent_ctx: Option<NodeId>,
    link_kind: u32,
) -> NodeId {
    let is_module = scope_nodes.get(scope_node_id).kind == ScopeNodeKind::Module;
    let (decl_kind, decl_flags) = if is_module {
        (DeclSymKind::K08, 3u8)
    } else {
        (DeclSymKind::K07, 1u8)
    };

    let decl_id = alloc_decl_node(decls, scopes, decl_kind, p_scope, decl_flags, parent_ctx);

    // Wire up back-links.
    scope_nodes.get_mut(scope_node_id).linked_decl = Some(decl_id);
    decls.get_mut(decl_id).link_param = link_kind;
    decls.get_mut(decl_id).scope_parent = Some(scope_node_id);

    match link_kind {
        1 => decls.get_mut(decl_id).ext_flags |= 0x0d,
        2 => decls.get_mut(decl_id).ext_flags |= 0x01,
        _ => {}
    }

    decl_id
}

// ---------------------------------------------------------------------------
// Scope block
// ---------------------------------------------------------------------------

/// A scope block — the list head for all declaration nodes in a single scope.
///
/// Layout (20 bytes):
/// * word[0] (bytes 0–3): pointer to head DeclNode in the symbol list.
/// * word[1] (bytes 4–7): reserved / zero at init.
/// * u16 at bytes 8–9, 10–11: reserved / zero at init.
/// * word[3] (bytes 12–15): hash invalidation sentinel = `0xffffffff`.
/// * word[4] (bytes 16–19): reserved / zero at init.
///
/// The sentinel is reset to 0xffffffff in `alloc_decl_node` each time a new
/// symbol is prepended to the scope, invalidating any cached hash lookup.
pub struct ScopeBlock {
    /// Head of the symbol list (DeclNode chain via `scope_chain`).
    /// `None` when empty.
    pub head: Option<NodeId>,
    /// Hash invalidation sentinel.  `0xffffffff` at init; reset by
    /// `alloc_decl_node` after prepend.
    pub hash_sentinel: u32,
}

/// Arena for scope-block nodes.
pub type ScopeBlockArena = Arena<ScopeBlock>;

/// Allocates a fresh, empty scope block, initialised to zero with the hash
/// sentinel set to `0xffffffff`.
///
/// Pushes into `ScopeBlockArena` and returns the `NodeId`; the caller binds
/// the id.
pub fn alloc_scope_block(arena: &mut ScopeBlockArena) -> NodeId {
    arena.alloc(ScopeBlock { head: None, hash_sentinel: u32::MAX })
}

// ---------------------------------------------------------------------------
// Declaration / symbol node family
// ---------------------------------------------------------------------------

/// Kind discriminant for declaration/symbol nodes created by `alloc_decl_node`.
///
/// Values are the `nodeKind` argument (0–26), indexing a byte size table.
/// Sizes (bytes):
///
/// K00→56, K01→68, K02→56, K03→46, K04→56, K05→88, K06→56,
/// K07→60, K08→124, K09→80, K0a→56, K0b→22, K0c→44, K0d→26,
/// K0e→48, K0f→0, K10→60, K11→66, K12→85, K13→73, K14→76,
/// K15→84, K16→45, K17→73, K18→78, K19→62, K1a→0.
///
/// Named variants where semantics are established from caller context:
/// `K08` = class-module declaration; `K09` = standard-module declaration
/// (module-level flag set on its children).
#[derive(Clone, Copy, Debug, PartialEq, Eq)]
#[repr(u8)]
pub enum DeclSymKind {
    K00 = 0x00, K01 = 0x01, K02 = 0x02, K03 = 0x03,
    K04 = 0x04, K05 = 0x05, K06 = 0x06,
    /// Regular scope declaration (size 60).
    K07 = 0x07,
    /// Class-module declaration (size 124).
    K08 = 0x08,
    /// Standard-module declaration (size 80).
    K09 = 0x09,
    K0a = 0x0a, K0b = 0x0b, K0c = 0x0c, K0d = 0x0d,
    K0e = 0x0e, K0f = 0x0f, K10 = 0x10, K11 = 0x11,
    K12 = 0x12, K13 = 0x13, K14 = 0x14, K15 = 0x15,
    K16 = 0x16, K17 = 0x17, K18 = 0x18, K19 = 0x19,
}

/// A declaration/symbol node.
///
/// VB6 uses variable-size raw allocations (22–124 bytes, kind-dependent) with
/// fields accessed by fixed byte offsets.  Here all fields are named;
/// variable-length lists (`children`, `sec_children`) replace the
/// tail-linked raw pointer chains.
///
/// Byte-offset → field mapping:
/// * offset 0: parent context (parent DeclNode ptr) → `parent`
/// * byte 8 (low nibble): node kind             → `kind`
/// * byte 9 (6 bits): flags                     → `flags`
/// * byte 17 (0x11): secondary flags            → `sec_flags`
/// * byte 18 (0x12): extended flags             → `ext_flags`
/// * offset 28 (word[7]): scope-block ptr       → `scope`
/// * offset 32 (0x20): scope-parent ref         → `scope_parent`
/// * offset 36 (word[9]): next-in-scope         → `scope_chain`
/// * offsets 40/44/48 (0x28/0x2c/0x30): child tail-link → `children`
/// * offsets 60/64 (0x3c/0x40): secondary child chain → `sec_children`
/// * offset 108 (0x6c): link param              → `link_param`
/// * offset 112 (0x70): type annotation         → `type_info`
pub struct DeclNode {
    pub kind: DeclSymKind,
    /// 6-bit flags field (byte 9; bit 6–7 unused).
    pub flags: u8,
    /// Flags byte at offset 0x10 (16).  Bit 3 set when building predefined
    /// nodes whose enclosing context has bit 0 of its ext_flags set.
    pub flags_10: u8,
    /// Secondary flags byte (offset 0x11).  Bit 2 = module-level symbol.
    pub sec_flags: u8,
    /// Extended flags byte (offset 0x12).  Set by `ast_node_link`.
    pub ext_flags: u8,
    /// Parent context (another DeclNode).  word[0].
    pub parent: Option<NodeId>,
    /// Owning scope block (into `ScopeBlockArena`).  word[7] (offset 28).
    pub scope: Option<NodeId>,
    /// Parent scope reference.  offset 0x20 (set by `ast_node_link`).
    pub scope_parent: Option<NodeId>,
    /// Next sibling in scope symbol list.  word[9] (offset 36).
    pub scope_chain: Option<NodeId>,
    /// Primary child list.  Replaces the tail-linked chain at offsets
    /// 0x28/0x2c/0x30 (chain nodes → `Vec<NodeId>`).
    pub children: Vec<NodeId>,
    /// Secondary child list (class-module only).  Replaces the chain at
    /// offsets 0x3c/0x40.
    pub sec_children: Vec<NodeId>,
    /// Short at offset 0x2c.  Written for member-ref nodes (K03) to store the
    /// field index.
    pub field_2c: u16,
    /// Short at offset 0x34.  Written for field-access nodes (K06) to store the
    /// field index.
    pub field_34: u16,
    /// Short at offset 0x38.  Written for let-accessor nodes (K07) to store the
    /// field index.
    pub field_38: u16,
    /// Field-type / slot table at offset 0x38 (for type/module parent decls,
    /// where offset 0x38 is a table pointer rather than the `field_38` short).
    /// Indexed by `0xd + slot` when building a K1a field-access node, where the
    /// entry is the member's resolved type descriptor. Empty until the
    /// type-layout pass that populates a type's member table has run; an absent
    /// entry reads as `0`.
    pub field_type_table: Vec<u32>,
    /// Short at offset 0x68.  Written for binding and proc-call nodes (K08) to
    /// store the field index.
    pub field_68: u16,
    /// Flags byte at offset 0x13.  Written for Implements-proc nodes (K08):
    /// bit 3 always set; bit 5 (0x20) set when `access_spec.flags & 2`.
    pub flags_13: u8,
    /// Integer at offset 0x6c; threaded from `ast_node_link`'s `link_kind`.
    pub link_param: u32,
    /// Type-annotation integer at offset 0x70.
    pub type_info: u32,
}

/// Arena for declaration/symbol nodes.
pub type DeclArena = Arena<DeclNode>;

/// Allocates a zero-initialised declaration node and wires it into:
/// 1. The owning scope block's prepend chain (if `scope` is `Some`).
/// 2. The parent DeclNode's primary child list (if `parent` is `Some`).
///
/// Special rules:
/// * If parent kind == `K08` (class module): the node is also appended to the
///   parent's `sec_children` list.
/// * If parent kind == `K09` (standard module): bit 2 of `sec_flags` is set
///   (marks the new node as module-level).
///
/// **`flags`** is masked to 6 bits (bits 6–7 are cleared).
pub fn alloc_decl_node(
    decls: &mut DeclArena,
    scopes: &mut ScopeBlockArena,
    kind: DeclSymKind,
    scope: Option<NodeId>,
    flags: u8,
    parent: Option<NodeId>,
) -> NodeId {
    let node_id = decls.alloc(DeclNode {
        kind,
        flags: flags & 0x3f,
        flags_10: 0,
        sec_flags: 0,
        ext_flags: 0,
        parent,
        scope,
        scope_parent: None,
        scope_chain: None,
        children: Vec::new(),
        sec_children: Vec::new(),
        field_2c: 0,
        field_34: 0,
        field_38: 0,
        field_type_table: Vec::new(),
        field_68: 0,
        flags_13: 0,
        link_param: 0,
        type_info: 0,
    });

    // Prepend into the scope block's symbol list.
    if let Some(scope_id) = scope {
        let sb = scopes.get_mut(scope_id);
        sb.hash_sentinel = u32::MAX;
        let prev_head = sb.head;
        sb.head = Some(node_id);
        decls.get_mut(node_id).scope_chain = prev_head;
    }

    // Append to parent's child list (and optional secondary list).
    if let Some(parent_id) = parent {
        let parent_kind = decls.get(parent_id).kind;
        decls.get_mut(parent_id).children.push(node_id);
        if parent_kind == DeclSymKind::K08 {
            decls.get_mut(parent_id).sec_children.push(node_id);
        }
        if parent_kind == DeclSymKind::K09 {
            decls.get_mut(node_id).sec_flags |= 4;
        }
    }

    node_id
}
