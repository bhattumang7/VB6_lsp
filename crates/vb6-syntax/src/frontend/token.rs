//! VB6 token types: [`Kw`], [`TokenKind`], [`Lit`], [`Span`], [`Token`].
//!
//! ## Grounding
//!
//! The 271 entries of `KEYWORD_TABLE` (`crates/vb6-core/src/frontend/keyword_table.rs`)
//! are the behavioral ground truth for what the scanner emits:
//!
//! * Table indices 1–225 are named VBA keywords/built-ins.
//! * Indices 226–270 are operators and punctuation characters that the
//!   scanner recognises and emits with their table index as the token ID.
//! * Indices 230–243 are unassigned reserved slots (name = "0") — no `Kw`
//!   variant is defined for them.
//! * Index 0 is the null/error sentinel — likewise omitted.
//! * Token 0x10f (271) is EOL — one past the table end, a synthetic sentinel
//!   emitted when the source is exhausted.
//!
//! The `w0` high 16 bits of a keyword-table entry encode operator precedence
//! and other parser-level attributes.  That information is preserved in
//! [`KEYWORD_TABLE`] and will be used by the Pratt parser; it is not
//! modelled here.
//!
//! ## Representation
//!
//! VB6 tracks position via pointer arithmetic into a u16 wide-char source array
//! and produces a raw numeric token id.  Our [`Token`] replaces the numeric id
//! with a typed [`TokenKind`], replaces the global source-cursor with owned
//! [`Span`] fields, and stores literal values inline instead of in separate
//! globals.  None of this is observable at the public boundary.

use crate::frontend::scanner::SymId;

// ── Span ──────────────────────────────────────────────────────────────────────

/// Byte-offset range within the source passed to the scanner for one logical
/// line.  `start + len` must not exceed the source length.
#[derive(Clone, Copy, Debug, Default, PartialEq, Eq)]
pub struct Span {
    /// Byte offset of the first character of the token.
    pub start: u32,
    /// Length of the token in bytes.
    pub len: u32,
}

impl Span {
    pub const DUMMY: Self = Self { start: 0, len: 0 };
}

// ── Lit ───────────────────────────────────────────────────────────────────────

/// A literal value carried by a token.
///
/// The numeric payload and string buffer are stored inline on the token rather
/// than in separate per-type globals.
#[derive(Clone, Debug, PartialEq)]
pub enum Lit {
    /// `Integer` (i16 range) — type declarator `%` or bare integer.
    Int(i32),
    /// `Long` (i32 range) — type declarator `&` or integer too large for i16.
    Long(i32),
    /// `Single` — type declarator `!` or exponent-form float.
    Single(f32),
    /// `Double` — type declarator `#` or default float.
    Double(f64),
    /// `Currency` — type declarator `@`, stored as 10000ths.
    Currency(i64),
    /// String literal — content between the outer double-quotes, with
    /// doubled `""` sequences already collapsed to single `"`.
    Str(Box<str>),
    /// Date literal — Julian day serial as f64 (VBA's internal date format).
    Date(f64),
}

// ── TypeSuffix ──────────────────────────────────────────────────────────────────

/// A type-declaration character (`% & ! # @ $`) attached to an identifier.
///
/// The identifier scanner consumes a trailing type-suffix character off the
/// source and records its class here, and the parser reads it when building the
/// name/type node.  `None` when the identifier carries no suffix.
#[derive(Clone, Copy, Debug, Default, PartialEq, Eq)]
pub enum TypeSuffix {
    /// No type-declaration suffix (class 0).
    #[default]
    None,
    /// `%` — Integer (class 1).
    Integer,
    /// `&` — Long (class 2).
    Long,
    /// `!` — Single (class 3).
    Single,
    /// `#` — Double (class 4).
    Double,
    /// `$` — String (class 5).
    String,
    /// `@` — Currency (class 6).
    Currency,
}

impl TypeSuffix {
    /// Classify a terminating byte as a type-declaration suffix.
    ///
    /// `!`→Single, `#`→Double, `$`→String, `%`→Integer, `&`→Long,
    /// `@`→Currency, else None.
    #[inline]
    pub fn from_byte(b: u8) -> Self {
        match b {
            b'%' => Self::Integer,
            b'&' => Self::Long,
            b'!' => Self::Single,
            b'#' => Self::Double,
            b'$' => Self::String,
            b'@' => Self::Currency,
            _ => Self::None,
        }
    }

    /// The `BuiltinType { kind }` code this suffix implies as a *declared* type,
    /// matching `parse_type_spec` (Integer=2, Long=3, Single=4, Double=5,
    /// Currency=6).  `String` declares a `StringType` node, so it returns `None`
    /// here and is handled separately; `None` (no suffix) likewise.
    #[inline]
    pub fn builtin_kind(self) -> Option<u32> {
        match self {
            Self::Integer => Some(2),
            Self::Long => Some(3),
            Self::Single => Some(4),
            Self::Double => Some(5),
            Self::Currency => Some(6),
            Self::String | Self::None => None,
        }
    }
}

// ── Kw ────────────────────────────────────────────────────────────────────────

/// A VBA keyword or operator, identified by its KEYWORD_TABLE index.
///
/// The discriminant equals the VB6 token id (`w0 & 0xffff`).  Use
/// [`Kw::token_id`] to recover it and [`Kw::from_token_id`] to convert the
/// other way.
///
/// **Naming rules for non-alphanumeric entries:**
/// * `$`-suffix forms: drop `$`, add `S` (e.g. `Dir$` → `DirS`).
/// * `#`-prefix conditional-compile directives: replace `#` with `Cc`
///   (e.g. `#Const` → `CcConst`).
/// * `VB_*` attribute keys: `VB_` → `Vb`, then CamelCase the rest.
/// * `LINEINPUT` → `LineInput`.
/// * Single-/two-letter words `B`, `BF`, `F` kept as-is.
/// * Operator characters get English names (see operator region below).
#[derive(Clone, Copy, Debug, PartialEq, Eq, Hash)]
#[repr(u16)]
pub enum Kw {
    // ── Named keywords / built-ins (token IDs 1–225) ──────────────────────
    Abs          = 1,
    Access       = 2,
    AddressOf    = 3,
    Alias        = 4,
    And          = 5,
    Any          = 6,
    Append       = 7,
    Array        = 8,
    As           = 9,
    Assert       = 10,
    B            = 11,
    Base         = 12,
    BF           = 13,
    Binary       = 14,
    Boolean      = 15,
    ByRef        = 16,
    Byte         = 17,
    ByVal        = 18,
    Call         = 19,
    Case         = 20,
    CBool        = 21,
    CByte        = 22,
    CCur         = 23,
    CDate        = 24,
    CDec         = 25,
    CDbl         = 26,
    CDecl        = 27,
    ChDir        = 28,
    CInt         = 29,
    Circle       = 30,
    CLng         = 31,
    Close        = 32,
    Compare      = 33,
    Const        = 34,
    CSng         = 35,
    CStr         = 36,
    CurDir       = 37,
    /// `CurDir$`
    CurDirS      = 38,
    CVar         = 39,
    CVDate       = 40,
    CVErr        = 41,
    Currency     = 42,
    Database     = 43,
    Date         = 44,
    /// `Date$`
    DateS        = 45,
    Debug        = 46,
    Decimal      = 47,
    Declare      = 48,
    DefBool      = 49,
    DefByte      = 50,
    DefCur       = 51,
    DefDate      = 52,
    DefDec       = 53,
    DefDbl       = 54,
    DefInt       = 55,
    DefLng       = 56,
    DefObj       = 57,
    DefSng       = 58,
    DefStr       = 59,
    DefVar       = 60,
    Dim          = 61,
    Dir          = 62,
    /// `Dir$`
    DirS         = 63,
    Do           = 64,
    DoEvents     = 65,
    Double       = 66,
    Each         = 67,
    Else         = 68,
    ElseIf       = 69,
    Empty        = 70,
    End          = 71,
    EndIf        = 72,
    Enum         = 73,
    Eqv          = 74,
    Erase        = 75,
    Error        = 76,
    /// `Error$`
    ErrorS       = 77,
    Event        = 78,
    Exit         = 79,
    Explicit     = 80,
    F            = 81,
    False        = 82,
    Fix          = 83,
    For          = 84,
    Format       = 85,
    /// `Format$`
    FormatS      = 86,
    FreeFile     = 87,
    Friend       = 88,
    Function     = 89,
    Get          = 90,
    Global       = 91,
    Go           = 92,
    GoSub        = 93,
    GoTo         = 94,
    If           = 95,
    Imp          = 96,
    Implements   = 97,
    In           = 98,
    Input        = 99,
    /// `Input$`
    InputS       = 100,
    InputB       = 101,
    /// `InputB$`
    InputBS      = 102,
    InStr        = 103,
    InStrB       = 104,
    Int          = 105,
    Integer      = 106,
    Is           = 107,
    LBound       = 108,
    Left         = 109,
    Len          = 110,
    LenB         = 111,
    Let          = 112,
    Lib          = 113,
    Like         = 114,
    Line         = 115,
    /// `LINEINPUT`
    LineInput    = 116,
    Load         = 117,
    Local        = 118,
    Lock         = 119,
    Long         = 120,
    Loop         = 121,
    LSet         = 122,
    Me           = 123,
    Mid          = 124,
    /// `Mid$`
    MidS         = 125,
    MidB         = 126,
    /// `MidB$`
    MidBS        = 127,
    Mod          = 128,
    Module       = 129,
    Name         = 130,
    New          = 131,
    Next         = 132,
    Not          = 133,
    Nothing      = 134,
    Null         = 135,
    Object       = 136,
    On           = 137,
    Open         = 138,
    Option       = 139,
    Optional     = 140,
    Or           = 141,
    Output       = 142,
    ParamArray   = 143,
    Preserve     = 144,
    Print        = 145,
    Private      = 146,
    Property     = 147,
    PSet         = 148,
    Public       = 149,
    Put          = 150,
    RaiseEvent   = 151,
    Random       = 152,
    Randomize    = 153,
    Read         = 154,
    ReDim        = 155,
    Rem          = 156,
    Resume       = 157,
    Return       = 158,
    RGB          = 159,
    RSet         = 160,
    Scale        = 161,
    Seek         = 162,
    Select       = 163,
    Set          = 164,
    Sgn          = 165,
    Shared       = 166,
    Single       = 167,
    Spc          = 168,
    Static       = 169,
    Step         = 170,
    Stop         = 171,
    StrComp      = 172,
    String       = 173,
    /// `String$`
    StringS      = 174,
    Sub          = 175,
    Tab          = 176,
    Text         = 177,
    Then         = 178,
    To           = 179,
    True         = 180,
    Type         = 181,
    TypeOf       = 182,
    UBound       = 183,
    Unload       = 184,
    Unlock       = 185,
    Unknown      = 186,
    Until        = 187,
    Variant      = 188,
    Wend         = 189,
    While        = 190,
    Width        = 191,
    With         = 192,
    WithEvents   = 193,
    Write        = 194,
    Xor          = 195,

    // ── Conditional-compilation directives (token IDs 196–200) ────────────
    /// `#Const`
    CcConst      = 196,
    /// `#Else`
    CcElse       = 197,
    /// `#ElseIf`
    CcElseIf     = 198,
    /// `#End`
    CcEnd        = 199,
    /// `#If`
    CcIf         = 200,

    // ── Attribute keys (token IDs 201–225) ────────────────────────────────
    Attribute             = 201,
    VbBase                = 202,
    VbControl             = 203,
    VbCreatable           = 204,
    VbCustomizable        = 205,
    VbDescription         = 206,
    VbExposed             = 207,
    VbExtKey              = 208,
    VbHelpId              = 209,
    VbInvokeFunc          = 210,
    VbInvokeProperty      = 211,
    VbInvokePropertyPut   = 212,
    VbInvokePropertyPutRef = 213,
    VbMemberFlags         = 214,
    VbName                = 215,
    VbPredeclaredId       = 216,
    VbProcData            = 217,
    VbTemplateDerived     = 218,
    VbVarDescription      = 219,
    VbVarHelpId           = 220,
    VbVarMemberFlags      = 221,
    VbVarProcData         = 222,
    VbUserMemId           = 223,
    VbVarUserMemId        = 224,
    VbGlobalNameSpace     = 225,

    // ── Operator / punctuation entries (token IDs 226–270) ────────────────
    //
    // These are emitted by the scanner when it encounters the corresponding
    // character.  They are assigned the table index as the token id;
    // the parser switches on these ids just like keyword ids.
    //
    // Two dual entries exist where VB6 uses different ids in different
    // contexts (both `,` and both `.`).

    /// `,` in a specific non-expression context (token 0xe2 = 226).
    CommaStmt    = 226,
    /// `.` in a specific non-expression context (token 0xe3 = 227).
    DotStmt      = 227,
    /// `"` — triggers string-literal scanning; this id is emitted if a bare
    /// double-quote appears in an unexpected position (token 0xe4 = 228).
    DQuote       = 228,
    /// `_` — line-continuation marker (token 0xe5 = 229).
    LineCont     = 229,

    // IDs 230–243 (0xe6–0xf3) are unassigned reserved slots; no variants.

    /// `!` — dictionary/member shorthand (token 0xf4 = 244).
    Bang         = 244,
    /// `#` — date-literal delimiter or conditional-compile marker
    ///        (token 0xf5 = 245).
    Hash         = 245,
    /// `&` — string concatenation operator (token 0xf6 = 246).
    Amp          = 246,
    /// `'` — comment start (token 0xf7 = 247).
    Apos         = 247,
    /// `(` (token 0xf8 = 248).
    LParen       = 248,
    /// `)` (token 0xf9 = 249).
    RParen       = 249,
    /// `*` — multiplication (token 0xfa = 250).
    Star         = 250,
    /// `+` — addition (token 0xfb = 251).
    Plus         = 251,
    /// `-` — subtraction / unary negation (token 0xfc = 252).
    Minus        = 252,
    /// `.` in expression context — member access (token 0xfd = 253).
    Dot          = 253,
    /// `/` — division (token 0xfe = 254).
    Slash        = 254,
    /// `:` — statement separator (token 0xff = 255).
    Colon        = 255,
    /// `;` (token 0x100 = 256).
    Semi         = 256,
    /// `<` — less-than (token 0x101 = 257).
    Lt           = 257,
    /// `<=` — less-or-equal (token 0x102 = 258).
    Le           = 258,
    /// `<>` — not-equal (token 0x103 = 259).
    Ne           = 259,
    /// `=` — assignment or equality test (token 0x104 = 260).
    Eq           = 260,
    /// `=<` — alternate spelling of `<=` (token 0x105 = 261).
    EqLt         = 261,
    /// `=>` — alternate spelling of `>=` (token 0x106 = 262).
    EqGt         = 262,
    /// `>` — greater-than (token 0x107 = 263).
    Gt           = 263,
    /// `><` — alternate spelling of `<>` (token 0x108 = 264).
    GtLt         = 264,
    /// `>=` — greater-or-equal (token 0x109 = 265).
    Ge           = 265,
    /// `?` — `Print` shortcut (token 0x10a = 266).
    Question     = 266,
    /// `\` — integer division (token 0x10b = 267).
    Backslash    = 267,
    /// `^` — exponentiation (token 0x10c = 268).
    Caret        = 268,
    /// `:=` — named-argument separator (token 0x10d = 269).
    ColonEq      = 269,
    /// `,` in expression context (token 0x10e = 270).
    Comma        = 270,
}

impl Kw {
    /// Returns the VB6 token id (`w0 & 0xffff` of the KEYWORD_TABLE entry).
    #[inline]
    pub fn token_id(self) -> u16 {
        self as u16
    }

    /// Converts a raw VB6 token id to the corresponding [`Kw`] variant, or
    /// `None` for unassigned slots (0 and 230–243).
    pub fn from_token_id(id: u16) -> Option<Self> {
        use Kw::*;
        Some(match id {
            1   => Abs,          2   => Access,       3   => AddressOf,
            4   => Alias,        5   => And,           6   => Any,
            7   => Append,       8   => Array,         9   => As,
            10  => Assert,       11  => B,             12  => Base,
            13  => BF,           14  => Binary,        15  => Boolean,
            16  => ByRef,        17  => Byte,          18  => ByVal,
            19  => Call,         20  => Case,          21  => CBool,
            22  => CByte,        23  => CCur,          24  => CDate,
            25  => CDec,         26  => CDbl,          27  => CDecl,
            28  => ChDir,        29  => CInt,          30  => Circle,
            31  => CLng,         32  => Close,         33  => Compare,
            34  => Const,        35  => CSng,          36  => CStr,
            37  => CurDir,       38  => CurDirS,       39  => CVar,
            40  => CVDate,       41  => CVErr,         42  => Currency,
            43  => Database,     44  => Date,          45  => DateS,
            46  => Debug,        47  => Decimal,       48  => Declare,
            49  => DefBool,      50  => DefByte,       51  => DefCur,
            52  => DefDate,      53  => DefDec,        54  => DefDbl,
            55  => DefInt,       56  => DefLng,        57  => DefObj,
            58  => DefSng,       59  => DefStr,        60  => DefVar,
            61  => Dim,          62  => Dir,           63  => DirS,
            64  => Do,           65  => DoEvents,      66  => Double,
            67  => Each,         68  => Else,          69  => ElseIf,
            70  => Empty,        71  => End,           72  => EndIf,
            73  => Enum,         74  => Eqv,           75  => Erase,
            76  => Error,        77  => ErrorS,        78  => Event,
            79  => Exit,         80  => Explicit,      81  => F,
            82  => False,        83  => Fix,           84  => For,
            85  => Format,       86  => FormatS,       87  => FreeFile,
            88  => Friend,       89  => Function,      90  => Get,
            91  => Global,       92  => Go,            93  => GoSub,
            94  => GoTo,         95  => If,            96  => Imp,
            97  => Implements,   98  => In,            99  => Input,
            100 => InputS,       101 => InputB,        102 => InputBS,
            103 => InStr,        104 => InStrB,        105 => Int,
            106 => Integer,      107 => Is,            108 => LBound,
            109 => Left,         110 => Len,           111 => LenB,
            112 => Let,          113 => Lib,           114 => Like,
            115 => Line,         116 => LineInput,     117 => Load,
            118 => Local,        119 => Lock,          120 => Long,
            121 => Loop,         122 => LSet,          123 => Me,
            124 => Mid,          125 => MidS,          126 => MidB,
            127 => MidBS,        128 => Mod,           129 => Module,
            130 => Name,         131 => New,           132 => Next,
            133 => Not,          134 => Nothing,       135 => Null,
            136 => Object,       137 => On,            138 => Open,
            139 => Option,       140 => Optional,      141 => Or,
            142 => Output,       143 => ParamArray,    144 => Preserve,
            145 => Print,        146 => Private,       147 => Property,
            148 => PSet,         149 => Public,        150 => Put,
            151 => RaiseEvent,   152 => Random,        153 => Randomize,
            154 => Read,         155 => ReDim,         156 => Rem,
            157 => Resume,       158 => Return,        159 => RGB,
            160 => RSet,         161 => Scale,         162 => Seek,
            163 => Select,       164 => Set,           165 => Sgn,
            166 => Shared,       167 => Single,        168 => Spc,
            169 => Static,       170 => Step,          171 => Stop,
            172 => StrComp,      173 => String,        174 => StringS,
            175 => Sub,          176 => Tab,           177 => Text,
            178 => Then,         179 => To,            180 => True,
            181 => Type,         182 => TypeOf,        183 => UBound,
            184 => Unload,       185 => Unlock,        186 => Unknown,
            187 => Until,        188 => Variant,       189 => Wend,
            190 => While,        191 => Width,         192 => With,
            193 => WithEvents,   194 => Write,         195 => Xor,
            196 => CcConst,      197 => CcElse,        198 => CcElseIf,
            199 => CcEnd,        200 => CcIf,
            201 => Attribute,
            202 => VbBase,            203 => VbControl,
            204 => VbCreatable,       205 => VbCustomizable,
            206 => VbDescription,     207 => VbExposed,
            208 => VbExtKey,          209 => VbHelpId,
            210 => VbInvokeFunc,      211 => VbInvokeProperty,
            212 => VbInvokePropertyPut,
            213 => VbInvokePropertyPutRef,
            214 => VbMemberFlags,     215 => VbName,
            216 => VbPredeclaredId,   217 => VbProcData,
            218 => VbTemplateDerived, 219 => VbVarDescription,
            220 => VbVarHelpId,       221 => VbVarMemberFlags,
            222 => VbVarProcData,     223 => VbUserMemId,
            224 => VbVarUserMemId,    225 => VbGlobalNameSpace,
            226 => CommaStmt,    227 => DotStmt,      228 => DQuote,
            229 => LineCont,
            // 230–243: unassigned reserved slots
            244 => Bang,         245 => Hash,         246 => Amp,
            247 => Apos,         248 => LParen,       249 => RParen,
            250 => Star,         251 => Plus,         252 => Minus,
            253 => Dot,          254 => Slash,        255 => Colon,
            256 => Semi,         257 => Lt,           258 => Le,
            259 => Ne,           260 => Eq,           261 => EqLt,
            262 => EqGt,         263 => Gt,           264 => GtLt,
            265 => Ge,           266 => Question,     267 => Backslash,
            268 => Caret,        269 => ColonEq,      270 => Comma,
            _ => return None,
        })
    }

    /// The canonical VB6 string form (e.g. `"And"`, `","`, `"<="`, `"#Const"`).
    pub fn name(self) -> &'static str {
        crate::frontend::keyword_table::KEYWORD_TABLE[self.token_id() as usize].name
    }
}

// ── TokenKind ─────────────────────────────────────────────────────────────────

/// The discriminant of a scanned token.
///
/// `Kw(kw)` covers the entire KEYWORD_TABLE token-id space (1–270).
/// `Ident` is an identifier whose interned name did not resolve to any table
/// entry (token id 0).
/// Literals, `Eol`, `Eof`, and `Error` are synthetic — they have no KEYWORD_TABLE
/// entry but appear in the token stream.
#[derive(Clone, Debug, PartialEq)]
pub enum TokenKind {
    /// A token that matched a KEYWORD_TABLE entry (any of the 256 valid
    /// ids 1–270, excluding unassigned slots 230–243).
    Kw(Kw),
    /// An identifier that did not resolve to a keyword (token id = 0
    /// after `intern_string` lookup returns a symbol whose `.token` field
    /// is 0).  The [`Token::sym`] field carries the interned [`SymId`].
    Ident,
    /// An integer literal; literal value in [`Token::lit`].
    IntLit,
    /// A long (32-bit integer) literal; literal value in [`Token::lit`].
    LongLit,
    /// A single-precision float literal; literal value in [`Token::lit`].
    SngLit,
    /// A double-precision float literal; literal value in [`Token::lit`].
    DblLit,
    /// A currency literal; literal value in [`Token::lit`].
    CurLit,
    /// A string literal; literal value in [`Token::lit`].
    StrLit,
    /// A date literal; literal value in [`Token::lit`].
    DateLit,
    /// End of logical line — synthetic token emitted when the source u16
    /// array is exhausted (sentinel 0x104, mapped to token 0x10f = 271).
    Eol,
    /// End of all source input (no more logical lines).
    Eof,
    /// An unrecognised character was encountered; scanning continued.
    Error,
}

// ── Token ─────────────────────────────────────────────────────────────────────

/// A single scanned token.
///
/// Corresponds to one element of VB6's token stream (token id + symbol record +
/// numeric payload + string buffer).  All per-token state is owned by this
/// struct instead of scattered across process globals.
#[derive(Clone, Debug, PartialEq)]
pub struct Token {
    /// What kind of token this is.
    pub kind: TokenKind,
    /// Interned symbol id — set for [`TokenKind::Kw`] and [`TokenKind::Ident`];
    /// `None` for literals and structural tokens.
    pub sym: Option<SymId>,
    /// Literal value — set for the `*Lit` variants; `None` otherwise.
    pub lit: Option<Lit>,
    /// Type-declaration suffix consumed off this identifier.
    /// `TypeSuffix::None` for everything except a suffixed
    /// `Ident`/`Kw` token (e.g. `count%`, `STATE_MIXED&`).
    pub type_suffix: TypeSuffix,
    /// Source position.
    pub span: Span,
}

impl Token {
    /// Convenience constructor for a keyword token without a literal.
    pub fn kw(kw: Kw, sym: SymId, span: Span) -> Self {
        Self { kind: TokenKind::Kw(kw), sym: Some(sym), lit: None, type_suffix: TypeSuffix::None, span }
    }

    /// Convenience constructor for an identifier token.
    pub fn ident(sym: SymId, span: Span) -> Self {
        Self { kind: TokenKind::Ident, sym: Some(sym), lit: None, type_suffix: TypeSuffix::None, span }
    }

    /// Convenience constructor for a literal token.
    pub fn lit(kind: TokenKind, value: Lit, span: Span) -> Self {
        Self { kind, sym: None, lit: Some(value), type_suffix: TypeSuffix::None, span }
    }

    /// End-of-line token.
    pub fn eol(pos: u32) -> Self {
        Self { kind: TokenKind::Eol, sym: None, lit: None, type_suffix: TypeSuffix::None, span: Span { start: pos, len: 0 } }
    }

    /// End-of-file token.
    pub fn eof() -> Self {
        Self { kind: TokenKind::Eof, sym: None, lit: None, type_suffix: TypeSuffix::None, span: Span::DUMMY }
    }

    /// Scan-error token.
    pub fn error(span: Span) -> Self {
        Self { kind: TokenKind::Error, sym: None, lit: None, type_suffix: TypeSuffix::None, span }
    }

    /// Operator / punctuation token — no symbol record needed.
    /// Use this for single-char and multi-char operators that carry no interned
    /// name (the parser only cares about the [`Kw`] discriminant).
    pub fn op(kw: Kw, span: Span) -> Self {
        Self { kind: TokenKind::Kw(kw), sym: None, lit: None, type_suffix: TypeSuffix::None, span }
    }

    /// Returns `true` for the two statement-terminator tokens: `Eol` and the
    /// `Apos` (`'`) keyword.  The parser loop exits on `0x10f` or `0xf7`.
    pub fn is_stmt_end(&self) -> bool {
        matches!(
            &self.kind,
            TokenKind::Eol
                | TokenKind::Eof
                | TokenKind::Kw(Kw::Apos)
                | TokenKind::Kw(Kw::Rem)
        )
    }
}
