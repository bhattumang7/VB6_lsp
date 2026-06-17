# ANTLR VB6 Grammar — Bugs

All entries in this file are bugs in the ANTLR grammar
(`VisualBasic6Lexer.g4` / `VisualBasic6Parser.g4`). They are grounded in the
VB6 keyword table (271 entries) in
`crates/vb6-syntax/src/frontend/keyword_table.rs`, the Rust
recursive-descent parser (used as a reference for correct VB6 behaviour), and
the ANTLR acceptance-test corpus in
`crates/vb6-core/tests/antlr_coverage.rs`.

---

## Summary Table

| # | File | Severity | Synopsis |
|---|------|----------|----------|
| L1 | Lexer | **Critical** | `#Const` directive has no lexer rule — entire CC-constant feature missing |
| L2 | Lexer | **Critical** | `:` statement separator without trailing space rejected |
| L3 | Lexer | High | Line-continuation `_` only recognised after a space character, not a tab |
| L4 | Lexer | High | `#End If` with extra or tab whitespace not tokenised as `MACRO_END_IF` |
| L5 | Lexer | High | `Rem` comment without a trailing space causes a parse error |
| L6 | Lexer | High | Numeric line labels (`10:`, `GoTo 100`) completely unparseable |
| L7 | Lexer | High | `1#`, `1!`, `1@`, `1E2` wrongly typed as `INTEGERLITERAL` |
| L8 | Lexer | Medium | `Currency` (and `Decimal`, `Any`) absent from keyword list and `baseType` |
| P1 | Parser | **Critical** | `+=` / `-=` accepted — not VB6 syntax (VBA7 only) |
| P2 | Parser | **Critical** | `Call Foo()` with empty parens fails — `argsCall` requires ≥ 1 argument |
| P3 | Parser | High | Unary `-` and `Not` consume too much — wrong precedence vs `*`/`/` and `And` |
| P4 | Parser | High | `#If` body inside a Sub uses `moduleBody`, not `block` |
| P5 | Parser | Medium | `vsStruct` accepts tuple expressions `(a, b, c)` — invalid VB6 |
| P6 | Parser | Medium | `For i As Integer = …` accepted — VB.NET-only inline type declaration |
| P7 | Parser | Medium | `ParamArray` allowed at call sites in `argCall` |
| P8 | Parser | Medium | `DefType` letter ranges accept multi-character identifiers |
| P9 | Parser | Medium | Bare `Print expr` (form-surface print, no `#` channel) causes parse error |
| P10 | Parser | Medium | `Empty` keyword missing from `literal` rule; parsed as an identifier |
| P11 | Parser | Low | `For` loop control variable cannot be a member access |
| P12 | Parser | High | `^` exponentiation is left-associative — VB6 requires right-associativity |
| P13 | Parser | **High** | Signed literals (`-2`, `-2.5`, `-&O7`) fold the sign into the literal — breaks unary precedence vs `^` |
| P14 | Parser | Medium | `Next i, j` (one `Next` closing nested loops) — only a single variable accepted |
| P15 | Parser | Low | Inline `If … Then` body is a single `blockStmt` — colon-separated statement lists rejected |
| L9 | Lexer | Low | Trailing-dot float `1.` not tokenised as `DOUBLELITERAL` |
| L10 | Lexer | Medium | `Go To` (space-separated) not tokenised — only contiguous `GoTo` |
| P16 | Parser | High | One-word `EndIf` rejected — valid VB6 (oracle-confirmed) |
| P17 | Parser | High | `If <cond> GoTo <label>` without `Then` rejected — valid VB6 (oracle-confirmed) |
| L11 | Lexer | **Critical** | `DATELITERAL` greedily spans two `#` on a line — breaks `Close #1, #2` (oracle-confirmed) |
| L12 | Lexer | High | File channel `#<var>`/`#<expr>` collapses into one `FILENUMBER` literal — loses the variable (oracle-confirmed) |
| L13 | Lexer | Medium | `Debug` keyword absent — `Debug.Print` only parses by accident, mismodelling the print output list |
| P18 | Parser | Medium | `Resume <line#>` / `Resume 0` rejected — numeric target valid VB6 (oracle-confirmed) |
| C1 | Correction | — | P9 is wrong: bare `Print "x"` is **rejected** in a standard module (oracle-confirmed) |
| C2 | Correction | — | L8 over-claims: `As Decimal` and `As Any` are **rejected** by VB6 (oracle-confirmed) |

---

## Lexer Bugs

### L1 — `#Const` conditional-compilation constant directive is completely absent

**File:** `VisualBasic6Lexer.g4`

The lexer defines `MACRO_IF`, `MACRO_ELSEIF`, `MACRO_ELSE`, `MACRO_END_IF`
but has **no** `MACRO_CONST` rule. The VB6 keyword table records
`#Const` as token 0xC4. It is used throughout real VB6 projects to define
conditional-compilation symbols.

```vb
#Const DEBUG = 1          ' no matching lexer rule → parse failure
#If DEBUG Then
    ' ...
#End If
```

`#Const` must be accepted both at module level and inside procedure bodies.
The Rust parser handles both (`CcConst` dispatched in `parse_stmt` and
`parse_module`).

**Fix:** Add `MACRO_CONST: HASH 'CONST';` to the lexer and a corresponding
parser rule that matches `MACRO_CONST WS ambiguousIdentifier WS? EQ WS? valueStmt`.

---

### L2 — Statement separator colon without trailing space rejected

**File:** `VisualBasic6Lexer.g4`, line 461

```antlr
NEWLINE: WS? ('\r'? '\n' | COLON ' ') WS?;
```

`COLON ' '` requires exactly one space after `:`. VB6 permits no space:

```vb
x = 1:y = 2           ' ':' tokenised as bare COLON, not NEWLINE → parse error
For i = 1 To 10:Next  ' same failure
```

The bare `:` becomes a `COLON` token that no parser rule consumes as a
statement separator, causing a cascade of errors.

**Fix:** Change `COLON ' '` to `COLON WS?` so an optional whitespace (or
nothing at all) is accepted after the separator colon.

---

### L3 — Line-continuation requires a space character, not a tab

**File:** `VisualBasic6Lexer.g4`, line 459

```antlr
LINE_CONTINUATION: ' ' '_' '\r'? '\n' -> skip;
```

The leading `' '` is a literal space. A tab before `_` is legal VB6
line-continuation syntax — the VB6 IDE preserves tabs in indented code:

```vb
Dim x As _
\t Integer   ' tab before _ → LINE_CONTINUATION not matched → parse error
```

**Fix:** Change `' ' '_'` to `[ \t] '_'` (or `WS '_'`).

---

### L4 — `#End If` with non-single-space whitespace fails

**File:** `VisualBasic6Lexer.g4`, line 209

```antlr
MACRO_END_IF: HASH 'END IF';
```

`'END IF'` is a fixed string with exactly one space. Two spaces, a tab, or
any other whitespace between `END` and `IF` will not produce a
`MACRO_END_IF` token. The `#End` is instead emitted as `HASH KEYWORD(END)`,
leaving the `#If` block permanently unclosed and cascading failures through
the rest of the file.

**Fix:** Split into separate tokens: `MACRO_END: HASH 'END';` and handle
the `IF` as a follow-on keyword in the parser rule. Apply the same fix to
the similarly constructed compound keyword tokens (`END_IF`, `END_SUB`, etc.)
if tabs/multi-space variants are expected.

---

### L5 — `Rem` comment without a space after the keyword is rejected

**File:** `VisualBasic6Lexer.g4`, line 463

```antlr
COMMENT: WS? ('\'' | COLON? REM ' ') ( LINE_CONTINUATION | ~ ('\n' | '\r'))* -> skip;
```

`REM ' '` requires a space after `Rem`. A bare `Rem` at end of line — legal
in VB6 — is not matched:

```vb
Sub F()
  Rem          ' no space after Rem → not consumed as COMMENT → parse error
End Sub
```

The unmatched `Rem` token is passed to the parser, where no `blockStmt`
rule accepts it, causing a parse error.

**Fix:** Change `REM ' '` to `REM (' ' | '\r' | '\n')?` so a trailing
space, newline, or nothing at all is accepted after `Rem`.

---

### L6 — Numeric line labels not lexable or parseable

**Files:** `VisualBasic6Lexer.g4` + `VisualBasic6Parser.g4`

The `lineLabel` parser rule is:

```antlr
lineLabel: ambiguousIdentifier COLON;
```

Numeric (integer) line labels from legacy VB6 code cannot be represented.
Neither a bare line-number prefix nor `GoTo` with a numeric target matches:

```vb
10    x = 1         ' line-number prefix — no matching rule
GoTo 100            ' target is a numeric label
100:  MsgBox x      ' numeric lineLabel — ambiguousIdentifier rejects integers
```

The Rust parser handles both forms explicitly (parser.rs lines 951–956):

```rust
TokenKind::IntLit | TokenKind::LongLit => {
    self.advance();                        // consume the line number
    let _ = self.eat(TokenKind::Kw(Kw::Colon)); // optional colon
    return self.parse_stmt(arena);
}
```

**Fix:** Add a second alternative to `lineLabel`:
`lineLabel: (ambiguousIdentifier | integerLiteral) COLON;`
and handle the numeric-target case in `goToStmt` / `goSubStmt`.

---

### L7 — `INTEGERLITERAL` misclassifies Single, Double, and Currency literals

**File:** `VisualBasic6Lexer.g4`, lines 438–443

```antlr
INTEGERLITERAL: [0-9]+ ('E' INTEGERLITERAL)* ( HASH | AMPERSAND | EXCLAMATIONMARK | AT)?;
```

Three problems in one rule:

**a) `#`, `!`, `@` type suffixes produce the wrong token kind.**
`1#` is a Double literal, `1!` is Single, `1@` is Currency. All three are
mis-tokenised as `INTEGERLITERAL` and fed into `integerLiteral` rules. Tools
that inspect the parse tree to infer types will be wrong.

**b) Exponent form creates an integer where VB6 requires a float.**
`1E2` (scientific notation with no decimal point) is a Single literal in
VB6 — it evaluates to 100.0. The grammar tokenises it as `INTEGERLITERAL`
because `('E' INTEGERLITERAL)*` absorbs the exponent.

**c) Recursive `INTEGERLITERAL` in the exponent is structurally invalid.**
`1E2E3` would be accepted (`1` then `E2` then `E3`). No such form exists
in VB6.

The Rust scanner emits distinct `IntLit`, `LongLit`, `SngLit`, `DblLit`,
`CurLit` token kinds based on the trailing suffix character
(`TypeSuffix::from_byte` in `token.rs`).

**Fix:** Remove `( HASH | EXCLAMATIONMARK | AT)?` from `INTEGERLITERAL`
(keep only `( AMPERSAND | PERCENT)?`). Remove the `('E' INTEGERLITERAL)*`
clause. Handle `[0-9]+ 'E' [0-9]+` as a `SINGLELITERAL` token, and route
`#`/`!`/`@` suffixes through the `DOUBLELITERAL`/`SINGLELITERAL`/`CURRENCYLITERAL`
rules.

---

### L8 — `Currency` type keyword absent from lexer and `baseType`

**File:** `VisualBasic6Lexer.g4` + `VisualBasic6Parser.g4`, line 781

`Currency` is VB6 token 0x2A (type-keyword marker `w1 = 0x00006810`).
It is absent from the lexer keyword list, and absent from the `baseType`
rule:

```antlr
baseType
    : BOOLEAN | BYTE | COLLECTION | DATE | DOUBLE | INTEGER | LONG
    | OBJECT | SINGLE | STRING | VARIANT
    ;
```

`Dim x As Currency` only parses accidentally — `Currency` falls back to
`IDENTIFIER` and routes through `complexType`, so it is treated as a
user-defined type name rather than a built-in. Any tool doing type inference
from the parse tree will see `Currency` as an unresolved user type.

Same omission for `Decimal` (token 0x2F) and `Any` (token 0x06, used in
`Declare` statements as `As Any`).

**Fix:** Add `CURRENCY: 'CURRENCY';` to the lexer, add `| CURRENCY` to
`baseType`, and add `CURRENCY` to `ambiguousKeyword`.

---

## Parser Bugs

### P1 — `+=` and `-=` accepted — not valid VB6

**File:** `VisualBasic6Parser.g4`, line 392

```antlr
letStmt
    : (LET WS)? implicitCallStmt_InStmt WS? (EQ | PLUS_EQ | MINUS_EQ) WS? valueStmt
    ;
```

`PLUS_EQ` (`+=`) and `MINUS_EQ` (`-=`) are **not part of the VB6 language**.
They were introduced in VBA 7 (Office 2010+). VB6 has no
compound-assignment operators — only plain `=`. The Rust parser's assignment
dispatch accepts only `Eq`. This grammar silently accepts source that VB6.EXE
would reject.

**Fix:** Change `(EQ | PLUS_EQ | MINUS_EQ)` to just `EQ`. Remove `PLUS_EQ`
and `MINUS_EQ` from the lexer if they serve no other purpose.

---

### P2 — `Call Foo()` with empty argument list fails

**File:** `VisualBasic6Parser.g4`, lines 673–681

```antlr
eCS_ProcedureCall
    : CALL WS ambiguousIdentifier typeHint? (WS? LPAREN WS? argsCall WS? RPAREN)?
    ;
```

The grammar comment says *"empty parantheses are removed"*, but `argsCall`
is **not optional** inside the paren group, and `argsCall` itself requires
at least one central `argCall`:

```antlr
argsCall
    : (argCall? WS? (COMMA | SEMICOLON) WS?)* argCall (WS? (COMMA | SEMICOLON) WS? argCall?)*
    ;
```

When ANTLR enters the optional group on seeing `(` and then finds `)`,
`argsCall` cannot match empty, leaving the `(` unconsumed and causing a
parse error. `Call Foo()` is valid VB6.

The same defect applies to `eCS_MemberProcedureCall`.

The Rust `parse_call_stmt` uses `parse_arg_list` which explicitly handles
zero arguments.

**Fix:** Inside the explicit-call paren group, make `argsCall` optional:
`(WS? LPAREN WS? argsCall? WS? RPAREN)?`.

---

### P3 — Unary `-`/`+` and `Not` have wrong precedence

**File:** `VisualBasic6Parser.g4`, lines 618, 625

```antlr
| (PLUS | MINUS) WS? valueStmt   # vsPlusMinus
…
| NOT (WS valueStmt | LPAREN WS? valueStmt WS? RPAREN)  # vsNot
```

In ANTLR4's left-recursion rewriting, *primary* alternatives (those not
starting with `valueStmt`) call the inner `valueStmt` with **no minimum
precedence**. The inner call therefore consumes every binary operator that
follows, regardless of the VB6 precedence table.

**Unary `-`** has higher precedence than `*`/`/` in VB6:

```
-x * y  →  grammar yields: -(x * y)   ✗
           VB6 correct:     (-x) * y   ✓
```

**`Not`** has higher precedence than `And` in VB6:

```
Not x And y  →  grammar yields: Not (x And y)   ✗
               VB6 correct:     (Not x) And y    ✓
```

The Rust Pratt parser assigns `prefix_bp(Minus) = 24` and
`prefix_bp(Not) = 11`. Since `*`/`/` have `left_bp = 22 < 24` and `And`
has `left_bp = 10 < 11`, they are correctly excluded from the unary
operand parse.

Note: `-2 ^ 2` is coincidentally correct in the ANTLR grammar (`^` is
listed at a higher position than unary `-`) because `^` has `left_bp = 25 > 24`
in the Rust parser, meaning exponentiation IS consumed by the unary-minus
operand — matching VB6 behavior.

**Fix:** The grammar cannot correctly express unary operator precedence
through rule ordering alone. Either use a Pratt-style approach or split
`valueStmt` into precedence-level sub-rules (an `exprNot` rule that only
allows comparison and above as its operand, an `exprNeg` rule that only
allows `^` and atoms as its operand, etc.).

---

### P4 — `#If` body inside a procedure uses `moduleBody`, not `block`

**File:** `VisualBasic6Parser.g4`, lines 415–425

```antlr
macroIfBlockStmt
    : MACRO_IF WS ifConditionStmt WS THEN NEWLINE+ (moduleBody NEWLINE+)?
    ;
macroElseIfBlockStmt
    : MACRO_ELSEIF WS ifConditionStmt WS THEN NEWLINE+ (moduleBody NEWLINE+)?
    ;
macroElseBlockStmt
    : MACRO_ELSE NEWLINE+ (moduleBody NEWLINE+)?
    ;
```

`macroIfThenElseStmt` is listed in `blockStmt` so it can appear inside a
Sub/Function body. However the content between `#If` and `#End If` is
always parsed as `moduleBody` (a sequence of Sub/Function/Type declarations),
never as `block` (statement-level code). Assignments, control flow, and
other statements that appear alone inside a conditional block in a procedure
are incorrectly required to be `moduleBodyElement`s.

The Rust parser dispatches `CcIf`/`CcConst`/`CcElse`/`CcEnd` through
`parse_stmt` inside `parse_block`, i.e. they are handled at statement level
and their bodies are parsed as statement blocks.

**Fix:** Introduce a context-sensitive `macroIfBody` rule that is either
`moduleBody` (at module level) or `block` (inside a procedure). Alternatively,
create two variants of `macroIfThenElseStmt` — one for module level and one
for statement level.

---

### P5 — `vsStruct` accepts invalid tuple expressions

**File:** `VisualBasic6Parser.g4`, line 612

```antlr
| LPAREN WS? valueStmt (WS? COMMA WS? valueStmt)* WS? RPAREN  # vsStruct
```

`(a, b, c)` — a parenthesised comma-separated list — is accepted as a valid
`valueStmt`. VB6 has no tuple or parenthesised list expression: `(a)` (single
value) is valid; `(a, b)` is not. This makes the grammar accept:

```vb
x = (a, b)        ' invalid VB6, parsed without error
If (a, b) Then …  ' invalid VB6, parsed without error
```

**Fix:** Change `vsStruct` to disallow the comma: drop the
`(WS? COMMA WS? valueStmt)*` part so it only wraps a single expression.

---

### P6 — `forNextStmt` accepts VB.NET inline `As` type declaration

**File:** `VisualBasic6Parser.g4`, line 335

```antlr
forNextStmt
    : FOR WS iCS_S_VariableOrProcedureCall typeHint? (WS asTypeClause)? WS? EQ …
    ;
```

`(WS asTypeClause)?` allows `For i As Integer = 1 To 10`. That is VB.NET
syntax — it is not valid VB6. In VB6 the loop variable must be pre-declared;
an `As` clause on the `For` line is rejected by VB6.EXE.

The Rust parser uses `parse_target_expr` for the loop variable, which has
no `As`-clause path.

**Fix:** Remove `(WS asTypeClause)?` from `forNextStmt`.

---

### P7 — `ParamArray` allowed at call sites in `argCall`

**File:** `VisualBasic6Parser.g4`, line 739

```antlr
argCall
    : ((BYVAL | BYREF | PARAMARRAY) WS)? valueStmt
    ;
```

`ByVal` and `ByRef` can appear at call sites in explicit-`Call` invocations.
`ParamArray` cannot — it is a declaration-only modifier. `Call Foo(ParamArray args)`
is accepted by the grammar but is a VB6 compile error.

**Fix:** Change the alternative to `((BYVAL | BYREF) WS)?`.

---

### P8 — `DefType` letter ranges accept multi-character identifiers

**File:** `VisualBasic6Parser.g4`, lines 819–821

```antlr
letterrange
    : certainIdentifier (WS? MINUS WS? certainIdentifier)?
    ;
```

`certainIdentifier` can match multi-character names. VB6 `DefType` ranges
are **single letters only**: `DefInt A-Z`. The grammar accepts
`DefInt Ab-Zz` as syntactically valid; VB6.EXE rejects it.

**Fix:** Replace `certainIdentifier` in `letterrange` with a single-letter
terminal: `IDENTIFIER` (with a semantic check that it is one letter), or add
a dedicated `singleLetter` lexer rule.

---

### P9 — Bare `Print expr` (form-surface output) causes a parse error

**File:** `VisualBasic6Parser.g4`, line 470

```antlr
printStmt
    : PRINT WS valueStmt WS? COMMA (WS? outputList)?
    ;
```

The first `valueStmt` is the file channel. `Print #1, "text"` works because
`#1` is a `FILENUMBER` literal (which is a `valueStmt`). However
`Print "Hello"` (calling the `Print` method on the current form — valid VB6)
cannot parse:

- `printStmt` requires a comma after the first value, so `Print "Hello"` fails there.
- `implicitCallStmt_InBlock` requires `certainIdentifier`, whose second
  alternative needs at least two tokens (`ambiguousKeyword (ambiguousKeyword | IDENTIFIER)+`),
  so a bare `PRINT` keyword as the first and only token also fails.

There is no fallback that successfully routes `Print "Hello"` to a valid
parse tree node.

**Fix:** Add an alternative before `printStmt` that handles `PRINT` without
a mandatory channel and comma, or widen `certainIdentifier` to accept a
single `ambiguousKeyword`.

---

### P10 — `Empty` keyword missing from `literal` rule

**Files:** `VisualBasic6Lexer.g4` + `VisualBasic6Parser.g4`, line 827

`Empty` is VB6 keyword token 0x46. It is the built-in value for an
uninitialized `Variant` variable, parallel to `Nothing` (objects) and
`Null` (database). It is absent from the lexer keyword list and from
`literal`:

```antlr
literal
    : COLORLITERAL | DATELITERAL | doubleLiteral | FILENUMBER
    | integerLiteral | octalLiteral | STRINGLITERAL
    | TRUE | FALSE | NOTHING | NULL_   ← Empty is missing
    ;
```

`x = Empty` parses `Empty` as an `IDENTIFIER` routed through
`iCS_S_VariableOrProcedureCall`. Any tool analysing the parse tree treats
`Empty` as a name-reference rather than a built-in literal constant.

**Fix:** Add `EMPTY: 'EMPTY';` to the lexer, add `| EMPTY` to the `literal`
rule, and add `EMPTY` to `ambiguousKeyword`.

---

### P11 — `For` loop control variable cannot be a member access

**File:** `VisualBasic6Parser.g4`, line 335

```antlr
forNextStmt
    : FOR WS iCS_S_VariableOrProcedureCall typeHint? …
    ;
```

`iCS_S_VariableOrProcedureCall` resolves only a simple identifier with an
optional type hint. The loop control variable cannot be a member expression:

```vb
For obj.Index = 0 To 10   ' valid VB6, fails grammar
```

The Rust parser uses `parse_target_expr` for the loop variable, which handles
member chains.

**Fix:** Change `iCS_S_VariableOrProcedureCall` to `implicitCallStmt_InStmt`
in `forNextStmt` so member access is allowed. (After also applying the P6 fix
to remove `(WS asTypeClause)?`, which would otherwise become even more
ambiguous.)

---

### P12 — `^` exponentiation is left-associative — VB6 requires right-associativity

**File:** `VisualBasic6Parser.g4`, line 617

```antlr
| valueStmt WS? POW WS? valueStmt   # vsPow
```

ANTLR4 defaults to **left-associativity** for binary rules of the form
`rule op rule`. `^` in VB6 is **right-associative**: the exponent is
evaluated right-to-left, matching mathematical convention and the VB6.EXE
runtime:

```vb
x = 2 ^ 3 ^ 2   ' VB6 evaluates as 2 ^ (3 ^ 2) = 2 ^ 9 = 512
                 ' ANTLR parses as  (2 ^ 3) ^ 2 = 8 ^ 2 = 64   ✗
```

The Rust Pratt parser encodes right-associativity via asymmetric binding
powers — `left_bp = 25`, `right_bp = 24` — so the right operand's recursive
call accepts another `^` at the same level:

```rust
TokenKind::Kw(Kw::Caret) => (25, 24, B::Pow),   // left > right → right-assoc
```

All other binary operators use `(2N, 2N+1)` (left < right → left-assoc).

**Fix:** Annotate the `vsPow` alternative with `<assoc=right>`:

```antlr
| <assoc=right> valueStmt WS? POW WS? valueStmt   # vsPow
```

---

## Second Deep Pass — Additional Findings

### P13 — Signed literals fold the sign into the literal, breaking unary precedence

**File:** `VisualBasic6Parser.g4`, lines 1020–1029

```antlr
integerLiteral : (PLUS | MINUS)* INTEGERLITERAL ;
octalLiteral   : (PLUS | MINUS)* OCTALLITERAL ;
doubleLiteral  : (PLUS | MINUS)* DOUBLELITERAL ;
```

A leading `-`/`+` is absorbed **into the literal node** (`vsLiteral`, the
first and highest-priority `valueStmt` alternative). This makes a negative
number bind tighter than every operator — including `^`, which in VB6 binds
tighter than unary minus:

```vb
x = -2 ^ 2   ' VB6:   -(2 ^ 2) = -4
              ' ANTLR:  (-2) ^ 2 =  4    ✗  (literal "-2" formed first, then ^)
```

The Rust parser never folds the sign into the literal — `-` is a prefix
operator with `prefix_bp = 24`, below `^`'s `left_bp = 25`, so `^` is pulled
into the operand and `-2^2` correctly parses as `-(2^2)`:

```rust
TokenKind::Kw(Minus) => Some((24, UnOpKind::Neg)),   // below ^ (25)
```

This is distinct from P3 (unary `-` consuming too much *downward*, toward
`*`/`/`) and from P12 (`^` associativity). Here the bug is *upward*: the sign
is glued onto the literal before any operator is considered, so even the
correctly-ordered `^` alternative cannot reach across it.

Two further consequences of the `(PLUS | MINUS)*` (Kleene star):

* Multiple signs are accepted: `--5`, `+-+5` parse as a single literal.
* `1 - -2` is fine, but `1 --2` (a literal `-2` with no operator between)
  also parses, masking what VB6 would treat as a syntax error.

**Fix:** Remove the sign prefix from `integerLiteral`, `octalLiteral`, and
`doubleLiteral`. Let unary `+`/`-` be handled solely by `vsPlusMinus` (after
that rule's own precedence is corrected per P3).

---

### P14 — `Next i, j` (single `Next` closing nested loops) rejected

**File:** `VisualBasic6Parser.g4`, line 337

```antlr
forNextStmt
    : FOR … NEXT (WS ambiguousIdentifier typeHint?)?
    ;
```

A single `Next` can close several nested `For` loops in VB6 by listing the
loop variables, outermost last:

```vb
For i = 1 To 10
    For j = 1 To 10
Next j, i        ' valid VB6 — closes both loops; grammar accepts only "Next j"
```

The trailing clause accepts at most one `ambiguousIdentifier`, so the `, i`
is left unconsumed and the parse fails. `forEachStmt` (line 329) has the
identical defect.

The Rust parser consumes a comma-separated variable list after `Next`
(`parse_for_stmt`, the `while … eat(Comma)` loop after `expect(Next)`).

**Fix:** Change the trailing clause to
`(WS ambiguousIdentifier typeHint? (WS? COMMA WS? ambiguousIdentifier typeHint?)*)?`.

---

### P15 — Inline `If … Then <stmt>` cannot hold a colon-separated statement list

**File:** `VisualBasic6Parser.g4`, line 359

```antlr
ifThenElseStmt
    : IF WS ifConditionStmt WS THEN WS blockStmt (WS ELSE WS blockStmt)?  # inlineIfThenElse
    | …
```

The single-line form allows exactly **one** `blockStmt` in the Then branch
(and one in Else). VB6 permits a colon-separated list on a single line:

```vb
If ok Then x = 1: y = 2          ' two statements — second is dropped/errors
If ok Then a = 1 Else b = 1: c = 1
```

Because the inline alternative takes a single `blockStmt`, the `:` separator
that follows cannot be attached, and the remainder of the line fails to
parse. (This compounds with L2, which already prevents the `:` from being
tokenised as a separator unless followed by a space.)

**Fix:** Allow a statement list in each inline branch, e.g.
`blockStmt (WS? COLON WS? blockStmt)*` — after first fixing the colon
tokenisation in L2.

---

### L9 — Trailing-dot float `1.` is not tokenised as a float

**File:** `VisualBasic6Lexer.g4`, lines 440–442

```antlr
DOUBLELITERAL:
    [0-9]* DOT [0-9]+ ('E' (PLUS | MINUS)? [0-9]+)* (HASH | AMPERSAND | EXCLAMATIONMARK | AT)?
;
```

`DOUBLELITERAL` requires at least one digit **after** the dot (`[0-9]+`). A
trailing-dot float such as `1.` therefore lexes as `INTEGERLITERAL(1)`
followed by a bare `DOT`, which no expression rule accepts in that position.

The Rust scanner (`scan_number`) sets `has_decimal` on seeing the dot and
does not require a following digit, so `1.` is scanned as one float token.

```vb
x = 1.          ' Rust: float 1.0 ;  ANTLR: INTEGERLITERAL + DOT → parse error
```

(Severity Low — VB6's IDE normalises `1.` to `1#` on entry, so it is rare in
saved source. Listed for parity with the Rust scanner.)

**Fix:** Change the fractional part to allow an empty tail when a dot is
present, e.g. add an alternative `[0-9]+ DOT (… suffixes …)` or relax to
`[0-9]* DOT [0-9]*` with a guard that at least one digit appears on either
side.

---

### L10 / P16 / P17 — Forms the real VB6 compiler accepts but the grammar rejects

Verified directly against the VB6 SP6 compiler running headless
(`VB6.EXE /make <proj.vbp> /out <log>`; exit 0 + "Build … succeeded" =
accept). A deliberately-malformed negative control (`x = 1 +* 2`) was
rejected with "Compile Error … Syntax error", confirming the harness
distinguishes accept from reject.

| Construct | VB6 compiler | Grammar |
|-----------|--------------|---------|
| `EndIf` (one word) | **accept** | reject — `END_IF: 'END IF'` requires the two-word spelling |
| `If x = 1 GoTo done` (no `Then`) | **accept** | reject — `ifThenElseStmt` requires `THEN` |
| `Go To done` (space) | **accept** | reject — no lexer rule for `Go` `To`; only `GOTO: 'GOTO'` |

These match the VB6 keyword table, which carries `EndIf` and a
standalone `Go`/`GoTo`/`GoSub` as distinct tokens. The Rust parser accepts all
three forms: one-word `EndIf` in `end_if`; the no-`Then` legacy form via the
`legacy_goto` branch in `parse_if_stmt`; and the space-separated `Go To` /
`Go Sub` via a `Kw::Go` branch in `parse_stmt` that defers to
`parse_goto_stmt` (which consumes the trailing `To`/`Sub`). `Go Sub` is
oracle-confirmed valid as well.

**P16 fix:** Add `END_IF_ONEWORD: 'ENDIF';` to the lexer (or accept `ENDIF`
in the parser's block-If terminator alongside `END_IF`).

**P17 fix:** Add a single-line alternative to `ifThenElseStmt` that allows
`IF WS ifConditionStmt WS (GOTO | GOSUB) WS valueStmt` with no `THEN`.

**L10 fix:** Add a `GOTO: 'GO' WS? 'TO';`-style rule (or a separate `GO`
token plus a parser rule accepting `GO WS TO`). The same applies to other
split forms VB6 tolerates.

---

## Correction to L6 (`GoTo 100`)

On re-examination, `goToStmt: GOTO WS valueStmt` parses a **numeric target**
correctly: `valueStmt → vsLiteral → literal → integerLiteral` accepts `100`.
So `GoTo 100` is **not** unparseable, contrary to the example in L6. The real,
confirmed defect in L6 is only on the **label-definition** side:

* `lineLabel : ambiguousIdentifier COLON` cannot represent a numeric label
  `100:` (an integer is not an `ambiguousIdentifier`).
* A bare line-number prefix (`10  x = 1`) has no matching rule.

L6's severity and label-definition analysis stand; only the `GoTo 100`
illustration should be struck.

---

## Third Pass — Oracle-Verified Findings (VB6.EXE /make)

All entries below were verified directly against the VB6 SP6 compiler running
headless (`VB6.EXE /make <proj.vbp> /out <log>`; exit 0 + "Build … succeeded" =
accept; "Compile Error …" = reject). Each case was a self-contained single
module so missing-reference/semantic errors could not be confused with syntax
errors.

### L11 — `DATELITERAL` greedily spans two `#` on a line

**File:** `VisualBasic6Lexer.g4`, line 434

```antlr
DATELITERAL: HASH (~ [#\r\n])* HASH;
```

`DATELITERAL` is defined **before** `FILENUMBER`. By maximal munch, any line
that carries two `#` tokens which are *not* an intended date pair is fused into
one bogus date literal — the rule body `(~ [#\r\n])*` happily eats commas,
spaces, and even string contents until the next `#`.

Oracle-verified valid VB6 that this breaks:

```vb
Close #1, #2, #3      ' lexes "#1, #" as a DATELITERAL, then "2, #3…" → parse error
Print #1, "a#b"       ' the '#' inside the string is grabbed: "#1, ""a#" → date literal
```

Both compile cleanly in VB6. This corrupts every multi-channel file statement
and any channel line containing a later `#` (including inside a string literal).

**Fix:** a `#…#` run is a date literal only when its body is actually a
date/time. Validate the content (numeric `m/d/y`, a `H:MM[:SS]` time, or a
month-name form such as `#December, 24, 2000#`) before committing; otherwise
the leading `#` is the file-number / operator token. Disambiguate against
`FILENUMBER` and the `#` operator rather than "anything between two hashes".

### L12 — File channel `#<var>` / `#<expr>` collapses into one `FILENUMBER` literal

**File:** `VisualBasic6Lexer.g4`, line 444

```antlr
FILENUMBER: HASH LETTERORDIGIT+;
```

`#f`, `#fnum`, `#FreeFile` all match as a single opaque literal token. The
channel in real code is almost always a **variable**:

```vb
f = FreeFile : Open "a" For Output As #f : Print #f, x : Close #f   ' all VB6-accepted
```

The grammar captures `f` as literal *text* inside a `FILENUMBER`, so the
variable reference is lost from the parse tree. Affects
`Print/Write/Input/Line Input/Get/Put/Seek/Width/Open/Close #var`. As a
corollary, a parenthesised channel `#(expr)` cannot lex at all (`(` is not a
`LETTERORDIGIT`).

**Fix:** drop `FILENUMBER`; lex `#` as its operator token and parse the channel
as a `valueStmt`.

### L13 — `Debug` keyword absent

**Files:** `VisualBasic6Lexer.g4` + `VisualBasic6Parser.g4`

`Debug` is not a lexer token and has no statement rule. `Debug.Print …` (valid
and ubiquitous) only parses by accident through the generic member-call path
(`iCS_B_MemberProcedureCall`), which models the arguments as a plain `argsCall`
and therefore does **not** model `Print`'s output-list grammar (`;` separators,
trailing `;`, `Spc(n)`, `Tab(n)`):

```vb
Debug.Print "hi"; 1; Tab(5); "x"   ' oracle-accepted; member-call path mismodels it
```

**Fix:** add a `Debug` object/statement and route `Debug.Print` through the same
`outputList` grammar used by `printStmt`.

### P18 — `Resume <line#>` / `Resume 0` rejected

**File:** `VisualBasic6Parser.g4`, line 516

```antlr
resumeStmt: RESUME (WS (NEXT | ambiguousIdentifier))?;
```

There is no numeric alternative, yet both numeric forms are valid VB6
(oracle-accepted):

```vb
Resume 0       ' retry the faulting statement
Resume 100     ' jump to numeric line label 100
```

This is the `Resume` facet of the numeric-line-label gap (see L6). The Rust
parser already handles all four `Resume` forms.

**Fix:** add an integer-literal alternative:
`RESUME (WS (NEXT | integerLiteral | ambiguousIdentifier))?`.

### C1 — Correction to P9: bare `Print` is rejected in a standard module

P9 claims `Print "Hello"` is "valid VB6". Oracle-verified: in a standard
(`.bas`) module VB6 **rejects** bare `Print "x"` (Compile error). Bare `Print`
is valid only where an implicit `Print` method exists — a `Form`, `Printer`, or
via `Debug.Print`. So the fix for P9 is **not** "always accept bare `Print`";
it is context-dependent, and the grammar should not unconditionally admit a
channel-less `Print` at statement level.

### C2 — Correction to L8: `Decimal` and `Any` are not general declared types

L8 proposes adding `Currency`, `Decimal`, and `Any` to `baseType`. Oracle-verified:

| Declaration | VB6 |
|-------------|-----|
| `Dim c As Currency` | **accept** |
| `Dim d As Decimal`  | **reject** |
| `Dim a As Any`      | **reject** |

So: add `Currency` to `baseType` (correct), but **`Decimal` is not a declarable
type** (it exists only as a `Variant` subtype) and **`As Any` is valid only
inside a `Declare`** parameter/return clause. Adding either to `baseType` would
itself be an over-acceptance bug.
