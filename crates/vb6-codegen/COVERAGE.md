# P-code emission coverage

Goal: every p-code sequence the original VB6 compiler can emit for a standard
(`CompilationType=-1`) module, our pipeline emits byte-identically.

Status legend:
- **emit** — produces bytes; covered by an exact test.
- **GAP** — reachable for real VB6 source, not yet emitting (must close).
- **ext** — only reachable via external type libraries / COM objects (needs a
  referenced type lib); out of scope for standalone-module coverage until then.
- **n/a** — not reachable for real source (degenerate decompiler artifact or
  internal error path); not counted against coverage.

## COM-free single-procedure surface

The full grammar-derived corpus (every literal form, operator × type,
statement, array/string/Variant op) for a single COM-free procedure is
byte-exact end-to-end, each cell backed by a committed exact-byte regression
test. This includes: all scalar types (Integer/Long/Single/Double/Currency/
Date/String/Boolean/Byte/Variant) load/store/coercion; all arithmetic/logical/
comparison operators; string concat/compare/fixed-length/Mid/LSet; 1-D and
multi-dim fixed and dynamic arrays with ReDim; Const folding; all control flow
(If/While/Do/For/Select Case incl. ranges and multi-value, GoTo/labels, Exit
For/Do, On Error GoTo); the statement line-number table; and per-procedure
Static locals.

## Expression / statement node opcodes

| op | what | status |
|----|------|--------|
| 0x0b–0x0e | unary/most expr leaves | emit |
| 0x0f | name-ref classify | emit (class 0/2/3); GAP class 1 (typed name reference — front-end wiring, no new back-end work needed); GAP class 4 (in-place node-flag mutation — blocked on our own immutable emit representation, not a missing back-end routine) |
| 0x10–0x15, 0x18, 0x1a | typed leaves/coercions | emit |
| 0x24 | (group) | emit |
| 0x2c | assignment statement | emit (common scalar: RHS + resolved store); GAP: dispatch-binding LHS, array/special LHS, object stack-arg pass-through, object-typed LHS, ByRef-init store, LHS member sub-expression |
| 0x2d | assignment (0x10-child) | emit (scalar); GAP: object-typed child fallthrough |
| 0x2f–0x3b | arithmetic/logical group | emit |
| 0x3e, 0x3f | compare group | emit |
| 0x41 | argument list | emit; GAP op-class 3 (not yet confirmed reachable for real source) |
| 0x42, 0x43 | dispatch type | emit (common + dispatch-binding sub-path) |
| 0x44–0x50 | builtins / conversions | emit |
| 0x51, 0x52 | operator classify | emit (class 0–5); n/a class 6/7 (degenerate) |
| 0x53–0x59 | group | emit |
| 0x5a | complex binary op | emit |
| 0x5c | group | emit |
| 0x5d | type-library-driven cast | ext |
| 0x5e, 0x5f | group | emit |
| 0x60 | member-reference coercion | emit (common path); GAP dispatch/late-bound sub-path + member sub-expression |
| 0x61 | call | emit (by-ref common); GAP: early-bound dispatch call, type-expression argument, ByRef coercion-sequence argument, ByVal (value-returning) call, Variant-result finalize |
| 0x63, 0x65–0x67 | group | emit |
| 0x68 | object child | emit (trailing word); GAP object-child attribute path |
| 0x69 | binary-operation setup | emit (common); GAP ByRef stack-init variant |
| 0x6a–0x6e | instruction group | emit |
| 0x72 | type-node builder | GAP (object/late-bound case only; a scalar/struct-only sub-port is a promising near-term target) |
| 0x73 | group | emit |
| 0x14(0x17), 0x5d | external type-library paths | ext |

## Reference / value-emitter sub-dispatch

| path | status |
|------|--------|
| reference kinds 1/2/3/5/6/7 | emit |
| reference kind 0xa | emit (the sub-case reachable from the currently-wired descriptor builders); GAP: a second sub-case, not reachable from any wired builder today |
| reference kind 4 | GAP — needs a descriptor field this sub-dispatch doesn't carry yet (now independently confirmed by two other routines) plus a finalize-emit mode variant; adding both safely (without risk to already-tested kinds) is a scoped follow-up, not yet done |
| value-emitter kinds 8/9/0xb + typed store | emit |
| value-emitter finalize / re-entry coercion tail | emit |
| store-with-conversion outside the currently-modeled type-offset range | GAP — confirmed reachable for some real type combinations (not just a table-extent formality); needs further work, not yet a table poke |
| binding-emit tail: type/expression coercion propagation | emit (as a standalone, buffer-verified routine) |
| binding-emit tail: member-record binding-descriptor construction | emit (common + a COM-bypass edge case); GAP: the live slot-table path (COM) |
| binding-emit tail: operator-descriptor construction for a resolved binding | emit (scratch-descriptor construction only); GAP: the final result-descriptor's fields, which depend on a side-table this pipeline has no model for |

## Front-half wiring

The binder resolves the common name-reference context and the resolver's
member-access classification (category 4) end-to-end when a construct
supplies it a resolved binding. Remaining GAPs in this area are a mix of (a)
genuinely open back-end research (object/COM member resolution) and (b)
front-end wiring where the back-end routine already exists but `lower.rs`
doesn't yet call it for the relevant source construct — see the per-op notes
above for which is which.

The binding-emit-tail routines above are similarly back-end-ready but not
yet reachable end-to-end. The resolver now accepts the externally-supplied
type descriptor its classification needs (closing the earlier parameter-
threading gap) and its method/object-binding gate is narrowed to only the
genuinely COM-dependent sub-case — but no currently-lowered VB6 construct
ever reaches the categories that route through these routines: this
pipeline's only wired member-record constructor (used by UDT field access)
always builds a record shape that classifies elsewhere. Reaching them
requires lowering a construct this project hasn't audited yet (Property
declarations, class methods, or similar — see below), not further back-end
work.

## Constructs outside the currently audited surface

These were never exercised by the COM-free single-procedure push and their
status is presently **unaudited GAP**, not confirmed `n/a`:

- Procedure calls beyond the by-ref common path (see 0x61 above), multi-
  procedure modules.
- User-defined types (`Type...End Type`): field access/assignment, arrays of
  UDTs, UDTs as parameters.
- Property Get/Let/Set declarations.
- Objects: `Set`, `New`, `Is`/`IsNot`, `TypeOf...Is`, `With`, `For Each`,
  member access (`.`/`!`), `Me`, `Nothing`, `AddressOf`.
- `WithEvents`, `RaiseEvent`, event declarations.
- `Declare` (DLL import) declarations.
- Optional parameters with default values; `ParamArray`.
- File I/O statements (`Open`/`Close`/`Print #`/`Input #`/`Get`/`Put`) and the
  `Debug` object.
- `RSet`; `MidB`/`MidB$` spellings; `ReDim Preserve`; multi-dimensional array
  element load (store is covered); non-`Long` multi-dimensional element
  store; module-level `Const` folding; name-form date literals.
- `GoSub`/`Return`; `On...GoTo`/`On...GoSub` (list form); `Resume`/
  `Resume Next`/`Resume <label>`; `Stop`; `End`; `Error n` as a statement;
  `Exit For`/`Exit Do` (status not independently re-verified since the
  original push).
- Module directives with no proc-body opcodes (`Option Explicit/Base/Compare`,
  `DefType`, `Attribute`).

## Already byte-exact end-to-end (source → p-code, verified vs the compiler)

Scalar `Dim`/global/parameter load+store (all 10 scalar types), arithmetic
(`+ - * / \ Mod ^`), logical (`And Or Xor Eqv Imp`), comparisons
(`= <> < > <= >=`), string ops (concat/compare/fixed-length/`Mid`/`LSet`),
1-D and multi-dim arrays (fixed + dynamic + `ReDim`), `Const` folding, Variant
scalar assignment, and full control flow (`If`/`While`/`Do`/`For`/
`Select Case`/`GoTo`/labels/`Exit For`/`Exit Do`/`On Error GoTo`) — backed by a
grammar-derived exact-byte regression corpus in
`crates/vb6-codegen/tests/pipeline_e2e.rs`.
