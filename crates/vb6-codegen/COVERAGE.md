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

## Procedure terminator (`0x14`)

Every procedure's p-code ends with exactly one implicit-return opcode `0x14`,
unconditionally appended by `lower_proc_pooled` (`lower/mod.rs`) after the
body and any line-table backpatching — regardless of what the last statement
was, and with NO deduplication when that statement is itself an explicit
`Exit Sub`/`Exit Function` (which already emits its own `0x14` via
`lower/stmt.rs`'s `ExitStmt` handling): a bare `Sub Main() : Exit Sub : End
Sub` compiles to two `0x14` bytes, not one.

Oracle-confirmed 2026-07-14 via six independent real-VB6-compiler recaptures
(scalar assign, class field access, Property Get/Let, class-method call with
args, a `GoTo`/line-tracking body, and a `Function` return) after discovering
the corpus-wide gap: 273 of the then-274 `tests/fixtures/*/expected.pcode`
files were missing this trailing byte (a capture-rig omission from whichever
tool produced them, not a real compiler behavior), while the emitter also
never appended one — a mutually-compensating gap that let every fixture pass
without ever exercising the real terminator. Fixed on both sides: the emitter
now appends it, and every affected `expected.pcode` was corrected (271 got a
single `0x14` appended; `e2e_exit_sub` — whose source is already exactly
`Exit Sub` — got a second `0x14` per the no-dedup finding above).

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
- Objects: `Is`/`IsNot`, `TypeOf...Is`, `With`, `For Each`, member access
  (`.`/`!`), `Me`, `AddressOf`. (`Set`/`New`/`Nothing` for a plain object
  local are now covered for the grounded shapes — see "`Set` assignment to a
  plain object local" below.)
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

## Class-member vtable dispatch slot numbering

A `Dim o As New Class1` member access (`o.Field`, `o.Prop`, `o.Method`) emits
a `0x0d <slot>` vtable call whose 2-byte slot operand is assigned by
`resolve_class_member_slots` (vb6-sema `binder.rs`): walk the class's members
in strict source-declaration order, base `0x1c` (the 7-method IDispatch/
IUnknown prefix × 4), stride 4 — a scalar field takes a Get and a Let slot, an
object-typed field additionally a Set slot, and a single property accessor or
a `Sub`/`Function` one slot each; a non-exposed `Private` member is not a
dispinterface member and consumes no slot.

This is a **closed-form rule matched to the real VB6 compiler's output, not a
port of a compiler routine** — the compiler assigns the final slot value in a
separate binary (its type-info layout engine), so no in-scope source computes
it and the only ground truth is the emitted bytes. The rule's OUTPUT was
audited byte-for-byte against the real compiler across a matrix of member
shapes and matches in every case (Outcome: closed form is correct):

- scalar fields of mixed types (`Long`/`Double`/`String`) — every scalar field
  is 2 slots regardless of type;
- 5 fields accessed in an order different from declaration — the slot value
  follows declaration order, not access order;
- `Get`-only, `Let`-only, and combined `Get`+`Let`+`Set` properties — each
  accessor is its own independent slot in source order, no slot sharing;
- an object-typed field interleaved between two scalar fields — 3 slots, and
  the following scalar's slots shift past all three;
- a `Sub` and a `Function` interleaved with fields and a property — 1 slot
  each, positionally;
- a bare `Private` field between two `Public` fields — 0 slots (subsequent
  members are **not** shifted);
- declaration-order permutations of the same members — slots follow source
  order (not alphabetical/type order).

Committed byte-exact fixtures (`tests/fixtures/`): `e2e_class_field_scalar_
access`, `e2e_class_multi_field_and_property`, `e2e_class_property_let_before_
get`, `e2e_class_property_set_object`, `e2e_class_field_private_skipped`,
`e2e_class_field_five_access_order`, `e2e_class_field_source_order`,
`e2e_class_object_field_between_scalars`, `e2e_class_method_mixed_members`,
`e2e_class_property_get_double`, `e2e_class_property_get_string`,
`e2e_class_property_get_object`, `e2e_class_property_let_double`,
`e2e_class_property_let_string`, plus the method-argument fixtures
(`e2e_class_method_*`).

The class-member vtable-dispatch scratch temp (`class_member_base`) is sized
to the actual VBA type read through a Get access (`class_get_temp_ctx`,
`lower/decl.rs`) — oracle-confirmed distinct read-back opcodes per type
(`Long` reads back with `0x6c`, `Double` with `0x6f`, both oracle-confirmed;
`oracle_bank/c1_get_double`). A proc mixing Get accesses of different sizes
(e.g. a `Long` Get and a `Double` Get in the same proc) is gated
(`UnsupportedType`) rather than guessed — no oracle capture shows how the
real compiler sizes a shared temp across mixed Get types in one proc.

A `String`-returning Get (`oracle_bank/c1_get_string`) uses a STRUCTURALLY
DIFFERENT mechanism, not just a different opcode: the temp is read back with
the steal opcode `0x3e` (push the temp's BSTR pointer AND zero the temp
slot — no separate release needed, since the temp no longer owns anything
after), and the assignment that consumes it uses the move-store `0x31`
(`ClassFieldRef::is_string`, `lower/expr.rs`; `value_is_class_get_string`,
`lower/assign.rs`) rather than the refcounted copy-store `0x43` a plain
string-variable source would get.

An `Object`-returning Get (`oracle_bank/c1_get_object`) is a THIRD distinct
temp read-back mechanism: opcode `0x51`, a plain 4-byte pointer read distinct
from both `0x6c` (the `RT_LOAD_BY_CTX` load used for an ORDINARY Object
*variable*, not a Get's out-temp) and String's steal-load `0x3e`
(`ClassFieldRef::is_object`, `lower/expr.rs`). Its only grounded client
spelling is `Set x = o.P` where `x` is a plain `Object`-typed local (`Dim x
As Object`) — a whole new target path (`object_typed_local` in
`lower/assign.rs`, distinct from `plain_object_local`'s specific-class-typed
match) feeds a shared `lower_set_from_class_get` helper that emits the Get
sequence then the refcounted AddRef-store `0x19`, the SAME store already
grounded for `Set o = New`/`Set o = Nothing`. A `Set x = o.P` target
declared as a SPECIFIC class type (`Dim x As Class1`, not plain `Object`) is
gated (`UnsupportedNode`) — no oracle capture grounds that shape.
`class_member_temp_ctx`/`count_class_get_temps` (`lower/decl.rs`) both scan
`SetAssign` as well as `Assign` now, so a proc whose ONLY class-member access
is a `Set`-spelled Get still reserves/sizes the shared temp correctly.

A `Double`-typed Property Let (`oracle_bank/c2_let_double`) stages its
argument with the FPU-aware store `0xfd 0xc9` (pop FPU-top, store as Double
with an overflow check) instead of the plain top-of-eval-stack store `0x59`
(`ClassFieldRef::is_double`, `lower/expr.rs`). The class-member scratch
temp's sizing scan (`class_member_temp_ctx`, `lower/decl.rs`) now covers
Property Let STAGING targets too, not just Get reads — a proc with only a
Double `Property Let` (no Get at all) still needs the shared temp sized to
8 bytes.

A `String`-typed Property Let (`oracle_bank/c2_let_string`) is different
again: the pushed value is COPY-STORED (`0x43`) into the shared temp
(properly owning/addref'ing the BSTR, unlike every other type's plain
value-staging), the call then receives the temp's ADDRESS (`0x04`) rather
than a staged value, and the temp is explicitly released (`0x2f`) after the
call returns (`ClassFieldRef::is_string` in `lower_class_field_store`).

**Two vtable-dispatch operands previously hardcoded were discovered to be
WRONG in general** while grounding this slice (both still happened to be
correct for every single-class, no-preceding-string-literal fixture shipped
before this slice, which is why they passed undetected):
- The object-resolve opcode's own operand (`0x24 <idx>`) is NOT a fixed `0`
  — it's a `ClassConstKind::Create` const-pool entry like `New`'s, and the
  pool is genuinely module-wide and SHARED with the string-literal pool (one
  flat sequential index space, confirmed by `c2_let_string`: the string
  literal claims index 0, pushing the class-create entry to index 1).
- The vtable-call opcode's own second operand (`0x0d <slot> <idx>`) is NOT a
  fixed `1` either — it's a NEW pool-entry kind, `ModuleConstEntry::
  MemberType`. Read only on the vtable call's error path (a type-mismatch
  message), per the `0x0d` handler's own disassembly, but its index is
  consumed unconditionally.

Both are now real pool interns (`intern_class_const`/
`intern_member_type_const`, `lower/mod.rs`). The object-resolve fix
(`0x24 <idx>`) was applied everywhere it appears, including
`lower_class_method_call` (`intrinsics.rs`) — it's the SAME opcode/mechanism
regardless of what kind of vtable call follows, so leaving it hardcoded
there would have been an equally-wrong latent bug.

**`MemberType` was originally modeled as deduped by the accessed member's own
type** (`e2e_class_multi_field_and_property`'s six same-typed `Long`
accesses all landing on one shared entry seemed to confirm this) — and
`lower_class_method_call`'s own hardcoded `01 00` was left unfixed pending a
later slice. Slice #7/#9 (2026-07-20) DISPROVED the type-keyed model with two
fresh real-compiler captures: a `Double` Get followed by a no-return `Sub`
call, and a `Long` Get followed by a `String` Get, EACH in one proc — both
pairs land on the SAME operand index despite differing types. `MemberType`
is actually a SINGLE per-module shared descriptor, lazily allocated once and
reused by every Get/Let/Set/Sub/Function vtable call regardless of type —
now a unit-payload `ModuleConstEntry::MemberType`, and `lower_class_method_
call`'s hardcode is fixed too.

**The class-member scratch temp area is likewise NOT one shared slot for the
whole proc** — the same two captures show the `Double` Get's temp and the
`Sub` call's `Long`-argument stage temp at two SEPARATE, non-overlapping
frame offsets (and the `Long`/`String` Get pair likewise), even though same-
TYPE repeated accesses still correctly share one slot. Now modeled as one
region PER DISTINCT frame type-context (`decl::ClassMemberRegion`,
`class_member_regions`), each sized to the max CONCURRENT slots that context
ever needs (a class-method call's own multiple same-type arguments, e.g.
`Three(Long,Long,Long)`, still correctly need 3 concurrent slots within
their one region). `LowerCtx::class_member_base: usize` became
`class_member_bases: HashMap<usize, usize>` (context → region base), with a
`class_member_slot(ctx, index)` accessor replacing the old flat indexing.

**Class-method call argument staging** buckets each staged argument by its
OWN type-context (`decl::bucket_method_call_slots`), in parameter
DECLARATION order (bucket-creation order) — oracle-confirmed via
`c4_sub_2arg_mixed_call`'s `DoIt(x As Long, s As String)`. A `Function`
result temp (value position) claims a slot in ITS OWN return-type bucket,
positioned after that bucket's own argument slots — confirmed via
`c5_func_string`'s result+2-String-args all sharing one bucket. A staged
`String` argument uses the SAME copy-store/by-address/post-call-release
convention already grounded for Property Let's `String` staging (`0x43`/
`0x04`.../`0x2f`), previously unhandled in the method-call staging path
(would have emitted a plain `0x59`, wrong for `String`).

**A class-method `Function` call in value position is grounded for `Long`,
`String`, and `Object` return types** (`e2e_class_function_string_return`,
`e2e_class_function_object_return`, plus the pre-existing `Long` coverage) —
`c5_func_string`/`c5_func_object` recompiled and re-extracted fresh this
session, byte-identical to the previously-banked capture, promoting both
from MEDIUM to HIGH confidence. The result temp's read-back opcode follows
the same per-type split already grounded for a class-member Get: `0x6c`
(`Long`), `0x3e` (`String`, a steal-load — no separate release, the temp is
left zeroed), `0x51` (`Object`, a plain pointer read). A `String` result is
consumed by the caller's move-store (ctx 9, `0x31`), the same rule already
used for a class-member Get returning String; an `Object` result requires
the `Set` spelling (`Set r = o.Method(...)`) and is consumed by the
AddRef-store `0x19` — a plain `Assign` to an `Object`-typed target from a
class-method call is rejected rather than guessed (real VB6 requires `Set`
for that shape). Two or more staged String argument temps needing release
share ONE bulk release (`0x32 <byte-len> <offsets…>`, the same opcode
already grounded for multi-temp concat cleanup) instead of one `0x2f`
apiece — oracle-confirmed by `c5_func_string`'s two-argument release, which
also proved release order is PARAMETER DECLARATION order, not push order.
For a `Function` call in value position, that release can't be emitted
inside `lower_class_method_call` itself — the byte order is `call →
result-load → caller's own store of the result → release` — so
`lower_class_method_call` returns the pending release list and the caller
(the plain-`Assign`/`Set`-assign lowering) emits it after its own store.

**A class-method `Double` parameter (VB6's default ByRef, no keyword needed)
passed a non-addressable source stages via the SAME FPU-aware store already
grounded for Property Let's Double staging** (`0xfd 0xc9 <offset>`), with no
follow-up address push (unlike `String`'s explicit `0x04` re-push) —
oracle-confirmed via `oracle_bank/c6_nonlong_arg_and_return`
(`e2e_class_function_nonlong_arg_and_return`), which ALSO combines this with
a `Double` return type in value position (result read-back generalized from
a hardcoded 3-way match to a proper `RT_LOAD_BY_CTX` table lookup, resolving
`Double` to `0x6f`; the consuming store needed no special-casing, unlike
`String`'s move-store carve-out — `load_store_ctx(Double)` already maps to
the correct `0x74`). This capture was originally banked at the wrong
preamble length (missing its leading 16 bytes, the argument's own
push+stage sequence) — a real instance of "recapture MEDIUMs before
landing" catching a genuinely wrong byte sequence, not a formality.

**An `Object`-typed parameter's argument being a SPECIFIC-class-typed `As
New` local** (`Dim y As New Class1` passed where the parameter is plain
`Object`) is a DIFFERENT source shape from the plain-`Object`-variable case
(`e2e_class_method_arg_variable`) — it reads via the SAME lazy-fetch
sequence already grounded for `Set o = otherObjLocal` (`04 <src>` LdAddr
then `56 <create-idx>`) rather than the generic coercion pipeline (which has
no notion of coercing a `UserDefined`-typed variable into an `Object`-tagged
node and would error `UnsupportedType`), then feeds the SAME `fd 9c
<temp_off>` staging store every Object argument uses. Oracle-confirmed:
`oracle_bank/c8_obj_byval_param` (`e2e_class_method_object_byval_param`).
Surprising asymmetry, confirmed by two independent fixtures: the lazy-fetch
temp IS released afterward (`0x1a <offset>`, the same opcode already
grounded for a Property Set's staged Object temp) but the plain-Object-
variable case's temp is NOT — releasing both unconditionally regressed the
already-shipped `e2e_class_method_arg_variable` fixture immediately, caught
by the full-workspace run before landing. Gated (not guessed): two-or-more
such release temps in one call (no bulk form confirmed for Object) and any
`Function`-in-value-position combination (no evidence for release-vs-
result-store ordering).

**A ByRef `Object` argument sourced from a specific-class-typed `As New`
local is a genuinely different mechanism from the ByVal case above, not a
variant of it** — oracle-confirmed: `oracle_bank/c8_obj_byref_param`
(`e2e_class_method_object_byref_param`, the LAST of the original 15-slice
fan-out manifest). The lazy-fetch load is identical, but the loaded value is
MOVE-STORED (`fc f8`, not `fd 9c`) into a temp, the temp's ADDRESS (not a
staged value) is pushed as the argument, and — because the callee can
reassign through a ByRef parameter — the temp's FINAL value is written BACK
into the source local after the call (`6c` load, `3d <idx>` coerce, `19`
AddRef-store) before the temp is released (`1a`, after the write-back here,
not immediately after the call like the ByVal case). A flagged, UNRESOLVED
open question from this capture: the write-back's `3d` coercion operand is
IDENTICAL to the call's own shared member-type-descriptor operand
(`intern_member_type_const`), not a fresh class-specific `TypeDesc` entry —
whether the two concepts are genuinely the same pool entry or this is a
coincidental single-class-fixture collision is undetermined with one data
point; do not extend the "member-type is class-independent" conclusion
further, and do not resolve this quietly, until a two-distinct-class
fixture exists to actually distinguish the hypotheses.

Codegen surface still gated (slot rule confirmed by capture, but the
surrounding load/store not yet lowerable, so no byte-exact fixture drives
them from this path): object-typed **field** Get/Set (`Set y = o.ObjField` /
`Set o.ObjField = y` return `UnsupportedNode`; only an explicit `Property Set`
is grounded), a bare 0-argument `Sub` **statement** call (`o.Method` with
the result discarded returns `UnsupportedNode`; a 0-arg `Function` in value
position, `x = o.Method()`, does lower), and `Set x = o.P`/`Set x =
o.Method()` against a SPECIFIC-class-typed target (`Dim x As Class1`, not
plain `Object` — see the "`Set` assignment to a plain object local" section
below).

## `Set` assignment to a plain object local

`Set localVar = v` where `localVar` is a bare `Dim x As ClassName` /
`Dim x As New ClassName` identifier (not `o.Field` — that's the vtable-
dispatch path above) is lowered in `lower_set_plain_object_local`
(`lower/assign.rs`). Three source forms are oracle-confirmed byte-exact, all
from one `Sub Main` exercising all three against a single class in sequence:

- `Set o = New ClassName` — `fd f4 <create-idx>` (construct) then `19 <dest>`
  (pop, AddRef, release-old-and-store).
- `Set o = otherObjLocal` where `otherObjLocal` was declared `As New` — `04
  <src>` (LdAddr) then `56 <create-idx>` (lazy-fetch: construct on first null
  access, otherwise a plain owned load) then `fc f8 <dest>` (pop,
  release-old-and-store, no AddRef — steals the reference `56` already owns).
- `Set o = Nothing` — `fc 63` (push literal 0) then `3d <type-idx>` (coerce to
  the target's declared class) then `19 <dest>` (AddRef-store; safe on null).

Each source form's class-constant-table operand (`<create-idx>`/`<type-idx>`)
is a per-procedure table, entries in first-use order, deduped by `(kind,
class symbol)` — confirmed a SINGLE shared sequence across the `Create`/
`TypeDesc` kinds by the oracle capture itself (the `New`-expr and the `As
New` lazy-fetch of the same class dedupe to index 0; the `Nothing` coercion
of that same class is a different kind and gets index 1, not a fresh 0).

Gated (no oracle capture distinguishes these from the shapes above, so they
return an error rather than guess): `New` of a class other than the target's
declared type; a Set-source local NOT declared `As New` (only the lazy-fetch
opcode is grounded — a plain object local's read-as-Set-source may use a
different, ungrounded opcode); a var-to-var copy across two different
classes.

Committed byte-exact fixture: `e2e_set_new_reassign_nothing`.
