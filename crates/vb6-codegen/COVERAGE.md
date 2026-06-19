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

The list below is the bound-node opcode dispatch in `emit.rs::emit_expr`
(opcodes `0x05..0x73`) plus the sub-dispatches and the front-half wiring.

## Expression / statement node opcodes

| op | what | status |
|----|------|--------|
| 0x0b–0x0e | unary/most expr leaves | emit |
| 0x0f | name-ref classify | emit (class 0/2/3); **GAP** class 1 (symbol var-type), class 4 (in-place node mutation) |
| 0x10–0x15, 0x18, 0x1a | typed leaves/coercions | emit |
| 0x24 | (group) | emit |
| 0x2c | assignment statement | emit (common scalar: RHS + resolved store); **GAP** dispatch/compound-op/array/object/ByRef-init sub-paths |
| 0x2d | assignment (0x10-child) | emit |
| 0x2f–0x3b | arithmetic/logical group | emit |
| 0x3e, 0x3f | compare group | emit |
| 0x41 | argument list | emit |
| 0x42, 0x43 | dispatch type | emit (common); **GAP** dispatch-binding sub-path (object) |
| 0x44–0x50 | builtins / conversions | emit |
| 0x51, 0x52 | operator classify | emit (class 0–5); **n/a** class 6/7 (degenerate) |
| 0x53–0x59 | group | emit |
| 0x5a | complex binary op | emit |
| 0x5c | group | emit |
| 0x5d | type-library-driven cast | **ext** |
| 0x5e, 0x5f | group | emit |
| 0x60 | member-reference coercion | emit (common path via resolver+binder+value-emitter); **GAP** dispatch/late-bound sub-path + member sub-expr |
| 0x61 | call | emit (by-ref common); **GAP** ByVal / dispatch / member numbering |
| 0x63, 0x65–0x67 | group | emit |
| 0x68 | object child | emit (trailing word); **GAP** object-child attr path |
| 0x69 | binary-operation setup | emit (operands + operator descriptor kinds 9/0xb, nOp 5/6); **GAP** kind-0xa + nOp 1-4 finalize (EbEmitExpressionOp opcode base, asm-needed) |
| 0x6a–0x6e | instruction group | emit |
| 0x72 | type-node builder | **GAP** (resolver / EbCreateTypeNode3) |
| 0x73 | group | emit |
| 0x14(0x17), 0x5d | external type-lib paths | **ext** |

## Reference / value-emitter sub-dispatch (`emit_reference`, `emit_value2` paths)

| path | status |
|------|--------|
| reference kinds 1/2 | emit |
| reference kinds 3–9 | **GAP** (resolver-built descriptor) |
| value-emitter kinds 8/9/0xb + typed store | emit |
| value-emitter kind-3 resolver-base finalize | **GAP** |

## Front-half wiring (the lever that unblocks most GAPs)

Progress: the binder front-half (`binder.rs`) now resolves the common
name-reference context (disc 1) → `(kind, byref)`, and resolver **category 4**
emits its descriptor end-to-end via `call_conv_descriptor` when given that
binding. Remaining binder work: disc 3/5/6 (document-context allocator), disc 4
(COM `ITypeInfo` slot), resolver categories 0xd/0xe/0xf (binding-emit tail), and
wiring `emit.rs` 0x60/0x2c/0x69 → `resolve_reference2` with the binder result.

The GAPs `0x2c / 0x60 / 0x69 / 0x72`, `0x0f` class 1/4, `0x42/0x43` dispatch,
`0x61` ByVal/dispatch, `0x68` object-child, and reference kinds 3–9 all bottom out
at the **resolver / declaration front-half**: `lower.rs` must build the symbol
records + context (`pNode[5]`) the ported resolver reads, then `emit.rs` must call
`resolve_reference2` → value-emitter. Ported so far (this work):
`heap.rs` (allocator+bags), `typenode.rs`, `decl.rs` (member/slot records),
`resolver.rs` (classifier + value cases + category-4 selection). Remaining:
- `EbResolveExprNode` (bind-result resolution of the context node) OR have
  `lower.rs` supply the resolved handle from `vb6-sema`.
- resolver categories 0xd/0xe/0xf (binding-emit tail) + the method-binding path.
- wire `emit.rs` 0x60/0x2c/0x69/0x72 → `resolve_reference2`.

## Already byte-exact end-to-end (source → p-code, verified vs the compiler)

Scalar `Dim`/global/parameter load+store (Integer/Long/Double), arithmetic
(`+ - * / And Or Xor`), comparisons (`= <> < > <= >=`), and control flow — 20
`pipeline_e2e` + 21 `oracle_pcode` exact-byte tests.
