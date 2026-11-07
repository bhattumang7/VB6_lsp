# Architecture

This repository holds two products that share one front end:

* **The language tooling**: an LSP server (`vb6-lsp`) and an MCP server (`vb6-mcp`) that answer
  editor and agent queries about VB6 code.
* **The compiler**: p-code generation (`vb6-codegen`), with the goal of compiling VB6 projects
  end to end.

Both are built on the same scanner, parser and binder, so a construct is understood the same way by
the editor and by the compiler.

## Crate graph

```
                     vb6-syntax            scanner, parser, AST, diagnostics
                         |
                     vb6-sema              name binder, type system
                         |
          vb6-formats    |                 .vbp .vbg .frm .frx .res readers and writers
              |     \    |
              |      vb6-engine            Session: hover, completion, rename, tokens, ...
              |          |
        vb6-typelib ---- vb6-project       open a project or group; one Session per project
                           |
   vb6-lsp (src/)  <-------+               tower-lsp server, MCP server, wire-type glue

   vb6-codegen   --> vb6-syntax, vb6-sema  p-code generation
   vb6-core      --> vb6-syntax, vb6-sema  compiler context, COM value layouts
   vb6-ast-derive                          proc macro behind the AST's children() walk
```

Arrows point from a crate to what it depends on. A library never depends on the language server.

| Crate | Owns |
|-------|------|
| `vb6-syntax` | Scanner, parser, AST, diagnostics, conditional compilation. The grammar is validated against the real compiler's behaviour, not against a generated grammar. |
| `vb6-sema` | Name resolution and types. `children()` on every AST node is derived, and an invariant test fails if a node kind is not covered by resolution. |
| `vb6-formats` | File formats: `.vbp`/`.vbg` projects, the designer section of `.frm`/`.ctl`/`.pag`/`.dob`, `.frx`/`.ctx` companion blobs (record readers and writers plus decoded values), `.res` resources. No dependency on the analysis. |
| `vb6-engine` | `Session`: every IDE query over a set of modules. Takes the form designer parser from `vb6-formats`. |
| `vb6-typelib` | Reads the type libraries a project references (`.tlb`, the TYPELIB resource of a DLL/OCX, and on Windows the registry and COM). |
| `vb6-project` | Opens a `.vbp` or `.vbg`: members, references, problems, one `Session` per project, path identity. |
| `vb6-codegen` | P-code generation for the whole module. See below. |
| `vb6-core` | The compiler context, the binder-phase node builder and the 32-bit COM value layouts. |
| `vb6-lsp` (root) | The two server binaries, document handling and the mapping of engine results to LSP and MCP types. |

## Testing

`cargo test` runs the language server and every library. Tests live in each crate's `tests/`
folder or in a `tests/` folder beside the source file they cover; none are written inline.
Tests check the exact result (the parsed tree, the exact messages and positions), not just that
something was reported.

Some tests run over a corpus of third-party VB6 projects that is too large to keep in the
repository. They look in `VB6_CORPUS_DIR` or `../re_lab/corpus` and report that they were skipped
when it is absent.

## The compiler

`vb6-codegen` is a separate workstream from the language tooling. It takes the bound syntax
tree of a module and produces p-code, and it is meant to grow until a project can be compiled
end to end. The rules that keep it growable:

* It depends only on `vb6-syntax` and `vb6-sema`, never on the engine or the server, so the
  compiler builds and tests without any editor machinery. Keep it that way.
* Projects are opened through `vb6-project` (which reads a `.vbp` or `.vbg`, resolves members and
  references, and binds each project) and form resources through `vb6-formats`. A compile front
  end belongs on top of those two crates, as another binary beside the language server, not
  inside it.
* A path that is not implemented yet is written as `todo!()`/`unimplemented!()` stating what it
  must emit, never as a guessed value.
* What is and is not yet supported is tracked in `crates/vb6-codegen/COVERAGE.md`, and the
  crate's tests run as part of `cargo test`.
