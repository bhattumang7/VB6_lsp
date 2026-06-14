#!/usr/bin/env pwsh
# Unified test runner: every test surface in one command.
#
#   Host (x86_64) : vb6-lsp server + analysis engine (syntax/sema/engine)
#   Structural    : frontend/sema tests (AST nodes, parser, spans, scanner,
#                   binding) — pure Rust, run on the host
#   32-bit (i686) : VB6 runtime emulation (vb6-core), which assumes 32-bit
#                   VARIANT/pointer layout
#
# The 32-bit half needs the i686-pc-windows-msvc target installed
# (`rustup target add i686-pc-windows-msvc`).
#
# Usage:  pwsh scripts/test.ps1   (extra args are forwarded to every cargo run)

$args_fwd = $args

Write-Host "== Host suite (x86_64): vb6-lsp + analysis engine ==" -ForegroundColor Cyan
cargo test @args_fwd
if ($LASTEXITCODE -ne 0) { Write-Host "Host suite FAILED." -ForegroundColor Red; exit $LASTEXITCODE }

Write-Host ""
Write-Host "== Structural suite (host): AST / parser / spans / scanner / sema ==" -ForegroundColor Cyan
cargo test-ast @args_fwd
if ($LASTEXITCODE -ne 0) { Write-Host "Structural suite FAILED." -ForegroundColor Red; exit $LASTEXITCODE }

Write-Host ""
Write-Host "== 32-bit suite (i686): VB6 runtime emulation ==" -ForegroundColor Cyan
cargo test-i686 @args_fwd
if ($LASTEXITCODE -ne 0) { Write-Host "32-bit suite FAILED." -ForegroundColor Red; exit $LASTEXITCODE }

Write-Host ""
Write-Host "All suites passed." -ForegroundColor Green
