#!/usr/bin/env pwsh
# Runs every test in the workspace (the language server and all the libraries).
#
# Usage:  pwsh scripts/test.ps1   (extra args are forwarded to cargo)

cargo test @args
exit $LASTEXITCODE
