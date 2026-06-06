#!/usr/bin/env bash
set -e
cd "$(dirname "$0")/.."
GCC="/c/Users/Umang/AppData/Local/Microsoft/WinGet/Packages/BrechtSanders.WinLibs.POSIX.UCRT_Microsoft.Winget.Source_8wekyb3d8bbwe/mingw64/bin/gcc.exe"
TS=/tmp/tree-sitter-0.24.4/lib
./node_modules/.bin/tree-sitter generate 2>&1 | grep -vE '^(Warning: unnecessary conflicts|  `)' || true
"$GCC" -O1 -o /tmp/vb6parse.exe -I "$TS/include" -I "$TS/src" -I src \
  test/harness.c src/parser.c src/scanner.c "$TS/src/lib.c"
echo "REBUILD_OK"
