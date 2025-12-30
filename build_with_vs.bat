@echo off
setlocal

REM Set Rust path
set PATH=C:\Program Files\Rust stable MSVC 1.92\bin;%PATH%

REM MSVC paths - using BuildTools installation
set MSVC_ROOT=C:\Program Files (x86)\Microsoft Visual Studio\2022\BuildTools\VC\Tools\MSVC\14.44.35207
set SDK_ROOT=C:\Program Files (x86)\Windows Kits\10
set LLVM_INCLUDE=C:\Program Files (x86)\Microsoft Visual Studio\2022\BuildTools\VC\Tools\Llvm\x64\lib\clang\19\include

REM Use standard x64 libs
set LIB=%MSVC_ROOT%\lib\x64;%SDK_ROOT%\Lib\10.0.19041.0\ucrt\x64;%SDK_ROOT%\Lib\10.0.19041.0\um\x64

REM Include paths - LLVM clang includes for stdbool.h etc
set INCLUDE=%LLVM_INCLUDE%;%SDK_ROOT%\Include\10.0.19041.0\ucrt;%SDK_ROOT%\Include\10.0.19041.0\um;%SDK_ROOT%\Include\10.0.19041.0\shared

REM Linker path
set PATH=%MSVC_ROOT%\bin\HostX64\x64;%PATH%

cd /d c:\projects\VB6_lsp

echo Building VB6 LSP...
echo INCLUDE=%INCLUDE%
echo.

cargo build --release 2>&1

echo.
if exist target\release\vb6-lsp.exe (
    echo SUCCESS: Binary at target\release\vb6-lsp.exe
) else (
    echo FAILED: Binary not found
)
