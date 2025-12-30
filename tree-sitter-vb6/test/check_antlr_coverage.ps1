# Check ANTLR example parse coverage
# This script verifies that all ANTLR VB6 examples parse without errors

$ErrorActionPreference = 'SilentlyContinue'

function Check-File {
    param($FilePath)

    Push-Location $PSScriptRoot\..
    $result = npx tree-sitter parse "test\antlr_examples\$FilePath" 2>&1 | Out-String
    Pop-Location

    return -not ($result -cmatch '\(ERROR')
}

$categories = @{
    'Calls' = 'calls'
    'Expressions' = 'expressions'
    'Forms' = 'forms'
    'Statements' = 'statements'
}

$totalFiles = 0
$cleanFiles = 0

foreach ($category in $categories.Keys | Sort-Object) {
    $path = Join-Path $PSScriptRoot "antlr_examples\$($categories[$category])"
    Write-Host "`n============================================================"
    Write-Host $category
    Write-Host "============================================================"

    $files = Get-ChildItem "$path\*.cls", "$path\*.vb", "$path\*.frm", "$path\*.bas" -ErrorAction SilentlyContinue

    $catClean = 0
    foreach ($file in $files | Sort-Object Name) {
        $totalFiles++
        $relativePath = "$($categories[$category])\$($file.Name)"
        $isClean = Check-File $relativePath

        if ($isClean) {
            $cleanFiles++
            $catClean++
            Write-Host "  [OK]  $($file.Name)" -ForegroundColor Green
        } else {
            Write-Host "  [ERR] $($file.Name)" -ForegroundColor Red
        }
    }

    if ($files.Count -gt 0) {
        $pct = [int](($catClean / $files.Count) * 100)
        $fileCount = $files.Count
        Write-Host "`n${category}: $catClean/$fileCount clean ($pct%)"
    }
}

Write-Host "`n============================================================"
Write-Host "Standalone Files"
Write-Host "============================================================"

$standalone = @('helloworld.vb', 'pr578.vb', 'test1.bas')
$standaloneClean = 0

foreach ($fileName in $standalone) {
    $filePath = Join-Path $PSScriptRoot "antlr_examples\$fileName"
    if (Test-Path $filePath) {
        $totalFiles++
        $isClean = Check-File $fileName

        if ($isClean) {
            $cleanFiles++
            $standaloneClean++
            Write-Host "  [OK]  $fileName" -ForegroundColor Green
        } else {
            Write-Host "  [ERR] $fileName" -ForegroundColor Red
        }
    }
}

Write-Host "`nStandalone: $standaloneClean/$($standalone.Count) clean"

Write-Host "`n============================================================"
Write-Host "OVERALL SUMMARY"
Write-Host "============================================================"
Write-Host "Total files: $totalFiles"
Write-Host "Clean parses (no ERROR nodes): $cleanFiles" -ForegroundColor Green
Write-Host "Files with errors: $($totalFiles - $cleanFiles)" -ForegroundColor $(if ($totalFiles - $cleanFiles -eq 0) { "Green" } else { "Yellow" })
if ($totalFiles -gt 0) {
    $pct = [int](($cleanFiles / $totalFiles) * 100)
    Write-Host "Success rate: $pct%" -ForegroundColor $(if ($pct -ge 95) { "Green" } elseif ($pct -ge 80) { "Yellow" } else { "Red" })
}
