<#
.SYNOPSIS
    Generate coverage + clippy reports, submit a SonarQube scan for the vb6_lsp
    project, and wait for the Quality Gate result.

.DESCRIPTION
    This project does NOT run its own SonarQube server. It reuses the shared
    container started from the sibling vba6_rs project
    (docker-compose.sonar.yml there, container "vba6_sonarqube" on port 9000) and
    is registered as a *second project* (sonar.projectKey=vb6_lsp). A single
    global analysis token authenticates analysis for every project on that server.

    Pipeline:
      1. cargo llvm-cov over the default workspace members → target/coverage/lcov.info
         (skip with -SkipCoverage to reuse an existing report).
      2. cargo clippy --message-format=json → target/clippy-report.json
         (cargo is not available inside the scanner container, so the Clippy
         sensor reads this pre-generated file).
      3. sonarsource/sonar-scanner-cli via Docker (no Java needed on the host).
      4. Poll the CE task API until analysis finishes.
      5. Fetch and print the Quality Gate status.

    Prerequisite: the shared SonarQube container must be up. Start it once from
    the vba6_rs project: `..\vba6_rs\scripts\start-sonar.ps1`.

.PARAMETER Token
    SonarQube analysis token. Resolution order: this parameter, then this
    project's .sonar-token, then $env:SONAR_TOKEN, then the shared
    ..\vba6_rs\.sonar-token (the global token minted for the same container).

.PARAMETER ServerUrl
    SonarQube server URL as seen from the host. Default: http://localhost:9000

.PARAMETER SkipCoverage
    Skip the cargo llvm-cov + clippy generation step and reuse existing reports.

.PARAMETER PollIntervalSeconds
    How often to poll the CE task API while waiting. Default: 5.

.PARAMETER TimeoutMinutes
    Give up polling after this many minutes. Default: 10.

.EXAMPLE
    .\scripts\sonar-scan.ps1

.EXAMPLE
    .\scripts\sonar-scan.ps1 -SkipCoverage
#>
[CmdletBinding()]
param(
    [string]$Token,
    [string]$ServerUrl = "http://localhost:9000",
    [switch]$SkipCoverage,
    [int]$PollIntervalSeconds = 5,
    [int]$TimeoutMinutes = 10
)

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

# ── Resolve token ─────────────────────────────────────────────────────────────
# Prefer a local token, then the environment, then the shared token created for
# the same container by vba6_rs's start-sonar.ps1 (a GLOBAL_ANALYSIS_TOKEN is
# valid for any project on the server).
if (-not $Token) {
    $localToken  = Join-Path $PSScriptRoot "..\.sonar-token"
    $sharedToken = Join-Path $PSScriptRoot "..\..\vba6_rs\.sonar-token"
    if (Test-Path $localToken) {
        $Token = (Get-Content $localToken -Raw).Trim()
    } elseif ($env:SONAR_TOKEN) {
        $Token = $env:SONAR_TOKEN
    } elseif (Test-Path $sharedToken) {
        $Token = (Get-Content $sharedToken -Raw).Trim()
        Write-Host "  Using shared token from vba6_rs\.sonar-token (same container)."
    }
}
if (-not $Token) {
    Write-Error "No SonarQube token found. Start the shared server via ..\vba6_rs\scripts\start-sonar.ps1, or set SONAR_TOKEN."
    exit 1
}

# Basic-auth header: token as username, empty password (SonarQube token auth).
$authHeader = @{
    Authorization = "Basic " + [Convert]::ToBase64String(
        [System.Text.Encoding]::ASCII.GetBytes("${Token}:")
    )
}

# ── Step 1 & 2: Coverage + Clippy generation ──────────────────────────────────
$coverageLcov = Join-Path $PSScriptRoot "..\target\coverage\lcov.info"
$clippyReport = Join-Path $PSScriptRoot "..\target\clippy-report.json"

if ($SkipCoverage) {
    Write-Host "`n==> Skipping coverage + clippy generation (-SkipCoverage set)."
    if (-not (Test-Path $coverageLcov)) {
        Write-Warning "  target\coverage\lcov.info not found — SonarQube will report 0% coverage."
    }
    if (-not (Test-Path $clippyReport)) {
        Write-Warning "  target\clippy-report.json not found — Clippy sensor will show no diagnostics."
    }
} else {
    Write-Host "`n==> Checking cargo-llvm-cov..."
    cargo llvm-cov --version 2>&1 | Out-Null
    if ($LASTEXITCODE -ne 0) {
        Write-Host "  Installing cargo-llvm-cov..."
        cargo install cargo-llvm-cov --locked
        if ($LASTEXITCODE -ne 0) { Write-Error "cargo install cargo-llvm-cov failed"; exit 1 }
    }

    Write-Host "==> Ensuring llvm-tools-preview component..."
    rustup component add llvm-tools-preview 2>&1 | Out-Null

    New-Item -ItemType Directory -Force -Path (Split-Path $coverageLcov) | Out-Null
    New-Item -ItemType Directory -Force -Path (Split-Path $clippyReport) | Out-Null

    # Generate the Clippy JSON report on the host — cargo is not available inside
    # the scanner container, so the Clippy sensor must read a pre-generated file.
    # Exit code is intentionally ignored: clippy warnings return non-zero but we
    # always want the report regardless of lint outcome. Runs over the default
    # workspace members (vb6-core is excluded from default-members and does not
    # build on the host).
    Write-Host "`n==> Running Clippy (generating report for SonarQube)..."
    cargo clippy --all-targets --message-format=json 2>$null | Set-Content -Path $clippyReport -Encoding utf8
    Write-Host "  Clippy report written to target\clippy-report.json"

    # Run the default-member test suite under llvm-cov and emit LCOV.
    # The packages are listed explicitly (rather than --workspace, which would
    # pull in the broken vb6-core crate). Listing every default member is
    # REQUIRED: a bare `cargo llvm-cov` only reports coverage for the root
    # vb6-lsp package, leaving the analysis crates with no LCOV data — SonarQube
    # would then count them as 0% covered and understate project coverage.
    Write-Host "`n==> Running tests with coverage (cargo llvm-cov)..."
    cargo llvm-cov --lcov --output-path $coverageLcov `
        -p vb6-lsp -p vb6-ast-derive -p vb6-syntax -p vb6-sema -p vb6-engine
    if ($LASTEXITCODE -ne 0) {
        Write-Error "cargo llvm-cov failed — fix test failures before scanning."
        exit 1
    }

    # Normalize SF: paths in the LCOV file. cargo llvm-cov writes absolute Windows
    # paths (SF:C:\projects\VB6_lsp\...); SonarQube resolves paths relative to the
    # project base dir (/usr/src), so it expects POSIX-relative paths.
    Write-Host "  Normalizing LCOV paths..."
    $lcovRoot = (Resolve-Path (Join-Path $PSScriptRoot "..")).Path + "\"
    (Get-Content $coverageLcov) | ForEach-Object {
        if ($_ -match "^SF:") {
            "SF:" + ($_ -replace "^SF:", "" `
                        -replace [regex]::Escape($lcovRoot), "" `
                        -replace '\\', '/')
        } else { $_ }
    } | Set-Content $coverageLcov -Encoding utf8
    Write-Host "  Coverage written to target\coverage\lcov.info"
}

# ── Step 3: Run sonar-scanner via Docker ─────────────────────────────────────
Write-Host "`n==> Running sonar-scanner (Docker)..."

# Convert Windows path to Docker-compatible POSIX path:
# C:\projects\VB6_lsp  →  /c/projects/VB6_lsp
$projectRoot = (Resolve-Path (Join-Path $PSScriptRoot "..")).Path
$dockerMount = "/" + $projectRoot[0].ToString().ToLower() + `
               ($projectRoot.Substring(2) -replace '\\', '/')

# Docker Desktop for Windows provides host.docker.internal for the host loopback.
$scannerUrl = $ServerUrl -replace 'localhost',    'host.docker.internal' `
                         -replace '127\.0\.0\.1', 'host.docker.internal'

# The Clippy sensor unconditionally runs `cargo --version` even when a
# pre-generated report path is set. Mount a minimal fake `cargo` stub to satisfy
# it. LF-only line endings are required.
$fakeCargoDir  = Join-Path $projectRoot "target\fake-cargo"
$fakeCargoFile = Join-Path $fakeCargoDir "cargo"
New-Item -ItemType Directory -Force -Path $fakeCargoDir | Out-Null
[System.IO.File]::WriteAllText($fakeCargoFile, "#!/bin/sh`necho 'cargo 1.87.0 (sonar-stub)'`n")
$fakeCargoMount = "/" + $fakeCargoFile[0].ToString().ToLower() + `
                  ($fakeCargoFile.Substring(2) -replace '\\', '/')

$scannerOutput = docker run --rm `
    -v "${dockerMount}:/usr/src" `
    -v "${fakeCargoMount}:/usr/local/bin/cargo" `
    -e "SONAR_HOST_URL=$scannerUrl" `
    -e "SONAR_TOKEN=$Token" `
    -e "SONAR_SCANNER_OPTS=-Dsonar.projectBaseDir=/usr/src" `
    sonarsource/sonar-scanner-cli 2>&1

$scannerOutput | ForEach-Object { Write-Host $_ }

if ($LASTEXITCODE -ne 0) {
    Write-Error "sonar-scanner exited with code $LASTEXITCODE"
    exit 1
}

# ── Step 4: Parse CE task ID from scanner log output ─────────────────────────
$taskLine = $scannerOutput | Where-Object { $_ -match "api/ce/task\?id=" } | Select-Object -Last 1
if (-not $taskLine -or $taskLine -notmatch "id=([a-f0-9-]{36})") {
    Write-Error "Could not find CE task ID in scanner output."
    exit 1
}
$ceTaskId = $Matches[1]

$dashLine     = $scannerOutput | Where-Object { $_ -match "dashboard\?id=" } | Select-Object -Last 1
$dashboardUrl = if ($dashLine -match "(https?://\S+dashboard\?id=\S+)") {
    $Matches[1] -replace 'host\.docker\.internal', 'localhost'
} else { "$ServerUrl/dashboard?id=vb6_lsp" }

Write-Host "`n==> CE Task ID : $ceTaskId"
Write-Host "    Dashboard  : $dashboardUrl"

# ── Step 5: Poll until the background task finishes ───────────────────────────
Write-Host "`n==> Polling analysis task (up to ${TimeoutMinutes} min)..."
$taskUrl  = "$ServerUrl/api/ce/task?id=$ceTaskId"
$deadline = (Get-Date).AddMinutes($TimeoutMinutes)
$taskResp = $null
$status   = "PENDING"

while ($status -notin @("SUCCESS", "FAILED", "CANCELLED")) {
    if ((Get-Date) -gt $deadline) {
        Write-Error "Timed out waiting for analysis after ${TimeoutMinutes} minutes."
        exit 1
    }
    Start-Sleep -Seconds $PollIntervalSeconds
    try {
        $taskResp = Invoke-RestMethod -Uri $taskUrl -Headers $authHeader -Method Get
        $status   = $taskResp.task.status
        Write-Host "  [$( Get-Date -Format 'HH:mm:ss' )]  $status"
    } catch {
        Write-Warning "  API poll failed: $_ (will retry)"
    }
}

if ($status -ne "SUCCESS") {
    Write-Error "Analysis ended with status: $status"
    if ($taskResp.task.errorMessage) { Write-Host "  Error: $($taskResp.task.errorMessage)" }
    exit 1
}

# ── Step 6: Quality Gate ──────────────────────────────────────────────────────
Write-Host "`n==> Fetching Quality Gate result..."
$analysisId = $taskResp.task.analysisId
$gateResp   = Invoke-RestMethod `
    -Uri "$ServerUrl/api/qualitygates/project_status?analysisId=$analysisId" `
    -Headers $authHeader -Method Get
$gateStatus = $gateResp.projectStatus.status

Write-Host ""
Write-Host ("=" * 60)
Write-Host "  Quality Gate : $gateStatus"
Write-Host "  Dashboard    : $dashboardUrl"
Write-Host ("=" * 60)

if ($gateStatus -eq "ERROR") {
    Write-Host "`n  Failed conditions:"
    $gateResp.projectStatus.conditions |
        Where-Object { $_.status -eq "ERROR" } |
        ForEach-Object {
            Write-Host "    x  $($_.metricKey): actual=$($_.actualValue)  threshold=$($_.errorThreshold)"
        }
    Write-Host ""
    exit 1
}

Write-Host ""
