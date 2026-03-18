<#
.SYNOPSIS
    Generate Management Group Hierarchy and Network Topology diagrams from audit JSON.

.DESCRIPTION
    PowerShell wrapper around Generate-StandaloneDiagrams.js.
    Generates PNG diagrams (or SVG if sharp is not installed) from the JSON
    output of Invoke-AzureTenancyAudit.ps1.

    Requires Node.js to be installed. Automatically runs npm install if the
    docx module is missing.

.PARAMETER JsonPath
    Path to the JSON file from Invoke-AzureTenancyAudit.ps1.

.PARAMETER OutputDir
    Directory for diagram output. Defaults to the same directory as the JSON file.

.PARAMETER Width
    Diagram width in pixels. Default: 1200.

.PARAMETER Format
    Output format: png or svg. Default: png (falls back to svg if sharp unavailable).

.PARAMETER MgOnly
    Generate only the Management Group hierarchy diagram.

.PARAMETER NetOnly
    Generate only the Network Topology diagram.

.EXAMPLE
    .\Generate-Diagrams.ps1 -JsonPath ".\output\Sense_Azure_ASIS_Data.json"

.EXAMPLE
    .\Generate-Diagrams.ps1 -JsonPath ".\output\data.json" -OutputDir ".\diagrams" -Width 1600

.EXAMPLE
    .\Generate-Diagrams.ps1 -JsonPath ".\output\data.json" -NetOnly

.EXAMPLE
    .\Generate-Diagrams.ps1 -JsonPath ".\output\data.json" -Format svg
#>

[CmdletBinding()]
param(
    [Parameter(Position = 0)]
    [string]$JsonPath,

    [string]$OutputDir,

    [int]$Width = 1200,

    [ValidateSet("png", "svg")]
    [string]$Format = "png",

    [switch]$MgOnly,

    [switch]$NetOnly,

    [string]$OutputPath = ".\AzureBuildDoc_Output"
)

$ErrorActionPreference = "Stop"
$scriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path

# ── Auto-detect JSON if not specified ────────────────────────
if (-not $JsonPath) {
    # Find the latest JSON in the default output directory
    if (Test-Path $OutputPath) {
        $latestDir = Get-ChildItem -Path $OutputPath -Directory | Sort-Object Name -Descending | Select-Object -First 1
        if ($latestDir) {
            $jsonFile = Get-ChildItem -Path $latestDir.FullName -Filter "*_Azure_ASIS_Data.json" | Select-Object -First 1
            if ($jsonFile) {
                $JsonPath = $jsonFile.FullName
                Write-Host "[OK] Auto-detected JSON: $JsonPath" -ForegroundColor Green
            }
        }
    }

    if (-not $JsonPath) {
        # Try current directory
        $jsonFile = Get-ChildItem -Path "." -Filter "*_Azure_ASIS_Data.json" -Recurse -Depth 2 | Sort-Object LastWriteTime -Descending | Select-Object -First 1
        if ($jsonFile) {
            $JsonPath = $jsonFile.FullName
            Write-Host "[OK] Auto-detected JSON: $JsonPath" -ForegroundColor Green
        }
    }

    if (-not $JsonPath) {
        Write-Host "[ERROR] No JSON file found. Either specify -JsonPath or run Invoke-AzureTenancyAudit.ps1 first." -ForegroundColor Red
        Write-Host "        Usage: .\Generate-Diagrams.ps1 -JsonPath 'path\to\data.json'" -ForegroundColor Yellow
        exit 1
    }
}

if (-not (Test-Path $JsonPath)) {
    Write-Host "[ERROR] JSON file not found: $JsonPath" -ForegroundColor Red
    exit 1
}

# ── Check Node.js ────────────────────────────────────────────
$nodeVersion = & node --version 2>$null
if (-not $nodeVersion) {
    Write-Host "[ERROR] Node.js is not installed. Install from https://nodejs.org" -ForegroundColor Red
    Write-Host "        Node.js is required to generate diagrams." -ForegroundColor Red
    exit 1
}
Write-Host "[OK] Node.js: $nodeVersion" -ForegroundColor Green

# ── Check npm modules ────────────────────────────────────────
$docxCheck = & node -e "try { require('docx'); console.log('OK'); } catch(e) { console.log('MISSING'); }" 2>$null
if ($docxCheck -ne "OK") {
    Write-Host "[INFO] Installing npm dependencies..." -ForegroundColor Yellow
    Push-Location $scriptDir
    & npm install 2>$null
    Pop-Location
}

# ── Build arguments ──────────────────────────────────────────
$generatorScript = Join-Path $scriptDir "Generate-StandaloneDiagrams.js"
if (-not (Test-Path $generatorScript)) {
    Write-Host "[ERROR] Generate-StandaloneDiagrams.js not found at: $generatorScript" -ForegroundColor Red
    exit 1
}

$nodeArgs = @($generatorScript, $JsonPath)

if ($OutputDir) {
    $nodeArgs += $OutputDir
}

$nodeArgs += "--width=$Width"
$nodeArgs += "--format=$Format"

if ($MgOnly) { $nodeArgs += "--mg-only" }
if ($NetOnly) { $nodeArgs += "--net-only" }

# ── Run ──────────────────────────────────────────────────────
Write-Host ""
& node $nodeArgs
$exitCode = $LASTEXITCODE

if ($exitCode -ne 0) {
    Write-Host "[ERROR] Diagram generation failed with exit code $exitCode" -ForegroundColor Red
    exit $exitCode
}

Write-Host ""
Write-Host "[OK] Diagram generation complete." -ForegroundColor Green
