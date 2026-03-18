<#
.SYNOPSIS
    Complete pipeline: Audit Azure tenancy → Generate Word document.
    
.DESCRIPTION
    Runs Invoke-AzureTenancyAudit.ps1 to collect data, then calls
    Generate-BuildDocument.js (Node.js) to produce the Word document.
    
.PARAMETER CustomerName
    Customer name for the report.

.PARAMETER IncludeEntra
    Include Entra ID data collection via Microsoft Graph.
    Omit this flag if you do not have access to Entra ID.
    Alias: -IncludeGraph (backward compatible).

.PARAMETER SubscriptionFilter
    Optional subscription IDs to scope collection.

.PARAMETER OutputPath
    Output directory. Defaults to .\AzureBuildDoc_Output

.EXAMPLE
    .\Build-AzureASISDocument.ps1 -CustomerName "Sense" -IncludeEntra
#>

[CmdletBinding()]
param(
    [Parameter(Mandatory = $true)]
    [string]$CustomerName,
    
    [string]$OutputPath = ".\AzureBuildDoc_Output",
    
    [Alias("IncludeGraph")]
    [switch]$IncludeEntra,
    
    [string[]]$SubscriptionFilter
)

$ErrorActionPreference = "Stop"
$scriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path

# Validate OutputPath is writable early
if (-not (Test-Path $OutputPath)) {
    try { New-Item -ItemType Directory -Path $OutputPath -Force | Out-Null }
    catch {
        Write-Host "[ERROR] Cannot create output directory: $OutputPath" -ForegroundColor Red
        Write-Host "        $($_.Exception.Message)" -ForegroundColor Red
        exit 1
    }
}
try {
    $testFile = Join-Path $OutputPath ".write_test_$(Get-Random)"
    [IO.File]::WriteAllText($testFile, "test")
    Remove-Item $testFile -Force
}
catch {
    Write-Host "[ERROR] Output path is not writable: $OutputPath" -ForegroundColor Red
    exit 1
}

# ============================================================
# PRE-FLIGHT: Check dependencies
# ============================================================
Write-Host "============================================================" -ForegroundColor Cyan
Write-Host "  Azure AS-IS Build Document Pipeline" -ForegroundColor Cyan
Write-Host "  Customer: $CustomerName" -ForegroundColor Cyan
Write-Host "  Entra ID: $(if ($IncludeEntra) { 'INCLUDED' } else { 'EXCLUDED (use -IncludeEntra to add)' })" -ForegroundColor $(if ($IncludeEntra) { "Green" } else { "Yellow" })
Write-Host "============================================================" -ForegroundColor Cyan
Write-Host ""

# Check Node.js
$nodeVersion = & node --version 2>$null
if (-not $nodeVersion) {
    Write-Host "[ERROR] Node.js is not installed. Install from https://nodejs.org" -ForegroundColor Red
    Write-Host "        Node.js is required to generate the Word document." -ForegroundColor Red
    exit 1
}
Write-Host "[OK] Node.js: $nodeVersion" -ForegroundColor Green

# Check docx npm module (must be installed locally in script directory)
$docxCheck = & node -e "try { require('docx'); console.log('OK'); } catch(e) { console.log('MISSING'); }" 2>$null
if ($docxCheck -ne "OK") {
    Write-Host "[INFO] Installing docx npm module locally..." -ForegroundColor Yellow
    Push-Location $scriptDir
    & npm install docx --save 2>$null
    Pop-Location
}
Write-Host "[OK] docx module available" -ForegroundColor Green

# Check Az module
$azModule = Get-Module -ListAvailable Az.Accounts
if (-not $azModule) {
    Write-Host "[ERROR] Az PowerShell module not installed." -ForegroundColor Red
    Write-Host "        Run: Install-Module Az -Scope CurrentUser" -ForegroundColor Red
    exit 1
}
Write-Host "[OK] Az module: $($azModule.Version)" -ForegroundColor Green

# Check ImportExcel
$ieModule = Get-Module -ListAvailable ImportExcel
if (-not $ieModule) {
    Write-Host "[WARN] ImportExcel module not installed. Excel export will be skipped." -ForegroundColor Yellow
    Write-Host "       Run: Install-Module ImportExcel -Scope CurrentUser" -ForegroundColor Yellow
} else {
    Write-Host "[OK] ImportExcel: $($ieModule.Version)" -ForegroundColor Green
}

# Check Graph modules (only if Entra is in scope)
if ($IncludeEntra) {
    $graphModule = Get-Module -ListAvailable Microsoft.Graph.Identity.DirectoryManagement
    if (-not $graphModule) {
        Write-Host "[WARN] Microsoft.Graph modules not installed. Entra ID collection may fail." -ForegroundColor Yellow
        Write-Host "       Run: Install-Module Microsoft.Graph.Identity.DirectoryManagement -Scope CurrentUser" -ForegroundColor Yellow
    } else {
        Write-Host "[OK] Microsoft.Graph: $($graphModule.Version)" -ForegroundColor Green
    }
}

Write-Host ""

# ============================================================
# STEP 1: Run the audit script
# ============================================================
Write-Host "STEP 1: Running Azure tenancy audit..." -ForegroundColor Cyan
Write-Host "============================================================" -ForegroundColor Cyan

$auditScript = Join-Path $scriptDir "Invoke-AzureTenancyAudit.ps1"
if (-not (Test-Path $auditScript)) {
    Write-Host "[ERROR] Invoke-AzureTenancyAudit.ps1 not found at: $auditScript" -ForegroundColor Red
    exit 1
}

$auditParams = @{
    CustomerName = $CustomerName
    OutputPath   = $OutputPath
}
if ($IncludeEntra) { $auditParams["IncludeEntra"] = $true }
if ($SubscriptionFilter) { $auditParams["SubscriptionFilter"] = $SubscriptionFilter }

& $auditScript @auditParams

# Find the JSON output
$latestDir = Get-ChildItem -Path $OutputPath -Directory | Sort-Object Name -Descending | Select-Object -First 1
if (-not $latestDir) {
    Write-Host "[ERROR] No output directory found. Audit may have failed." -ForegroundColor Red
    exit 1
}

$jsonFile = Get-ChildItem -Path $latestDir.FullName -Filter "*_Azure_ASIS_Data.json" | Select-Object -First 1
if (-not $jsonFile) {
    Write-Host "[ERROR] JSON output not found in $($latestDir.FullName)" -ForegroundColor Red
    exit 1
}

Write-Host ""
Write-Host "[OK] Audit data: $($jsonFile.FullName)" -ForegroundColor Green

# ============================================================
# STEP 2: Generate Word document
# ============================================================
Write-Host ""
Write-Host "STEP 2: Generating Word document..." -ForegroundColor Cyan
Write-Host "============================================================" -ForegroundColor Cyan

$generatorScript = Join-Path $scriptDir "Generate-BuildDocument.js"
if (-not (Test-Path $generatorScript)) {
    Write-Host "[ERROR] Generate-BuildDocument.js not found at: $generatorScript" -ForegroundColor Red
    exit 1
}

$docxOutput = Join-Path $latestDir.FullName "${CustomerName}_Azure_ASIS_Build_Document.docx"

& node $generatorScript $jsonFile.FullName $docxOutput

if (Test-Path $docxOutput) {
    Write-Host ""
    Write-Host "[OK] Word document generated: $docxOutput" -ForegroundColor Green
} else {
    Write-Host "[ERROR] Word document generation failed." -ForegroundColor Red
    exit 1
}

# ============================================================
# SUMMARY
# ============================================================
Write-Host ""
Write-Host "============================================================" -ForegroundColor Cyan
Write-Host "  BUILD DOCUMENT PIPELINE COMPLETE" -ForegroundColor Cyan
Write-Host "============================================================" -ForegroundColor Cyan
Write-Host ""
Write-Host "OUTPUT FILES:" -ForegroundColor Green
Write-Host "  JSON:   $($jsonFile.FullName)" 
Write-Host "  Excel:  $(Join-Path $latestDir.FullName "${CustomerName}_Azure_ASIS_BuildDoc.xlsx")"
Write-Host "  Word:   $docxOutput"
Write-Host ""
Write-Host "NEXT STEPS:" -ForegroundColor Yellow
Write-Host "  1. Open the Word document and right-click the Table of Contents > Update Field"
Write-Host "  2. Review the Gap Analysis section (Section 15) for critical items"
Write-Host "  3. Review the Operational Compliance Framework (Section 17)"
Write-Host "  4. Manually add: LogicMonitor details, SMTP config, on-prem firewall rules"
Write-Host "  5. Agree backup governance, cost management, and incident management with customer"
Write-Host ""
