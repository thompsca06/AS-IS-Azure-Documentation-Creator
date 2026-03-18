<#
.SYNOPSIS
    One-time setup for Azure Cloud Shell. Clone repo + install dependencies.

.DESCRIPTION
    Run this once in Cloud Shell to pull the toolkit from GitHub and install npm packages.
    After setup, use Build-AzureASISDocument.ps1 to run the full pipeline.

.EXAMPLE
    # First time setup:
    Invoke-WebRequest -Uri "https://raw.githubusercontent.com/thompsca06/AS-IS-Azure-Documentation-Creator/main/Setup-CloudShell.ps1" -OutFile Setup-CloudShell.ps1
    .\Setup-CloudShell.ps1

    # Or if you've already cloned:
    cd ~/AS-IS-Azure-Documentation-Creator
    .\Setup-CloudShell.ps1 -SkipClone
#>

[CmdletBinding()]
param(
    [switch]$SkipClone
)

$ErrorActionPreference = "Stop"
$repoUrl = "https://github.com/thompsca06/AS-IS-Azure-Documentation-Creator.git"
$repoDir = "$HOME/AS-IS-Azure-Documentation-Creator"

Write-Host "============================================================" -ForegroundColor Cyan
Write-Host "  AS-IS Azure Documentation Creator - Cloud Shell Setup" -ForegroundColor Cyan
Write-Host "============================================================" -ForegroundColor Cyan
Write-Host ""

# ── Step 1: Clone repo ──────────────────────────────────────
if (-not $SkipClone) {
    if (Test-Path $repoDir) {
        Write-Host "[INFO] Repo already exists at $repoDir - pulling latest..." -ForegroundColor Yellow
        Push-Location $repoDir
        & git pull
        Pop-Location
    } else {
        Write-Host "[INFO] Cloning repository..." -ForegroundColor Cyan
        & git clone $repoUrl $repoDir
    }
}

Set-Location $repoDir
Write-Host "[OK] Working directory: $repoDir" -ForegroundColor Green

# ── Step 2: Check Node.js ───────────────────────────────────
$nodeVersion = & node --version 2>$null
if (-not $nodeVersion) {
    Write-Host "[ERROR] Node.js not found. Cloud Shell should have it pre-installed." -ForegroundColor Red
    Write-Host "        Try: nvm install --lts" -ForegroundColor Yellow
    exit 1
}
Write-Host "[OK] Node.js: $nodeVersion" -ForegroundColor Green

# ── Step 3: Install npm packages ────────────────────────────
Write-Host "[INFO] Installing npm packages..." -ForegroundColor Cyan
& npm install --omit=optional 2>&1 | Select-Object -Last 3
Write-Host "[OK] npm packages installed (docx)" -ForegroundColor Green

# ── Step 4: Check PowerShell modules ────────────────────────
$modules = @(
    @{ Name = "Az.Accounts"; Required = $true },
    @{ Name = "Az.Resources"; Required = $true },
    @{ Name = "Az.Network"; Required = $true },
    @{ Name = "Az.Compute"; Required = $true },
    @{ Name = "Az.Storage"; Required = $true },
    @{ Name = "Az.RecoveryServices"; Required = $true },
    @{ Name = "Az.Monitor"; Required = $true },
    @{ Name = "Az.Security"; Required = $true },
    @{ Name = "Az.PolicyInsights"; Required = $true },
    @{ Name = "Az.KeyVault"; Required = $true },
    @{ Name = "ImportExcel"; Required = $false },
    @{ Name = "Microsoft.Graph.Identity.DirectoryManagement"; Required = $false }
)

$missing = @()
foreach ($mod in $modules) {
    $installed = Get-Module -ListAvailable $mod.Name -ErrorAction SilentlyContinue
    if ($installed) {
        Write-Host "[OK] $($mod.Name): $($installed.Version)" -ForegroundColor Green
    } elseif ($mod.Required) {
        Write-Host "[MISSING] $($mod.Name) - REQUIRED" -ForegroundColor Red
        $missing += $mod.Name
    } else {
        Write-Host "[SKIP] $($mod.Name) - optional" -ForegroundColor Yellow
    }
}

if ($missing.Count -gt 0) {
    Write-Host ""
    Write-Host "[WARN] Missing required modules. Install with:" -ForegroundColor Yellow
    foreach ($m in $missing) {
        Write-Host "  Install-Module $m -Scope CurrentUser -Force" -ForegroundColor Yellow
    }
}

# ── Step 5: Check Azure connection ──────────────────────────
$ctx = Get-AzContext -ErrorAction SilentlyContinue
if ($ctx) {
    Write-Host "[OK] Connected to Azure as: $($ctx.Account.Id)" -ForegroundColor Green
    Write-Host "     Tenant: $($ctx.Tenant.Id)" -ForegroundColor Green
} else {
    Write-Host "[INFO] Not connected to Azure. Run Connect-AzAccount first." -ForegroundColor Yellow
}

# ── Done ────────────────────────────────────────────────────
Write-Host ""
Write-Host "============================================================" -ForegroundColor Cyan
Write-Host "  SETUP COMPLETE" -ForegroundColor Green
Write-Host "============================================================" -ForegroundColor Cyan
Write-Host ""
Write-Host "USAGE:" -ForegroundColor White
Write-Host "  # Full pipeline (audit + Word doc):" -ForegroundColor Gray
Write-Host "  .\Build-AzureASISDocument.ps1 -CustomerName 'CustomerName' -IncludeEntra" -ForegroundColor White
Write-Host ""
Write-Host "  # Without Entra ID:" -ForegroundColor Gray
Write-Host "  .\Build-AzureASISDocument.ps1 -CustomerName 'CustomerName'" -ForegroundColor White
Write-Host ""
Write-Host "  # Generate diagrams only:" -ForegroundColor Gray
Write-Host "  .\Generate-Diagrams.ps1" -ForegroundColor White
Write-Host ""
Write-Host "  # Scoped to specific subscriptions:" -ForegroundColor Gray
Write-Host '  .\Build-AzureASISDocument.ps1 -CustomerName "Name" -SubscriptionFilter @("sub-id-1","sub-id-2")' -ForegroundColor White
Write-Host ""
