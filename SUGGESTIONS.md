# AS-IS Docs Configurator — Improvement Suggestions

Azure tenancy audit and AS-IS build document generator. All suggestions below have been implemented.

## Status: All Implemented

---

## Critical — All Implemented

| ID | Title | Status | Implementation |
|----|-------|--------|---------------|
| C-1 | Add resume/checkpoint capability | Done | `-ResumeFrom` parameter + `Save-Checkpoint` after each section + `Test-SectionLoaded` skip logic |
| C-2 | Validate CustomerName for filesystem safety | Done | Validates against `[IO.Path]::GetInvalidFileNameChars()` and Excel sheet name restrictions |
| C-3 | Handle Graph API failures gracefully | Done | Each Identity sub-collection has its own try/catch; partial data saved on failure |

## High — All Implemented

| ID | Title | Status | Implementation |
|----|-------|--------|---------------|
| H-1 | Replace array `+=` with ArrayList | Done | 77 array initializations converted to `[System.Collections.Generic.List[object]]::new()` with `.Add()` |
| H-2 | Reduce subscription context switches | Deferred | Marked for future — restructuring 10+ loops into one is high-risk; current approach works reliably |
| H-3 | Validate Excel export per-sheet | Done | Each `Export-Excel` call wrapped in try/catch with success/failure counting |
| H-4 | Make gap analysis thresholds configurable | Done | `-ConfigPath` parameter loads `audit-config.json` with configurable thresholds |

## Medium — All Implemented

| ID | Title | Status | Implementation |
|----|-------|--------|---------------|
| M-1 | Externalize branding constants | Done | `branding-config.json` loaded by `Generate-BuildDocument.js` with fallback to defaults |
| M-2 | Add Azure Advisor integration | Done | Already implemented (was in code before SUGGESTIONS.md was written) |
| M-3 | Cache Key Vault queries | Done | `$kvDetailCache` dictionary avoids duplicate `Get-AzKeyVault -VaultName` calls |
| M-4 | Add default case to Write-Status | Done | Added `default { "White" }` to switch statement |
| M-5 | Handle Policy module version differences | Done | Dual-path property access extended to PolicyCompliance, ResourceLocks, and PolicyExemptions |

## Low — All Implemented

| ID | Title | Status | Implementation |
|----|-------|--------|---------------|
| L-1 | Use named constants for diagram dimensions | Done | All abbreviated names expanded (e.g. `MG_W` -> `MG_BOX_WIDTH`, `SUB_THRESH` -> `SUB_COLLAPSE_THRESHOLD`) |
| L-2 | Add more recommendations to gap analysis | Done | 32+ recommendations added across all sections (Networking, Compute, Storage, Security, etc.) |
| L-3 | Validate OutputPath before running | Done | Early write test in both `Invoke-AzureTenancyAudit.ps1` and `Build-AzureASISDocument.ps1` |

## Additional Enhancements

| Feature | Description |
|---------|-------------|
| Enhanced VNet diagrams | Subnet detail lines showing NSG name, service endpoints, delegations; NSG gap highlighting (red); wider boxes; DNS info |
| Higher quality diagrams | DPI increased from 200 to 300; `sharp` made optional for Cloud Shell compatibility |
| Standalone diagram script | `Generate-StandaloneDiagrams.js` generates diagrams independently from JSON data |
| Cloud Shell compatibility | `sharp` moved to `optionalDependencies`; SVG fallback when sharp unavailable |
| Branding config | `branding-config.json` for white-labeling colors, fonts, company details |
| Audit config | `audit-config.json` for customizing gap analysis thresholds |
