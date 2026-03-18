# Azure Tenancy AS-IS Build Document - Audit Toolkit

## Overview

This toolkit automates the collection of Azure tenancy configuration data to populate the AS-IS Build Document template. It interrogates all aspects of an Azure environment via PowerShell Az modules and Microsoft Graph API.

## Files

| File | Purpose |
|------|---------|
| `Build-AzureASISDocument.ps1` | **One-click pipeline** - runs audit + generates Word doc |
| `Invoke-AzureTenancyAudit.ps1` | PowerShell script - collects all Azure data → JSON + Excel |
| `Generate-BuildDocument.js` | Node.js script - reads JSON → produces populated Word doc |
| `Generate-Diagrams.js` | SVG/PNG diagram generator for MG hierarchy and network topology |
| `Generate-StandaloneDiagrams.js` | **Standalone diagram generator** - produces diagrams from JSON |
| `branding-config.json` | Configurable colours, fonts, page layout, company details |
| `audit-config.json` | Configurable gap analysis thresholds and standards |
| `Azure_AS-IS_Build_Document_Template.docx` | Blank Word template (for manual use if preferred) |

## Quick Start (One Command)

```powershell
# Full audit including Entra ID (Azure AD)
.\Build-AzureASISDocument.ps1 -CustomerName "Sense" -IncludeEntra

# Azure resources only (no Entra ID access needed)
.\Build-AzureASISDocument.ps1 -CustomerName "Sense"
```

This runs both steps automatically and produces JSON, Excel, and Word outputs.

> **Note:** Use `-IncludeEntra` only when you have access to the Entra ID tenant (Microsoft Graph). Without the flag, the Entra ID section is marked "Not in scope" and no gap is logged. The old `-IncludeGraph` flag still works as an alias.

## Prerequisites

### PowerShell Modules

```powershell
# Core Azure modules
Install-Module Az -Scope CurrentUser -Force

# Excel export
Install-Module ImportExcel -Scope CurrentUser -Force

# Entra ID / Identity (optional but recommended)
Install-Module Microsoft.Graph.Identity.DirectoryManagement -Scope CurrentUser -Force
Install-Module Microsoft.Graph.Identity.SignIns -Scope CurrentUser -Force
Install-Module Microsoft.Graph.Applications -Scope CurrentUser -Force
Install-Module Microsoft.Graph.DirectoryObjects -Scope CurrentUser -Force

# AVD (if applicable)
Install-Module Az.DesktopVirtualization -Scope CurrentUser -Force
```

### Permissions Required

| Scope | Permission | Purpose |
|-------|-----------|---------|
| Azure RBAC | **Reader** on all subscriptions | Read all Azure resource config |
| Azure RBAC | **Management Group Reader** | Read MG hierarchy |
| Entra ID | **Global Reader** (or equiv) | Read tenant config |
| Graph API | User.Read.All | Read user/group data |
| Graph API | Policy.Read.All | Read Conditional Access |
| Graph API | Application.Read.All | Read app registrations |
| Graph API | RoleManagement.Read.Directory | Read directory roles |
| Graph API | Directory.Read.All | Read tenant details |

## Usage

### Basic (Azure resources only — no Entra ID)

```powershell
.\Invoke-AzureTenancyAudit.ps1 -CustomerName "CustomerName"
```

### Full audit including Entra ID

```powershell
.\Invoke-AzureTenancyAudit.ps1 -CustomerName "CustomerName" -IncludeEntra
```

### Scoped to specific subscriptions

```powershell
.\Invoke-AzureTenancyAudit.ps1 -CustomerName "CustomerName" -IncludeEntra -SubscriptionFilter @("sub-id-1", "sub-id-2")
```

### Custom output location

```powershell
.\Invoke-AzureTenancyAudit.ps1 -CustomerName "CustomerName" -OutputPath "C:\Reports\Azure"
```

### With custom thresholds

```powershell
.\Invoke-AzureTenancyAudit.ps1 -CustomerName "CustomerName" -ConfigPath ".\audit-config.json"
```

### Resume interrupted audit

```powershell
# If an audit was interrupted, resume from the checkpoint:
.\Invoke-AzureTenancyAudit.ps1 -CustomerName "CustomerName" -ResumeFrom ".\AzureBuildDoc_Output\20260226_143000\_checkpoint.json"
```

### Generate diagrams only

```powershell
# Generate standalone diagrams from existing JSON data
node Generate-StandaloneDiagrams.js ".\AzureBuildDoc_Output\20260226_143000\Sense_Azure_ASIS_Data.json"

# Custom width and output directory
node Generate-StandaloneDiagrams.js data.json ./diagrams --width=1600

# Network diagram only
node Generate-StandaloneDiagrams.js data.json --net-only

# MG hierarchy only
node Generate-StandaloneDiagrams.js data.json --mg-only
```

## Output

The script produces two files in a timestamped subdirectory:

### 1. JSON (`CustomerName_Azure_ASIS_Data.json`)
Complete structured data export for programmatic use or feeding into document generation.

### 2. Excel (`CustomerName_Azure_ASIS_BuildDoc.xlsx`)
Multi-sheet workbook with colour-coded tabs:

| Sheet | Contents |
|-------|----------|
| Summary | Collection metadata and resource counts |
| Mgmt Groups | Management group hierarchy |
| Subscriptions | All active subscriptions |
| Conditional Access | CA policies (if -IncludeEntra) |
| Directory Roles | Entra ID role assignments (if -IncludeEntra) |
| Virtual Networks | All VNets with address spaces and DNS |
| Subnets | All subnets with NSG/UDR associations |
| VNet Peerings | Peering configuration and status |
| NSG Rules | All custom NSG rules |
| Route Tables | All UDR routes |
| VPN Gateways | Gateway configuration |
| VPN Connections | Connection status |
| Bastions | Azure Bastion instances |
| Private DNS | Private DNS zones and links |
| Public IPs | All public IP addresses |
| Virtual Machines | All VMs with specs, IPs, disk config, tags |
| VM Extensions | Installed extensions per VM |
| AVD Host Pools | Azure Virtual Desktop pools |
| AVD Session Hosts | Session host status |
| Storage Accounts | All storage with security settings |
| Recovery Vaults | Backup vault configuration |
| Backup Policies | Retention and schedule details |
| Backup Items | Protected resources |
| **Unprotected VMs** | **Running VMs with no backup** |
| Defender Plans | Microsoft Defender for Cloud status |
| Secure Scores | Security posture scores |
| Key Vaults | Key vault configuration |
| RBAC Assignments | All role assignments |
| Policy Assignments | Azure Policy assignments |
| Policy Compliance | Compliance state summary |
| Tag Audit Summary | Tagging percentage per subscription |
| Tag Key Usage | All tag keys with sample values |
| Log Analytics | Workspaces and retention |
| Alert Rules | Configured metric alerts |
| Action Groups | Alert notification targets |
| Budgets | Cost management budgets |
| Resource Groups | All RGs with resource counts |
| **GAP ANALYSIS** | **All identified gaps with severity** |
| **RECOMMENDATIONS** | **Improvement recommendations** |
| **OPERATIONAL COMPLIANCE** | **Compliance framework checklist** |

### Gap Analysis Auto-Detection

The script automatically identifies and flags:

- **Critical**: VMs without backup, NSG rules allowing RDP/SSH from internet, VPN connections not in Connected state, no Conditional Access policies, Defender for Servers not enabled
- **High**: Subnets without NSGs, missing standard Azure policies, CAF non-compliance, tagging below 50%, no budgets configured
- **Medium**: Storage accounts with public access, TLS below 1.2, expired app registration credentials, direct user RBAC assignments

### Operational Compliance Framework

The OPERATIONAL COMPLIANCE sheet covers the items you flagged:

1. **Backup governance alignment** - retention vs compliance, roles & responsibilities
2. **Monitoring & incident management** - LogicMonitor integration, alert routing, SLAs
3. **Cost management** - budgets, reservations, review cadence
4. **Security posture** - Defender status, Secure Score, remediation
5. **Conditional Access** - MFA enforcement, access controls
6. **Tagging standard** - consistency, mandatory tags, enforcement
7. **Policy governance** - applied vs missing standard policies
8. **Disaster recovery** - ASR status, RPO/RTO, testing
9. **Change management** - process, approvals, communication
10. **Naming convention** - consistency analysis

## What Still Needs Manual Review

The script cannot capture everything. These items need manual investigation:

| Item | Why Manual | Where to Check |
|------|-----------|---------------|
| LogicMonitor integration | Third-party tool, no Azure API | LogicMonitor portal |
| SMTP relay config | Application-level config | Server config / Exchange |
| On-prem firewall rules | Not accessible via Azure API | FortiGate / on-prem mgmt |
| Incident management process | Process, not technical config | Existing SOPs / agreements |
| Backup roles & responsibilities | Organisational, not technical | RACI matrix / contracts |
| Compliance requirements | Customer-specific | Customer governance docs |
| DR testing schedule | Process-based | Operational runbooks |
| ExpressRoute provider details | Contract-based | Provider portal / contracts |

## Populating the Word Document

### Option 1: Automatic (Recommended)

The pipeline handles everything:

```powershell
# Full pipeline - audit + Word doc in one go (with Entra ID)
.\Build-AzureASISDocument.ps1 -CustomerName "Sense" -IncludeEntra

# Without Entra ID (no Graph access needed)
.\Build-AzureASISDocument.ps1 -CustomerName "Sense"
```

### Option 2: Two-Step (Audit first, generate later)

```powershell
# Step 1: Collect data (with Entra ID)
.\Invoke-AzureTenancyAudit.ps1 -CustomerName "Sense" -IncludeEntra

# Step 1 (alt): Collect data without Entra ID
.\Invoke-AzureTenancyAudit.ps1 -CustomerName "Sense"

# Step 2: Generate Word doc from JSON
node Generate-BuildDocument.js ".\AzureBuildDoc_Output\20260226_143000\Sense_Azure_ASIS_Data.json"
```

### Option 3: Manual

1. Run the audit script to collect data
2. Open the Excel workbook and blank Word template side by side
3. Copy data from Excel sheets into the Word template tables

### Word Document Sections (Auto-Generated)

The Generate-BuildDocument.js script produces a fully populated Word document with:

| Section | Content |
|---------|---------|
| 1. Executive Summary | Environment-at-a-glance metrics table |
| 2. Entra ID & Identity | Tenant details, CA policies (colour-coded), directory roles, app registrations |
| 3. Management Groups & Subscriptions | Full MG hierarchy, subscription inventory |
| 4. Networking | VNets, subnets (NSG gaps flagged), peerings, NSG rules per NSG, UDRs, VPN gateways, connections, Bastion, Private DNS, public IPs |
| 5. Compute | VMs with full specs, disk config, extensions summary, AVD if deployed |
| 6. Storage | Storage accounts with security posture |
| 7. Backup & Recovery | Vaults, policies with retention, protected items, **UNPROTECTED VMs highlighted**, governance alignment checklist |
| 8. Security | Defender plans per subscription (colour-coded), Secure Scores, Key Vaults |
| 9. RBAC | Role assignment summary by role type |
| 10. Governance & Policy | Policy assignments, **missing standard policies flagged**, compliance summary |
| 11. Tagging Standard | Coverage per subscription, tag keys in use, **missing standard tags** |
| 12. Monitoring & Logging | Log Analytics, alert rules, action groups, **LogicMonitor manual checklist** |
| 13. Cost Management | Budgets (or flagged as missing), **cost management framework checklist** |
| 14. Resource Groups | Full RG inventory |
| 15. **Gap Analysis** | All gaps grouped by Critical/High/Medium with colour-coded status |
| 16. **Recommendations** | All improvement recommendations |
| 17. **Operational Compliance Framework** | 10-point checklist with status, actions, and owners |

All status columns are colour-coded: red (critical/not configured), amber (needs attention), green (configured/healthy).

## Configuration Files

### `audit-config.json` — Gap Analysis Thresholds

Customise the gap analysis thresholds and standard policies. The audit script loads this via `-ConfigPath`:

| Key | Default | Description |
|-----|---------|-------------|
| `maxGlobalAdmins` | 5 | Flag when Global Admin count exceeds this |
| `minGlobalAdmins` | 2 | Flag when Global Admin count is below this |
| `cafKeywordThreshold` | 3 | Minimum CAF keywords required in MG names |
| `lowTaggingThreshold` | 50 | Flag subscriptions with tagging below this % |
| `directUserAssignmentThreshold` | 5 | Flag when direct user RBAC assignments exceed this |
| `standardPolicies` | Array | List of Azure Policies expected to be applied |
| `expectedTags` | Array | List of standard tag keys expected on resources |

### `branding-config.json` — White-Labelling

Customise the Word document branding (colours, fonts, company details):

| Section | Keys | Description |
|---------|------|-------------|
| `colors` | primary, accent, dark, light, etc. | Hex colours (without #) |
| `page` | width, height, margin | Page dimensions in DXA units |
| `fonts` | heading, body | Font family names |
| `company` | name, website, classification | Company details for headers/footers |

## Running in Azure Cloud Shell

The toolkit works in Azure Cloud Shell with one consideration: the `sharp` module (used for PNG diagram conversion) may not be available. When `sharp` is missing, diagrams are output as SVG instead of PNG.

```bash
# In Cloud Shell, install only docx (sharp will be skipped automatically)
npm install

# Run the audit
pwsh ./Invoke-AzureTenancyAudit.ps1 -CustomerName "Customer" -IncludeEntra

# Generate diagrams (outputs SVG when sharp unavailable)
node Generate-StandaloneDiagrams.js ./AzureBuildDoc_Output/latest/Customer_Azure_ASIS_Data.json
```

## Checkpoint / Resume

If the audit script is interrupted (network error, timeout, etc.), it automatically saves checkpoint files. To resume:

```powershell
# Find the checkpoint file in your output directory
.\Invoke-AzureTenancyAudit.ps1 -CustomerName "Customer" -ResumeFrom ".\AzureBuildDoc_Output\20260226_143000\_checkpoint.json"
```

Sections already collected will be skipped, and the audit continues from where it left off. The checkpoint file is automatically deleted on successful completion.

## Troubleshooting

```powershell
# If modules aren't found
Get-InstalledModule Az*
Get-InstalledModule Microsoft.Graph*

# If permissions are insufficient
Get-AzRoleAssignment -SignInName (Get-AzContext).Account.Id

# If Graph connection fails
Connect-MgGraph -Scopes "User.Read.All" -TenantId "your-tenant-id"
```
