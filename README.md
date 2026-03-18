# AS-IS Azure Documentation Creator

Automated Azure tenancy audit toolkit that collects configuration data and generates comprehensive AS-IS Build Documents with diagrams.

## What It Does

Runs a single command against an Azure tenant and produces:

- **Word Document** (.docx) — 19-section AS-IS Build Document with colour-coded tables, gap analysis, and recommendations
- **Excel Workbook** (.xlsx) — 40+ sheet detailed data export with conditional formatting
- **JSON Data** (.json) — Complete structured data for programmatic use
- **Network Topology Diagram** — Hub-spoke VNet layout with subnet details, NSG status, and appliance badges
- **Management Group Hierarchy Diagram** — Tree view of MG structure with subscription placement
- **ZIP Package** — All outputs bundled for easy download

Everything is packaged into a single ZIP file at the end.

---

## Quick Start — Azure Cloud Shell

### 1. Clone and Setup (one time)

```powershell
git clone https://github.com/thompsca06/AS-IS-Azure-Documentation-Creator.git
cd AS-IS-Azure-Documentation-Creator
.\Setup-CloudShell.ps1
```

This clones the repo, installs Node.js packages, checks PowerShell modules, and verifies your Azure connection.

### 2. Run the Audit

```powershell
# Azure resources only (recommended first run)
.\Build-AzureASISDocument.ps1 -CustomerName "Contoso"

# Include Entra ID (requires Graph API permissions — will prompt for device login)
.\Build-AzureASISDocument.ps1 -CustomerName "Contoso" -IncludeEntra

# Scope to specific subscriptions
.\Build-AzureASISDocument.ps1 -CustomerName "Contoso" -SubscriptionFilter @("sub-id-1","sub-id-2")
```

### 3. Download the Output

When complete, you'll see:

```
OUTPUT FILES:
  JSON:   ./AzureBuildDoc_Output/20260318_143000/Contoso_Azure_ASIS_Data.json
  Excel:  ./AzureBuildDoc_Output/20260318_143000/Contoso_Azure_ASIS_BuildDoc.xlsx
  Word:   ./AzureBuildDoc_Output/20260318_143000/Contoso_Azure_ASIS_Build_Document.docx
  ZIP:    ./AzureBuildDoc_Output/Contoso_Azure_ASIS_20260318.zip
```

In Cloud Shell, click the **Upload/Download** button in the toolbar, select **Download**, and enter the ZIP path.

---

## Quick Start — Local Machine (Windows)

### 1. Prerequisites

```powershell
# Install required modules
Install-Module Az -Scope CurrentUser -Force
Install-Module ImportExcel -Scope CurrentUser -Force

# Optional: Entra ID modules (for -IncludeEntra)
Install-Module Microsoft.Graph.Identity.DirectoryManagement -Scope CurrentUser -Force
Install-Module Microsoft.Graph.Identity.SignIns -Scope CurrentUser -Force
Install-Module Microsoft.Graph.Applications -Scope CurrentUser -Force

# Node.js (download from https://nodejs.org)
# Then install npm packages:
npm install
```

### 2. Connect and Run

```powershell
Connect-AzAccount
.\Build-AzureASISDocument.ps1 -CustomerName "Contoso" -IncludeEntra
```

---

## Files

| File | Purpose |
|------|---------|
| `Build-AzureASISDocument.ps1` | **One-click pipeline** — audit + Word doc + diagrams + ZIP |
| `Invoke-AzureTenancyAudit.ps1` | Data collection script — all Azure resources → JSON + Excel |
| `Generate-BuildDocument.js` | Word document generator — reads JSON → 19-section .docx |
| `Generate-Diagrams.js` | Diagram engine — MG hierarchy + network topology (SVG/PNG) |
| `Generate-Diagrams.ps1` | **PowerShell wrapper** for standalone diagram generation |
| `Generate-StandaloneDiagrams.js` | Node.js CLI for standalone diagram generation |
| `Setup-CloudShell.ps1` | **Cloud Shell setup** — clone, install deps, verify environment |
| `audit-config.json` | Configurable gap analysis thresholds |
| `branding-config.json` | White-labelling (colours, fonts, company details) |
| `Azure_AS-IS_Build_Document_Template.docx` | Blank Word template (manual use) |

---

## Pipeline Steps

The `Build-AzureASISDocument.ps1` pipeline runs 4 steps:

| Step | What It Does |
|------|-------------|
| **1. Audit** | Runs `Invoke-AzureTenancyAudit.ps1` — collects all Azure resource data across every subscription, exports JSON + Excel, runs gap analysis |
| **2. Word Doc** | Runs `Generate-BuildDocument.js` — reads the JSON and produces a fully populated, branded Word document |
| **3. Diagrams** | Runs `Generate-StandaloneDiagrams.js` — creates MG hierarchy and network topology diagrams |
| **4. ZIP** | Packages all output files into a single ZIP for download |

Dependencies (ImportExcel, docx npm module) are **auto-installed** if missing.

---

## Usage Examples

### Generate Diagrams Only

If you've already run the audit and just want to regenerate diagrams:

```powershell
# Auto-detects latest JSON from output folder
.\Generate-Diagrams.ps1

# Network diagram only, wider
.\Generate-Diagrams.ps1 -NetOnly -Width 1600

# MG hierarchy only
.\Generate-Diagrams.ps1 -MgOnly

# SVG format (useful in Cloud Shell without sharp)
.\Generate-Diagrams.ps1 -Format svg

# Explicit JSON path
.\Generate-Diagrams.ps1 -JsonPath ".\output\data.json" -OutputDir ".\diagrams"
```

### Two-Step Workflow (Audit First, Generate Later)

```powershell
# Step 1: Collect data
.\Invoke-AzureTenancyAudit.ps1 -CustomerName "Contoso"

# Step 2: Generate Word doc from JSON
node Generate-BuildDocument.js ".\AzureBuildDoc_Output\20260318_143000\Contoso_Azure_ASIS_Data.json"
```

### Resume an Interrupted Audit

The audit saves checkpoints after each section. If interrupted:

```powershell
.\Invoke-AzureTenancyAudit.ps1 -CustomerName "Contoso" -ResumeFrom ".\AzureBuildDoc_Output\20260318_143000\_checkpoint.json"
```

### Custom Gap Analysis Thresholds

```powershell
.\Invoke-AzureTenancyAudit.ps1 -CustomerName "Contoso" -ConfigPath ".\audit-config.json"
```

---

## Permissions Required

| Scope | Permission | Purpose | Required? |
|-------|-----------|---------|-----------|
| Azure RBAC | **Reader** on target subscriptions | Read all resource configuration | Yes |
| Azure RBAC | **Management Group Reader** | Read MG hierarchy | Yes |
| Entra ID | **Global Reader** (or equivalent) | Read tenant config | Only with `-IncludeEntra` |
| Graph API | User.Read.All, Policy.Read.All | CA policies, users/groups | Only with `-IncludeEntra` |
| Graph API | Application.Read.All | App registrations | Only with `-IncludeEntra` |
| Graph API | RoleManagement.Read.Directory | Directory roles | Only with `-IncludeEntra` |
| Graph API | Directory.Read.All | Tenant details | Only with `-IncludeEntra` |

> **Cloud Shell Note:** Azure resources are accessible via the Cloud Shell managed identity. Entra ID (Graph API) requires a separate device login — the script will prompt you with a URL and code.

---

## Output — Word Document Sections

The generated Word document contains 19 sections:

| # | Section | Content |
|---|---------|---------|
| 1 | Executive Summary | Environment-at-a-glance metrics |
| 2 | Entra ID & Identity | Tenant details, CA policies, directory roles, app registrations, groups, licenses |
| 3 | Management Groups & Subscriptions | MG hierarchy with diagram, subscription inventory |
| 4 | Networking | VNets, subnets, peerings, NSG rules, route tables, VPN, Bastion, DNS, public IPs, firewalls, AppGW, LB, NAT GW, ExpressRoute, DDoS, private endpoints, NSG flow logs — with topology diagram |
| 5 | Compute | VMs, disk config, extensions, AVD, managed disks, snapshots, VMSS |
| 6 | Storage | Storage accounts with security posture |
| 7 | Backup & Recovery | Vaults, policies, protected items, **unprotected VMs**, ASR |
| 8 | Security | Defender plans, Secure Scores, Key Vaults, Defender recommendations |
| 9 | RBAC | Role assignment summary by type |
| 10 | Governance & Policy | Policy assignments, **missing standard policies**, compliance, locks, custom roles, exemptions |
| 11 | Tagging | Coverage per subscription, tag keys, **missing standard tags** |
| 12 | Monitoring & Logging | Log Analytics, alerts, action groups, App Insights, diagnostics, query rules |
| 13 | Cost Management | Budgets or flagged as missing |
| 14 | Resource Groups | Full inventory |
| 15 | Application Services | App Service Plans, Web Apps, Function Apps |
| 16 | Databases | Azure SQL, Cosmos DB |
| 17 | Containers | AKS, Container Registries |
| 18 | Automation & Advisor | Automation Accounts, Runbooks, Advisor recommendations |
| 19 | Resource Summary | Resource Graph type counts |

Plus: **Gap Analysis**, **Recommendations**, and **Operational Compliance Framework** sections.

All status columns are colour-coded: red (critical), amber (warning), green (healthy).

---

## Output — Excel Workbook

40+ sheets including all resource types, plus:

- **GAP ANALYSIS** — All identified gaps with severity (colour-coded)
- **RECOMMENDATIONS** — 30+ improvement recommendations
- **OPERATIONAL COMPLIANCE** — 10-point compliance framework checklist

---

## Output — Diagrams

### Network Topology
- Hub-spoke layout (auto-detected from peering data)
- Each VNet shows: address space, DNS servers, all subnets with CIDR
- Subnet details: NSG name (red warning if missing), service endpoints, delegations
- Appliance badges: Bastion, Firewall, VPN GW, AppGW, Load Balancer
- Peering lines, VPN connections, ExpressRoute circuits
- Connection dots at endpoints

### Management Group Hierarchy
- Tree layout from tenant root
- Shows subscriptions under each MG
- Collapses when >4 subscriptions per MG

---

## Configuration

### `audit-config.json` — Gap Analysis Thresholds

| Key | Default | Description |
|-----|---------|-------------|
| `maxGlobalAdmins` | 5 | Flag when Global Admin count exceeds this |
| `minGlobalAdmins` | 2 | Flag when count is below this |
| `cafKeywordThreshold` | 3 | Min CAF keywords in MG names |
| `lowTaggingThreshold` | 50 | Flag subscriptions below this % tagged |
| `directUserAssignmentThreshold` | 5 | Flag when direct user RBAC assignments exceed this |
| `standardPolicies` | Array | Azure Policies expected to be applied |
| `expectedTags` | Array | Tag keys expected on resources |

### `branding-config.json` — White-Labelling

| Section | Description |
|---------|-------------|
| `colors` | Hex colours for headings, tables, status indicators |
| `page` | Page dimensions (DXA units) |
| `fonts` | Heading and body font families |
| `company` | Company name, website, document classification |

---

## What Still Needs Manual Review

| Item | Why | Where to Check |
|------|-----|---------------|
| LogicMonitor integration | Third-party, no Azure API | LogicMonitor portal |
| SMTP relay config | Application-level | Server config / Exchange |
| On-prem firewall rules | Not Azure API accessible | FortiGate / on-prem mgmt |
| Incident management process | Organisational process | SOPs / agreements |
| Backup roles & responsibilities | Organisational | RACI matrix / contracts |
| Compliance requirements | Customer-specific | Governance docs |
| DR testing schedule | Process-based | Operational runbooks |
| ExpressRoute provider details | Contract-based | Provider portal |

---

## Troubleshooting

### Cloud Shell: "Output path is not writable"
```powershell
git pull   # get latest fix
```

### Modules not found
```powershell
Get-InstalledModule Az*
Get-InstalledModule Microsoft.Graph*
Get-InstalledModule ImportExcel
```

### Entra ID: Device login prompt
This is expected in Cloud Shell — Graph API requires separate authentication. Either:
- Complete the device login (open URL, enter code)
- Or skip Entra: remove `-IncludeEntra` flag

### Graph connection fails
```powershell
Connect-MgGraph -Scopes "User.Read.All","Policy.Read.All" -TenantId "your-tenant-id"
```

### Diagrams: sharp not available
Normal in Cloud Shell. Diagrams output as SVG instead of PNG. Use `-Format svg` explicitly or the fallback is automatic.

### Update to latest version
```powershell
cd ~/AS-IS-Azure-Documentation-Creator
git pull
.\Setup-CloudShell.ps1 -SkipClone
```
