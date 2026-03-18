#Requires -Modules Az.Accounts, Az.Resources, Az.Network, Az.Compute, Az.Storage, Az.RecoveryServices, Az.Monitor, Az.Security, Az.PolicyInsights, Az.KeyVault
<#
.SYNOPSIS
    Azure Tenancy AS-IS Build Document - Data Collection Script
    
.DESCRIPTION
    Interrogates ALL aspects of an Azure tenancy and exports structured JSON + Excel
    that can be used to auto-populate the AS-IS Build Document template.
    
    Covers: Management Groups, Subscriptions, Identity (Entra ID via Graph),
    Networking, Compute, Storage, Backup, Security (Defender), Governance (Policy),
    Monitoring, Cost Management, Tags, Naming, RBAC, and Operational config.

.NOTES
    Author:  Tieva Ltd
    Version: 1.0
    Date:    February 2026
    
    PREREQUISITES:
    - Az PowerShell modules: Install-Module Az -Scope CurrentUser
    - ImportExcel module:    Install-Module ImportExcel -Scope CurrentUser
    - Microsoft.Graph modules (for Entra ID):
        Install-Module Microsoft.Graph.Identity.DirectoryManagement -Scope CurrentUser
        Install-Module Microsoft.Graph.Identity.SignIns -Scope CurrentUser
        Install-Module Microsoft.Graph.Applications -Scope CurrentUser
    - Permissions: Reader on all subscriptions, Global Reader or equiv in Entra ID
    - For Graph API: User.Read.All, Policy.Read.All, Application.Read.All, 
                     RoleManagement.Read.Directory, AuditLog.Read.All

.PARAMETER OutputPath
    Directory to save output files. Defaults to .\AzureBuildDoc_Output

.PARAMETER IncludeEntra
    Switch to include Entra ID (Azure AD) data collection via Microsoft Graph.
    Requires Graph modules and appropriate permissions.
    When omitted, the Entra ID section is marked as 'Not in scope' and no gap is logged.
    Alias: -IncludeGraph (backward compatible).

.PARAMETER SubscriptionFilter
    Optional array of subscription IDs to scope the collection.
    If omitted, all accessible subscriptions are interrogated.

.PARAMETER CustomerName
    Customer name for the report header.

.EXAMPLE
    .\Invoke-AzureTenancyAudit.ps1 -CustomerName "Contoso" -IncludeEntra
    .\Invoke-AzureTenancyAudit.ps1 -CustomerName "Sense" -SubscriptionFilter @("sub-id-1","sub-id-2")
#>

[CmdletBinding()]
param(
    [Parameter(Mandatory = $true)]
    [string]$CustomerName,
    
    [string]$OutputPath = ".\AzureBuildDoc_Output",
    
    [Alias("IncludeGraph")]
    [switch]$IncludeEntra,
    
    [string[]]$SubscriptionFilter,

    [string]$ConfigPath,

    [string]$ResumeFrom
)

# ============================================================
# VALIDATE CUSTOMERNAME FOR FILESYSTEM SAFETY
# ============================================================
$invalidFileChars = [IO.Path]::GetInvalidFileNameChars()
if ($CustomerName.IndexOfAny($invalidFileChars) -ge 0) {
    Write-Error "CustomerName '$CustomerName' contains characters that are invalid in file names. Remove characters: / \ : * ? `" < > |"
    exit 1
}
# Excel sheet names have additional restrictions (max 31 chars, no [ ] : * ? / \)
$excelInvalid = '[', ']', ':', '*', '?', '/', '\'
foreach ($ch in $excelInvalid) {
    if ($CustomerName.Contains($ch)) {
        Write-Error "CustomerName '$CustomerName' contains '$ch' which is invalid in Excel sheet names."
        exit 1
    }
}
if ($CustomerName.Length -gt 31) {
    Write-Warning "CustomerName is longer than 31 characters. It will be truncated in Excel sheet names."
}

# ============================================================
# LOAD AUDIT CONFIGURATION (thresholds, standard policies, etc.)
# ============================================================
$auditConfig = @{
    maxGlobalAdmins                = 5
    minGlobalAdmins                = 2
    cafKeywordThreshold            = 3
    lowTaggingThreshold            = 50
    directUserAssignmentThreshold  = 5
    standardPolicies               = @(
        "Network interfaces should not have public IPs",
        "Azure Backup should be enabled for Virtual Machines",
        "Subnets should have a Network Security Group",
        "Allowed locations",
        "Audit VMs that do not use managed disks",
        "Key vaults should have soft delete enabled"
    )
    expectedTags                   = @("Environment", "Owner", "CostCenter", "Application", "Criticality")
}

if ($ConfigPath -and (Test-Path $ConfigPath)) {
    try {
        $userConfig = Get-Content $ConfigPath -Raw | ConvertFrom-Json
        foreach ($prop in $userConfig.PSObject.Properties) {
            $auditConfig[$prop.Name] = $prop.Value
        }
        Write-Host "[OK] Loaded audit config from: $ConfigPath" -ForegroundColor Green
    }
    catch {
        Write-Warning "Failed to load config from '$ConfigPath': $($_.Exception.Message). Using defaults."
    }
}

# ============================================================
# CHECKPOINT / RESUME FUNCTIONS
# ============================================================
function Save-Checkpoint {
    try {
        $checkpointPath = Join-Path $OutputPath "_checkpoint.json"
        $tenancyData | ConvertTo-Json -Depth 10 -Compress:$false | Out-File -FilePath $checkpointPath -Encoding UTF8
    }
    catch {
        Write-Host "[WARN] Failed to save checkpoint: $($_.Exception.Message)" -ForegroundColor Yellow
    }
}

function Test-SectionLoaded {
    param([string]$SectionName)
    return ($tenancyData.Sections.Contains($SectionName) -and $null -ne $tenancyData.Sections[$SectionName])
}

# ============================================================
# INITIALISATION
# ============================================================
$ErrorActionPreference = "Continue"
$WarningPreference = "SilentlyContinue"
$ProgressPreference = "SilentlyContinue"
$PSDefaultParameterValues['Invoke-WebRequest:UseBasicParsing'] = $true

$timestamp = Get-Date -Format "yyyyMMdd_HHmmss"
$OutputPath = Join-Path $OutputPath $timestamp
New-Item -ItemType Directory -Path $OutputPath -Force | Out-Null

# Validate output path is writable
try {
    $testFile = Join-Path $OutputPath ".write_test_$(Get-Random)"
    "test" | Out-File -FilePath $testFile -Force
    Remove-Item $testFile -Force
}
catch {
    Write-Error "Output path '$OutputPath' is not writable: $($_.Exception.Message)"
    exit 1
}

# Master data collection object
$tenancyData = [ordered]@{
    CollectionDate     = (Get-Date -Format "dd/MM/yyyy HH:mm:ss")
    CustomerName       = $CustomerName
    CollectedBy        = (Get-AzContext).Account.Id
    TenantId           = (Get-AzContext).Tenant.Id
    EntraIncluded      = [bool]$IncludeEntra   # Whether Entra ID (Graph) was in scope
    EntraStatus        = if ($IncludeEntra) { "Pending" } else { "Not in scope" }
    Sections           = [ordered]@{}
    Gaps               = [System.Collections.ArrayList]::new()   # Track gaps/issues
    Recommendations    = [System.Collections.ArrayList]::new()   # Track recommendations
}

# Load checkpoint if resuming
if ($ResumeFrom -and (Test-Path $ResumeFrom)) {
    try {
        $resumeData = Get-Content $ResumeFrom -Raw | ConvertFrom-Json
        foreach ($prop in $resumeData.Sections.PSObject.Properties) {
            $tenancyData.Sections[$prop.Name] = $prop.Value
        }
        if ($resumeData.Gaps) {
            foreach ($g in $resumeData.Gaps) { $null = $tenancyData.Gaps.Add($g) }
        }
        if ($resumeData.Recommendations) {
            foreach ($r in $resumeData.Recommendations) { $null = $tenancyData.Recommendations.Add($r) }
        }
        Write-Host "[OK] Resumed from checkpoint with $($tenancyData.Sections.Count) sections loaded" -ForegroundColor Green
    }
    catch {
        Write-Warning "Failed to load checkpoint from '$ResumeFrom': $($_.Exception.Message). Starting fresh."
    }
}

function Write-Status {
    param([string]$Section, [string]$Message, [string]$Level = "Info")
    $colour = switch ($Level) {
        "Info"    { "Cyan" }
        "Warning" { "Yellow" }
        "Error"   { "Red" }
        "Success" { "Green" }
        default   { "White" }
    }
    Write-Host "[$Level] [$Section] $Message" -ForegroundColor $colour
}

function Add-Gap {
    param([string]$Section, [string]$Description, [string]$Severity = "Medium")
    $null = $tenancyData.Gaps.Add([PSCustomObject]@{
        Section     = $Section
        Description = $Description
        Severity    = $Severity
        Status      = "Requires Review"
    }))
}

function Add-Recommendation {
    param([string]$Section, [string]$Description, [string]$Priority = "Medium")
    $null = $tenancyData.Recommendations.Add([PSCustomObject]@{
        Section     = $Section
        Description = $Description
        Priority    = $Priority
    }))
}

# ============================================================
# PRE-FLIGHT CHECKS
# ============================================================
Write-Status "Init" "Starting Azure Tenancy Audit for: $CustomerName"
Write-Status "Init" "Output directory: $OutputPath"

$context = Get-AzContext
if (-not $context) {
    Write-Status "Init" "Not logged in. Running Connect-AzAccount..." "Warning"
    Connect-AzAccount
    $context = Get-AzContext
}
Write-Status "Init" "Connected as: $($context.Account.Id)" "Success"
Write-Status "Init" "Tenant: $($context.Tenant.Id)"

# ============================================================
# SECTION 1: MANAGEMENT GROUPS
# ============================================================
if (Test-SectionLoaded "ManagementGroups") {
    Write-Status "MgmtGroups" "Loaded from checkpoint - skipping" "Info"
} else {

Write-Status "MgmtGroups" "Collecting management group hierarchy..."

try {
    $allMGs = Get-AzManagementGroup -ErrorAction Stop
    $mgHierarchy = [System.Collections.Generic.List[object]]::new()
    
    foreach ($mg in $allMGs) {
        $mgDetail = Get-AzManagementGroup -GroupName $mg.Name -Expand -ErrorAction SilentlyContinue
        
        $mgObj = [PSCustomObject]@{
            DisplayName      = $mg.DisplayName
            Name             = $mg.Name
            Id               = $mg.Id
            ParentId         = if ($mgDetail.ParentId) { $mgDetail.ParentId } else { "Root" }
            ParentName       = if ($mgDetail.ParentDisplayName) { $mgDetail.ParentDisplayName } else { "Tenant Root" }
            ChildCount       = if ($mgDetail.Children) { $mgDetail.Children.Count } else { 0 }
            Subscriptions    = @()
            ChildMGs         = @()
        }
        
        if ($mgDetail.Children) {
            foreach ($child in $mgDetail.Children) {
                if ($child.Type -eq "/subscriptions") {
                    $mgObj.Subscriptions += $child.DisplayName
                }
                elseif ($child.Type -match "managementGroups") {
                    $mgObj.ChildMGs += $child.DisplayName
                }
            }
        }
        $null = $mgHierarchy.Add($mgObj)
    }
    
    $tenancyData.Sections["ManagementGroups"] = $mgHierarchy
Save-Checkpoint
}
    Write-Status "MgmtGroups" "Found $($mgHierarchy.Count) management groups" "Success"
    
    # Gap check: Is there a proper CAF structure?
    $cafKeywords = @("platform", "identity", "connectivity", "management", "landing", "corp", "online", "sandbox")
    $mgNames = ($mgHierarchy | ForEach-Object { $_.DisplayName.ToLower() }) -join " "
    $cafMatch = $cafKeywords | Where-Object { $mgNames -match $_ }
    if ($cafMatch.Count -lt $auditConfig.cafKeywordThreshold) {
        Add-Gap "ManagementGroups" "Management group hierarchy does not appear to follow CAF Landing Zone structure. Only $($cafMatch.Count)/8 CAF keywords found." "High"
    }
}
catch {
    Write-Status "MgmtGroups" "Failed to collect management groups: $($_.Exception.Message)" "Error"
    Add-Gap "ManagementGroups" "Unable to collect management group data - check permissions (Management Group Reader required)" "High"
}

# ============================================================
# SECTION 2: SUBSCRIPTIONS
# ============================================================
Write-Status "Subscriptions" "Collecting subscription details..."

$allSubs = Get-AzSubscription -TenantId $context.Tenant.Id | Where-Object { $_.State -eq "Enabled" }

if ($SubscriptionFilter) {
    # Explicit filter provided via parameter
    $allSubs = $allSubs | Where-Object { $_.Id -in $SubscriptionFilter }
}
elseif ($allSubs.Count -gt 1) {
    # Interactive selection when multiple subscriptions exist
    Write-Status "Subscriptions" "Found $($allSubs.Count) active subscriptions. Select which to audit..." "Info"

    $selectedSubs = $null
    try {
        # Try Out-GridView (works in Windows PowerShell and PS7 with module)
        $selectedSubs = $allSubs |
            Select-Object Name, Id, State |
            Out-GridView -Title "Select subscriptions to audit (hold Ctrl to multi-select)" -PassThru
    }
    catch {
        # Fallback: numbered console list
        Write-Status "Subscriptions" "GridView not available, using console selection..." "Warning"
    }

    if (-not $selectedSubs) {
        # Console fallback
        Write-Host ""
        Write-Host "  Available Subscriptions:" -ForegroundColor Cyan
        for ($i = 0; $i -lt $allSubs.Count; $i++) {
            Write-Host "    [$($i + 1)] $($allSubs[$i].Name)  ($($allSubs[$i].Id))"
        }
        Write-Host ""
        $selection = Read-Host "  Enter subscription numbers to audit (comma-separated, e.g. 1,3,5) or 'all'"

        if ($selection.Trim().ToLower() -ne "all") {
            $indices = $selection -split "," | ForEach-Object { [int]$_.Trim() - 1 }
            $selectedSubs = $indices | ForEach-Object { $allSubs[$_] | Select-Object Name, Id, State }
        }
    }

    if ($selectedSubs) {
        $allSubs = $allSubs | Where-Object { $_.Id -in $selectedSubs.Id }
        Write-Status "Subscriptions" "Selected $($allSubs.Count) subscription(s) for audit" "Success"
    }
    else {
        Write-Status "Subscriptions" "No selection made - auditing all $($allSubs.Count) subscriptions" "Warning"
    }
}

$subData = foreach ($sub in $allSubs) {
    [PSCustomObject]@{
        SubscriptionName = $sub.Name
        SubscriptionId   = $sub.Id
        State            = $sub.State
        TenantId         = $sub.TenantId
        OfferType        = $sub.SubscriptionPolicies.QuotaId
        SpendingLimit    = $sub.SubscriptionPolicies.SpendingLimit
    }
}
$tenancyData.Sections["Subscriptions"] = $subData
Save-Checkpoint
Write-Status "Subscriptions" "Found $($allSubs.Count) active subscriptions" "Success"

# ============================================================
# SECTION 3: ENTRA ID / IDENTITY (via Graph)
# ============================================================
if ($IncludeEntra) {
    Write-Status "Identity" "Collecting Entra ID configuration via Microsoft Graph..."
    
    try {
        # Connect to Graph - will prompt for auth if needed
        Connect-MgGraph -Scopes "User.Read.All","Policy.Read.All","Application.Read.All","RoleManagement.Read.Directory","Directory.Read.All" -NoWelcome -ErrorAction Stop
        
        $graphContext = Get-MgContext
        Write-Status "Identity" "Connected to Graph as: $($graphContext.Account)" "Success"
        
        # Tenant details
        $org = Get-MgOrganization
        $domains = Get-MgDomain
        
        $tenantInfo = [PSCustomObject]@{
            TenantName          = $org.DisplayName
            TenantId            = $org.Id
            PrimaryDomain       = ($domains | Where-Object { $_.IsDefault }).Id
            VerifiedDomains     = ($domains | Where-Object { $_.IsVerified } | ForEach-Object { $_.Id }) -join ", "
            AllDomains          = ($domains | ForEach-Object { "$($_.Id) (Verified: $($_.IsVerified))" }) -join "; "
            CreatedDateTime     = $org.CreatedDateTime
        }
        
        # Conditional Access Policies
        Write-Status "Identity" "Collecting Conditional Access policies..."
        $caPolicies = [System.Collections.Generic.List[object]]::new()
        try {
            $rawPolicies = Get-MgIdentityConditionalAccessPolicy -All
            foreach ($policy in $rawPolicies) {
                $null = $caPolicies.Add([PSCustomObject]@{
                    DisplayName         = $policy.DisplayName
                    State               = $policy.State
                    CreatedDateTime     = $policy.CreatedDateTime
                    ModifiedDateTime    = $policy.ModifiedDateTime
                    IncludeUsers        = ($policy.Conditions.Users.IncludeUsers) -join ", "
                    IncludeGroups       = ($policy.Conditions.Users.IncludeGroups) -join ", "
                    ExcludeUsers        = ($policy.Conditions.Users.ExcludeUsers) -join ", "
                    ExcludeGroups       = ($policy.Conditions.Users.ExcludeGroups) -join ", "
                    IncludeApps         = ($policy.Conditions.Applications.IncludeApplications) -join ", "
                    ExcludeApps         = ($policy.Conditions.Applications.ExcludeApplications) -join ", "
                    ClientAppTypes      = ($policy.Conditions.ClientAppTypes) -join ", "
                    Platforms           = ($policy.Conditions.Platforms.IncludePlatforms) -join ", "
                    Locations           = ($policy.Conditions.Locations.IncludeLocations) -join ", "
                    GrantControls       = ($policy.GrantControls.BuiltInControls) -join ", "
                    GrantOperator       = $policy.GrantControls.Operator
                    SessionControls     = if ($policy.SessionControls) { "Configured" } else { "None" }
                })
            }
            
            # Gap checks for CA
            $enabledPolicies = $caPolicies | Where-Object { $_.State -eq "enabled" }
            $mfaPolicies = $caPolicies | Where-Object { $_.GrantControls -match "mfa" -and $_.State -eq "enabled" }
            
            if ($enabledPolicies.Count -eq 0) {
                Add-Gap "Identity-ConditionalAccess" "No Conditional Access policies are enabled. MFA and access controls are not enforced." "Critical"
            }
            if ($mfaPolicies.Count -eq 0) {
                Add-Gap "Identity-ConditionalAccess" "No Conditional Access policy enforces MFA. This is a critical security gap." "Critical"
            }
            
            # Check for common baseline policies
            $allUsersCA = $caPolicies | Where-Object { $_.IncludeUsers -match "All" -and $_.State -eq "enabled" }
            if ($allUsersCA.Count -eq 0) {
                Add-Gap "Identity-ConditionalAccess" "No CA policy targets 'All Users'. Consider baseline policies for all users." "High"
            }
        }
        catch {
            Write-Status "Identity" "Cannot collect CA policies: $($_.Exception.Message)" "Warning"
            Add-Gap "Identity-ConditionalAccess" "Unable to collect Conditional Access policies - requires Policy.Read.All permission" "High"
        }
        
        # App Registrations & Enterprise Apps
        Write-Status "Identity" "Collecting App Registrations..."
        $appRegs = [System.Collections.Generic.List[object]]::new()
        try {
            $rawApps = Get-MgApplication -All -Property DisplayName,AppId,SignInAudience,CreatedDateTime,PasswordCredentials,KeyCredentials
            foreach ($app in $rawApps) {
                $expiredCreds = @()
                foreach ($pwd in $app.PasswordCredentials) {
                    if ($pwd.EndDateTime -lt (Get-Date)) { $expiredCreds += "Secret: $($pwd.DisplayName)" }
                }
                foreach ($key in $app.KeyCredentials) {
                    if ($key.EndDateTime -lt (Get-Date)) { $expiredCreds += "Cert: $($key.DisplayName)" }
                }
                
                $null = $appRegs.Add([PSCustomObject]@{
                    DisplayName     = $app.DisplayName
                    AppId           = $app.AppId
                    SignInAudience  = $app.SignInAudience
                    CreatedDateTime = $app.CreatedDateTime
                    SecretCount     = $app.PasswordCredentials.Count
                    CertCount       = $app.KeyCredentials.Count
                    ExpiredCreds    = if ($expiredCreds.Count -gt 0) { $expiredCreds -join "; " } else { "None" }
                })
            }
            
            $expiredApps = $appRegs | Where-Object { $_.ExpiredCreds -ne "None" }
            if ($expiredApps.Count -gt 0) {
                Add-Gap "Identity-AppRegistrations" "$($expiredApps.Count) app registrations have expired credentials. Review and rotate." "Medium"
            }
        }
        catch {
            Write-Status "Identity" "Cannot collect app registrations: $($_.Exception.Message)" "Warning"
        }

        # Enterprise Apps (Service Principals)
        Write-Status "Identity" "Collecting Enterprise Applications..."
        $enterpriseApps = [System.Collections.Generic.List[object]]::new()
        try {
            $rawSPs = Get-MgServicePrincipal -All -Property DisplayName,AppId,ServicePrincipalType,AccountEnabled,PreferredSingleSignOnMode
            foreach ($sp in $rawSPs | Where-Object { $_.ServicePrincipalType -eq "Application" }) {
                $null = $enterpriseApps.Add([PSCustomObject]@{
                    DisplayName     = $sp.DisplayName
                    AppId           = $sp.AppId
                    Type            = $sp.ServicePrincipalType
                    Enabled         = $sp.AccountEnabled
                    SSOMode         = if ($sp.PreferredSingleSignOnMode) { $sp.PreferredSingleSignOnMode } else { "None" }
                })
            }
        }
        catch {
            Write-Status "Identity" "Cannot collect enterprise apps: $($_.Exception.Message)" "Warning"
        }

        # Directory Roles
        Write-Status "Identity" "Collecting directory role assignments..."
        $directoryRoles = [System.Collections.Generic.List[object]]::new()
        try {
            $roles = Get-MgDirectoryRole -All
            foreach ($role in $roles) {
                $members = Get-MgDirectoryRoleMember -DirectoryRoleId $role.Id
                foreach ($member in $members) {
                    $null = $directoryRoles.Add([PSCustomObject]@{
                        RoleName        = $role.DisplayName
                        MemberName      = $member.AdditionalProperties.displayName
                        MemberType      = $member.AdditionalProperties.'@odata.type'
                        MemberId        = $member.Id
                    })
                }
            }
            
            # Check for Global Admins count
            $globalAdmins = $directoryRoles | Where-Object { $_.RoleName -eq "Global Administrator" }
            if ($globalAdmins.Count -gt $auditConfig.maxGlobalAdmins) {
                Add-Gap "Identity-Roles" "There are $($globalAdmins.Count) Global Administrators. Best practice is 2-4 maximum." "High"
            }
            if ($globalAdmins.Count -lt $auditConfig.minGlobalAdmins) {
                Add-Gap "Identity-Roles" "Only 1 Global Administrator. Recommend at least 2 for redundancy." "Medium"
            }
        }
        catch {
            Write-Status "Identity" "Cannot collect directory roles: $($_.Exception.Message)" "Warning"
        }

        # Named Locations
        $namedLocations = [System.Collections.Generic.List[object]]::new()
        try {
            $nls = Get-MgIdentityConditionalAccessNamedLocation -All -ErrorAction SilentlyContinue
            foreach ($nl in $nls) {
                $null = $namedLocations.Add([PSCustomObject]@{
                    Name         = $nl.DisplayName
                    Type         = if ($nl.AdditionalProperties.'@odata.type' -match "ipNamed") { "IP Ranges" } else { "Countries" }
                    IsTrusted    = if ($nl.AdditionalProperties.isTrusted) { $true } else { $false }
                    Details      = if ($nl.AdditionalProperties.ipRanges) { ($nl.AdditionalProperties.ipRanges | ForEach-Object { $_.cidrAddress }) -join "; " } elseif ($nl.AdditionalProperties.countriesAndRegions) { ($nl.AdditionalProperties.countriesAndRegions) -join ", " } else { "" }
                })
            }
        } catch { }

        # Groups Summary
        $groupSummary = [System.Collections.Generic.List[object]]::new()
        try {
            $groups = Get-MgGroup -All -Property DisplayName, GroupTypes, SecurityEnabled, MailEnabled, MembershipRule -ErrorAction SilentlyContinue
            $securityGroups = ($groups | Where-Object { $_.SecurityEnabled -and -not ($_.GroupTypes -contains "Unified") }).Count
            $m365Groups = ($groups | Where-Object { $_.GroupTypes -contains "Unified" }).Count
            $dynamicGroups = ($groups | Where-Object { $_.GroupTypes -contains "DynamicMembership" }).Count
            $groupSummary = [PSCustomObject]@{
                TotalGroups     = $groups.Count
                SecurityGroups  = $securityGroups
                M365Groups      = $m365Groups
                DynamicGroups   = $dynamicGroups
            }
        } catch { }

        # License Summary
        $licenseSummary = [System.Collections.Generic.List[object]]::new()
        try {
            $skus = Get-MgSubscribedSku -All -ErrorAction SilentlyContinue
            foreach ($sku in $skus) {
                $null = $licenseSummary.Add([PSCustomObject]@{
                    SKUName         = $sku.SkuPartNumber
                    Total           = $sku.PrepaidUnits.Enabled
                    Consumed        = $sku.ConsumedUnits
                    Available       = $sku.PrepaidUnits.Enabled - $sku.ConsumedUnits
                })
            }
        } catch { }

        $tenancyData.Sections["Identity"] = [ordered]@{
            TenantInfo          = $tenantInfo
            ConditionalAccess   = $caPolicies
            AppRegistrations    = $appRegs
            EnterpriseApps      = $enterpriseApps
            DirectoryRoles      = $directoryRoles
            NamedLocations      = $namedLocations
            GroupSummary        = $groupSummary
            Licenses            = $licenseSummary
        }
        
        $tenancyData.EntraStatus = "Collected"
        Write-Status "Identity" "Entra ID collection complete. CA Policies: $($caPolicies.Count), Apps: $($appRegs.Count)" "Success"

        Disconnect-MgGraph -ErrorAction SilentlyContinue
    }
    catch {
        $tenancyData.EntraStatus = "Failed"
        Write-Status "Identity" "Graph connection failed: $($_.Exception.Message)" "Error"
        Add-Gap "Identity" "Unable to connect to Microsoft Graph. Entra ID data not collected. Re-run with -IncludeEntra to retry." "High"
        # Save partial identity data even on failure
        if (-not $tenancyData.Sections.Contains("Identity")) {
            $tenancyData.Sections["Identity"] = [ordered]@{
                TenantInfo        = if ($tenantInfo) { $tenantInfo } else { $null }
                ConditionalAccess = if ($caPolicies) { $caPolicies } else { @() }
                AppRegistrations  = if ($appRegs) { $appRegs } else { @() }
                EnterpriseApps    = if ($enterpriseApps) { $enterpriseApps } else { @() }
                DirectoryRoles    = if ($directoryRoles) { $directoryRoles } else { @() }
                NamedLocations    = if ($namedLocations) { $namedLocations } else { @() }
                GroupSummary      = if ($groupSummary) { $groupSummary } else { @() }
                Licenses          = if ($licenseSummary) { $licenseSummary } else { @() }
            }
        }
        Save-Checkpoint
    }
}
else {
    Write-Status "Identity" "Entra ID not in scope (-IncludeEntra not specified)" "Info"
    # No gap logged - Entra was intentionally excluded
}

# ============================================================
# SECTION 4: NETWORKING (per subscription)
# ============================================================
if (Test-SectionLoaded "Networking") {
    Write-Status "Networking" "Loaded from checkpoint - skipping" "Info"
} else {

Write-Status "Networking" "Collecting networking configuration across all subscriptions..."

$allVNets = [System.Collections.Generic.List[object]]::new()
$allSubnets = [System.Collections.Generic.List[object]]::new()
$allNSGs = [System.Collections.Generic.List[object]]::new()
$allRouteTables = [System.Collections.Generic.List[object]]::new()
$allPeerings = [System.Collections.Generic.List[object]]::new()
$allVPNGateways = [System.Collections.Generic.List[object]]::new()
$allLocalNetGateways = [System.Collections.Generic.List[object]]::new()
$allConnections = [System.Collections.Generic.List[object]]::new()
$allBastions = [System.Collections.Generic.List[object]]::new()
$allPrivateDNSZones = [System.Collections.Generic.List[object]]::new()
$allPublicIPs = [System.Collections.Generic.List[object]]::new()
$allNVAs = [System.Collections.Generic.List[object]]::new()
$allFirewalls = [System.Collections.Generic.List[object]]::new()
$allAppGateways = [System.Collections.Generic.List[object]]::new()
$allLoadBalancers = [System.Collections.Generic.List[object]]::new()
$allNatGateways = [System.Collections.Generic.List[object]]::new()
$allExpressRoutes = [System.Collections.Generic.List[object]]::new()
$allDdosPlans = [System.Collections.Generic.List[object]]::new()
$allPrivateEndpoints = [System.Collections.Generic.List[object]]::new()
$allNsgFlowLogs = [System.Collections.Generic.List[object]]::new()

foreach ($sub in $allSubs) {
    Set-AzContext -SubscriptionId $sub.Id -Force | Out-Null
    Write-Status "Networking" "  Scanning subscription: $($sub.Name)..."
    
    # Virtual Networks
    $vnets = Get-AzVirtualNetwork
    foreach ($vnet in $vnets) {
        $null = $allVNets.Add([PSCustomObject]@{
            VNetName        = $vnet.Name
            ResourceGroup   = $vnet.ResourceGroupName
            Subscription    = $sub.Name
            Location        = $vnet.Location
            AddressSpace    = ($vnet.AddressSpace.AddressPrefixes) -join ", "
            DNSServers      = if ($vnet.DhcpOptions.DnsServers) { ($vnet.DhcpOptions.DnsServers) -join ", " } else { "Azure Default" }
            SubnetCount     = $vnet.Subnets.Count
            PeeringCount    = $vnet.VirtualNetworkPeerings.Count
        })
        
        # Subnets
        foreach ($snet in $vnet.Subnets) {
            $nsgName = if ($snet.NetworkSecurityGroup) { ($snet.NetworkSecurityGroup.Id -split "/")[-1] } else { "None" }
            $rtName = if ($snet.RouteTable) { ($snet.RouteTable.Id -split "/")[-1] } else { "None" }
            
            $null = $allSubnets.Add([PSCustomObject]@{
                SubnetName       = $snet.Name
                VNetName         = $vnet.Name
                Subscription     = $sub.Name
                AddressPrefix    = ($snet.AddressPrefix) -join ", "
                NSG              = $nsgName
                RouteTable       = $rtName
                ServiceEndpoints = if ($snet.ServiceEndpoints) { ($snet.ServiceEndpoints.Service) -join ", " } else { "None" }
                Delegations      = if ($snet.Delegations) { ($snet.Delegations.ServiceName) -join ", " } else { "None" }
                PrivateEndpointPolicy = $snet.PrivateEndpointNetworkPolicies
            })
            
            # Gap: Subnet without NSG (except special subnets)
            $specialSubnets = @("GatewaySubnet", "AzureBastionSubnet", "AzureFirewallSubnet", "AzureFirewallManagementSubnet", "RouteServerSubnet")
            if ($nsgName -eq "None" -and $snet.Name -notin $specialSubnets) {
                Add-Gap "Networking-NSG" "Subnet '$($snet.Name)' in VNet '$($vnet.Name)' has no NSG associated." "High"
            }
        }
        
        # Peerings
        foreach ($peer in $vnet.VirtualNetworkPeerings) {
            $remoteVNet = ($peer.RemoteVirtualNetwork.Id -split "/")[-1]
            $null = $allPeerings.Add([PSCustomObject]@{
                PeeringName          = $peer.Name
                SourceVNet           = $vnet.Name
                SourceSubscription   = $sub.Name
                RemoteVNet           = $remoteVNet
                PeeringState         = $peer.PeeringState
                AllowGatewayTransit  = $peer.AllowGatewayTransit
                UseRemoteGateways    = $peer.UseRemoteGateways
                AllowForwardedTraffic = $peer.AllowForwardedTraffic
                AllowVNetAccess      = $peer.AllowVirtualNetworkAccess
            })
            
            if ($peer.PeeringState -ne "Connected") {
                Add-Gap "Networking-Peering" "VNet peering '$($peer.Name)' is in state '$($peer.PeeringState)' - not Connected." "Critical"
            }
        }
    }
    
    # NSGs
    $nsgs = Get-AzNetworkSecurityGroup
    foreach ($nsg in $nsgs) {
        foreach ($rule in $nsg.SecurityRules) {
            $null = $allNSGs.Add([PSCustomObject]@{
                NSGName              = $nsg.Name
                ResourceGroup        = $nsg.ResourceGroupName
                Subscription         = $sub.Name
                RuleName             = $rule.Name
                Direction            = $rule.Direction
                Priority             = $rule.Priority
                Access               = $rule.Access
                Protocol             = $rule.Protocol
                SourceAddress        = ($rule.SourceAddressPrefix) -join ", "
                SourcePort           = ($rule.SourcePortRange) -join ", "
                DestAddress          = ($rule.DestinationAddressPrefix) -join ", "
                DestPort             = ($rule.DestinationPortRange) -join ", "
            })
        }
        
        # Gap: NSG with allow-all inbound from internet
        $dangerousRules = $nsg.SecurityRules | Where-Object {
            $_.Direction -eq "Inbound" -and $_.Access -eq "Allow" -and 
            ($_.SourceAddressPrefix -eq "*" -or $_.SourceAddressPrefix -eq "Internet") -and
            ($_.DestinationPortRange -eq "*" -or $_.DestinationPortRange -eq "3389" -or $_.DestinationPortRange -eq "22")
        }
        if ($dangerousRules) {
            Add-Gap "Networking-NSG" "NSG '$($nsg.Name)' has rules allowing RDP/SSH/All from Internet: $($dangerousRules.Name -join ', ')" "Critical"
        }
    }
    
    # Route Tables
    $routeTables = Get-AzRouteTable
    foreach ($rt in $routeTables) {
        foreach ($route in $rt.Routes) {
            $null = $allRouteTables.Add([PSCustomObject]@{
                RouteTableName   = $rt.Name
                ResourceGroup    = $rt.ResourceGroupName
                Subscription     = $sub.Name
                RouteName        = $route.Name
                AddressPrefix    = $route.AddressPrefix
                NextHopType      = $route.NextHopType
                NextHopIP        = $route.NextHopIpAddress
                AssociatedSubnets = if ($rt.Subnets) { ($rt.Subnets.Id | ForEach-Object { ($_ -split "/")[-1] }) -join ", " } else { "None" }
            })
        }
    }
    
    # VPN Gateways
    $vpnGws = Get-AzVirtualNetworkGateway -ResourceGroupName * -ErrorAction SilentlyContinue
    foreach ($gw in $vpnGws) {
        $null = $allVPNGateways.Add([PSCustomObject]@{
            GatewayName      = $gw.Name
            ResourceGroup    = $gw.ResourceGroupName
            Subscription     = $sub.Name
            GatewayType      = $gw.GatewayType
            VpnType          = $gw.VpnType
            SKU              = $gw.Sku.Name
            ActiveActive     = $gw.ActiveActive
            EnableBGP        = $gw.EnableBgp
            BGPAsn           = $gw.BgpSettings.Asn
            PublicIPs        = ($gw.IpConfigurations | ForEach-Object { 
                if ($_.PublicIpAddress) { ($_.PublicIpAddress.Id -split "/")[-1] } 
            }) -join ", "
            Location         = $gw.Location
        })
    }
    
    # Local Network Gateways
    $lngs = Get-AzLocalNetworkGateway -ResourceGroupName * -ErrorAction SilentlyContinue
    foreach ($lng in $lngs) {
        $null = $allLocalNetGateways.Add([PSCustomObject]@{
            Name              = $lng.Name
            ResourceGroup     = $lng.ResourceGroupName
            Subscription      = $sub.Name
            GatewayIPAddress  = $lng.GatewayIpAddress
            AddressSpaces     = ($lng.LocalNetworkAddressSpace.AddressPrefixes) -join ", "
            BGPAsn            = $lng.BgpSettings.Asn
            BGPPeeringAddress = $lng.BgpSettings.BgpPeeringAddress
        })
    }
    
    # VPN Connections
    $connections = Get-AzVirtualNetworkGatewayConnection -ResourceGroupName * -ErrorAction SilentlyContinue
    foreach ($conn in $connections) {
        $null = $allConnections.Add([PSCustomObject]@{
            ConnectionName    = $conn.Name
            ResourceGroup     = $conn.ResourceGroupName
            Subscription      = $sub.Name
            ConnectionType    = $conn.ConnectionType
            ConnectionStatus  = $conn.ConnectionStatus
            VPNGateway        = if ($conn.VirtualNetworkGateway1) { ($conn.VirtualNetworkGateway1.Id -split "/")[-1] } else { "" }
            LocalNetGateway   = if ($conn.LocalNetworkGateway2) { ($conn.LocalNetworkGateway2.Id -split "/")[-1] } else { "" }
            RoutingWeight     = $conn.RoutingWeight
            Protocol          = $conn.ConnectionProtocol
        })
        
        if ($conn.ConnectionStatus -ne "Connected") {
            Add-Gap "Networking-VPN" "VPN connection '$($conn.Name)' status is '$($conn.ConnectionStatus)' - not Connected." "Critical"
        }
    }
    
    # Azure Bastion
    try {
        $bastions = Get-AzBastion -ErrorAction SilentlyContinue
        foreach ($bastion in $bastions) {
            $null = $allBastions.Add([PSCustomObject]@{
                BastionName      = $bastion.Name
                ResourceGroup    = $bastion.ResourceGroupName
                Subscription     = $sub.Name
                Location         = $bastion.Location
                SKU              = $bastion.Sku.Name
                DNSName          = $bastion.DnsName
            })
        }
    }
    catch { }
    
    # Private DNS Zones
    $dnsZones = Get-AzPrivateDnsZone -ErrorAction SilentlyContinue
    foreach ($zone in $dnsZones) {
        $links = Get-AzPrivateDnsVirtualNetworkLink -ResourceGroupName $zone.ResourceGroupName -ZoneName $zone.Name -ErrorAction SilentlyContinue
        $null = $allPrivateDNSZones.Add([PSCustomObject]@{
            ZoneName         = $zone.Name
            ResourceGroup    = $zone.ResourceGroupName
            Subscription     = $sub.Name
            RecordSetCount   = $zone.NumberOfRecordSets
            LinkedVNets      = ($links | ForEach-Object { ($_.VirtualNetworkId -split "/")[-1] }) -join ", "
            AutoRegistration = ($links | Where-Object { $_.RegistrationEnabled } | ForEach-Object { ($_.VirtualNetworkId -split "/")[-1] }) -join ", "
        })
    }
    
    # Public IPs
    $pips = Get-AzPublicIpAddress
    foreach ($pip in $pips) {
        $null = $allPublicIPs.Add([PSCustomObject]@{
            Name             = $pip.Name
            ResourceGroup    = $pip.ResourceGroupName
            Subscription     = $sub.Name
            IPAddress        = $pip.IpAddress
            AllocationMethod = $pip.PublicIpAllocationMethod
            SKU              = $pip.Sku.Name
            AssociatedTo     = if ($pip.IpConfiguration) { ($pip.IpConfiguration.Id -split "/")[8] + "/" + ($pip.IpConfiguration.Id -split "/")[-3] } else { "Unassociated" }
            Location         = $pip.Location
        })
        
        if (-not $pip.IpConfiguration) {
            Add-Recommendation "Networking" "Public IP '$($pip.Name)' is unassociated - consider removing to save cost."
        }
    }

    # Azure Firewalls
    $firewalls = Get-AzFirewall -ErrorAction SilentlyContinue
    foreach ($fw in $firewalls) {
        $null = $allFirewalls.Add([PSCustomObject]@{
            FirewallName    = $fw.Name
            ResourceGroup   = $fw.ResourceGroupName
            Subscription    = $sub.Name
            Location        = $fw.Location
            SKU             = if ($fw.Sku) { $fw.Sku.Tier } else { "N/A" }
            ThreatIntelMode = $fw.ThreatIntelMode
            PolicyName      = if ($fw.FirewallPolicy) { ($fw.FirewallPolicy.Id -split "/")[-1] } else { "None" }
            Zones           = if ($fw.Zones) { ($fw.Zones) -join ", " } else { "None" }
        })
    }

    # Application Gateways
    $appGws = Get-AzApplicationGateway -ErrorAction SilentlyContinue
    foreach ($agw in $appGws) {
        $null = $allAppGateways.Add([PSCustomObject]@{
            Name             = $agw.Name
            ResourceGroup    = $agw.ResourceGroupName
            Subscription     = $sub.Name
            Location         = $agw.Location
            SKU              = $agw.Sku.Name
            Tier             = $agw.Sku.Tier
            WAFEnabled       = if ($agw.WebApplicationFirewallConfiguration -or $agw.FirewallPolicy) { $true } else { $false }
            ListenerCount    = $agw.HttpListeners.Count
            BackendPoolCount = $agw.BackendAddressPools.Count
            Zones            = if ($agw.Zones) { ($agw.Zones) -join ", " } else { "None" }
        })
    }

    # Load Balancers
    $lbs = Get-AzLoadBalancer -ErrorAction SilentlyContinue
    foreach ($lb in $lbs) {
        $lbType = "Internal"
        foreach ($feip in $lb.FrontendIpConfigurations) {
            if ($feip.PublicIpAddress) { $lbType = "Public"; break }
        }
        $null = $allLoadBalancers.Add([PSCustomObject]@{
            Name             = $lb.Name
            ResourceGroup    = $lb.ResourceGroupName
            Subscription     = $sub.Name
            SKU              = $lb.Sku.Name
            Type             = $lbType
            FrontendCount    = $lb.FrontendIpConfigurations.Count
            BackendPoolCount = $lb.BackendAddressPools.Count
            RuleCount        = $lb.LoadBalancingRules.Count
            ProbeCount       = $lb.Probes.Count
        })
    }

    # NAT Gateways
    $natGws = Get-AzNatGateway -ErrorAction SilentlyContinue
    foreach ($ng in $natGws) {
        $null = $allNatGateways.Add([PSCustomObject]@{
            Name            = $ng.Name
            ResourceGroup   = $ng.ResourceGroupName
            Subscription    = $sub.Name
            Location        = $ng.Location
            PublicIPs       = if ($ng.PublicIpAddresses) { ($ng.PublicIpAddresses | ForEach-Object { ($_.Id -split "/")[-1] }) -join ", " } else { "None" }
            IdleTimeout     = $ng.IdleTimeoutInMinutes
            Zones           = if ($ng.Zones) { ($ng.Zones) -join ", " } else { "None" }
        })
    }

    # ExpressRoute Circuits
    $erCircuits = Get-AzExpressRouteCircuit -ErrorAction SilentlyContinue
    foreach ($er in $erCircuits) {
        $null = $allExpressRoutes.Add([PSCustomObject]@{
            Name            = $er.Name
            ResourceGroup   = $er.ResourceGroupName
            Subscription    = $sub.Name
            Location        = $er.Location
            Provider        = $er.ServiceProviderProperties.ServiceProviderName
            Bandwidth       = "$($er.ServiceProviderProperties.BandwidthInMbps) Mbps"
            SKU             = "$($er.Sku.Tier) / $($er.Sku.Family)"
            CircuitState    = $er.CircuitProvisioningState
            PeeringCount    = $er.Peerings.Count
        })
    }

    # DDoS Protection Plans
    $ddosPlans = Get-AzDdosProtectionPlan -ErrorAction SilentlyContinue
    foreach ($ddos in $ddosPlans) {
        $null = $allDdosPlans.Add([PSCustomObject]@{
            Name            = $ddos.Name
            ResourceGroup   = $ddos.ResourceGroupName
            Subscription    = $sub.Name
            Location        = $ddos.Location
            ProtectedVNets  = if ($ddos.VirtualNetworks) { $ddos.VirtualNetworks.Count } else { 0 }
        })
    }

    # Private Endpoints
    $pes = Get-AzPrivateEndpoint -ErrorAction SilentlyContinue
    foreach ($pe in $pes) {
        $targetType = ""
        $targetName = ""
        if ($pe.PrivateLinkServiceConnections -and $pe.PrivateLinkServiceConnections.Count -gt 0) {
            $conn = $pe.PrivateLinkServiceConnections[0]
            $targetType = ($conn.PrivateLinkServiceId -split "/")[-2]
            $targetName = ($conn.PrivateLinkServiceId -split "/")[-1]
        }
        $null = $allPrivateEndpoints.Add([PSCustomObject]@{
            Name            = $pe.Name
            ResourceGroup   = $pe.ResourceGroupName
            Subscription    = $sub.Name
            Location        = $pe.Location
            TargetType      = $targetType
            TargetResource  = $targetName
            Subnet          = if ($pe.Subnet) { ($pe.Subnet.Id -split "/")[-1] } else { "N/A" }
            Status          = if ($pe.PrivateLinkServiceConnections[0].PrivateLinkServiceConnectionState) { $pe.PrivateLinkServiceConnections[0].PrivateLinkServiceConnectionState.Status } else { "N/A" }
        })
    }

    # NSG Flow Logs
    try {
        $watchers = Get-AzNetworkWatcher -ErrorAction SilentlyContinue | Where-Object { $_.Location -in ($vnets.Location | Select-Object -Unique) }
        foreach ($watcher in $watchers) {
            $flowLogs = Get-AzNetworkWatcherFlowLog -NetworkWatcher $watcher -ErrorAction SilentlyContinue
            foreach ($fl in $flowLogs) {
                $null = $allNsgFlowLogs.Add([PSCustomObject]@{
                    FlowLogName      = $fl.Name
                    NSG              = if ($fl.TargetResourceId) { ($fl.TargetResourceId -split "/")[-1] } else { "N/A" }
                    Subscription     = $sub.Name
                    Enabled          = $fl.Enabled
                    StorageAccount   = if ($fl.StorageId) { ($fl.StorageId -split "/")[-1] } else { "None" }
                    RetentionDays    = if ($fl.RetentionPolicy) { $fl.RetentionPolicy.Days } else { 0 }
                    TrafficAnalytics = if ($fl.FlowAnalyticsConfiguration -and $fl.FlowAnalyticsConfiguration.NetworkWatcherFlowAnalyticsConfiguration.Enabled) { $true } else { $false }
                })
            }
        }
    } catch { }
}

$tenancyData.Sections["Networking"] = [ordered]@{
}
Save-Checkpoint
    VirtualNetworks      = $allVNets
    Subnets              = $allSubnets
    NSGRules             = $allNSGs
    RouteTables          = $allRouteTables
    VNetPeerings         = $allPeerings
    VPNGateways          = $allVPNGateways
    LocalNetGateways     = $allLocalNetGateways
    VPNConnections       = $allConnections
    Bastions             = $allBastions
    PrivateDNSZones      = $allPrivateDNSZones
    PublicIPs            = $allPublicIPs
    Firewalls            = $allFirewalls
    ApplicationGateways  = $allAppGateways
    LoadBalancers        = $allLoadBalancers
    NATGateways          = $allNatGateways
    ExpressRoutes        = $allExpressRoutes
    DDoSPlans            = $allDdosPlans
    PrivateEndpoints     = $allPrivateEndpoints
    NSGFlowLogs          = $allNsgFlowLogs
}

Write-Status "Networking" "VNets: $($allVNets.Count), Subnets: $($allSubnets.Count), NSG Rules: $($allNSGs.Count), Peerings: $($allPeerings.Count)" "Success"

Add-Recommendation "Networking" "Enable Azure DDoS Protection Standard for internet-facing VNets to protect against DDoS attacks." "Medium"
Add-Recommendation "Networking" "Enable NSG Flow Logs on all NSGs for network traffic analysis and troubleshooting." "Medium"
Add-Recommendation "Networking" "Use Private Endpoints for PaaS services to eliminate public network exposure." "Medium"
Add-Recommendation "Networking" "Implement hub-spoke network topology with Azure Firewall for centralised network security." "Medium"

# ============================================================
# SECTION 5: COMPUTE (VMs, AVD, etc.)
# ============================================================
if (Test-SectionLoaded "Compute") {
    Write-Status "Compute" "Loaded from checkpoint - skipping" "Info"
} else {

Write-Status "Compute" "Collecting compute resources across all subscriptions..."

$allVMs = [System.Collections.Generic.List[object]]::new()
$allVMExtensions = [System.Collections.Generic.List[object]]::new()
$allAvailSets = [System.Collections.Generic.List[object]]::new()
$allAVDHostPools = [System.Collections.Generic.List[object]]::new()
$allAVDSessionHosts = [System.Collections.Generic.List[object]]::new()
$allAVDAppGroups = [System.Collections.Generic.List[object]]::new()
$allManagedDisks = [System.Collections.Generic.List[object]]::new()
$allSnapshots = [System.Collections.Generic.List[object]]::new()
$allVMSS = [System.Collections.Generic.List[object]]::new()

foreach ($sub in $allSubs) {
    Set-AzContext -SubscriptionId $sub.Id -Force | Out-Null
    Write-Status "Compute" "  Scanning subscription: $($sub.Name)..."
    
    # Virtual Machines
    $vms = Get-AzVM -Status
    foreach ($vm in $vms) {
        $vmDetail = Get-AzVM -ResourceGroupName $vm.ResourceGroupName -Name $vm.Name
        $nics = @()
        $privateIPs = @()
        foreach ($nicRef in $vmDetail.NetworkProfile.NetworkInterfaces) {
            $nic = Get-AzNetworkInterface -ResourceId $nicRef.Id -ErrorAction SilentlyContinue
            if ($nic) {
                foreach ($ipConfig in $nic.IpConfigurations) {
                    $privateIPs += $ipConfig.PrivateIpAddress
                }
            }
        }
        
        # Disk info
        $osDisk = $vmDetail.StorageProfile.OsDisk
        $dataDisks = $vmDetail.StorageProfile.DataDisks
        $diskInfo = "OS: $($osDisk.ManagedDisk.StorageAccountType)"
        if ($dataDisks.Count -gt 0) {
            $diskInfo += " | Data: " + ($dataDisks | ForEach-Object { "$($_.Name)($($_.DiskSizeGB)GB)" }) -join ", "
        }
        
        $null = $allVMs.Add([PSCustomObject]@{
            VMName           = $vm.Name
            ResourceGroup    = $vm.ResourceGroupName
            Subscription     = $sub.Name
            Location         = $vm.Location
            VMSize           = $vm.HardwareProfile.VmSize
            PowerState       = ($vm.Statuses | Where-Object { $_.Code -like "PowerState/*" }).DisplayStatus
            OSType           = $vmDetail.StorageProfile.OsDisk.OsType
            OSImage          = "$($vmDetail.StorageProfile.ImageReference.Publisher)/$($vmDetail.StorageProfile.ImageReference.Offer)/$($vmDetail.StorageProfile.ImageReference.Sku)"
            PrivateIPs       = ($privateIPs) -join ", "
            AvailabilitySet  = if ($vmDetail.AvailabilitySetReference) { ($vmDetail.AvailabilitySetReference.Id -split "/")[-1] } else { "None" }
            Zone             = if ($vmDetail.Zones) { ($vmDetail.Zones) -join "," } else { "None" }
            DiskConfig       = $diskInfo
            DataDiskCount    = $dataDisks.Count
            BootDiagnostics  = if ($vmDetail.DiagnosticsProfile.BootDiagnostics.Enabled) { "Enabled" } else { "Disabled" }
            Tags             = if ($vmDetail.Tags) { ($vmDetail.Tags.GetEnumerator() | ForEach-Object { "$($_.Key)=$($_.Value)" }) -join "; " } else { "None" }
        })
        
        # VM Extensions
        $extensions = Get-AzVMExtension -ResourceGroupName $vm.ResourceGroupName -VMName $vm.Name -ErrorAction SilentlyContinue
        foreach ($ext in $extensions) {
            $null = $allVMExtensions.Add([PSCustomObject]@{
                VMName           = $vm.Name
                Subscription     = $sub.Name
                ExtensionName    = $ext.Name
                Publisher        = $ext.Publisher
                Type             = $ext.ExtensionType
                Version          = $ext.TypeHandlerVersion
                ProvisioningState = $ext.ProvisioningState
            })
        }
        
        # Gap: VM without boot diagnostics
        if (-not $vmDetail.DiagnosticsProfile.BootDiagnostics.Enabled) {
            Add-Recommendation "Compute" "VM '$($vm.Name)' has boot diagnostics disabled. Enable for troubleshooting."
        }
    }
    
    # Availability Sets
    $avSets = Get-AzAvailabilitySet -ErrorAction SilentlyContinue
    foreach ($avSet in $avSets) {
        $null = $allAvailSets.Add([PSCustomObject]@{
            Name              = $avSet.Name
            ResourceGroup     = $avSet.ResourceGroupName
            Subscription      = $sub.Name
            Location          = $avSet.Location
            FaultDomains      = $avSet.PlatformFaultDomainCount
            UpdateDomains     = $avSet.PlatformUpdateDomainCount
            VMCount           = $avSet.VirtualMachinesReferences.Count
            SKU               = $avSet.Sku
        })
    }
    
    # Managed Disks
    $disks = Get-AzDisk -ErrorAction SilentlyContinue
    foreach ($disk in $disks) {
        $null = $allManagedDisks.Add([PSCustomObject]@{
            DiskName         = $disk.Name
            ResourceGroup    = $disk.ResourceGroupName
            Subscription     = $sub.Name
            Location         = $disk.Location
            SizeGB           = $disk.DiskSizeGB
            SKU              = $disk.Sku.Name
            State            = $disk.DiskState
            EncryptionType   = if ($disk.Encryption) { $disk.Encryption.Type } else { "PlatformManaged" }
            AttachedTo       = if ($disk.ManagedBy) { ($disk.ManagedBy -split "/")[-1] } else { "Unattached" }
            Zone             = if ($disk.Zones) { ($disk.Zones) -join ", " } else { "None" }
        })
    }

    # Snapshots
    $snaps = Get-AzSnapshot -ErrorAction SilentlyContinue
    foreach ($snap in $snaps) {
        $null = $allSnapshots.Add([PSCustomObject]@{
            SnapshotName    = $snap.Name
            ResourceGroup   = $snap.ResourceGroupName
            Subscription    = $sub.Name
            Location        = $snap.Location
            SizeGB          = $snap.DiskSizeGB
            SourceDisk      = if ($snap.CreationData.SourceResourceId) { ($snap.CreationData.SourceResourceId -split "/")[-1] } else { "N/A" }
            Incremental     = $snap.Incremental
            TimeCreated     = if ($snap.TimeCreated) { $snap.TimeCreated.ToString("yyyy-MM-dd") } else { "N/A" }
        })
    }

    # VM Scale Sets
    $vmssList = Get-AzVmss -ErrorAction SilentlyContinue
    foreach ($vmss in $vmssList) {
        $null = $allVMSS.Add([PSCustomObject]@{
            Name            = $vmss.Name
            ResourceGroup   = $vmss.ResourceGroupName
            Subscription    = $sub.Name
            Location        = $vmss.Location
            SKU             = $vmss.Sku.Name
            Capacity        = $vmss.Sku.Capacity
            UpgradePolicy   = $vmss.UpgradePolicy.Mode
            Zones           = if ($vmss.Zones) { ($vmss.Zones) -join ", " } else { "None" }
            OrchestrationMode = $vmss.OrchestrationMode
        })
    }

    # Azure Virtual Desktop
    try {
        $hostPools = Get-AzWvdHostPool -ErrorAction SilentlyContinue
        foreach ($hp in $hostPools) {
            $rgName = ($hp.Id -split "/")[4]
            $null = $allAVDHostPools.Add([PSCustomObject]@{
                HostPoolName     = $hp.Name
                ResourceGroup    = $rgName
                Subscription     = $sub.Name
                Location         = $hp.Location
                HostPoolType     = $hp.HostPoolType
                LoadBalancerType = $hp.LoadBalancerType
                MaxSessionLimit  = $hp.MaxSessionLimit
                ValidationEnv    = $hp.ValidationEnvironment
                StartVMOnConnect = $hp.StartVMOnConnect
                PreferredAppGroupType = $hp.PreferredAppGroupType
            })
            
            # Session Hosts
            $sessionHosts = Get-AzWvdSessionHost -ResourceGroupName $rgName -HostPoolName $hp.Name -ErrorAction SilentlyContinue
            foreach ($sh in $sessionHosts) {
                $null = $allAVDSessionHosts.Add([PSCustomObject]@{
                    HostPoolName     = $hp.Name
                    SessionHostName  = $sh.Name
                    Status           = $sh.Status
                    UpdateState      = $sh.UpdateState
                    LastHeartBeat    = $sh.LastHeartBeat
                    Sessions         = $sh.Session
                    AllowNewSession  = $sh.AllowNewSession
                    OSVersion        = $sh.OsVersion
                    AgentVersion     = $sh.AgentVersion
                })
            }
        }
        
        # App Groups
        $appGroups = Get-AzWvdApplicationGroup -ErrorAction SilentlyContinue
        foreach ($ag in $appGroups) {
            $null = $allAVDAppGroups.Add([PSCustomObject]@{
                AppGroupName     = $ag.Name
                ResourceGroup    = ($ag.Id -split "/")[4]
                Subscription     = $sub.Name
                HostPoolRef      = ($ag.HostPoolArmPath -split "/")[-1]
                AppGroupType     = $ag.ApplicationGroupType
                FriendlyName     = $ag.FriendlyName
            })
        }
    }
    catch { }
}

$tenancyData.Sections["Compute"] = [ordered]@{
}
Save-Checkpoint
    VirtualMachines    = $allVMs
    VMExtensions       = $allVMExtensions
    AvailabilitySets   = $allAvailSets
    ManagedDisks       = $allManagedDisks
    Snapshots          = $allSnapshots
    VMScaleSets        = $allVMSS
    AVDHostPools       = $allAVDHostPools
    AVDSessionHosts    = $allAVDSessionHosts
    AVDAppGroups       = $allAVDAppGroups
}

Write-Status "Compute" "VMs: $($allVMs.Count), Extensions: $($allVMExtensions.Count), AVD Pools: $($allAVDHostPools.Count)" "Success"

Add-Recommendation "Compute" "Deploy VMs in Availability Zones for zone-level resilience where supported." "Medium"
Add-Recommendation "Compute" "Deploy Azure Monitor Agent (AMA) to all VMs for unified monitoring." "High"
Add-Recommendation "Compute" "Review VM sizing against Azure Advisor recommendations to optimise cost and performance." "Medium"
Add-Recommendation "Compute" "Consider Azure Hybrid Benefit for Windows Server and SQL Server VMs to reduce licensing costs." "Low"

# Gap: VMs without tags
$untaggedVMs = $allVMs | Where-Object { $_.Tags -eq "None" }
if ($untaggedVMs.Count -gt 0) {
    Add-Gap "Compute-Tags" "$($untaggedVMs.Count) VMs have no tags applied. Tagging standard not enforced." "High"
}

# ============================================================
# SECTION 6: STORAGE
# ============================================================
if (Test-SectionLoaded "Storage") {
    Write-Status "Storage" "Loaded from checkpoint - skipping" "Info"
} else {

Write-Status "Storage" "Collecting storage accounts..."

$allStorageAccounts = [System.Collections.Generic.List[object]]::new()

foreach ($sub in $allSubs) {
    Set-AzContext -SubscriptionId $sub.Id -Force | Out-Null
    
    $storageAccounts = Get-AzStorageAccount
    foreach ($sa in $storageAccounts) {
        $null = $allStorageAccounts.Add([PSCustomObject]@{
            AccountName      = $sa.StorageAccountName
            ResourceGroup    = $sa.ResourceGroupName
            Subscription     = $sub.Name
            Location         = $sa.PrimaryLocation
            SKU              = $sa.Sku.Name
            Kind             = $sa.Kind
            AccessTier       = $sa.AccessTier
            HTTPSOnly        = $sa.EnableHttpsTrafficOnly
            MinTLS           = $sa.MinimumTlsVersion
            PublicAccess     = $sa.AllowBlobPublicAccess
            NetworkRuleSet   = $sa.NetworkRuleSet.DefaultAction
            PrivateEndpoints = if ($sa.PrivateEndpointConnections) { $sa.PrivateEndpointConnections.Count } else { 0 }
            Tags             = if ($sa.Tags) { ($sa.Tags.GetEnumerator() | ForEach-Object { "$($_.Key)=$($_.Value)" }) -join "; " } else { "None" }
        })
        
        # Security gaps
        if ($sa.AllowBlobPublicAccess -eq $true) {
            Add-Gap "Storage-Security" "Storage account '$($sa.StorageAccountName)' allows public blob access." "High"
        }
        if ($sa.MinimumTlsVersion -ne "TLS1_2") {
            Add-Gap "Storage-Security" "Storage account '$($sa.StorageAccountName)' minimum TLS is '$($sa.MinimumTlsVersion)' - should be TLS1_2." "Medium"
        }
        if ($sa.NetworkRuleSet.DefaultAction -eq "Allow") {
            Add-Gap "Storage-Security" "Storage account '$($sa.StorageAccountName)' allows public network access by default." "Medium"
        }
    }
}

$tenancyData.Sections["Storage"] = $allStorageAccounts
Save-Checkpoint
}
Write-Status "Storage" "Found $($allStorageAccounts.Count) storage accounts" "Success"

Add-Recommendation "Storage" "Enable soft delete for blob storage to protect against accidental deletion." "Medium"
Add-Recommendation "Storage" "Configure lifecycle management policies on storage accounts for cost optimisation." "Low"
Add-Recommendation "Storage" "Use Private Endpoints for storage accounts to restrict access to VNet only." "Medium"

# ============================================================
# SECTION 7: BACKUP & RECOVERY
# ============================================================
if (Test-SectionLoaded "Backup") {
    Write-Status "Backup" "Loaded from checkpoint - skipping" "Info"
} else {

Write-Status "Backup" "Collecting backup configuration..."

$allVaults = [System.Collections.Generic.List[object]]::new()
$allBackupPolicies = [System.Collections.Generic.List[object]]::new()
$allBackupItems = [System.Collections.Generic.List[object]]::new()
$allASRItems = [System.Collections.Generic.List[object]]::new()
$allASRPolicies = [System.Collections.Generic.List[object]]::new()

foreach ($sub in $allSubs) {
    Set-AzContext -SubscriptionId $sub.Id -Force | Out-Null

    $vaults = Get-AzRecoveryServicesVault -ErrorAction SilentlyContinue
    foreach ($vault in $vaults) {
        Set-AzRecoveryServicesVaultContext -Vault $vault
        
        $vaultProps = Get-AzRecoveryServicesVaultProperty -VaultId $vault.ID -ErrorAction SilentlyContinue
        $backupProps = Get-AzRecoveryServicesBackupProperty -Vault $vault -ErrorAction SilentlyContinue

        # Soft-delete: try VaultProperty first, then fall back to vault SecuritySettings
        $softDel = "Unknown"
        if ($vaultProps -and $vaultProps.SoftDeleteFeatureState) {
            $softDel = $vaultProps.SoftDeleteFeatureState
        } elseif ($vault.Properties.SecuritySettings.SoftDeleteSettings.SoftDeleteState) {
            $softDel = $vault.Properties.SecuritySettings.SoftDeleteSettings.SoftDeleteState
        }

        # Immutability: try VaultProperty, then vault SecuritySettings
        $immut = "Unknown"
        if ($vaultProps -and $vaultProps.ImmutabilityState) {
            $immut = $vaultProps.ImmutabilityState
        } elseif ($vault.Properties.SecuritySettings.ImmutabilitySettings.State) {
            $immut = $vault.Properties.SecuritySettings.ImmutabilitySettings.State
        }

        $null = $allVaults.Add([PSCustomObject]@{
            VaultName        = $vault.Name
            ResourceGroup    = $vault.ResourceGroupName
            Subscription     = $sub.Name
            Location         = $vault.Location
            Redundancy       = $backupProps.BackupStorageRedundancy
            SoftDelete       = $softDel
            Immutability     = $immut
            CrossRegionRestore = if ($backupProps) { $backupProps.CrossRegionRestore } else { "Unknown" }
        })
        
        # Backup Policies
        $policies = Get-AzRecoveryServicesBackupProtectionPolicy -ErrorAction SilentlyContinue
        foreach ($policy in $policies) {
            $null = $allBackupPolicies.Add([PSCustomObject]@{
                PolicyName       = $policy.Name
                VaultName        = $vault.Name
                Subscription     = $sub.Name
                WorkloadType     = $policy.WorkloadType
                BackupTime       = if ($policy.SchedulePolicy.ScheduleRunTimes) { ($policy.SchedulePolicy.ScheduleRunTimes | ForEach-Object { $_.ToString("HH:mm") }) -join ", " } else { "" }
                Frequency        = if ($policy.SchedulePolicy.ScheduleRunFrequency) { $policy.SchedulePolicy.ScheduleRunFrequency } else { "" }
                DailyRetention   = if ($policy.RetentionPolicy.DailySchedule) { "$($policy.RetentionPolicy.DailySchedule.DurationCountInDays) days" } else { "N/A" }
                WeeklyRetention  = if ($policy.RetentionPolicy.WeeklySchedule) { "$($policy.RetentionPolicy.WeeklySchedule.DurationCountInWeeks) weeks" } else { "N/A" }
                MonthlyRetention = if ($policy.RetentionPolicy.MonthlySchedule) { "$($policy.RetentionPolicy.MonthlySchedule.DurationCountInMonths) months" } else { "N/A" }
                YearlyRetention  = if ($policy.RetentionPolicy.YearlySchedule) { "$($policy.RetentionPolicy.YearlySchedule.DurationCountInYears) years" } else { "N/A" }
                InstantRestore   = if ($policy.InstantRpRetentionRangeInDays) { "$($policy.InstantRpRetentionRangeInDays) days" } else { "N/A" }
            })
        }
        
        # Backup Items — build policy lookup from collected policies
        $policyLookup = @{}
        foreach ($pol in $policies) {
            if ($pol.Id) { $policyLookup[$pol.Id] = $pol.Name }
        }

        $containers = Get-AzRecoveryServicesBackupContainer -ContainerType AzureVM -ErrorAction SilentlyContinue
        foreach ($container in $containers) {
            $items = Get-AzRecoveryServicesBackupItem -Container $container -WorkloadType AzureVM -ErrorAction SilentlyContinue
            foreach ($item in $items) {
                # Extract clean VM name from container path (e.g. "VM;iaasvmcontainerv2;rg-name;VMName")
                $vmDisplayName = if ($item.Name -match ";") { ($item.Name -split ";")[-1] } else { $item.Name }

                # Resolve policy name: try PolicyId lookup, then ProtectionPolicyId path split, then PolicyName property
                $resolvedPolicy = "N/A"
                if ($item.PolicyId -and $policyLookup.ContainsKey($item.PolicyId)) {
                    $resolvedPolicy = $policyLookup[$item.PolicyId]
                } elseif ($item.ProtectionPolicyId) {
                    $resolvedPolicy = ($item.ProtectionPolicyId -split "/")[-1]
                } elseif ($item.PolicyName) {
                    $resolvedPolicy = $item.PolicyName
                }

                $null = $allBackupItems.Add([PSCustomObject]@{
                    VMName           = $vmDisplayName
                    VaultName        = $vault.Name
                    Subscription     = $sub.Name
                    PolicyName       = $resolvedPolicy
                    ProtectionStatus = $item.ProtectionStatus
                    LastBackupStatus = $item.LastBackupStatus
                    LastBackupTime   = $item.LastBackupTime
                    HealthStatus     = $item.HealthStatus
                })
            }
        }
        
        # ASR Replicated Items
        try {
            $asrFabrics = Get-AzRecoveryServicesAsrFabric -ErrorAction SilentlyContinue
            foreach ($fabric in $asrFabrics) {
                $containers = Get-AzRecoveryServicesAsrProtectionContainer -Fabric $fabric -ErrorAction SilentlyContinue
                foreach ($container in $containers) {
                    $replItems = Get-AzRecoveryServicesAsrReplicationProtectedItem -ProtectionContainer $container -ErrorAction SilentlyContinue
                    foreach ($item in $replItems) {
                        $null = $allASRItems.Add([PSCustomObject]@{
                            FriendlyName    = $item.FriendlyName
                            Subscription    = $sub.Name
                            Vault           = $vault.Name
                            ReplicationHealth = $item.ReplicationHealth
                            ProtectionState = $item.ProtectionState
                            ActiveLocation  = $item.ActiveLocation
                            TestFailoverState = $item.TestFailoverState
                        })
                    }
                }
            }

            # ASR Policies
            $asrPolicies = Get-AzRecoveryServicesAsrPolicy -ErrorAction SilentlyContinue
            foreach ($pol in $asrPolicies) {
                $null = $allASRPolicies.Add([PSCustomObject]@{
                    PolicyName          = $pol.Name
                    Vault               = $vault.Name
                    Subscription        = $sub.Name
                    ReplicationProvider = $pol.ReplicationProvider
                    RecoveryPointRetention = if ($pol.ReplicationProviderSettings) { $pol.ReplicationProviderSettings.RecoveryPointHistory } else { "N/A" }
                    AppConsistentFreq   = if ($pol.ReplicationProviderSettings) { $pol.ReplicationProviderSettings.ApplicationConsistentSnapshotFrequencyInHours } else { "N/A" }
                })
            }
        } catch { }

        # Gaps: Soft delete / immutability
        if ($vaultProps -and $vaultProps.SoftDeleteFeatureState -ne "Enabled") {
            Add-Gap "Backup-Security" "Recovery vault '$($vault.Name)' does not have soft delete enabled." "High"
        }
        if ($vaultProps -and $vaultProps.ImmutabilityState -ne "Locked" -and $vaultProps.ImmutabilityState -ne "Unlocked") {
            Add-Recommendation "Backup" "Recovery vault '$($vault.Name)' does not have immutability enabled. Consider for ransomware protection."
        }
    }
}

# Check VMs without backup
$backedUpVMs = $allBackupItems | ForEach-Object { $_.VMName.Split(";")[-1] }
$unbacked = $allVMs | Where-Object { $_.VMName -notin $backedUpVMs -and $_.PowerState -eq "VM running" }
if ($unbacked.Count -gt 0) {
    Add-Gap "Backup-Coverage" "$($unbacked.Count) running VMs have no backup configured: $($unbacked.VMName -join ', ')" "Critical"
}

$tenancyData.Sections["Backup"] = [ordered]@{
}
Save-Checkpoint
    RecoveryVaults     = $allVaults
    BackupPolicies     = $allBackupPolicies
    BackupItems        = $allBackupItems
    UnprotectedVMs     = $unbacked | Select-Object VMName, Subscription, VMSize
    ASRReplicatedItems = $allASRItems
    ASRPolicies        = $allASRPolicies
}

Write-Status "Backup" "Vaults: $($allVaults.Count), Policies: $($allBackupPolicies.Count), Protected Items: $($allBackupItems.Count), Unprotected VMs: $($unbacked.Count)" "Success"

Add-Recommendation "Backup" "Implement immutability on Recovery Services vaults for ransomware protection." "High"
Add-Recommendation "Backup" "Test backup restore procedures quarterly to validate recoverability." "High"
Add-Recommendation "Backup" "Consider Azure Site Recovery for critical workloads requiring low RPO/RTO." "Medium"

# ============================================================
# SECTION 8: SECURITY (Defender for Cloud)
# ============================================================
if (Test-SectionLoaded "Security") {
    Write-Status "Security" "Loaded from checkpoint - skipping" "Info"
} else {

Write-Status "Security" "Collecting security configuration..."

$allDefenderPlans = [System.Collections.Generic.List[object]]::new()
$allSecureScores = [System.Collections.Generic.List[object]]::new()
$allKeyVaults = [System.Collections.Generic.List[object]]::new()
$allDefenderRecs = [System.Collections.Generic.List[object]]::new()

foreach ($sub in $allSubs) {
    Set-AzContext -SubscriptionId $sub.Id -Force | Out-Null

    # Defender Plans
    try {
        $pricings = Get-AzSecurityPricing -ErrorAction SilentlyContinue
        foreach ($pricing in $pricings) {
            $null = $allDefenderPlans.Add([PSCustomObject]@{
                Subscription     = $sub.Name
                ResourceType     = $pricing.Name
                PricingTier      = $pricing.PricingTier
                SubPlan          = $pricing.SubPlan
                FreeTrialRemaining = $pricing.FreeTrialRemainingTime
            })
        }
    }
    catch { }
    
    # Secure Score
    try {
        $scores = Get-AzSecuritySecureScore -ErrorAction SilentlyContinue
        foreach ($score in $scores) {
            $null = $allSecureScores.Add([PSCustomObject]@{
                Subscription     = $sub.Name
                DisplayName      = $score.DisplayName
                CurrentScore     = $score.CurrentScore
                MaxScore         = $score.MaxScore
                Percentage       = if ($score.MaxScore -gt 0) { [math]::Round(($score.CurrentScore / $score.MaxScore) * 100, 1) } else { 0 }
            })
        }
    }
    catch { }
    
    # Key Vaults (detail call cached per vault to avoid duplicate API calls)
    $kvs = Get-AzKeyVault -ErrorAction SilentlyContinue
    $kvDetailCache = @{}
    foreach ($kv in $kvs) {
        if (-not $kvDetailCache.ContainsKey($kv.VaultName)) {
            $kvDetailCache[$kv.VaultName] = Get-AzKeyVault -VaultName $kv.VaultName -ErrorAction SilentlyContinue
        }
        $kvDetail = $kvDetailCache[$kv.VaultName]
        if ($kvDetail) {
            $null = $allKeyVaults.Add([PSCustomObject]@{
                VaultName        = $kvDetail.VaultName
                ResourceGroup    = $kvDetail.ResourceGroupName
                Subscription     = $sub.Name
                Location         = $kvDetail.Location
                SKU              = $kvDetail.Sku
                SoftDelete       = $kvDetail.EnableSoftDelete
                PurgeProtection  = $kvDetail.EnablePurgeProtection
                RBACAuth         = $kvDetail.EnableRbacAuthorization
                PublicAccess     = $kvDetail.PublicNetworkAccess
                PrivateEndpoints = if ($kvDetail.PrivateEndpointConnections) { $kvDetail.PrivateEndpointConnections.Count } else { 0 }
            })
            
            if (-not $kvDetail.EnablePurgeProtection) {
                Add-Gap "Security-KeyVault" "Key Vault '$($kvDetail.VaultName)' does not have purge protection enabled." "Medium"
            }
        }
    }

    # Defender Recommendations
    try {
        $assessments = Get-AzSecurityAssessment -ErrorAction SilentlyContinue
        foreach ($assess in $assessments) {
            if ($assess.Status.Code -ne "Healthy") {
                $null = $allDefenderRecs.Add([PSCustomObject]@{
                    AssessmentName  = $assess.DisplayName
                    Subscription    = $sub.Name
                    Status          = $assess.Status.Code
                    Severity        = $assess.Metadata.Severity
                    ResourceType    = if ($assess.ResourceDetails.Id) { ($assess.ResourceDetails.Id -split "/")[-2] } else { "N/A" }
                })
            }
        }
    } catch { }
}

# Check if Defender for Servers is enabled
$serversDefender = $allDefenderPlans | Where-Object { $_.ResourceType -eq "VirtualMachines" -and $_.PricingTier -eq "Standard" }
if ($serversDefender.Count -eq 0) {
    Add-Gap "Security-Defender" "Microsoft Defender for Servers is not enabled on any subscription." "High"
}

$tenancyData.Sections["Security"] = [ordered]@{
}
Save-Checkpoint
    DefenderPlans           = $allDefenderPlans
    SecureScores            = $allSecureScores
    KeyVaults               = $allKeyVaults
    DefenderRecommendations = $allDefenderRecs
}

Write-Status "Security" "Defender plans: $($allDefenderPlans.Count), Key Vaults: $($allKeyVaults.Count)" "Success"

Add-Recommendation "Security" "Enable Microsoft Defender for Cloud on all resource types for comprehensive protection." "High"
Add-Recommendation "Security" "Remediate Secure Score recommendations to improve overall security posture." "High"
Add-Recommendation "Security" "Migrate Key Vaults to RBAC authorisation model for fine-grained access control." "Medium"

# ============================================================
# SECTION 9: GOVERNANCE & POLICY
# ============================================================
if (Test-SectionLoaded "Policy") {
    Write-Status "Policy" "Loaded from checkpoint - skipping" "Info"
} else {

Write-Status "Policy" "Collecting Azure Policy assignments and compliance..."

$allPolicyAssignments = [System.Collections.Generic.List[object]]::new()
$allPolicyCompliance = [System.Collections.Generic.List[object]]::new()
$allResourceLocks = [System.Collections.Generic.List[object]]::new()
$allCustomRoles = [System.Collections.Generic.List[object]]::new()
$allPolicyExemptions = [System.Collections.Generic.List[object]]::new()

foreach ($sub in $allSubs) {
    Set-AzContext -SubscriptionId $sub.Id -Force | Out-Null

    # Policy Assignments (dual-path: new Az module exposes properties directly, old module nests under .Properties)
    $assignments = Get-AzPolicyAssignment -ErrorAction SilentlyContinue
    foreach ($pa in $assignments) {
        # Detect Az module version - new modules expose properties directly on the object
        $paDisplayName      = if ($pa.DisplayName)                   { $pa.DisplayName }                   elseif ($pa.Properties.DisplayName)      { $pa.Properties.DisplayName }      else { $pa.Name }
        $paScope            = if ($pa.Scope)                         { $pa.Scope }                         elseif ($pa.Properties.Scope)             { $pa.Properties.Scope }             else { "" }
        $paDefId            = if ($pa.PolicyDefinitionId)            { $pa.PolicyDefinitionId }            elseif ($pa.Properties.PolicyDefinitionId){ $pa.Properties.PolicyDefinitionId } else { "" }
        $paEnforcement      = if ($null -ne $pa.EnforcementMode)     { $pa.EnforcementMode }               elseif ($pa.Properties.EnforcementMode)  { $pa.Properties.EnforcementMode }  else { "" }
        $paDescription      = if ($pa.Description)                   { $pa.Description }                   elseif ($pa.Properties.Description)      { $pa.Properties.Description }      else { "" }
        $paEffect           = if ($pa.Parameters.effect.value)       { $pa.Parameters.effect.value }       elseif ($pa.Properties.Parameters.effect.value) { $pa.Properties.Parameters.effect.value } else { "" }

        $null = $allPolicyAssignments.Add([PSCustomObject]@{
            AssignmentName   = $pa.Name
            DisplayName      = $paDisplayName
            Subscription     = $sub.Name
            Scope            = $paScope
            PolicyDefinition = if ($paDefId) { ($paDefId -split "/")[-1] } else { "" }
            IsInitiative     = $paDefId -match "policySetDefinitions"
            EnforcementMode  = $paEnforcement
            Effect           = $paEffect
            Description      = $paDescription
        })
    }
    
    # Policy Compliance Summary (handles both old and new Az module property structures)
    try {
        $compliance = Get-AzPolicyStateSummary -ErrorAction SilentlyContinue
        # New Az module: $compliance.PolicyAssignments; Old: $compliance.Results or different structure
        $cAssignments = if ($compliance.PolicyAssignments) { $compliance.PolicyAssignments } elseif ($compliance.policyAssignments) { $compliance.policyAssignments } else { @() }
        foreach ($c in $cAssignments) {
            $cPaId = if ($c.PolicyAssignmentId) { $c.PolicyAssignmentId } elseif ($c.policyAssignmentId) { $c.policyAssignmentId } else { "" }
            $cNonComplRes = if ($null -ne $c.Results.NonCompliantResources) { $c.Results.NonCompliantResources } elseif ($null -ne $c.results.nonCompliantResources) { $c.results.nonCompliantResources } else { 0 }
            $cNonComplPol = if ($null -ne $c.Results.NonCompliantPolicies) { $c.Results.NonCompliantPolicies } elseif ($null -ne $c.results.nonCompliantPolicies) { $c.results.nonCompliantPolicies } else { 0 }
            $cComplRes = 0
            try {
                $resDetails = if ($c.Results.ResourceDetails) { $c.Results.ResourceDetails } elseif ($c.results.resourceDetails) { $c.results.resourceDetails } else { @() }
                $cComplRes = ($resDetails | Where-Object { $_.ComplianceState -eq "compliant" -or $_.complianceState -eq "compliant" } | Measure-Object).Count
            } catch { }
            $null = $allPolicyCompliance.Add([PSCustomObject]@{
                Subscription         = $sub.Name
                PolicyAssignment     = if ($cPaId) { ($cPaId -split "/")[-1] } else { "" }
                NonCompliantResources = $cNonComplRes
                NonCompliantPolicies = $cNonComplPol
                CompliantResources   = $cComplRes
            })
        }
    }
    catch { }

    # Resource Locks (handles both old and new Az module property structures)
    $locks = Get-AzResourceLock -ErrorAction SilentlyContinue
    foreach ($lock in $locks) {
        $lockLevel = if ($lock.Properties.Level) { $lock.Properties.Level } elseif ($lock.Level) { $lock.Level } else { "Unknown" }
        $lockNotes = if ($lock.Properties.Notes) { $lock.Properties.Notes } elseif ($lock.Notes) { $lock.Notes } else { "" }
        $lockScope = if ($lock.ResourceId) { $lock.ResourceId } elseif ($lock.Properties.ResourceId) { $lock.Properties.ResourceId } else { "" }
        $null = $allResourceLocks.Add([PSCustomObject]@{
            LockName    = $lock.Name
            LockLevel   = $lockLevel
            Scope       = $lockScope
            ScopeType   = if ($lockScope -match "/resourceGroups/[^/]+$") { "ResourceGroup" } elseif ($lockScope -match "/providers/") { "Resource" } else { "Subscription" }
            Subscription = $sub.Name
            Notes       = $lockNotes
        })
    }

    # Custom RBAC Roles
    $customRoles = Get-AzRoleDefinition -Custom -ErrorAction SilentlyContinue
    foreach ($role in $customRoles) {
        $null = $allCustomRoles.Add([PSCustomObject]@{
            RoleName        = $role.Name
            Description     = $role.Description
            AssignableScopes = ($role.AssignableScopes) -join "; "
            PermissionCount = ($role.Actions + $role.DataActions).Count
        })
    }

    # Policy Exemptions
    $exemptions = Get-AzPolicyExemption -ErrorAction SilentlyContinue
    foreach ($ex in $exemptions) {
        $exDisplayName = if ($ex.DisplayName) { $ex.DisplayName } elseif ($ex.Properties.DisplayName) { $ex.Properties.DisplayName } else { $ex.Name }
        $exCategory = if ($ex.ExemptionCategory) { $ex.ExemptionCategory } elseif ($ex.Properties.ExemptionCategory) { $ex.Properties.ExemptionCategory } else { "N/A" }
        $exExpiry = if ($ex.ExpiresOn) { $ex.ExpiresOn.ToString("yyyy-MM-dd") } elseif ($ex.Properties.ExpiresOn) { $ex.Properties.ExpiresOn.ToString("yyyy-MM-dd") } else { "No Expiry" }
        $null = $allPolicyExemptions.Add([PSCustomObject]@{
            ExemptionName   = $exDisplayName
            Subscription    = $sub.Name
            Category        = $exCategory
            ExpiresOn       = $exExpiry
        })
    }
}

# Check for standard policies
$standardPolicies = $auditConfig.standardPolicies
$appliedPolicyNames = $allPolicyAssignments | ForEach-Object { $_.DisplayName }
$missingStandard = $standardPolicies | Where-Object { $appliedPolicyNames -notcontains $_ }
if ($missingStandard.Count -gt 0) {
    Add-Gap "Governance-Policy" "Standard policies not applied: $($missingStandard -join '; ')" "High"
}

$tenancyData.Sections["Policy"] = [ordered]@{
}
Save-Checkpoint
    PolicyAssignments   = $allPolicyAssignments
    PolicyCompliance    = $allPolicyCompliance
    MissingStdPolicies  = $missingStandard
    ResourceLocks       = $allResourceLocks
    CustomRoles         = $allCustomRoles
    PolicyExemptions    = $allPolicyExemptions
}

Write-Status "Policy" "Policy assignments: $($allPolicyAssignments.Count), Compliance records: $($allPolicyCompliance.Count)" "Success"

Add-Recommendation "Governance-Policy" "Apply Azure Policy initiatives at management group level for consistent governance." "High"
Add-Recommendation "Governance-Policy" "Apply CanNotDelete resource locks on production resource groups." "Medium"
Add-Recommendation "Governance-Policy" "Set expiry dates on all policy exemptions to ensure periodic review." "Low"

# ============================================================
# SECTION 10: RBAC
# ============================================================
if (Test-SectionLoaded "RBAC") {
    Write-Status "RBAC" "Loaded from checkpoint - skipping" "Info"
} else {

Write-Status "RBAC" "Collecting role assignments..."

$allRBAC = [System.Collections.Generic.List[object]]::new()

foreach ($sub in $allSubs) {
    Set-AzContext -SubscriptionId $sub.Id -Force | Out-Null
    
    $roleAssignments = Get-AzRoleAssignment -ErrorAction SilentlyContinue
    foreach ($ra in $roleAssignments) {
        $null = $allRBAC.Add([PSCustomObject]@{
            Subscription     = $sub.Name
            Scope            = $ra.Scope
            ScopeLevel       = if ($ra.Scope -match "resourceGroups") { "ResourceGroup" }
                              elseif ($ra.Scope -match "subscriptions" -and $ra.Scope -notmatch "resourceGroups") { "Subscription" }
                              elseif ($ra.Scope -match "managementGroups") { "ManagementGroup" }
                              else { "Resource" }
            RoleName         = $ra.RoleDefinitionName
            PrincipalName    = $ra.DisplayName
            PrincipalType    = $ra.ObjectType
            PrincipalId      = $ra.ObjectId
        })
    }
}

# Gap: Direct user assignments (should use groups)
$directUserAssignments = $allRBAC | Where-Object { $_.PrincipalType -eq "User" -and $_.ScopeLevel -in @("Subscription", "ManagementGroup") }
if ($directUserAssignments.Count -gt $auditConfig.directUserAssignmentThreshold) {
    Add-Gap "RBAC" "$($directUserAssignments.Count) direct user role assignments found at subscription/MG level. Best practice is to assign roles to groups." "Medium"
}

$tenancyData.Sections["RBAC"] = $allRBAC
Save-Checkpoint
}
Write-Status "RBAC" "Found $($allRBAC.Count) role assignments" "Success"

Add-Recommendation "RBAC" "Implement Privileged Identity Management (PIM) for all privileged roles." "High"
Add-Recommendation "RBAC" "Define custom RBAC roles following the principle of least privilege." "Medium"
Add-Recommendation "RBAC" "Review and remove stale role assignments quarterly." "Low"

# ============================================================
# SECTION 11: TAGS AUDIT
# ============================================================
if (Test-SectionLoaded "Tags") {
    Write-Status "Tags" "Loaded from checkpoint - skipping" "Info"
} else {

Write-Status "Tags" "Auditing resource tagging across all subscriptions..."

$tagAudit = [System.Collections.Generic.List[object]]::new()
$allResourceTags = [System.Collections.Generic.List[object]]::new()

foreach ($sub in $allSubs) {
    Set-AzContext -SubscriptionId $sub.Id -Force | Out-Null
    
    $resources = Get-AzResource -ErrorAction SilentlyContinue
    $totalResources = $resources.Count
    $taggedResources = ($resources | Where-Object { $_.Tags -and $_.Tags.Count -gt 0 }).Count
    $untaggedResources = $totalResources - $taggedResources
    
    $null = $tagAudit.Add([PSCustomObject]@{
        Subscription       = $sub.Name
        TotalResources     = $totalResources
        TaggedResources    = $taggedResources
        UntaggedResources  = $untaggedResources
        TaggingPercentage  = if ($totalResources -gt 0) { [math]::Round(($taggedResources / $totalResources) * 100, 1) } else { 0 }
    })
    
    # Collect all unique tag keys and sample values
    foreach ($resource in $resources | Where-Object { $_.Tags }) {
        foreach ($tag in $resource.Tags.GetEnumerator()) {
            $null = $allResourceTags.Add([PSCustomObject]@{
                Subscription  = $sub.Name
                ResourceName  = $resource.Name
                ResourceType  = $resource.ResourceType
                TagKey        = $tag.Key
                TagValue      = $tag.Value
            })
        }
    }
}

# Tag consistency analysis
$uniqueTagKeys = $allResourceTags | Select-Object -ExpandProperty TagKey -Unique | Sort-Object
$tagKeyUsage = $uniqueTagKeys | ForEach-Object {
    $key = $_
    $usage = ($allResourceTags | Where-Object { $_.TagKey -eq $key } | Measure-Object).Count
    [PSCustomObject]@{
        TagKey      = $key
        UsageCount  = $usage
        SampleValues = ($allResourceTags | Where-Object { $_.TagKey -eq $key } | Select-Object -ExpandProperty TagValue -Unique | Select-Object -First 5) -join "; "
    }
}

# Check for standard tag keys
$expectedTags = $auditConfig.expectedTags
$missingTags = $expectedTags | Where-Object { $_ -notin $uniqueTagKeys }
if ($missingTags.Count -gt 0) {
    Add-Gap "Governance-Tags" "Standard tags not found in any resources: $($missingTags -join ', ')" "Medium"
}

$lowTaggingSubs = $tagAudit | Where-Object { $_.TaggingPercentage -lt $auditConfig.lowTaggingThreshold }
if ($lowTaggingSubs.Count -gt 0) {
    Add-Gap "Governance-Tags" "Subscriptions with <50% tagged resources: $($lowTaggingSubs.Subscription -join ', ')" "High"
}

$tenancyData.Sections["Tags"] = [ordered]@{
}
Save-Checkpoint
    TagAuditSummary  = $tagAudit
    TagKeyUsage      = $tagKeyUsage
    MissingStdTags   = $missingTags
    UniqueTagKeys    = $uniqueTagKeys
}

Write-Status "Tags" "Unique tag keys: $($uniqueTagKeys.Count), Overall tagging: $(($tagAudit | Measure-Object -Property TaggingPercentage -Average).Average)%" "Success"

Add-Recommendation "Governance-Tags" "Use Azure Policy to enforce mandatory tags on resource creation." "Medium"
Add-Recommendation "Governance-Tags" "Implement automated tag remediation for untagged resources." "Low"

# ============================================================
# SECTION 12: MONITORING & LOGGING
# ============================================================
if (Test-SectionLoaded "Monitoring") {
    Write-Status "Monitoring" "Loaded from checkpoint - skipping" "Info"
} else {

Write-Status "Monitoring" "Collecting monitoring configuration..."

$allLogWorkspaces = [System.Collections.Generic.List[object]]::new()
$allAlertRules = [System.Collections.Generic.List[object]]::new()
$allActionGroups = [System.Collections.Generic.List[object]]::new()
$allAppInsights = [System.Collections.Generic.List[object]]::new()
$allDiagSettings = [System.Collections.Generic.List[object]]::new()
$allScheduledQueryRules = [System.Collections.Generic.List[object]]::new()
$allActivityLogAlerts = [System.Collections.Generic.List[object]]::new()

foreach ($sub in $allSubs) {
    Set-AzContext -SubscriptionId $sub.Id -Force | Out-Null

    # Log Analytics Workspaces
    $workspaces = Get-AzOperationalInsightsWorkspace -ErrorAction SilentlyContinue
    foreach ($ws in $workspaces) {
        $null = $allLogWorkspaces.Add([PSCustomObject]@{
            WorkspaceName    = $ws.Name
            ResourceGroup    = $ws.ResourceGroupName
            Subscription     = $sub.Name
            Location         = $ws.Location
            SKU              = $ws.Sku
            RetentionDays    = $ws.RetentionInDays
            DailyCapGB       = $ws.WorkspaceCapping.DailyQuotaGb
            CustomerId       = $ws.CustomerId
        })
    }
    
    # Alert Rules
    $alerts = Get-AzMetricAlertRuleV2 -ErrorAction SilentlyContinue
    foreach ($alert in $alerts) {
        $null = $allAlertRules.Add([PSCustomObject]@{
            AlertName        = $alert.Name
            ResourceGroup    = $alert.ResourceGroupName
            Subscription     = $sub.Name
            Severity         = $alert.Severity
            Enabled          = $alert.Enabled
            Description      = $alert.Description
            TargetResource   = ($alert.Scopes | ForEach-Object { ($_ -split "/")[-1] }) -join ", "
            ActionGroups     = ($alert.Actions | ForEach-Object { ($_.ActionGroupId -split "/")[-1] }) -join ", "
        })
    }
    
    # Action Groups
    $actionGroups = Get-AzActionGroup -ErrorAction SilentlyContinue
    foreach ($ag in $actionGroups) {
        $null = $allActionGroups.Add([PSCustomObject]@{
            ActionGroupName  = $ag.Name
            ResourceGroup    = $ag.ResourceGroupName
            Subscription     = $sub.Name
            EmailReceivers   = ($ag.EmailReceiver | ForEach-Object { $_.EmailAddress }) -join ", "
            SMSReceivers     = ($ag.SmsReceiver | ForEach-Object { $_.PhoneNumber }) -join ", "
            WebhookReceivers = ($ag.WebhookReceiver | ForEach-Object { $_.Name }) -join ", "
            Enabled          = $ag.Enabled
        })
    }

    # Application Insights
    try {
        $appInsights = Get-AzApplicationInsights -ErrorAction SilentlyContinue
        foreach ($ai in $appInsights) {
            $null = $allAppInsights.Add([PSCustomObject]@{
                Name            = $ai.Name
                ResourceGroup   = $ai.ResourceGroupName
                Subscription    = $sub.Name
                AppType         = $ai.ApplicationType
                WorkspaceId     = if ($ai.WorkspaceResourceId) { ($ai.WorkspaceResourceId -split "/")[-1] } else { "Classic" }
                RetentionDays   = $ai.RetentionInDays
                IngestionMode   = $ai.IngestionMode
            })
        }
    } catch { }

    # Diagnostic Settings (sample key resource types)
    try {
        $keyResources = @()
        $keyResources += Get-AzKeyVault -ErrorAction SilentlyContinue | ForEach-Object { $_.ResourceId }
        $keyResources += Get-AzNetworkSecurityGroup -ErrorAction SilentlyContinue | ForEach-Object { $_.Id }

        foreach ($resId in $keyResources) {
            $diagSettings = Get-AzDiagnosticSetting -ResourceId $resId -ErrorAction SilentlyContinue
            foreach ($ds in $diagSettings) {
                $null = $allDiagSettings.Add([PSCustomObject]@{
                    ResourceName     = ($resId -split "/")[-1]
                    ResourceType     = ($resId -split "/")[-2]
                    Subscription     = $sub.Name
                    DiagSettingName  = $ds.Name
                    WorkspaceId      = if ($ds.WorkspaceId) { ($ds.WorkspaceId -split "/")[-1] } else { "None" }
                    StorageAccount   = if ($ds.StorageAccountId) { ($ds.StorageAccountId -split "/")[-1] } else { "None" }
                    EventHub         = if ($ds.EventHubAuthorizationRuleId) { "Configured" } else { "None" }
                })
            }
        }
    } catch { }

    # Scheduled Query Rules (log-based alerts)
    try {
        $queryRules = Get-AzScheduledQueryRule -ErrorAction SilentlyContinue
        foreach ($qr in $queryRules) {
            $null = $allScheduledQueryRules.Add([PSCustomObject]@{
                Name            = $qr.Name
                ResourceGroup   = $qr.ResourceGroupName
                Subscription    = $sub.Name
                Severity        = $qr.Severity
                Enabled         = $qr.Enabled
                Description     = if ($qr.Description) { $qr.Description.Substring(0, [Math]::Min(100, $qr.Description.Length)) } else { "" }
            })
        }
    } catch { }

    # Activity Log Alerts
    try {
        $actAlerts = Get-AzActivityLogAlert -ErrorAction SilentlyContinue
        foreach ($ala in $actAlerts) {
            $null = $allActivityLogAlerts.Add([PSCustomObject]@{
                Name            = $ala.Name
                ResourceGroup   = $ala.ResourceGroupName
                Subscription    = $sub.Name
                Enabled         = $ala.Enabled
                Description     = if ($ala.Description) { $ala.Description } else { "" }
            })
        }
    } catch { }
}

# Gap: No Log Analytics workspace
if ($allLogWorkspaces.Count -eq 0) {
    Add-Gap "Monitoring" "No Log Analytics workspace found. Azure Monitor and diagnostics are not configured." "Critical"
}

# Gap: No alert rules
if ($allAlertRules.Count -eq 0) {
    Add-Gap "Monitoring" "No metric alert rules configured. No proactive alerting is in place." "High"
}

$tenancyData.Sections["Monitoring"] = [ordered]@{
}
Save-Checkpoint
    LogWorkspaces       = $allLogWorkspaces
    AlertRules          = $allAlertRules
    ActionGroups        = $allActionGroups
    AppInsights         = $allAppInsights
    DiagnosticSettings  = $allDiagSettings
    ScheduledQueryRules = $allScheduledQueryRules
    ActivityLogAlerts   = $allActivityLogAlerts
}

Write-Status "Monitoring" "Workspaces: $($allLogWorkspaces.Count), Alerts: $($allAlertRules.Count), Action Groups: $($allActionGroups.Count)" "Success"

Add-Recommendation "Monitoring" "Configure diagnostic settings on all key Azure resources." "High"
Add-Recommendation "Monitoring" "Set up Azure Service Health alerts for platform notifications." "Medium"
Add-Recommendation "Monitoring" "Deploy Application Insights for all web applications." "Medium"
Add-Recommendation "Monitoring" "Create log-based alert rules for critical failure scenarios." "Medium"

# ============================================================
# SECTION 13: COST MANAGEMENT
# ============================================================
if (Test-SectionLoaded "CostManagement") {
    Write-Status "CostMgmt" "Loaded from checkpoint - skipping" "Info"
} else {

Write-Status "CostMgmt" "Collecting cost management configuration..."

$allBudgets = [System.Collections.Generic.List[object]]::new()

foreach ($sub in $allSubs) {
    Set-AzContext -SubscriptionId $sub.Id -Force | Out-Null
    
    try {
        $budgets = Get-AzConsumptionBudget -ErrorAction SilentlyContinue
        foreach ($budget in $budgets) {
            $null = $allBudgets.Add([PSCustomObject]@{
                BudgetName       = $budget.Name
                Subscription     = $sub.Name
                Amount           = $budget.Amount
                TimeGrain        = $budget.TimeGrain
                StartDate        = $budget.TimePeriod.StartDate
                EndDate          = $budget.TimePeriod.EndDate
                CurrentSpend     = $budget.CurrentSpend.Amount
                Notifications    = ($budget.Notification.GetEnumerator() | ForEach-Object { "$($_.Key): $($_.Value.Threshold)%" }) -join "; "
            })
        }
    }
    catch { }
}

if ($allBudgets.Count -eq 0) {
    Add-Gap "CostManagement" "No Azure budgets are configured on any subscription. Cost overruns will not be detected." "High"
    Add-Recommendation "CostManagement" "Configure budgets on each subscription with alert thresholds at 80% and 100%."
}

$tenancyData.Sections["CostManagement"] = $allBudgets
Save-Checkpoint
}
Write-Status "CostMgmt" "Budgets: $($allBudgets.Count)" "Success"

Add-Recommendation "CostManagement" "Review Azure Advisor cost recommendations monthly." "Medium"
Add-Recommendation "CostManagement" "Evaluate Azure Reserved Instances for stable, long-running workloads." "Medium"
Add-Recommendation "CostManagement" "Remove unattached managed disks and unused public IPs to reduce cost." "Low"

# ============================================================
# SECTION 14: RESOURCE GROUPS & NAMING AUDIT
# ============================================================
if (Test-SectionLoaded "ResourceGroups") {
    Write-Status "ResourceGroups" "Loaded from checkpoint - skipping" "Info"
} else {

Write-Status "ResourceGroups" "Collecting resource groups and analysing naming conventions..."

$allResourceGroups = [System.Collections.Generic.List[object]]::new()

foreach ($sub in $allSubs) {
    Set-AzContext -SubscriptionId $sub.Id -Force | Out-Null
    
    $rgs = Get-AzResourceGroup
    foreach ($rg in $rgs) {
        $resourceCount = (Get-AzResource -ResourceGroupName $rg.ResourceGroupName -ErrorAction SilentlyContinue | Measure-Object).Count
        $null = $allResourceGroups.Add([PSCustomObject]@{
            ResourceGroupName = $rg.ResourceGroupName
            Subscription      = $sub.Name
            Location          = $rg.Location
            ResourceCount     = $resourceCount
            Tags              = if ($rg.Tags) { ($rg.Tags.GetEnumerator() | ForEach-Object { "$($_.Key)=$($_.Value)" }) -join "; " } else { "None" }
            ProvisioningState = $rg.ProvisioningState
        })
    }
}

# Naming convention analysis
$namingPatterns = @{
    "rg-*"   = ($allResourceGroups | Where-Object { $_.ResourceGroupName -match "^rg-" } | Measure-Object).Count
    "Other"  = ($allResourceGroups | Where-Object { $_.ResourceGroupName -notmatch "^rg-" } | Measure-Object).Count
}

if ($namingPatterns["Other"] -gt $namingPatterns["rg-*"]) {
    Add-Gap "Governance-Naming" "Most resource groups do not follow 'rg-' prefix naming convention. Inconsistent naming detected." "Medium"
}

$tenancyData.Sections["ResourceGroups"] = $allResourceGroups
Save-Checkpoint
}
Write-Status "ResourceGroups" "Found $($allResourceGroups.Count) resource groups" "Success"

# ============================================================
# SECTION: APPLICATION SERVICES (PaaS)
# ============================================================
if (Test-SectionLoaded "AppServices") {
    Write-Status "AppServices" "Loaded from checkpoint - skipping" "Info"
} else {

Write-Status "AppServices" "Collecting App Services..."

$allAppServicePlans = [System.Collections.Generic.List[object]]::new()
$allWebApps = [System.Collections.Generic.List[object]]::new()
$allFunctionApps = [System.Collections.Generic.List[object]]::new()

foreach ($sub in $allSubs) {
    Set-AzContext -SubscriptionId $sub.Id -Force | Out-Null

    try {
        # App Service Plans
        $plans = Get-AzAppServicePlan -ErrorAction SilentlyContinue
        foreach ($plan in $plans) {
            $null = $allAppServicePlans.Add([PSCustomObject]@{
                PlanName        = $plan.Name
                ResourceGroup   = $plan.ResourceGroup
                Subscription    = $sub.Name
                Location        = $plan.Location
                SKU             = $plan.Sku.Name
                Tier            = $plan.Sku.Tier
                OS              = if ($plan.Kind -match "linux") { "Linux" } else { "Windows" }
                WorkerCount     = $plan.Sku.Capacity
                ZoneRedundant   = if ($plan.ZoneRedundant) { $true } else { $false }
            })
        }

        # Web Apps
        $webApps = Get-AzWebApp -ErrorAction SilentlyContinue
        foreach ($wa in $webApps) {
            $siteConfig = $wa.SiteConfig
            $null = $allWebApps.Add([PSCustomObject]@{
                AppName         = $wa.Name
                ResourceGroup   = $wa.ResourceGroup
                Subscription    = $sub.Name
                AppServicePlan  = if ($wa.ServerFarmId) { ($wa.ServerFarmId -split "/")[-1] } else { "N/A" }
                State           = $wa.State
                HTTPSOnly       = $wa.HttpsOnly
                MinTLS          = if ($siteConfig) { $siteConfig.MinTlsVersion } else { "N/A" }
                Runtime         = if ($siteConfig -and $siteConfig.LinuxFxVersion) { $siteConfig.LinuxFxVersion } elseif ($siteConfig -and $siteConfig.NetFrameworkVersion) { "dotnet $($siteConfig.NetFrameworkVersion)" } else { "N/A" }
                VNetIntegration = if ($wa.VirtualNetworkSubnetId) { ($wa.VirtualNetworkSubnetId -split "/")[-1] } else { "None" }
                ManagedIdentity = if ($wa.Identity) { $wa.Identity.Type } else { "None" }
                CustomDomains   = if ($wa.HostNames) { ($wa.HostNames | Where-Object { $_ -notmatch "\.azurewebsites\.net$" }) -join ", " } else { "None" }
            })
        }

        # Function Apps
        $funcApps = Get-AzFunctionApp -ErrorAction SilentlyContinue
        foreach ($fa in $funcApps) {
            $null = $allFunctionApps.Add([PSCustomObject]@{
                AppName         = $fa.Name
                ResourceGroup   = $fa.ResourceGroup
                Subscription    = $sub.Name
                Runtime         = $fa.Runtime
                RuntimeVersion  = $fa.RuntimeVersion
                OSType          = $fa.OSType
                PlanType        = if ($fa.AppServicePlan) { ($fa.AppServicePlan -split "/")[-1] } else { "Consumption" }
                State           = $fa.State
                ManagedIdentity = if ($fa.IdentityType) { $fa.IdentityType } else { "None" }
            })
        }
    } catch {
        Write-Status "AppServices" "  Error in $($sub.Name): $($_.Exception.Message)" "Warning"
    }
}

$tenancyData.Sections["AppServices"] = [ordered]@{
}
Save-Checkpoint
    AppServicePlans = $allAppServicePlans
    WebApps         = $allWebApps
    FunctionApps    = $allFunctionApps
}

Write-Status "AppServices" "Plans: $($allAppServicePlans.Count), Web Apps: $($allWebApps.Count), Function Apps: $($allFunctionApps.Count)" "Success"

# ============================================================
# SECTION: DATABASES
# ============================================================
if (Test-SectionLoaded "Databases") {
    Write-Status "Databases" "Loaded from checkpoint - skipping" "Info"
} else {

Write-Status "Databases" "Collecting database services..."

$allSqlServers = [System.Collections.Generic.List[object]]::new()
$allSqlDatabases = [System.Collections.Generic.List[object]]::new()
$allCosmosAccounts = [System.Collections.Generic.List[object]]::new()

foreach ($sub in $allSubs) {
    Set-AzContext -SubscriptionId $sub.Id -Force | Out-Null

    try {
        # Azure SQL Servers
        $sqlServers = Get-AzSqlServer -ErrorAction SilentlyContinue
        foreach ($srv in $sqlServers) {
            $null = $allSqlServers.Add([PSCustomObject]@{
                ServerName      = $srv.ServerName
                ResourceGroup   = $srv.ResourceGroupName
                Subscription    = $sub.Name
                Location        = $srv.Location
                Version         = $srv.ServerVersion
                AdminLogin      = $srv.SqlAdministratorLogin
                PublicAccess     = $srv.PublicNetworkAccess
                MinTLS          = $srv.MinimalTlsVersion
            })

            # Databases per server
            $dbs = Get-AzSqlDatabase -ServerName $srv.ServerName -ResourceGroupName $srv.ResourceGroupName -ErrorAction SilentlyContinue
            foreach ($db in $dbs | Where-Object { $_.DatabaseName -ne "master" }) {
                $null = $allSqlDatabases.Add([PSCustomObject]@{
                    DatabaseName    = $db.DatabaseName
                    ServerName      = $srv.ServerName
                    Subscription    = $sub.Name
                    Edition         = $db.Edition
                    ServiceTier     = $db.CurrentServiceObjectiveName
                    MaxSizeGB       = [math]::Round($db.MaxSizeBytes / 1GB, 2)
                    ZoneRedundant   = $db.ZoneRedundant
                    Status          = $db.Status
                    EarliestRestore = if ($db.EarliestRestoreDate) { $db.EarliestRestoreDate.ToString("yyyy-MM-dd") } else { "N/A" }
                })
            }
        }
    } catch {
        Write-Status "Databases" "  SQL error in $($sub.Name): $($_.Exception.Message)" "Warning"
    }

    try {
        # Cosmos DB — iterate resource groups (cmdlet requires -ResourceGroupName)
        $rgs = Get-AzResourceGroup -ErrorAction SilentlyContinue
        foreach ($rg in $rgs) {
            $cosmosAccounts = Get-AzCosmosDBAccount -ResourceGroupName $rg.ResourceGroupName -ErrorAction SilentlyContinue
            foreach ($ca in $cosmosAccounts) {
                $null = $allCosmosAccounts.Add([PSCustomObject]@{
                    AccountName      = $ca.Name
                    ResourceGroup    = $ca.ResourceGroupName
                    Subscription     = $sub.Name
                    Location         = $ca.Location
                    Kind             = $ca.Kind
                    ConsistencyLevel = $ca.ConsistencyPolicy.DefaultConsistencyLevel
                    Locations        = ($ca.Locations | ForEach-Object { "$($_.LocationName)(Priority:$($_.FailoverPriority))" }) -join "; "
                    BackupPolicy     = if ($ca.BackupPolicy) { $ca.BackupPolicy.BackupType } else { "N/A" }
                    AutoFailover     = $ca.EnableAutomaticFailover
                    PublicAccess     = $ca.PublicNetworkAccess
                })
            }
        }
    } catch {
        Write-Status "Databases" "  Cosmos error in $($sub.Name): $($_.Exception.Message)" "Warning"
    }
}

$tenancyData.Sections["Databases"] = [ordered]@{
}
Save-Checkpoint
    SQLServers    = $allSqlServers
    SQLDatabases  = $allSqlDatabases
    CosmosDB      = $allCosmosAccounts
}

Write-Status "Databases" "SQL Servers: $($allSqlServers.Count), SQL DBs: $($allSqlDatabases.Count), Cosmos: $($allCosmosAccounts.Count)" "Success"

# ============================================================
# SECTION: CONTAINERS
# ============================================================
if (Test-SectionLoaded "Containers") {
    Write-Status "Containers" "Loaded from checkpoint - skipping" "Info"
} else {

Write-Status "Containers" "Collecting container services..."

$allAKSClusters = [System.Collections.Generic.List[object]]::new()
$allACRs = [System.Collections.Generic.List[object]]::new()

foreach ($sub in $allSubs) {
    Set-AzContext -SubscriptionId $sub.Id -Force | Out-Null

    try {
        $aksClusters = Get-AzAksCluster -ErrorAction SilentlyContinue
        foreach ($aks in $aksClusters) {
            $null = $allAKSClusters.Add([PSCustomObject]@{
                ClusterName     = $aks.Name
                ResourceGroup   = $aks.ResourceGroupName
                Subscription    = $sub.Name
                Location        = $aks.Location
                Version         = $aks.KubernetesVersion
                NodePools       = ($aks.AgentPoolProfiles | ForEach-Object { "$($_.Name)($($_.Count)x$($_.VmSize))" }) -join "; "
                NetworkPlugin   = $aks.NetworkProfile.NetworkPlugin
                RBACEnabled     = $aks.EnableRBAC
                PrivateCluster  = if ($aks.ApiServerAccessProfile) { $aks.ApiServerAccessProfile.EnablePrivateCluster } else { $false }
                ManagedIdentity = if ($aks.Identity) { $aks.Identity.Type } else { "None" }
            })
        }
    } catch {
        Write-Status "Containers" "  AKS error in $($sub.Name): $($_.Exception.Message)" "Warning"
    }

    try {
        $acrs = Get-AzContainerRegistry -ErrorAction SilentlyContinue
        foreach ($acr in $acrs) {
            $null = $allACRs.Add([PSCustomObject]@{
                RegistryName    = $acr.Name
                ResourceGroup   = $acr.ResourceGroupName
                Subscription    = $sub.Name
                Location        = $acr.Location
                SKU             = $acr.SkuName
                AdminEnabled    = $acr.AdminUserEnabled
                LoginServer     = $acr.LoginServer
                PublicAccess    = $acr.PublicNetworkAccess
            })
        }
    } catch {
        Write-Status "Containers" "  ACR error in $($sub.Name): $($_.Exception.Message)" "Warning"
    }
}

$tenancyData.Sections["Containers"] = [ordered]@{
}
Save-Checkpoint
    AKSClusters         = $allAKSClusters
    ContainerRegistries  = $allACRs
}

Write-Status "Containers" "AKS: $($allAKSClusters.Count), ACR: $($allACRs.Count)" "Success"

# ============================================================
# SECTION: AUTOMATION
# ============================================================
if (Test-SectionLoaded "Automation") {
    Write-Status "Automation" "Loaded from checkpoint - skipping" "Info"
} else {

Write-Status "Automation" "Collecting Automation Accounts..."

$allAutomationAccounts = [System.Collections.Generic.List[object]]::new()
$allRunbooks = [System.Collections.Generic.List[object]]::new()

foreach ($sub in $allSubs) {
    Set-AzContext -SubscriptionId $sub.Id -Force | Out-Null
    try {
        $accounts = Get-AzAutomationAccount -ErrorAction SilentlyContinue
        foreach ($aa in $accounts) {
            $null = $allAutomationAccounts.Add([PSCustomObject]@{
                AccountName     = $aa.AutomationAccountName
                ResourceGroup   = $aa.ResourceGroupName
                Subscription    = $sub.Name
                Location        = $aa.Location
                State           = $aa.State
                ManagedIdentity = if ($aa.Identity) { $aa.Identity.Type } else { "None" }
            })

            $rbs = Get-AzAutomationRunbook -AutomationAccountName $aa.AutomationAccountName -ResourceGroupName $aa.ResourceGroupName -ErrorAction SilentlyContinue
            foreach ($rb in $rbs) {
                $null = $allRunbooks.Add([PSCustomObject]@{
                    RunbookName      = $rb.Name
                    AutomationAccount = $aa.AutomationAccountName
                    Subscription     = $sub.Name
                    RunbookType      = $rb.RunbookType
                    State            = $rb.State
                    LastModified     = if ($rb.LastModifiedTime) { $rb.LastModifiedTime.ToString("yyyy-MM-dd") } else { "N/A" }
                })
            }
        }
    } catch {
        Write-Status "Automation" "  Error in $($sub.Name): $($_.Exception.Message)" "Warning"
    }
}

$tenancyData.Sections["Automation"] = [ordered]@{
}
Save-Checkpoint
    AutomationAccounts = $allAutomationAccounts
    Runbooks           = $allRunbooks
}

Write-Status "Automation" "Accounts: $($allAutomationAccounts.Count), Runbooks: $($allRunbooks.Count)" "Success"

# ============================================================
# SECTION: AZURE ADVISOR
# ============================================================
if (Test-SectionLoaded "Advisor") {
    Write-Status "Advisor" "Loaded from checkpoint - skipping" "Info"
} else {

Write-Status "Advisor" "Collecting Azure Advisor recommendations..."

$allAdvisorRecs = [System.Collections.Generic.List[object]]::new()

foreach ($sub in $allSubs) {
    Set-AzContext -SubscriptionId $sub.Id -Force | Out-Null
    try {
        $recs = Get-AzAdvisorRecommendation -ErrorAction SilentlyContinue
        foreach ($rec in $recs) {
            $null = $allAdvisorRecs.Add([PSCustomObject]@{
                Category        = $rec.Category
                Impact          = $rec.Impact
                Problem         = $rec.ShortDescription.Problem
                Solution        = $rec.ShortDescription.Solution
                ResourceType    = if ($rec.ImpactedField) { $rec.ImpactedField } else { "N/A" }
                ResourceName    = if ($rec.ImpactedValue) { $rec.ImpactedValue } else { "N/A" }
                Subscription    = $sub.Name
            })
        }
    } catch {
        Write-Status "Advisor" "  Error in $($sub.Name): $($_.Exception.Message)" "Warning"
    }
}

$tenancyData.Sections["Advisor"] = $allAdvisorRecs
Save-Checkpoint
}

Write-Status "Advisor" "Recommendations: $($allAdvisorRecs.Count)" "Success"

# ============================================================
# SECTION: RESOURCE GRAPH SUMMARY
# ============================================================
Write-Status "ResourceSummary" "Querying Resource Graph for complete resource inventory..."

$allResourceTypes = [System.Collections.Generic.List[object]]::new()
try {
    $graphQuery = "Resources | summarize Count=count() by type | order by Count desc"
    $graphResults = Search-AzGraph -Query $graphQuery -First 1000 -ErrorAction SilentlyContinue
    foreach ($r in $graphResults) {
        $null = $allResourceTypes.Add([PSCustomObject]@{
            ResourceType = $r.type
            Count        = $r.Count
        })
    }
} catch {
    Write-Status "ResourceSummary" "Resource Graph query failed: $($_.Exception.Message)" "Warning"
}

$tenancyData.Sections["ResourceSummary"] = $allResourceTypes
Save-Checkpoint

Write-Status "ResourceSummary" "Resource types found: $($allResourceTypes.Count)" "Success"

# ============================================================
# SECTION 15: OPERATIONAL COMPLIANCE FRAMEWORK CHECK
# ============================================================
Write-Status "Compliance" "Evaluating operational compliance framework..."

$complianceChecklist = [System.Collections.Generic.List[object]]::new()

# Check 1: Backup governance alignment
$null = $complianceChecklist.Add([PSCustomObject]@{
    Area        = "Backup & Recovery"
    Check       = "Backup policies aligned to retention/compliance requirements"
    Status      = if ($allBackupPolicies.Count -gt 0) { "Policies Found - Review Needed" } else { "NOT CONFIGURED" }
    Detail      = "Found $($allBackupPolicies.Count) backup policies across $($allVaults.Count) vaults. Retention alignment with governance-level compliance policies needs confirming."
    Action      = "Confirm backup retention periods align with customer data retention policies. Confirm roles and responsibilities for backup management."
    Owner       = "TBD"
    Priority    = "High"
})

# Check 2: Monitoring & Incident Management
$null = $complianceChecklist.Add([PSCustomObject]@{
    Area        = "Monitoring & Incidents"
    Check       = "Monitoring solution deployed and incident management integrated"
    Status      = if ($allAlertRules.Count -gt 0 -or $allLogWorkspaces.Count -gt 0) { "Partially Configured" } else { "NOT CONFIGURED" }
    Detail      = "Log Analytics: $($allLogWorkspaces.Count), Alert Rules: $($allAlertRules.Count), Action Groups: $($allActionGroups.Count). Third-party monitoring (e.g. LogicMonitor) integration status unknown - requires manual verification."
    Action      = "Document LogicMonitor integration, confirm alert routing, define incident severity matrix, confirm who receives alerts and SLA for response."
    Owner       = "TBD"
    Priority    = "High"
})

# Check 3: Cost Management
$null = $complianceChecklist.Add([PSCustomObject]@{
    Area        = "Cost Management"
    Check       = "Budgets, alerts, and cost governance in place"
    Status      = if ($allBudgets.Count -gt 0) { "Budgets Configured" } else { "NOT CONFIGURED" }
    Detail      = "Found $($allBudgets.Count) budgets. Reserved instances and savings plans need manual review via Azure portal."
    Action      = "Agree budget thresholds per subscription. Configure alerts at 80%/100%. Review Azure Advisor cost recommendations. Establish monthly cost review cadence."
    Owner       = "TBD"
    Priority    = "Medium"
})

# Check 4: Security posture
$null = $complianceChecklist.Add([PSCustomObject]@{
    Area        = "Security"
    Check       = "Defender for Cloud enabled with appropriate plans"
    Status      = if ($serversDefender.Count -gt 0) { "Defender Enabled" } else { "NOT ENABLED" }
    Detail      = "Defender for Servers: $(if ($serversDefender.Count -gt 0) { 'Enabled' } else { 'Disabled' }). Secure Score data collected."
    Action      = "Enable Defender for Servers (minimum Plan 1). Review and remediate Secure Score recommendations. Establish security review cadence."
    Owner       = "TBD"
    Priority    = "Critical"
})

# Check 5: Conditional Access
$null = $complianceChecklist.Add([PSCustomObject]@{
    Area        = "Identity & Access"
    Check       = "Conditional Access policies enforcing MFA and access controls"
    Status      = if ($IncludeEntra) {
        $caCount = ($tenancyData.Sections["Identity"].ConditionalAccess | Where-Object { $_.State -eq "enabled" }).Count
        if ($caCount -gt 0) { "$caCount policies enabled" } else { "NO POLICIES ENABLED" }
    } else { "NOT IN SCOPE (Entra excluded)" }
    Detail      = "Conditional Access audit requires Graph API. Run with -IncludeEntra to assess CA policies, MFA enforcement, and identity security posture."
    Action      = "Review and confirm CA policies cover: All users MFA, Admin MFA, Device compliance, Location-based access, Legacy auth block."
    Owner       = "TBD"
    Priority    = "Critical"
})

# Check 6: Tagging
$null = $complianceChecklist.Add([PSCustomObject]@{
    Area        = "Governance - Tagging"
    Check       = "Consistent tagging standard applied"
    Status      = if ($lowTaggingSubs.Count -eq 0 -and $tagAudit.Count -gt 0) { "Adequate" } else { "NEEDS IMPROVEMENT" }
    Detail      = "Average tagging: $(($tagAudit | Measure-Object -Property TaggingPercentage -Average).Average)%. Missing standard tags: $(if ($missingTags.Count -gt 0) { $missingTags -join ', ' } else { 'None' })."
    Action      = "Agree tagging taxonomy. Apply Azure Policy to enforce mandatory tags. Remediate existing untagged resources."
    Owner       = "TBD"
    Priority    = "Medium"
})

# Check 7: Policy governance
$null = $complianceChecklist.Add([PSCustomObject]@{
    Area        = "Governance - Policy"
    Check       = "Azure Policy initiatives applied and compliant"
    Status      = if ($allPolicyAssignments.Count -gt 0) { "$($allPolicyAssignments.Count) assignments" } else { "NO POLICIES" }
    Detail      = "Policy assignments: $($allPolicyAssignments.Count). Missing standard policies: $(if ($missingStandard.Count -gt 0) { $missingStandard -join '; ' } else { 'None' })."
    Action      = "Review policy compliance dashboard. Remediate non-compliant resources. Confirm all standard Landing Zone policies are applied."
    Owner       = "TBD"
    Priority    = "High"
})

# Check 8: DR
$null = $complianceChecklist.Add([PSCustomObject]@{
    Area        = "Disaster Recovery"
    Check       = "DR strategy defined and tested"
    Status      = "Requires Manual Review"
    Detail      = "Azure Site Recovery configuration cannot be fully assessed via script. Backup-based recovery is the minimum. Full DR requires ASR or equivalent."
    Action      = "Confirm if ASR is required. Define RPO/RTO targets. Establish DR testing schedule. Document recovery procedures."
    Owner       = "TBD"
    Priority    = "High"
})

# Check 9: Change Management
$null = $complianceChecklist.Add([PSCustomObject]@{
    Area        = "Operational - Change Mgmt"
    Check       = "Change management process agreed and documented"
    Status      = "Requires Manual Review"
    Detail      = "Cannot be assessed via script. Requires confirmation of change control process, approval workflows, and communication procedures."
    Action      = "Agree change management framework. Define change types (standard/normal/emergency). Establish CAB process. Document rollback procedures."
    Owner       = "TBD"
    Priority    = "Medium"
})

# Check 10: Naming convention
$null = $complianceChecklist.Add([PSCustomObject]@{
    Area        = "Governance - Naming"
    Check       = "Consistent naming convention applied to all resources"
    Status      = if ($namingPatterns["rg-*"] -gt $namingPatterns["Other"]) { "Mostly Consistent" } else { "INCONSISTENT" }
    Detail      = "Resource groups with 'rg-' prefix: $($namingPatterns['rg-*']), without: $($namingPatterns['Other'])."
    Action      = "Agree and document naming convention. Apply Azure Policy to enforce naming patterns where possible."
    Owner       = "TBD"
    Priority    = "Low"
})

$tenancyData.Sections["OperationalCompliance"] = $complianceChecklist
Save-Checkpoint

# ============================================================
# EXPORT: JSON
# ============================================================
Write-Status "Export" "Exporting master JSON..."

$jsonPath = Join-Path $OutputPath "$($CustomerName)_Azure_ASIS_Data.json"

# Convert to JSON-safe format (handle nested objects)
$jsonExport = $tenancyData | ConvertTo-Json -Depth 10 -Compress:$false
$jsonExport | Out-File -FilePath $jsonPath -Encoding UTF8
Write-Status "Export" "JSON saved: $jsonPath" "Success"

# ============================================================
# EXPORT: EXCEL (Multi-sheet workbook)
# ============================================================
Write-Status "Export" "Exporting Excel workbook..."

$excelPath = Join-Path $OutputPath "$($CustomerName)_Azure_ASIS_BuildDoc.xlsx"

$excelSheetsOK = 0
$excelSheetsFailed = 0

try {
    # Summary sheet
    $summaryData = [PSCustomObject]@{
        CustomerName     = $CustomerName
        CollectionDate   = $tenancyData.CollectionDate
        CollectedBy      = $tenancyData.CollectedBy
        TenantId         = $tenancyData.TenantId
        EntraIDScope     = $tenancyData.EntraStatus
        Subscriptions    = $allSubs.Count
        ManagementGroups = if ($tenancyData.Sections["ManagementGroups"]) { $tenancyData.Sections["ManagementGroups"].Count } else { "N/A" }
        VirtualNetworks  = $allVNets.Count
        VirtualMachines  = $allVMs.Count
        StorageAccounts  = $allStorageAccounts.Count
        RecoveryVaults   = $allVaults.Count
        PolicyAssignments = $allPolicyAssignments.Count
        GapsIdentified   = $tenancyData.Gaps.Count
        Recommendations  = $tenancyData.Recommendations.Count
    }
    try { $summaryData | Export-Excel -Path $excelPath -WorksheetName "Summary" -AutoSize -TableName "Summary" -Title "$CustomerName - Azure AS-IS Audit Summary" ; $excelSheetsOK++ } catch { Write-Status "Export" "Failed to export sheet 'Summary': $($_.Exception.Message)" "Warning"; $excelSheetsFailed++ }
    
    # Management Groups
    if ($tenancyData.Sections["ManagementGroups"]) {
        $tenancyData.Sections["ManagementGroups"] | Select-Object DisplayName, Name, ParentName, @{N="Subscriptions";E={$_.Subscriptions -join ", "}}, @{N="ChildMGs";E={$_.ChildMGs -join ", "}} |
            try { Export-Excel -Path $excelPath -WorksheetName "Mgmt Groups" -AutoSize -TableName "MgmtGroups" -Append ; $excelSheetsOK++ } catch { Write-Status "Export" "Failed to export sheet 'Mgmt Groups': $($_.Exception.Message)" "Warning"; $excelSheetsFailed++ }
    }
    
    # Subscriptions
    try { $subData | Export-Excel -Path $excelPath -WorksheetName "Subscriptions" -AutoSize -TableName "Subscriptions" -Append ; $excelSheetsOK++ } catch { Write-Status "Export" "Failed to export sheet 'Subscriptions': $($_.Exception.Message)" "Warning"; $excelSheetsFailed++ }
    
    # Identity (if collected)
    if ($tenancyData.Sections["Identity"]) {
        if ($tenancyData.Sections["Identity"].ConditionalAccess) {
            $tenancyData.Sections["Identity"].ConditionalAccess |
                try { Export-Excel -Path $excelPath -WorksheetName "Conditional Access" -AutoSize -TableName "ConditionalAccess" -Append ; $excelSheetsOK++ } catch { Write-Status "Export" "Failed to export sheet 'Conditional Access': $($_.Exception.Message)" "Warning"; $excelSheetsFailed++ }
        }
        if ($tenancyData.Sections["Identity"].DirectoryRoles) {
            $tenancyData.Sections["Identity"].DirectoryRoles |
                try { Export-Excel -Path $excelPath -WorksheetName "Directory Roles" -AutoSize -TableName "DirectoryRoles" -Append ; $excelSheetsOK++ } catch { Write-Status "Export" "Failed to export sheet 'Directory Roles': $($_.Exception.Message)" "Warning"; $excelSheetsFailed++ }
        }
        if ($tenancyData.Sections["Identity"].AppRegistrations -and $tenancyData.Sections["Identity"].AppRegistrations.Count -gt 0) {
            $tenancyData.Sections["Identity"].AppRegistrations |
                try { Export-Excel -Path $excelPath -WorksheetName "App Registrations" -AutoSize -TableName "AppRegistrations" -Append ; $excelSheetsOK++ } catch { Write-Status "Export" "Failed to export sheet 'App Registrations': $($_.Exception.Message)" "Warning"; $excelSheetsFailed++ }
        }
    }
    
    # Networking sheets
    try { $allVNets | Export-Excel -Path $excelPath -WorksheetName "Virtual Networks" -AutoSize -TableName "VNets" -Append ; $excelSheetsOK++ } catch { Write-Status "Export" "Failed to export sheet 'Virtual Networks': $($_.Exception.Message)" "Warning"; $excelSheetsFailed++ }
    try { $allSubnets | Export-Excel -Path $excelPath -WorksheetName "Subnets" -AutoSize -TableName "Subnets" -Append ; $excelSheetsOK++ } catch { Write-Status "Export" "Failed to export sheet 'Subnets': $($_.Exception.Message)" "Warning"; $excelSheetsFailed++ }
    if ($allPeerings.Count -gt 0) { try { $allPeerings | Export-Excel -Path $excelPath -WorksheetName "VNet Peerings" -AutoSize -TableName "Peerings" -Append } ; $excelSheetsOK++ } catch { Write-Status "Export" "Failed to export sheet 'VNet Peerings': $($_.Exception.Message)" "Warning"; $excelSheetsFailed++ }
    if ($allNSGs.Count -gt 0) { try { $allNSGs | Export-Excel -Path $excelPath -WorksheetName "NSG Rules" -AutoSize -TableName "NSGRules" -Append } ; $excelSheetsOK++ } catch { Write-Status "Export" "Failed to export sheet 'NSG Rules': $($_.Exception.Message)" "Warning"; $excelSheetsFailed++ }
    if ($allRouteTables.Count -gt 0) { try { $allRouteTables | Export-Excel -Path $excelPath -WorksheetName "Route Tables" -AutoSize -TableName "RouteTables" -Append } ; $excelSheetsOK++ } catch { Write-Status "Export" "Failed to export sheet 'Route Tables': $($_.Exception.Message)" "Warning"; $excelSheetsFailed++ }
    if ($allVPNGateways.Count -gt 0) { try { $allVPNGateways | Export-Excel -Path $excelPath -WorksheetName "VPN Gateways" -AutoSize -TableName "VPNGateways" -Append } ; $excelSheetsOK++ } catch { Write-Status "Export" "Failed to export sheet 'VPN Gateways': $($_.Exception.Message)" "Warning"; $excelSheetsFailed++ }
    if ($allLocalNetGateways.Count -gt 0) { try { $allLocalNetGateways | Export-Excel -Path $excelPath -WorksheetName "Local Net Gateways" -AutoSize -TableName "LNGs" -Append } ; $excelSheetsOK++ } catch { Write-Status "Export" "Failed to export sheet 'Local Net Gateways': $($_.Exception.Message)" "Warning"; $excelSheetsFailed++ }
    if ($allConnections.Count -gt 0) { try { $allConnections | Export-Excel -Path $excelPath -WorksheetName "VPN Connections" -AutoSize -TableName "VPNConns" -Append } ; $excelSheetsOK++ } catch { Write-Status "Export" "Failed to export sheet 'VPN Connections': $($_.Exception.Message)" "Warning"; $excelSheetsFailed++ }
    if ($allBastions.Count -gt 0) { try { $allBastions | Export-Excel -Path $excelPath -WorksheetName "Bastions" -AutoSize -TableName "Bastions" -Append } ; $excelSheetsOK++ } catch { Write-Status "Export" "Failed to export sheet 'Bastions': $($_.Exception.Message)" "Warning"; $excelSheetsFailed++ }
    if ($allPrivateDNSZones.Count -gt 0) { try { $allPrivateDNSZones | Export-Excel -Path $excelPath -WorksheetName "Private DNS" -AutoSize -TableName "PrivateDNS" -Append } ; $excelSheetsOK++ } catch { Write-Status "Export" "Failed to export sheet 'Private DNS': $($_.Exception.Message)" "Warning"; $excelSheetsFailed++ }
    try { $allPublicIPs | Export-Excel -Path $excelPath -WorksheetName "Public IPs" -AutoSize -TableName "PublicIPs" -Append ; $excelSheetsOK++ } catch { Write-Status "Export" "Failed to export sheet 'Public IPs': $($_.Exception.Message)" "Warning"; $excelSheetsFailed++ }
    if ($allFirewalls.Count -gt 0) { try { $allFirewalls | Export-Excel -Path $excelPath -WorksheetName "Firewalls" -AutoSize -TableName "Firewalls" -Append } ; $excelSheetsOK++ } catch { Write-Status "Export" "Failed to export sheet 'Firewalls': $($_.Exception.Message)" "Warning"; $excelSheetsFailed++ }
    if ($allAppGateways.Count -gt 0) { try { $allAppGateways | Export-Excel -Path $excelPath -WorksheetName "App Gateways" -AutoSize -TableName "AppGateways" -Append } ; $excelSheetsOK++ } catch { Write-Status "Export" "Failed to export sheet 'App Gateways': $($_.Exception.Message)" "Warning"; $excelSheetsFailed++ }
    if ($allLoadBalancers.Count -gt 0) { try { $allLoadBalancers | Export-Excel -Path $excelPath -WorksheetName "Load Balancers" -AutoSize -TableName "LoadBalancers" -Append } ; $excelSheetsOK++ } catch { Write-Status "Export" "Failed to export sheet 'Load Balancers': $($_.Exception.Message)" "Warning"; $excelSheetsFailed++ }
    if ($allNatGateways.Count -gt 0) { try { $allNatGateways | Export-Excel -Path $excelPath -WorksheetName "NAT Gateways" -AutoSize -TableName "NATGateways" -Append } ; $excelSheetsOK++ } catch { Write-Status "Export" "Failed to export sheet 'NAT Gateways': $($_.Exception.Message)" "Warning"; $excelSheetsFailed++ }
    if ($allExpressRoutes.Count -gt 0) { try { $allExpressRoutes | Export-Excel -Path $excelPath -WorksheetName "ExpressRoute" -AutoSize -TableName "ExpressRoutes" -Append } ; $excelSheetsOK++ } catch { Write-Status "Export" "Failed to export sheet 'ExpressRoute': $($_.Exception.Message)" "Warning"; $excelSheetsFailed++ }
    if ($allDdosPlans.Count -gt 0) { try { $allDdosPlans | Export-Excel -Path $excelPath -WorksheetName "DDoS Plans" -AutoSize -TableName "DdosPlans" -Append } ; $excelSheetsOK++ } catch { Write-Status "Export" "Failed to export sheet 'DDoS Plans': $($_.Exception.Message)" "Warning"; $excelSheetsFailed++ }
    if ($allPrivateEndpoints.Count -gt 0) { try { $allPrivateEndpoints | Export-Excel -Path $excelPath -WorksheetName "Private Endpoints" -AutoSize -TableName "PrivateEndpoints" -Append } ; $excelSheetsOK++ } catch { Write-Status "Export" "Failed to export sheet 'Private Endpoints': $($_.Exception.Message)" "Warning"; $excelSheetsFailed++ }
    if ($allNsgFlowLogs.Count -gt 0) { try { $allNsgFlowLogs | Export-Excel -Path $excelPath -WorksheetName "NSG Flow Logs" -AutoSize -TableName "NsgFlowLogs" -Append } ; $excelSheetsOK++ } catch { Write-Status "Export" "Failed to export sheet 'NSG Flow Logs': $($_.Exception.Message)" "Warning"; $excelSheetsFailed++ }

    # Compute
    try { $allVMs | Export-Excel -Path $excelPath -WorksheetName "Virtual Machines" -AutoSize -TableName "VMs" -Append ; $excelSheetsOK++ } catch { Write-Status "Export" "Failed to export sheet 'Virtual Machines': $($_.Exception.Message)" "Warning"; $excelSheetsFailed++ }
    if ($allVMExtensions.Count -gt 0) { try { $allVMExtensions | Export-Excel -Path $excelPath -WorksheetName "VM Extensions" -AutoSize -TableName "VMExtensions" -Append } ; $excelSheetsOK++ } catch { Write-Status "Export" "Failed to export sheet 'VM Extensions': $($_.Exception.Message)" "Warning"; $excelSheetsFailed++ }
    if ($allAVDHostPools.Count -gt 0) { try { $allAVDHostPools | Export-Excel -Path $excelPath -WorksheetName "AVD Host Pools" -AutoSize -TableName "AVDPools" -Append } ; $excelSheetsOK++ } catch { Write-Status "Export" "Failed to export sheet 'AVD Host Pools': $($_.Exception.Message)" "Warning"; $excelSheetsFailed++ }
    if ($allAVDSessionHosts.Count -gt 0) { try { $allAVDSessionHosts | Export-Excel -Path $excelPath -WorksheetName "AVD Session Hosts" -AutoSize -TableName "AVDHosts" -Append } ; $excelSheetsOK++ } catch { Write-Status "Export" "Failed to export sheet 'AVD Session Hosts': $($_.Exception.Message)" "Warning"; $excelSheetsFailed++ }
    if ($allManagedDisks.Count -gt 0) { try { $allManagedDisks | Export-Excel -Path $excelPath -WorksheetName "Managed Disks" -AutoSize -TableName "ManagedDisks" -Append } ; $excelSheetsOK++ } catch { Write-Status "Export" "Failed to export sheet 'Managed Disks': $($_.Exception.Message)" "Warning"; $excelSheetsFailed++ }
    if ($allSnapshots.Count -gt 0) { try { $allSnapshots | Export-Excel -Path $excelPath -WorksheetName "Snapshots" -AutoSize -TableName "Snapshots" -Append } ; $excelSheetsOK++ } catch { Write-Status "Export" "Failed to export sheet 'Snapshots': $($_.Exception.Message)" "Warning"; $excelSheetsFailed++ }
    if ($allVMSS.Count -gt 0) { try { $allVMSS | Export-Excel -Path $excelPath -WorksheetName "VM Scale Sets" -AutoSize -TableName "VMSS" -Append } ; $excelSheetsOK++ } catch { Write-Status "Export" "Failed to export sheet 'VM Scale Sets': $($_.Exception.Message)" "Warning"; $excelSheetsFailed++ }

    # Storage
    try { $allStorageAccounts | Export-Excel -Path $excelPath -WorksheetName "Storage Accounts" -AutoSize -TableName "StorageAccts" -Append ; $excelSheetsOK++ } catch { Write-Status "Export" "Failed to export sheet 'Storage Accounts': $($_.Exception.Message)" "Warning"; $excelSheetsFailed++ }
    
    # Backup
    if ($allVaults.Count -gt 0) { try { $allVaults | Export-Excel -Path $excelPath -WorksheetName "Recovery Vaults" -AutoSize -TableName "RSVaults" -Append } ; $excelSheetsOK++ } catch { Write-Status "Export" "Failed to export sheet 'Recovery Vaults': $($_.Exception.Message)" "Warning"; $excelSheetsFailed++ }
    if ($allBackupPolicies.Count -gt 0) { try { $allBackupPolicies | Export-Excel -Path $excelPath -WorksheetName "Backup Policies" -AutoSize -TableName "BackupPolicies" -Append } ; $excelSheetsOK++ } catch { Write-Status "Export" "Failed to export sheet 'Backup Policies': $($_.Exception.Message)" "Warning"; $excelSheetsFailed++ }
    if ($allBackupItems.Count -gt 0) { try { $allBackupItems | Export-Excel -Path $excelPath -WorksheetName "Backup Items" -AutoSize -TableName "BackupItems" -Append } ; $excelSheetsOK++ } catch { Write-Status "Export" "Failed to export sheet 'Backup Items': $($_.Exception.Message)" "Warning"; $excelSheetsFailed++ }
    if ($unbacked.Count -gt 0) {
        $unbacked | Select-Object VMName, Subscription, VMSize |
            try { Export-Excel -Path $excelPath -WorksheetName "Unprotected VMs" -AutoSize -TableName "UnprotectedVMs" -Append ; $excelSheetsOK++ } catch { Write-Status "Export" "Failed to export sheet 'Unprotected VMs': $($_.Exception.Message)" "Warning"; $excelSheetsFailed++ }
    }
    if ($allASRItems.Count -gt 0) { try { $allASRItems | Export-Excel -Path $excelPath -WorksheetName "ASR Replicated Items" -AutoSize -TableName "ASRItems" -Append } ; $excelSheetsOK++ } catch { Write-Status "Export" "Failed to export sheet 'ASR Replicated Items': $($_.Exception.Message)" "Warning"; $excelSheetsFailed++ }
    if ($allASRPolicies.Count -gt 0) { try { $allASRPolicies | Export-Excel -Path $excelPath -WorksheetName "ASR Policies" -AutoSize -TableName "ASRPolicies" -Append } ; $excelSheetsOK++ } catch { Write-Status "Export" "Failed to export sheet 'ASR Policies': $($_.Exception.Message)" "Warning"; $excelSheetsFailed++ }

    # Security
    if ($allDefenderPlans.Count -gt 0) { try { $allDefenderPlans | Export-Excel -Path $excelPath -WorksheetName "Defender Plans" -AutoSize -TableName "DefenderPlans" -Append } ; $excelSheetsOK++ } catch { Write-Status "Export" "Failed to export sheet 'Defender Plans': $($_.Exception.Message)" "Warning"; $excelSheetsFailed++ }
    if ($allSecureScores.Count -gt 0) { try { $allSecureScores | Export-Excel -Path $excelPath -WorksheetName "Secure Scores" -AutoSize -TableName "SecureScores" -Append } ; $excelSheetsOK++ } catch { Write-Status "Export" "Failed to export sheet 'Secure Scores': $($_.Exception.Message)" "Warning"; $excelSheetsFailed++ }
    if ($allKeyVaults.Count -gt 0) { try { $allKeyVaults | Export-Excel -Path $excelPath -WorksheetName "Key Vaults" -AutoSize -TableName "KeyVaults" -Append } ; $excelSheetsOK++ } catch { Write-Status "Export" "Failed to export sheet 'Key Vaults': $($_.Exception.Message)" "Warning"; $excelSheetsFailed++ }
    if ($allDefenderRecs.Count -gt 0) { try { $allDefenderRecs | Export-Excel -Path $excelPath -WorksheetName "Defender Recs" -AutoSize -TableName "DefenderRecs" -Append } ; $excelSheetsOK++ } catch { Write-Status "Export" "Failed to export sheet 'Defender Recs': $($_.Exception.Message)" "Warning"; $excelSheetsFailed++ }

    # RBAC
    try { $allRBAC | Export-Excel -Path $excelPath -WorksheetName "RBAC Assignments" -AutoSize -TableName "RBAC" -Append ; $excelSheetsOK++ } catch { Write-Status "Export" "Failed to export sheet 'RBAC Assignments': $($_.Exception.Message)" "Warning"; $excelSheetsFailed++ }
    
    # Policy
    if ($allPolicyAssignments.Count -gt 0) { try { $allPolicyAssignments | Export-Excel -Path $excelPath -WorksheetName "Policy Assignments" -AutoSize -TableName "PolicyAssign" -Append } ; $excelSheetsOK++ } catch { Write-Status "Export" "Failed to export sheet 'Policy Assignments': $($_.Exception.Message)" "Warning"; $excelSheetsFailed++ }
    if ($allPolicyCompliance.Count -gt 0) { try { $allPolicyCompliance | Export-Excel -Path $excelPath -WorksheetName "Policy Compliance" -AutoSize -TableName "PolicyCompliance" -Append } ; $excelSheetsOK++ } catch { Write-Status "Export" "Failed to export sheet 'Policy Compliance': $($_.Exception.Message)" "Warning"; $excelSheetsFailed++ }
    if ($allResourceLocks.Count -gt 0) { try { $allResourceLocks | Export-Excel -Path $excelPath -WorksheetName "Resource Locks" -AutoSize -TableName "ResourceLocks" -Append } ; $excelSheetsOK++ } catch { Write-Status "Export" "Failed to export sheet 'Resource Locks': $($_.Exception.Message)" "Warning"; $excelSheetsFailed++ }
    if ($allCustomRoles.Count -gt 0) { try { $allCustomRoles | Export-Excel -Path $excelPath -WorksheetName "Custom Roles" -AutoSize -TableName "CustomRoles" -Append } ; $excelSheetsOK++ } catch { Write-Status "Export" "Failed to export sheet 'Custom Roles': $($_.Exception.Message)" "Warning"; $excelSheetsFailed++ }
    if ($allPolicyExemptions.Count -gt 0) { try { $allPolicyExemptions | Export-Excel -Path $excelPath -WorksheetName "Policy Exemptions" -AutoSize -TableName "PolicyExemptions" -Append } ; $excelSheetsOK++ } catch { Write-Status "Export" "Failed to export sheet 'Policy Exemptions': $($_.Exception.Message)" "Warning"; $excelSheetsFailed++ }

    # Tags
    try { $tagAudit | Export-Excel -Path $excelPath -WorksheetName "Tag Audit Summary" -AutoSize -TableName "TagAudit" -Append ; $excelSheetsOK++ } catch { Write-Status "Export" "Failed to export sheet 'Tag Audit Summary': $($_.Exception.Message)" "Warning"; $excelSheetsFailed++ }
    if ($tagKeyUsage) { try { $tagKeyUsage | Export-Excel -Path $excelPath -WorksheetName "Tag Key Usage" -AutoSize -TableName "TagUsage" -Append } ; $excelSheetsOK++ } catch { Write-Status "Export" "Failed to export sheet 'Tag Key Usage': $($_.Exception.Message)" "Warning"; $excelSheetsFailed++ }
    
    # Monitoring
    if ($allLogWorkspaces.Count -gt 0) { try { $allLogWorkspaces | Export-Excel -Path $excelPath -WorksheetName "Log Analytics" -AutoSize -TableName "LogWorkspaces" -Append } ; $excelSheetsOK++ } catch { Write-Status "Export" "Failed to export sheet 'Log Analytics': $($_.Exception.Message)" "Warning"; $excelSheetsFailed++ }
    if ($allAlertRules.Count -gt 0) { try { $allAlertRules | Export-Excel -Path $excelPath -WorksheetName "Alert Rules" -AutoSize -TableName "AlertRules" -Append } ; $excelSheetsOK++ } catch { Write-Status "Export" "Failed to export sheet 'Alert Rules': $($_.Exception.Message)" "Warning"; $excelSheetsFailed++ }
    if ($allActionGroups.Count -gt 0) { try { $allActionGroups | Export-Excel -Path $excelPath -WorksheetName "Action Groups" -AutoSize -TableName "ActionGroups" -Append } ; $excelSheetsOK++ } catch { Write-Status "Export" "Failed to export sheet 'Action Groups': $($_.Exception.Message)" "Warning"; $excelSheetsFailed++ }
    if ($allAppInsights.Count -gt 0) { try { $allAppInsights | Export-Excel -Path $excelPath -WorksheetName "App Insights" -AutoSize -TableName "AppInsights" -Append } ; $excelSheetsOK++ } catch { Write-Status "Export" "Failed to export sheet 'App Insights': $($_.Exception.Message)" "Warning"; $excelSheetsFailed++ }
    if ($allDiagSettings.Count -gt 0) { try { $allDiagSettings | Export-Excel -Path $excelPath -WorksheetName "Diagnostic Settings" -AutoSize -TableName "DiagSettings" -Append } ; $excelSheetsOK++ } catch { Write-Status "Export" "Failed to export sheet 'Diagnostic Settings': $($_.Exception.Message)" "Warning"; $excelSheetsFailed++ }
    if ($allScheduledQueryRules.Count -gt 0) { try { $allScheduledQueryRules | Export-Excel -Path $excelPath -WorksheetName "Scheduled Query Rules" -AutoSize -TableName "ScheduledQueryRules" -Append } ; $excelSheetsOK++ } catch { Write-Status "Export" "Failed to export sheet 'Scheduled Query Rules': $($_.Exception.Message)" "Warning"; $excelSheetsFailed++ }
    if ($allActivityLogAlerts.Count -gt 0) { try { $allActivityLogAlerts | Export-Excel -Path $excelPath -WorksheetName "Activity Log Alerts" -AutoSize -TableName "ActivityLogAlerts" -Append } ; $excelSheetsOK++ } catch { Write-Status "Export" "Failed to export sheet 'Activity Log Alerts': $($_.Exception.Message)" "Warning"; $excelSheetsFailed++ }

    # Cost
    if ($allBudgets.Count -gt 0) { try { $allBudgets | Export-Excel -Path $excelPath -WorksheetName "Budgets" -AutoSize -TableName "Budgets" -Append } ; $excelSheetsOK++ } catch { Write-Status "Export" "Failed to export sheet 'Budgets': $($_.Exception.Message)" "Warning"; $excelSheetsFailed++ }
    
    # Resource Groups
    try { $allResourceGroups | Export-Excel -Path $excelPath -WorksheetName "Resource Groups" -AutoSize -TableName "ResourceGroups" -Append ; $excelSheetsOK++ } catch { Write-Status "Export" "Failed to export sheet 'Resource Groups': $($_.Exception.Message)" "Warning"; $excelSheetsFailed++ }

    # App Services
    if ($allAppServicePlans.Count -gt 0) { try { $allAppServicePlans | Export-Excel -Path $excelPath -WorksheetName "App Service Plans" -AutoSize -TableName "AppServicePlans" -Append } ; $excelSheetsOK++ } catch { Write-Status "Export" "Failed to export sheet 'App Service Plans': $($_.Exception.Message)" "Warning"; $excelSheetsFailed++ }
    if ($allWebApps.Count -gt 0) { try { $allWebApps | Export-Excel -Path $excelPath -WorksheetName "Web Apps" -AutoSize -TableName "WebApps" -Append } ; $excelSheetsOK++ } catch { Write-Status "Export" "Failed to export sheet 'Web Apps': $($_.Exception.Message)" "Warning"; $excelSheetsFailed++ }
    if ($allFunctionApps.Count -gt 0) { try { $allFunctionApps | Export-Excel -Path $excelPath -WorksheetName "Function Apps" -AutoSize -TableName "FunctionApps" -Append } ; $excelSheetsOK++ } catch { Write-Status "Export" "Failed to export sheet 'Function Apps': $($_.Exception.Message)" "Warning"; $excelSheetsFailed++ }

    # Databases
    if ($allSqlServers.Count -gt 0) { try { $allSqlServers | Export-Excel -Path $excelPath -WorksheetName "SQL Servers" -AutoSize -TableName "SqlServers" -Append } ; $excelSheetsOK++ } catch { Write-Status "Export" "Failed to export sheet 'SQL Servers': $($_.Exception.Message)" "Warning"; $excelSheetsFailed++ }
    if ($allSqlDatabases.Count -gt 0) { try { $allSqlDatabases | Export-Excel -Path $excelPath -WorksheetName "SQL Databases" -AutoSize -TableName "SqlDatabases" -Append } ; $excelSheetsOK++ } catch { Write-Status "Export" "Failed to export sheet 'SQL Databases': $($_.Exception.Message)" "Warning"; $excelSheetsFailed++ }
    if ($allCosmosAccounts.Count -gt 0) { try { $allCosmosAccounts | Export-Excel -Path $excelPath -WorksheetName "Cosmos DB" -AutoSize -TableName "CosmosAccounts" -Append } ; $excelSheetsOK++ } catch { Write-Status "Export" "Failed to export sheet 'Cosmos DB': $($_.Exception.Message)" "Warning"; $excelSheetsFailed++ }

    # Containers
    if ($allAKSClusters.Count -gt 0) { try { $allAKSClusters | Export-Excel -Path $excelPath -WorksheetName "AKS Clusters" -AutoSize -TableName "AKSClusters" -Append } ; $excelSheetsOK++ } catch { Write-Status "Export" "Failed to export sheet 'AKS Clusters': $($_.Exception.Message)" "Warning"; $excelSheetsFailed++ }
    if ($allACRs.Count -gt 0) { try { $allACRs | Export-Excel -Path $excelPath -WorksheetName "Container Registries" -AutoSize -TableName "ACRs" -Append } ; $excelSheetsOK++ } catch { Write-Status "Export" "Failed to export sheet 'Container Registries': $($_.Exception.Message)" "Warning"; $excelSheetsFailed++ }

    # Automation
    if ($allAutomationAccounts.Count -gt 0) { try { $allAutomationAccounts | Export-Excel -Path $excelPath -WorksheetName "Automation Accounts" -AutoSize -TableName "AutomationAccounts" -Append } ; $excelSheetsOK++ } catch { Write-Status "Export" "Failed to export sheet 'Automation Accounts': $($_.Exception.Message)" "Warning"; $excelSheetsFailed++ }
    if ($allRunbooks.Count -gt 0) { try { $allRunbooks | Export-Excel -Path $excelPath -WorksheetName "Runbooks" -AutoSize -TableName "Runbooks" -Append } ; $excelSheetsOK++ } catch { Write-Status "Export" "Failed to export sheet 'Runbooks': $($_.Exception.Message)" "Warning"; $excelSheetsFailed++ }

    # Advisor
    if ($allAdvisorRecs.Count -gt 0) { try { $allAdvisorRecs | Export-Excel -Path $excelPath -WorksheetName "Advisor Recs" -AutoSize -TableName "AdvisorRecs" -Append } ; $excelSheetsOK++ } catch { Write-Status "Export" "Failed to export sheet 'Advisor Recs': $($_.Exception.Message)" "Warning"; $excelSheetsFailed++ }

    # Resource Graph Summary
    if ($allResourceTypes.Count -gt 0) { try { $allResourceTypes | Export-Excel -Path $excelPath -WorksheetName "Resource Types" -AutoSize -TableName "ResourceTypes" -Append } ; $excelSheetsOK++ } catch { Write-Status "Export" "Failed to export sheet 'Resource Types': $($_.Exception.Message)" "Warning"; $excelSheetsFailed++ }

    # Identity Extensions (if collected)
    if ($tenancyData.Sections["Identity"]) {
        if ($namedLocations.Count -gt 0) { try { $namedLocations | Export-Excel -Path $excelPath -WorksheetName "Named Locations" -AutoSize -TableName "NamedLocations" -Append } ; $excelSheetsOK++ } catch { Write-Status "Export" "Failed to export sheet 'Named Locations': $($_.Exception.Message)" "Warning"; $excelSheetsFailed++ }
        if ($licenseSummary.Count -gt 0) { try { $licenseSummary | Export-Excel -Path $excelPath -WorksheetName "Licenses" -AutoSize -TableName "Licenses" -Append } ; $excelSheetsOK++ } catch { Write-Status "Export" "Failed to export sheet 'Licenses': $($_.Exception.Message)" "Warning"; $excelSheetsFailed++ }
    }

    # ====== GAP ANALYSIS & COMPLIANCE (Critical sheets) ======
    
    # Gaps
    if ($tenancyData.Gaps.Count -gt 0) {
        try { $tenancyData.Gaps | Export-Excel -Path $excelPath -WorksheetName "GAP ANALYSIS" -AutoSize -TableName "Gaps" -Append ` ; $excelSheetsOK++ } catch { Write-Status "Export" "Failed to export sheet 'GAP ANALYSIS': $($_.Exception.Message)" "Warning"; $excelSheetsFailed++ }
            -ConditionalText $(
                New-ConditionalText -Text "Critical" -BackgroundColor Red -ConditionalTextColor White
                New-ConditionalText -Text "High" -BackgroundColor Orange -ConditionalTextColor White
                New-ConditionalText -Text "Medium" -BackgroundColor Yellow
            )
    }
    
    # Recommendations
    if ($tenancyData.Recommendations.Count -gt 0) {
        try { $tenancyData.Recommendations | Export-Excel -Path $excelPath -WorksheetName "RECOMMENDATIONS" -AutoSize -TableName "Recommendations" -Append ; $excelSheetsOK++ } catch { Write-Status "Export" "Failed to export sheet 'RECOMMENDATIONS': $($_.Exception.Message)" "Warning"; $excelSheetsFailed++ }
    }
    
    # Operational Compliance Framework
    try { $complianceChecklist | Export-Excel -Path $excelPath -WorksheetName "OPERATIONAL COMPLIANCE" -AutoSize -TableName "OpCompliance" -Append ` ; $excelSheetsOK++ } catch { Write-Status "Export" "Failed to export sheet 'OPERATIONAL COMPLIANCE': $($_.Exception.Message)" "Warning"; $excelSheetsFailed++ }
        -ConditionalText $(
            New-ConditionalText -Text "NOT CONFIGURED" -BackgroundColor Red -ConditionalTextColor White
            New-ConditionalText -Text "NOT ENABLED" -BackgroundColor Red -ConditionalTextColor White
            New-ConditionalText -Text "NOT ASSESSED" -BackgroundColor Orange -ConditionalTextColor White
            New-ConditionalText -Text "NEEDS IMPROVEMENT" -BackgroundColor Yellow
            New-ConditionalText -Text "INCONSISTENT" -BackgroundColor Yellow
            New-ConditionalText -Text "Critical" -BackgroundColor Red -ConditionalTextColor White
        )
    
    Write-Status "Export" "Excel saved: $excelPath ($excelSheetsOK sheets OK, $excelSheetsFailed failed)" "Success"
}
catch {
    Write-Status "Export" "Excel export failed: $($_.Exception.Message). Ensure ImportExcel module is installed." "Error"
    Write-Status "Export" "Install with: Install-Module ImportExcel -Scope CurrentUser" "Warning"
}

# ============================================================
# EXPORT: CONSOLE SUMMARY
# ============================================================

Write-Host ""
Write-Host "============================================================" -ForegroundColor Cyan
Write-Host "  AZURE TENANCY AUDIT COMPLETE - $CustomerName" -ForegroundColor Cyan
Write-Host "============================================================" -ForegroundColor Cyan
Write-Host ""
Write-Host "COLLECTION SUMMARY:" -ForegroundColor White
Write-Host "  Tenant ID:           $($tenancyData.TenantId)"
Write-Host "  Subscriptions:       $($allSubs.Count)"
Write-Host "  Management Groups:   $(if ($tenancyData.Sections['ManagementGroups']) { $tenancyData.Sections['ManagementGroups'].Count } else { 'N/A' })"
Write-Host "  Virtual Networks:    $($allVNets.Count)"
Write-Host "  Subnets:             $($allSubnets.Count)"
Write-Host "  VMs:                 $($allVMs.Count)"
Write-Host "  Storage Accounts:    $($allStorageAccounts.Count)"
Write-Host "  Recovery Vaults:     $($allVaults.Count)"
Write-Host "  Policy Assignments:  $($allPolicyAssignments.Count)"
Write-Host "  RBAC Assignments:    $($allRBAC.Count)"
Write-Host ""

if ($tenancyData.Gaps.Count -gt 0) {
    Write-Host "GAPS IDENTIFIED: $($tenancyData.Gaps.Count)" -ForegroundColor Red
    $criticalGaps = $tenancyData.Gaps | Where-Object { $_.Severity -eq "Critical" }
    $highGaps = $tenancyData.Gaps | Where-Object { $_.Severity -eq "High" }
    $medGaps = $tenancyData.Gaps | Where-Object { $_.Severity -eq "Medium" }
    Write-Host "  Critical: $($criticalGaps.Count)" -ForegroundColor Red
    Write-Host "  High:     $($highGaps.Count)" -ForegroundColor Yellow
    Write-Host "  Medium:   $($medGaps.Count)" -ForegroundColor White
    Write-Host ""
    
    if ($criticalGaps.Count -gt 0) {
        Write-Host "CRITICAL GAPS:" -ForegroundColor Red
        foreach ($gap in $criticalGaps) {
            Write-Host "  [!] [$($gap.Section)] $($gap.Description)" -ForegroundColor Red
        }
        Write-Host ""
    }
}

if ($tenancyData.Recommendations.Count -gt 0) {
    Write-Host "RECOMMENDATIONS: $($tenancyData.Recommendations.Count)" -ForegroundColor Yellow
    foreach ($rec in $tenancyData.Recommendations | Select-Object -First 5) {
        Write-Host "  [-] [$($rec.Section)] $($rec.Description)" -ForegroundColor Yellow
    }
    if ($tenancyData.Recommendations.Count -gt 5) {
        Write-Host "  ... and $($tenancyData.Recommendations.Count - 5) more (see Excel)" -ForegroundColor Gray
    }
    Write-Host ""
}

Write-Host "OUTPUT FILES:" -ForegroundColor Green
Write-Host "  JSON:  $jsonPath"
Write-Host "  Excel: $excelPath"
Write-Host ""
Write-Host "NEXT STEPS:" -ForegroundColor Cyan
Write-Host "  1. Review the GAP ANALYSIS and OPERATIONAL COMPLIANCE sheets in Excel"
Write-Host "  2. Use the JSON data to auto-populate the Word template"
Write-Host "  3. Manually verify: LogicMonitor integration, SMTP config, on-prem connectivity"
Write-Host "  4. Confirm backup retention aligns with governance/compliance policies"
Write-Host "  5. Agree operational compliance framework with customer"
Write-Host ""
Write-Host "============================================================" -ForegroundColor Cyan

# Clean up checkpoint file on successful completion
$checkpointPath = Join-Path $OutputPath "_checkpoint.json"
if (Test-Path $checkpointPath) {
    Remove-Item $checkpointPath -Force -ErrorAction SilentlyContinue
    Write-Host "[OK] Checkpoint file removed (audit completed successfully)" -ForegroundColor Green
}
