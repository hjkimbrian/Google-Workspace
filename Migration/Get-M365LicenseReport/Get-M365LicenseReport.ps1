<#
.SYNOPSIS
    Exports a Microsoft 365 user and license inventory to help scope a Google Workspace migration.

.DESCRIPTION
    Connects to Microsoft Graph using certificate-based (app-only) authentication and
    produces a CSV with every user and their assigned M365 licenses, account status, and
    relevant attributes. This helps determine:

      - How many users need Google Workspace licenses
      - Which users are active vs. disabled (skip disabled accounts)
      - Which products are in use (Exchange, Teams, OneDrive, SharePoint, etc.)
      - Guest accounts and service accounts to exclude from migration

    Also produces a license-summary report showing total assigned counts per SKU.

.PARAMETER ParamFile
    Path to MigrationConnectionParams.json from Setup-PnPCertAuth.ps1.

.PARAMETER TenantName
    M365 tenant name. Not needed if -ParamFile is used.

.PARAMETER ClientId
    Azure AD Application client ID. Not needed if -ParamFile is used.

.PARAMETER CertificatePath
    Path to the .pfx certificate.

.PARAMETER CertificatePassword
    SecureString password for the certificate.

.PARAMETER TenantId
    Azure AD Tenant ID (GUID). Recommended for cert-auth to Graph.

.PARAMETER OutputCsv
    Path for the user+license CSV. Defaults to .\M365-LicenseReport-<timestamp>.csv

.PARAMETER OutputSummaryCsv
    Path for the license summary CSV. Defaults to .\M365-LicenseSummary-<timestamp>.csv

.PARAMETER IncludeGuests
    Include guest (B2B) accounts in the report.

.PARAMETER IncludeDisabled
    Include disabled/blocked accounts in the report.

.EXAMPLE
    .\Get-M365LicenseReport.ps1 -ParamFile ".\MigrationConnectionParams.json"

.EXAMPLE
    .\Get-M365LicenseReport.ps1 `
        -TenantName "contoso" `
        -TenantId "xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx" `
        -ClientId  "yyyyyyyy-yyyy-yyyy-yyyy-yyyyyyyyyyyy" `
        -CertificatePath ".\PnP-Migration-App.pfx"

.NOTES
    Requires: Microsoft.Graph.Users and Microsoft.Graph.Identity.DirectoryManagement modules
    OR the all-in-one Microsoft.Graph module.

    App Registration permissions required (application, not delegated):
      - User.Read.All
      - Directory.Read.All
      - Organization.Read.All
#>

[CmdletBinding()]
param(
    [string]$ParamFile,

    [string]$TenantName,
    [string]$TenantId,
    [string]$ClientId,
    [string]$CertificatePath,
    [SecureString]$CertificatePassword,

    [string]$OutputCsv,
    [string]$OutputSummaryCsv,

    [switch]$IncludeGuests,
    [switch]$IncludeDisabled
)

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

#region ── Load Parameters ───────────────────────────────────────────────────

if ($ParamFile) {
    if (-not (Test-Path $ParamFile)) { throw "ParamFile not found: $ParamFile" }
    $p = Get-Content $ParamFile -Raw | ConvertFrom-Json
    if (-not $TenantName)      { $TenantName      = $p.TenantName }
    if (-not $ClientId)        { $ClientId        = $p.ClientId }
    if (-not $CertificatePath) { $CertificatePath = $p.CertificatePath }
    if (-not $TenantId)        { $TenantId        = $p.TenantId }
}

foreach ($var in @('ClientId','CertificatePath')) {
    if (-not (Get-Variable $var -ValueOnly -ErrorAction SilentlyContinue)) {
        throw "Missing required parameter: -$var (or supply -ParamFile)"
    }
}

if (-not $TenantId -and $TenantName) {
    $TenantId = if ($TenantName -match '\.') { $TenantName } else { "$TenantName.onmicrosoft.com" }
}

if (-not $TenantId) {
    throw "Provide -TenantId (GUID) or -TenantName, or set TenantId in the ParamFile."
}

if (-not $OutputCsv) {
    $ts = Get-Date -Format "yyyyMMdd-HHmmss"
    $OutputCsv = Join-Path $PSScriptRoot "M365-LicenseReport-$ts.csv"
}
if (-not $OutputSummaryCsv) {
    $ts2 = if ($ts) { $ts } else { Get-Date -Format "yyyyMMdd-HHmmss" }
    $OutputSummaryCsv = Join-Path $PSScriptRoot "M365-LicenseSummary-$ts2.csv"
}

#endregion

#region ── Ensure Graph Module ───────────────────────────────────────────────

Write-Host "[INFO] Checking Microsoft Graph PowerShell module..." -ForegroundColor Cyan

if (-not (Get-Module -ListAvailable -Name "Microsoft.Graph.Users")) {
    Write-Host "       Installing Microsoft.Graph (this may take a few minutes)..." -ForegroundColor Yellow
    Install-Module Microsoft.Graph -Scope CurrentUser -Force -AllowClobber
}

Import-Module Microsoft.Graph.Authentication -ErrorAction SilentlyContinue
Import-Module Microsoft.Graph.Users          -ErrorAction SilentlyContinue
Import-Module Microsoft.Graph.Identity.DirectoryManagement -ErrorAction SilentlyContinue

Write-Host "[OK]   Graph module ready" -ForegroundColor Green

#endregion

#region ── Connect to Graph ──────────────────────────────────────────────────

Write-Host "[INFO] Connecting to Microsoft Graph (app-only)..." -ForegroundColor Cyan

$cert = if ($CertificatePassword) {
    [System.Security.Cryptography.X509Certificates.X509Certificate2]::new(
        $CertificatePath,
        $CertificatePassword
    )
} else {
    [System.Security.Cryptography.X509Certificates.X509Certificate2]::new($CertificatePath)
}

Connect-MgGraph `
    -ClientId   $ClientId `
    -TenantId   $TenantId `
    -Certificate $cert `
    -NoWelcome

Write-Host "[OK]   Connected to Microsoft Graph" -ForegroundColor Green

#endregion

#region ── Friendly SKU Name Map ─────────────────────────────────────────────

# Map of common SkuPartNumber → human-readable name
$skuFriendlyNames = @{
    "SPE_E3"                      = "Microsoft 365 E3"
    "SPE_E5"                      = "Microsoft 365 E5"
    "ENTERPRISEPREMIUM"           = "Office 365 E3"
    "ENTERPRISEPREMIUM_NOPSTNCONF" = "Office 365 E3 No PSTN"
    "ENTERPRISEPACK"              = "Office 365 E3"
    "ENTERPRISEWITHSCAL"          = "Office 365 E4"
    "DEVELOPERPACK_E5"            = "Microsoft 365 E5 Developer"
    "O365_BUSINESS_PREMIUM"       = "Microsoft 365 Business Premium"
    "O365_BUSINESS_ESSENTIALS"    = "Microsoft 365 Business Basic"
    "O365_BUSINESS"               = "Microsoft 365 Apps for Business"
    "BUSINESS_VOICE_MED2_TELCO"   = "Microsoft 365 Business Voice"
    "EXCHANGESTANDARD"            = "Exchange Online Plan 1"
    "EXCHANGEENTERPRISE"          = "Exchange Online Plan 2"
    "TEAMS_ESSENTIALS"            = "Microsoft Teams Essentials"
    "TEAMS_EXPLORATORY"           = "Microsoft Teams Exploratory"
    "FLOW_FREE"                   = "Power Automate Free"
    "POWER_BI_STANDARD"           = "Power BI Free"
    "POWER_BI_PRO"                = "Power BI Pro"
    "PROJECTPREMIUM"              = "Project Plan 5"
    "VISIOCLIENT"                 = "Visio Plan 2"
    "AAD_PREMIUM"                 = "Azure AD Premium P1"
    "AAD_PREMIUM_P2"              = "Azure AD Premium P2"
}

function Get-FriendlySkuName([string]$SkuPartNumber) {
    if ($skuFriendlyNames.ContainsKey($SkuPartNumber)) {
        return $skuFriendlyNames[$SkuPartNumber]
    }
    return $SkuPartNumber
}

#endregion

#region ── Retrieve All Users ────────────────────────────────────────────────

Write-Host "[INFO] Retrieving users from Microsoft Graph..." -ForegroundColor Cyan

$userFilter = "userType eq 'Member'"
if ($IncludeGuests) { $userFilter = $null }  # no filter = all users

$graphParams = @{
    All      = $true
    Property = @(
        "id","displayName","userPrincipalName","mail","jobTitle","department",
        "officeLocation","usageLocation","accountEnabled","userType",
        "createdDateTime","lastPasswordChangeDateTime","onPremisesSyncEnabled",
        "assignedLicenses","assignedPlans","proxyAddresses","mailNickname"
    )
}
if ($userFilter) { $graphParams['Filter'] = $userFilter }

$allUsers = Get-MgUser @graphParams

Write-Host "[INFO] Retrieved $($allUsers.Count) users" -ForegroundColor Cyan

# Filter disabled accounts unless requested
if (-not $IncludeDisabled) {
    $allUsers = $allUsers | Where-Object { $_.AccountEnabled -eq $true }
    Write-Host "[INFO] After filtering disabled: $($allUsers.Count) active users" -ForegroundColor Cyan
}

#endregion

#region ── Retrieve License SKU Names ────────────────────────────────────────

Write-Host "[INFO] Retrieving tenant license subscriptions..." -ForegroundColor Cyan
$subscribedSkus = Get-MgSubscribedSku -All

$skuIdToName = @{}
foreach ($sku in $subscribedSkus) {
    $skuIdToName[$sku.SkuId] = Get-FriendlySkuName -SkuPartNumber $sku.SkuPartNumber
}

#endregion

#region ── Build Report ──────────────────────────────────────────────────────

Write-Host "[INFO] Building license report..." -ForegroundColor Cyan

$results = [System.Collections.Generic.List[PSCustomObject]]::new()
$counter  = 0

foreach ($user in $allUsers) {
    $counter++
    if ($counter % 100 -eq 0) {
        Write-Progress -Activity "Processing users" `
            -Status "$counter / $($allUsers.Count)" `
            -PercentComplete (($counter / $allUsers.Count) * 100)
    }

    $assignedLicenseNames = $user.AssignedLicenses |
        ForEach-Object { if ($skuIdToName[$_.SkuId]) { $skuIdToName[$_.SkuId] } else { $_.SkuId } }

    $hasExchange    = $user.AssignedPlans | Where-Object { $_.Service -eq 'exchange' -and $_.CapabilityStatus -eq 'Enabled' }
    $hasSharePoint  = $user.AssignedPlans | Where-Object { $_.Service -eq 'SharePoint' -and $_.CapabilityStatus -eq 'Enabled' }
    $hasOneDrive    = $user.AssignedPlans | Where-Object { $_.Service -eq 'SharePoint' -and $_.CapabilityStatus -eq 'Enabled' }
    $hasTeams       = $user.AssignedPlans | Where-Object { $_.Service -eq 'TeamspaceAPI' -and $_.CapabilityStatus -eq 'Enabled' }

    $proxySmtp = $user.ProxyAddresses |
        Where-Object { $_ -match '^smtp:' -and $_ -notmatch '^SMTP:' } |
        ForEach-Object { $_ -replace '^smtp:','' }

    $results.Add([PSCustomObject]@{
        DisplayName               = $user.DisplayName
        UserPrincipalName         = $user.UserPrincipalName
        Mail                      = $user.Mail
        MailAliases               = $proxySmtp -join '; '
        UserType                  = $user.UserType
        AccountEnabled            = $user.AccountEnabled
        JobTitle                  = $user.JobTitle
        Department                = $user.Department
        OfficeLocation            = $user.OfficeLocation
        UsageLocation             = $user.UsageLocation
        OnPremisesSynced          = $user.OnPremisesSyncEnabled
        CreatedDateTime           = $user.CreatedDateTime
        LastPasswordChange        = $user.LastPasswordChangeDateTime
        LicenseCount              = $user.AssignedLicenses.Count
        AssignedLicenses          = $assignedLicenseNames -join '; '
        HasExchangeLicense        = [bool]$hasExchange
        HasSharePointLicense      = [bool]$hasSharePoint
        HasTeamsLicense           = [bool]$hasTeams
        ObjectId                  = $user.Id
    })
}

Write-Progress -Activity "Processing users" -Completed

#endregion

#region ── License Summary ───────────────────────────────────────────────────

$licenseSummary = [System.Collections.Generic.List[PSCustomObject]]::new()

foreach ($sku in $subscribedSkus) {
    $friendlyName = Get-FriendlySkuName -SkuPartNumber $sku.SkuPartNumber
    $licenseSummary.Add([PSCustomObject]@{
        SkuPartNumber    = $sku.SkuPartNumber
        FriendlyName     = $friendlyName
        TotalPurchased   = $sku.PrepaidUnits.Enabled
        Assigned         = $sku.ConsumedUnits
        Available        = $sku.PrepaidUnits.Enabled - $sku.ConsumedUnits
        Suspended        = $sku.PrepaidUnits.Suspended
        Warning          = $sku.PrepaidUnits.Warning
    })
}

#endregion

#region ── Export and Summary ────────────────────────────────────────────────

$results       | Export-Csv -Path $OutputCsv        -NoTypeInformation -Encoding UTF8
$licenseSummary | Export-Csv -Path $OutputSummaryCsv -NoTypeInformation -Encoding UTF8

$licensed   = ($results | Where-Object { $_.LicenseCount -gt 0 } | Measure-Object).Count
$unlicensed = ($results | Where-Object { $_.LicenseCount -eq 0 } | Measure-Object).Count
$exUsers    = ($results | Where-Object { $_.HasExchangeLicense } | Measure-Object).Count

Write-Host "`n===== USER AND LICENSE SUMMARY =====" -ForegroundColor Cyan
Write-Host "Total users (active)    : $($results.Count)"
Write-Host "Licensed users          : $licensed"
Write-Host "Unlicensed users        : $unlicensed"
Write-Host "Users with Exchange     : $exUsers  (these need Gmail licenses)"

Write-Host "`nLicense SKUs in tenant:" -ForegroundColor Cyan
$licenseSummary | Format-Table FriendlyName, TotalPurchased, Assigned, Available -AutoSize

Write-Host "[OK] User report     : $OutputCsv"          -ForegroundColor Green
Write-Host "[OK] License summary : $OutputSummaryCsv"   -ForegroundColor Green

Disconnect-MgGraph

#endregion
