<#
.SYNOPSIS
    Exports all Exchange Online distribution groups, mail-enabled security groups,
    and Microsoft 365 Groups with membership counts and settings.

.DESCRIPTION
    Connects to Exchange Online using certificate-based (app-only) authentication
    and produces a CSV containing every group and its key attributes:

      - Group type (DistributionGroup, MailEnabledSecurity, M365Group, DynamicDistributionGroup)
      - Primary SMTP address and alias list
      - Member count
      - Owners
      - Whether the group is externally accessible
      - Delivery management settings (who can send to it)
      - Whether the group is hidden from the GAL

    Also produces a separate CSV of all group memberships (group → member) for
    recreating groups in Google Workspace as Google Groups.

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

.PARAMETER OutputCsv
    Path for the groups summary CSV. Defaults to .\Exchange-Groups-<timestamp>.csv

.PARAMETER OutputMembershipCsv
    Path for the detailed membership CSV. Defaults to .\Exchange-GroupMemberships-<timestamp>.csv

.PARAMETER IncludeM365Groups
    Include Microsoft 365 Groups (Teams, Planner, SharePoint-connected groups).

.PARAMETER ExpandNestedGroups
    Attempt to resolve nested group membership recursively (slower).

.EXAMPLE
    .\Get-ExchangeDistributionGroups.ps1 -ParamFile ".\MigrationConnectionParams.json"

.EXAMPLE
    .\Get-ExchangeDistributionGroups.ps1 -ParamFile ".\MigrationConnectionParams.json" `
        -IncludeM365Groups -ExpandNestedGroups

.NOTES
    Requires: ExchangeOnlineManagement module v3+
    App Registration: Exchange.ManageAsApp permission + Exchange Administrator role
#>

[CmdletBinding()]
param(
    [string]$ParamFile,

    [string]$TenantName,
    [string]$ClientId,
    [string]$CertificatePath,
    [SecureString]$CertificatePassword,

    [string]$OutputCsv,
    [string]$OutputMembershipCsv,

    [switch]$IncludeM365Groups,
    [switch]$ExpandNestedGroups
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
}

foreach ($var in @('TenantName','ClientId','CertificatePath')) {
    if (-not (Get-Variable $var -ValueOnly -ErrorAction SilentlyContinue)) {
        throw "Missing required parameter: -$var (or supply -ParamFile)"
    }
}

if (-not $OutputCsv) {
    $ts = Get-Date -Format "yyyyMMdd-HHmmss"
    $OutputCsv = Join-Path $PSScriptRoot "Exchange-Groups-$ts.csv"
}
if (-not $OutputMembershipCsv) {
    $ts = if ($ts) { $ts } else { Get-Date -Format "yyyyMMdd-HHmmss" }
    $OutputMembershipCsv = Join-Path $PSScriptRoot "Exchange-GroupMemberships-$ts.csv"
}

$tenantDomain = if ($TenantName -match '\.') { $TenantName } else { "$TenantName.onmicrosoft.com" }

#endregion

#region ── Connect to Exchange Online ────────────────────────────────────────

Write-Host "[INFO] Connecting to Exchange Online..." -ForegroundColor Cyan

$connectParams = @{
    AppId            = $ClientId
    Organization     = $tenantDomain
    CertificateFilePath = $CertificatePath
}
if ($CertificatePassword) { $connectParams['CertificatePassword'] = $CertificatePassword }

Connect-ExchangeOnline @connectParams -ShowBanner:$false
Write-Host "[OK]   Connected" -ForegroundColor Green

#endregion

#region ── Enumerate Groups ──────────────────────────────────────────────────

Write-Host "[INFO] Retrieving distribution groups..." -ForegroundColor Cyan

$groupSummary  = [System.Collections.Generic.List[PSCustomObject]]::new()
$membershipRows = [System.Collections.Generic.List[PSCustomObject]]::new()

# ── Distribution Groups & Mail-Enabled Security Groups ─────────────────────

$distGroups = Get-DistributionGroup -ResultSize Unlimited
Write-Host "[INFO] Found $($distGroups.Count) distribution/mail-enabled security groups" -ForegroundColor Cyan

$counter = 0
foreach ($grp in $distGroups) {
    $counter++
    Write-Progress -Activity "Processing distribution groups" `
        -Status "$counter / $($distGroups.Count): $($grp.PrimarySmtpAddress)" `
        -PercentComplete (($counter / $distGroups.Count) * 100)

    # Get members
    $members = @()
    try {
        $members = Get-DistributionGroupMember -Identity $grp.Identity -ResultSize Unlimited -ErrorAction Stop
    } catch {
        Write-Warning "  Could not get members for $($grp.PrimarySmtpAddress): $_"
    }

    $owners = if ($grp.ManagedBy) { ($grp.ManagedBy | ForEach-Object { $_ }) -join '; ' } else { '' }

    $groupSummary.Add([PSCustomObject]@{
        DisplayName           = $grp.DisplayName
        PrimarySmtpAddress    = $grp.PrimarySmtpAddress
        Alias                 = $grp.Alias
        GroupType             = $grp.RecipientTypeDetails
        MemberCount           = $members.Count
        Owners                = $owners
        HiddenFromGAL         = $grp.HiddenFromAddressListsEnabled
        RequireSenderAuth     = $grp.RequireSenderAuthenticationEnabled
        MemberJoinRestriction = $grp.MemberJoinRestriction
        MemberDepartRestriction = $grp.MemberDepartRestriction
        SendModerationEnabled = $grp.ModerationEnabled
        EmailAddresses        = ($grp.EmailAddresses | Where-Object { $_ -notmatch '^X500:' }) -join '; '
        Notes                 = $grp.Notes
    })

    foreach ($m in $members) {
        $membershipRows.Add([PSCustomObject]@{
            GroupSmtpAddress   = $grp.PrimarySmtpAddress
            GroupDisplayName   = $grp.DisplayName
            MemberSmtpAddress  = $m.PrimarySmtpAddress
            MemberDisplayName  = $m.DisplayName
            MemberType         = $m.RecipientType
        })
    }
}

Write-Progress -Activity "Processing distribution groups" -Completed

# ── Dynamic Distribution Groups ────────────────────────────────────────────

$dynGroups = Get-DynamicDistributionGroup -ResultSize Unlimited
Write-Host "[INFO] Found $($dynGroups.Count) dynamic distribution groups" -ForegroundColor Cyan

foreach ($grp in $dynGroups) {
    $groupSummary.Add([PSCustomObject]@{
        DisplayName           = $grp.DisplayName
        PrimarySmtpAddress    = $grp.PrimarySmtpAddress
        Alias                 = $grp.Alias
        GroupType             = "DynamicDistributionGroup"
        MemberCount           = "(dynamic)"
        Owners                = ($grp.ManagedBy -join '; ')
        HiddenFromGAL         = $grp.HiddenFromAddressListsEnabled
        RequireSenderAuth     = $grp.RequireSenderAuthenticationEnabled
        MemberJoinRestriction = "Dynamic"
        MemberDepartRestriction = "Dynamic"
        SendModerationEnabled = $grp.ModerationEnabled
        EmailAddresses        = ($grp.EmailAddresses | Where-Object { $_ -notmatch '^X500:' }) -join '; '
        Notes                 = "RecipientFilter: $($grp.RecipientFilter)"
    })
}

# ── Microsoft 365 Groups (optional) ────────────────────────────────────────

if ($IncludeM365Groups) {
    Write-Host "[INFO] Retrieving Microsoft 365 Groups..." -ForegroundColor Cyan
    $m365Groups = Get-UnifiedGroup -ResultSize Unlimited

    Write-Host "[INFO] Found $($m365Groups.Count) Microsoft 365 Groups" -ForegroundColor Cyan

    $counter = 0
    foreach ($grp in $m365Groups) {
        $counter++
        Write-Progress -Activity "Processing M365 Groups" `
            -Status "$counter / $($m365Groups.Count): $($grp.PrimarySmtpAddress)" `
            -PercentComplete (($counter / $m365Groups.Count) * 100)

        $members = @()
        try {
            $members = Get-UnifiedGroupLinks -Identity $grp.Identity -LinkType Members -ResultSize Unlimited -ErrorAction Stop
        } catch { }

        $owners = @()
        try {
            $owners = Get-UnifiedGroupLinks -Identity $grp.Identity -LinkType Owners -ResultSize Unlimited -ErrorAction Stop
        } catch { }

        $groupSummary.Add([PSCustomObject]@{
            DisplayName           = $grp.DisplayName
            PrimarySmtpAddress    = $grp.PrimarySmtpAddress
            Alias                 = $grp.Alias
            GroupType             = "M365Group"
            MemberCount           = $members.Count
            Owners                = ($owners.PrimarySmtpAddress -join '; ')
            HiddenFromGAL         = $grp.HiddenFromAddressListsEnabled
            RequireSenderAuth     = $grp.RequireSenderAuthenticationEnabled
            MemberJoinRestriction = if ($grp.AccessType -eq 'Private') { 'ApprovalRequired' } else { 'Open' }
            MemberDepartRestriction = "Open"
            SendModerationEnabled = $grp.ModerationEnabled
            EmailAddresses        = ($grp.EmailAddresses | Where-Object { $_ -notmatch '^X500:' }) -join '; '
            Notes                 = "SharePointUrl: $($grp.SharePointSiteUrl)"
        })

        foreach ($m in $members) {
            $membershipRows.Add([PSCustomObject]@{
                GroupSmtpAddress  = $grp.PrimarySmtpAddress
                GroupDisplayName  = $grp.DisplayName
                MemberSmtpAddress = $m.PrimarySmtpAddress
                MemberDisplayName = $m.DisplayName
                MemberType        = "M365GroupMember"
            })
        }
    }

    Write-Progress -Activity "Processing M365 Groups" -Completed
}

#endregion

#region ── Export and Summary ────────────────────────────────────────────────

$groupSummary   | Export-Csv -Path $OutputCsv           -NoTypeInformation -Encoding UTF8
$membershipRows | Export-Csv -Path $OutputMembershipCsv -NoTypeInformation -Encoding UTF8

$byType = $groupSummary | Group-Object GroupType | ForEach-Object {
    [PSCustomObject]@{ Type = $_.Name; Count = $_.Count }
}

Write-Host "`n===== GROUP SUMMARY =====" -ForegroundColor Cyan
$byType | Format-Table -AutoSize

Write-Host "Total groups      : $($groupSummary.Count)"
Write-Host "Total memberships : $($membershipRows.Count)"
Write-Host "`n[OK] Groups report  : $OutputCsv" -ForegroundColor Green
Write-Host "[OK] Memberships CSV: $OutputMembershipCsv" -ForegroundColor Green

Disconnect-ExchangeOnline -Confirm:$false

#endregion
