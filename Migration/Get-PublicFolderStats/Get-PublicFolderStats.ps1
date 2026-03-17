<#
.SYNOPSIS
    Exports top-level mail-enabled public folders with message counts and permissions.

.DESCRIPTION
    Connects to Exchange Online using certificate-based (app-only) authentication and:

    1. Enumerates all top-level mail-enabled public folders (those with an SMTP address)
    2. Retrieves item counts and storage sizes per folder
    3. Enumerates client permissions per folder

    Produces two CSVs:
      - Public folder statistics (one row per folder)
      - Public folder permissions (one row per folder + user/group pairing)

    Scope: Top-level mail-enabled public folders only (not subfolders). A folder
    is "mail-enabled" if it has a PrimarySmtpAddress assigned. If you need stats
    for non-mail-enabled folders, remove the mail-enabled filter.

.PARAMETER ParamFile
    Path to MigrationConnectionParams.json from Setup-PnPCertAuth.ps1.

.PARAMETER TenantName / ClientId / CertificatePath
    Explicit connection parameters if not using -ParamFile.

.PARAMETER CertificatePassword
    SecureString password for the .pfx certificate.

.PARAMETER OutputStatsCsv
    Path for the stats CSV. Defaults to .\PublicFolder-Stats-<ts>.csv

.PARAMETER OutputPermissionsCsv
    Path for the permissions CSV. Defaults to .\PublicFolder-Permissions-<ts>.csv

.PARAMETER IncludeDefaultAndAnonymous
    By default, the built-in "Default" and "Anonymous" permission entries are
    excluded. Use this switch to include them.

.EXAMPLE
    .\Get-PublicFolderStats.ps1 -ParamFile ".\MigrationConnectionParams.json"

.EXAMPLE
    .\Get-PublicFolderStats.ps1 -ParamFile ".\MigrationConnectionParams.json" `
        -IncludeDefaultAndAnonymous

.NOTES
    Requires: ExchangeOnlineManagement module v3+
    App permissions: Exchange.ManageAsApp + Exchange Administrator role on service principal.

    Public folder cmdlets (Get-PublicFolder, Get-PublicFolderStatistics,
    Get-MailPublicFolder, Get-PublicFolderClientPermission) are available in
    Exchange Online with the correct role assignment.
#>

[CmdletBinding()]
param(
    [string]$ParamFile,

    [string]$TenantName,
    [string]$ClientId,
    [string]$CertificatePath,
    [SecureString]$CertificatePassword,

    [string]$OutputStatsCsv,
    [string]$OutputPermissionsCsv,

    [switch]$IncludeDefaultAndAnonymous
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

foreach ($v in @('TenantName','ClientId','CertificatePath')) {
    if (-not (Get-Variable $v -ValueOnly -ErrorAction SilentlyContinue)) {
        throw "Missing required parameter: -$v (or use -ParamFile)"
    }
}

$ts = Get-Date -Format "yyyyMMdd-HHmmss"
if (-not $OutputStatsCsv)       { $OutputStatsCsv       = Join-Path $PSScriptRoot "PublicFolder-Stats-$ts.csv" }
if (-not $OutputPermissionsCsv) { $OutputPermissionsCsv = Join-Path $PSScriptRoot "PublicFolder-Permissions-$ts.csv" }

$tenantDomain = if ($TenantName -match '\.') { $TenantName } else { "$TenantName.onmicrosoft.com" }

$excludePrincipals = @('Default','Anonymous')

#endregion

#region ── Connect ───────────────────────────────────────────────────────────

Write-Host "[INFO] Connecting to Exchange Online..." -ForegroundColor Cyan

$connectParams = @{
    AppId               = $ClientId
    Organization        = $tenantDomain
    CertificateFilePath = $CertificatePath
}
if ($CertificatePassword) { $connectParams['CertificatePassword'] = $CertificatePassword }

Connect-ExchangeOnline @connectParams -ShowBanner:$false
Write-Host "[OK]   Connected" -ForegroundColor Green

#endregion

#region ── Get Mail-Enabled Public Folders ────────────────────────────────────

Write-Host "[INFO] Retrieving mail-enabled public folders..." -ForegroundColor Cyan

# Get all mail-enabled public folders (have an SMTP address)
$mailPFs = Get-MailPublicFolder -ResultSize Unlimited

# Limit to top-level only: identity path has exactly one backslash segment, e.g. "\FolderName"
$topLevelMailPFs = $mailPFs | Where-Object {
    # Identity looks like "\FolderName" — split on \ gives 2 parts with empty first element
    ($_.Identity -split '\\').Count -le 2
}

Write-Host "[INFO] Found $($topLevelMailPFs.Count) top-level mail-enabled public folders" -ForegroundColor Cyan

if ($topLevelMailPFs.Count -eq 0) {
    Write-Warning "No top-level mail-enabled public folders found. The tenant may have no public folders, or they may all be nested."
}

#endregion

#region ── Collect Stats and Permissions ─────────────────────────────────────

$statsRows  = [System.Collections.Generic.List[PSCustomObject]]::new()
$permRows   = [System.Collections.Generic.List[PSCustomObject]]::new()
$counter    = 0

foreach ($mpf in $topLevelMailPFs) {
    $counter++
    Write-Progress -Activity "Processing public folders" `
        -Status "$counter / $($topLevelMailPFs.Count): $($mpf.PrimarySmtpAddress)" `
        -PercentComplete (($counter / $topLevelMailPFs.Count) * 100)

    # ── Statistics ─────────────────────────────────────────────────────────

    $stats = $null
    try {
        $stats = Get-PublicFolderStatistics -Identity $mpf.Identity -ErrorAction Stop
    } catch {
        Write-Warning "  Could not get statistics for '$($mpf.Identity)': $_"
    }

    $totalBytes   = if ($stats -and $stats.TotalItemSize.Value)   { $stats.TotalItemSize.Value.ToBytes()   } else { 0 }
    $deletedBytes = if ($stats -and $stats.TotalDeletedItemSize.Value) { $stats.TotalDeletedItemSize.Value.ToBytes() } else { 0 }

    $statsRows.Add([PSCustomObject]@{
        FolderName          = $mpf.Name
        FolderPath          = $mpf.Identity
        PrimarySmtpAddress  = $mpf.PrimarySmtpAddress
        Alias               = $mpf.Alias
        EmailAddresses      = ($mpf.EmailAddresses | Where-Object { $_ -notmatch '^X500:' }) -join '; '
        HiddenFromGAL       = $mpf.HiddenFromAddressListsEnabled
        ItemCount           = if ($stats) { $stats.ItemCount } else { "ERROR" }
        TotalSizeMB         = [math]::Round($totalBytes / 1MB, 2)
        TotalSizeGB         = [math]::Round($totalBytes / 1GB, 3)
        DeletedItemCount    = if ($stats) { $stats.DeletedItemCount } else { "ERROR" }
        DeletedSizeMB       = [math]::Round($deletedBytes / 1MB, 2)
        LastModifiedTime    = if ($stats) { $stats.LastModifiedTime } else { $null }
        CreationTime        = if ($stats) { $stats.CreationTime } else { $null }
    })

    # ── Permissions ────────────────────────────────────────────────────────

    $perms = @()
    try {
        $perms = Get-PublicFolderClientPermission -Identity $mpf.Identity -ErrorAction Stop
    } catch {
        Write-Warning "  Could not get permissions for '$($mpf.Identity)': $_"
    }

    foreach ($perm in $perms) {
        $userName = $perm.User.ToString()

        # Skip Default/Anonymous unless requested
        if (-not $IncludeDefaultAndAnonymous -and $userName -in $excludePrincipals) {
            continue
        }

        # Resolve SMTP address for the principal where possible
        $principalSmtp = $null
        if ($userName -notin $excludePrincipals -and $userName -notmatch '^NT ') {
            try {
                $recipient = Get-EXORecipient -Identity $userName -ErrorAction Stop
                $principalSmtp = $recipient.PrimarySmtpAddress
            } catch { }
        }

        $permRows.Add([PSCustomObject]@{
            FolderName         = $mpf.Name
            FolderPath         = $mpf.Identity
            FolderSmtp         = $mpf.PrimarySmtpAddress
            Principal          = $userName
            PrincipalEmail     = $principalSmtp
            AccessRights       = ($perm.AccessRights -join ', ')
            IsDefault          = $userName -eq 'Default'
            IsAnonymous        = $userName -eq 'Anonymous'
        })
    }
}

Write-Progress -Activity "Processing public folders" -Completed

#endregion

#region ── Export and Summary ────────────────────────────────────────────────

$statsRows | Export-Csv -Path $OutputStatsCsv       -NoTypeInformation -Encoding UTF8
$permRows  | Export-Csv -Path $OutputPermissionsCsv -NoTypeInformation -Encoding UTF8

$totalItems  = ($statsRows | Where-Object { $_.ItemCount -match '^\d' } | Measure-Object ItemCount -Sum).Sum
$totalGB     = [math]::Round(($statsRows | Measure-Object TotalSizeGB -Sum).Sum, 2)

Write-Host "`n===== PUBLIC FOLDER SUMMARY =====" -ForegroundColor Cyan
Write-Host "Mail-enabled top-level folders : $($statsRows.Count)"
Write-Host "Total message count            : $totalItems"
Write-Host "Total storage                  : $totalGB GB"
Write-Host "Total permission rows          : $($permRows.Count)"

# Access rights distribution
$rightsBreakdown = $permRows | Group-Object AccessRights | ForEach-Object {
    [PSCustomObject]@{ AccessRights = $_.Name; Count = $_.Count }
} | Sort-Object Count -Descending

Write-Host "`nPermission type breakdown:" -ForegroundColor Cyan
$rightsBreakdown | Format-Table -AutoSize

Write-Host "[OK] Stats CSV       : $OutputStatsCsv"       -ForegroundColor Green
Write-Host "[OK] Permissions CSV : $OutputPermissionsCsv" -ForegroundColor Green

Disconnect-ExchangeOnline -Confirm:$false

#endregion
