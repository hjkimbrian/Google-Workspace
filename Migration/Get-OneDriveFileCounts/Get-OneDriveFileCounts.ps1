<#
.SYNOPSIS
    Reports file counts and storage usage for all OneDrive for Business accounts.

.DESCRIPTION
    Uses PnP PowerShell with certificate-based (app-only) authentication to enumerate
    every user's OneDrive for Business and report:

      - Owner display name, UPN, and site URL
      - Storage used (MB / GB)
      - Total file count
      - Last activity date
      - Account status (Active / Recycled)

    This data is critical for sizing a OneDrive-to-Google Drive migration.

.PARAMETER ParamFile
    Path to MigrationConnectionParams.json from Setup-PnPCertAuth.ps1.

.PARAMETER TenantName
    M365 tenant name (e.g. "contoso"). Not needed if -ParamFile is used.

.PARAMETER ClientId
    Azure AD Application (client) ID. Not needed if -ParamFile is used.

.PARAMETER CertificatePath
    Path to the .pfx certificate.

.PARAMETER CertificatePassword
    SecureString password for the certificate.

.PARAMETER OutputCsv
    Path for the output CSV. Defaults to .\OneDrive-FileCounts-<timestamp>.csv

.PARAMETER SkipFileCount
    If set, returns storage sizes from the SPO admin API only (fast),
    without enumerating individual files per OneDrive.

.PARAMETER UserFilter
    Optional UPN wildcard to limit which users are processed.
    Example: "*@contoso.com"

.PARAMETER MinStorageMB
    Only include OneDrives with at least this many MB used. Useful to skip
    empty/unused accounts. Default: 0 (include all).

.EXAMPLE
    .\Get-OneDriveFileCounts.ps1 -ParamFile ".\MigrationConnectionParams.json"

.EXAMPLE
    .\Get-OneDriveFileCounts.ps1 -ParamFile ".\MigrationConnectionParams.json" -SkipFileCount

.EXAMPLE
    .\Get-OneDriveFileCounts.ps1 -ParamFile ".\MigrationConnectionParams.json" `
        -MinStorageMB 100 -SkipFileCount

.NOTES
    Requires: PnP.PowerShell module
    App Registration permissions required:
      - Sites.FullControl.All (SharePoint)
      - User.Read.All (Microsoft Graph)
#>

[CmdletBinding()]
param(
    [string]$ParamFile,

    [string]$TenantName,
    [string]$ClientId,
    [string]$CertificatePath,
    [SecureString]$CertificatePassword,

    [string]$OutputCsv,

    [switch]$SkipFileCount,

    [string]$UserFilter,

    [int]$MinStorageMB = 0
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
    $OutputCsv = Join-Path $PSScriptRoot "OneDrive-FileCounts-$ts.csv"
}

$tenantDomain = if ($TenantName -match '\.') { $TenantName } else { "$TenantName.onmicrosoft.com" }
$adminUrl     = "https://$TenantName-admin.sharepoint.com"

#endregion

#region ── Helper: Count files in a OneDrive ─────────────────────────────────

function Get-OneDriveFileCount {
    param(
        [string]$SiteUrl,
        [string]$ClientId,
        [string]$CertPath,
        [SecureString]$CertPassword,
        [string]$Tenant
    )

    $connectParams = @{
        Url            = $SiteUrl
        ClientId       = $ClientId
        CertificatePath = $CertPath
        Tenant         = $Tenant
    }
    if ($CertPassword) { $connectParams['CertificatePassword'] = $CertPassword }

    Connect-PnPOnline @connectParams -ErrorAction Stop

    try {
        $docs = Get-PnPList -Identity "Documents" -ErrorAction Stop
        $items = Get-PnPListItem -List $docs -Fields "FSObjType" -PageSize 5000 |
                    Where-Object { $_["FSObjType"] -eq 0 }
        $count = ($items | Measure-Object).Count
    } catch {
        $count = -1
    } finally {
        Disconnect-PnPOnline
    }
    return $count
}

#endregion

#region ── Connect to SPO Admin ──────────────────────────────────────────────

Write-Host "[INFO] Connecting to SharePoint Online admin center..." -ForegroundColor Cyan

$connectParams = @{
    Url            = $adminUrl
    ClientId       = $ClientId
    CertificatePath = $CertificatePath
    Tenant         = $tenantDomain
}
if ($CertificatePassword) { $connectParams['CertificatePassword'] = $CertificatePassword }

Connect-PnPOnline @connectParams
Write-Host "[OK]   Connected" -ForegroundColor Green

#endregion

#region ── Enumerate OneDrive Sites ──────────────────────────────────────────

Write-Host "[INFO] Retrieving OneDrive for Business accounts..." -ForegroundColor Cyan

$oneDrives = Get-PnPTenantSite -IncludeOneDriveSites -Filter "Url -like '-my.sharepoint.com/personal/'" -Detailed

if ($UserFilter) {
    # Convert UPN filter (user@contoso.com) to OneDrive URL pattern (user_contoso_com)
    $urlFragment = ($UserFilter -replace '@','_' -replace '\.','_' -replace '\*','*').ToLower()
    $oneDrives = $oneDrives | Where-Object { $_.Url -like "*personal/$urlFragment" }
}

if ($MinStorageMB -gt 0) {
    $oneDrives = $oneDrives | Where-Object { $_.StorageUsageCurrent -ge $MinStorageMB }
}

Write-Host "[INFO] Found $($oneDrives.Count) OneDrive accounts to process" -ForegroundColor Cyan

#endregion

#region ── Collect Stats ─────────────────────────────────────────────────────

$results = [System.Collections.Generic.List[PSCustomObject]]::new()
$counter  = 0

foreach ($od in $oneDrives) {
    $counter++
    Write-Progress -Activity "Analyzing OneDrive accounts" `
        -Status "$counter / $($oneDrives.Count): $($od.Owner)" `
        -PercentComplete (($counter / $oneDrives.Count) * 100)

    $storageMB = [math]::Round($od.StorageUsageCurrent, 2)
    $storageGB = [math]::Round($storageMB / 1024, 3)
    $quotaGB   = [math]::Round($od.StorageMaximumLevel / 1024, 1)

    # Extract UPN from URL: /personal/firstname_lastname_contoso_com
    $urlPart = $od.Url -replace '.*\/personal\/', ''
    $ownerUpn = if ($od.Owner) { $od.Owner } else {
        # Reconstruct UPN from URL segment
        $parts = $urlPart -split '_'
        if ($parts.Count -ge 3) {
            "$($parts[0..($parts.Count-3)] -join '.')@$($parts[-2]).$($parts[-1])"
        } else { $urlPart }
    }

    $row = [PSCustomObject]@{
        OwnerDisplayName   = $od.Title
        OwnerUPN           = $ownerUpn
        SiteUrl            = $od.Url
        StorageUsedMB      = $storageMB
        StorageUsedGB      = $storageGB
        StorageQuotaGB     = $quotaGB
        StorageUsedPercent = if ($od.StorageMaximumLevel -gt 0) { [math]::Round(($od.StorageUsageCurrent / $od.StorageMaximumLevel) * 100, 1) } else { 0 }
        LastContentModified = $od.LastContentModifiedDate
        Status             = $od.Status
        TotalFileCount     = $null
        Error              = $null
    }

    if (-not $SkipFileCount -and $storageMB -gt 0) {
        try {
            $fileCount = Get-OneDriveFileCount `
                -SiteUrl $od.Url `
                -ClientId $ClientId `
                -CertPath $CertificatePath `
                -CertPassword $CertificatePassword `
                -Tenant $tenantDomain

            $row.TotalFileCount = if ($fileCount -eq -1) { "ERROR" } else { $fileCount }
        } catch {
            $row.TotalFileCount = "ERROR"
            $row.Error = $_.Exception.Message
            Write-Warning "  Failed to count files for: $($od.Url)"
        }
    }

    $results.Add($row)
}

Write-Progress -Activity "Analyzing OneDrive accounts" -Completed

#endregion

#region ── Export and Summary ────────────────────────────────────────────────

$results | Export-Csv -Path $OutputCsv -NoTypeInformation -Encoding UTF8

$totalAccounts = $results.Count
$totalGB       = [math]::Round(($results | Measure-Object StorageUsedGB -Sum).Sum, 2)
$emptyAccounts = ($results | Where-Object { $_.StorageUsedMB -eq 0 } | Measure-Object).Count
$activeAccounts= $totalAccounts - $emptyAccounts

if (-not $SkipFileCount) {
    $totalFiles = ($results | Where-Object { $_.TotalFileCount -match '^\d' } | Measure-Object TotalFileCount -Sum).Sum
}

Write-Host "`n===== ONEDRIVE SUMMARY =====" -ForegroundColor Cyan
Write-Host "Total OneDrive accounts : $totalAccounts"
Write-Host "Active (non-empty)      : $activeAccounts"
Write-Host "Empty accounts          : $emptyAccounts"
Write-Host "Total storage used      : $totalGB GB"
if (-not $SkipFileCount) {
    Write-Host "Total files             : $totalFiles"
}

# Size distribution buckets
$buckets = @(
    @{ Name = "Empty (0 MB)";    Min = 0;      Max = 0 }
    @{ Name = "< 1 GB";          Min = 0.001;  Max = 1 }
    @{ Name = "1–5 GB";          Min = 1;      Max = 5 }
    @{ Name = "5–15 GB";         Min = 5;      Max = 15 }
    @{ Name = "15–50 GB";        Min = 15;     Max = 50 }
    @{ Name = "> 50 GB";         Min = 50;     Max = [double]::MaxValue }
)

Write-Host "`nSize distribution:" -ForegroundColor Cyan
foreach ($b in $buckets) {
    $count = ($results | Where-Object { $_.StorageUsedGB -ge $b.Min -and $_.StorageUsedGB -lt $b.Max } | Measure-Object).Count
    Write-Host ("  {0,-18} : {1,5} accounts" -f $b.Name, $count)
}

Write-Host "`nTop 10 accounts by storage:" -ForegroundColor Cyan
$results | Sort-Object StorageUsedGB -Descending | Select-Object -First 10 |
    Format-Table OwnerUPN, StorageUsedGB, TotalFileCount -AutoSize

Write-Host "[OK] Report saved to: $OutputCsv" -ForegroundColor Green

Disconnect-PnPOnline

#endregion
