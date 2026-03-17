<#
.SYNOPSIS
    Reports file counts and storage usage across all SharePoint Online site collections.

.DESCRIPTION
    Uses PnP PowerShell with certificate-based (app-only) authentication to enumerate
    every SharePoint Online site collection and report:

      - Site URL and title
      - Storage used (MB / GB) and quota
      - Document library count
      - Total file count across all document libraries
      - Last activity date

    Optionally drills into each library for per-library file counts. Results are
    written to a CSV for migration sizing.

    Note: For tenants with many sites, use -SkipLibraryDetail for a fast summary
    pass that uses the SharePoint Admin API quota data instead of enumerating files.

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
    Path for the output CSV. Defaults to .\SPO-FileCounts-<timestamp>.csv

.PARAMETER SkipLibraryDetail
    If set, uses site-level storage quota data only (fast) instead of
    enumerating individual files per library (accurate but slow).

.PARAMETER ExcludeSystemSites
    Skip personal OneDrive sites and system sites (Search, Mysite host, etc.).
    Use Get-OneDriveFileCounts.ps1 for OneDrive inventory instead.

.PARAMETER SiteFilter
    Optional wildcard pattern to limit which sites are processed.
    Example: "https://contoso.sharepoint.com/sites/Project*"

.EXAMPLE
    .\Get-SharePointFileCounts.ps1 -ParamFile ".\MigrationConnectionParams.json"

.EXAMPLE
    .\Get-SharePointFileCounts.ps1 -ParamFile ".\MigrationConnectionParams.json" -SkipLibraryDetail

.EXAMPLE
    .\Get-SharePointFileCounts.ps1 -ParamFile ".\MigrationConnectionParams.json" `
        -SiteFilter "*/sites/Marketing*"

.NOTES
    Requires: PnP.PowerShell module
    The Azure AD app must have the following application permissions (granted via admin consent):
      - Sites.FullControl.All (SharePoint)
      - Sites.Read.All (Microsoft Graph)
#>

[CmdletBinding()]
param(
    [string]$ParamFile,

    [string]$TenantName,
    [string]$ClientId,
    [string]$CertificatePath,
    [SecureString]$CertificatePassword,

    [string]$OutputCsv,

    [switch]$SkipLibraryDetail,

    [switch]$ExcludeSystemSites,

    [string]$SiteFilter
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
    $OutputCsv = Join-Path $PSScriptRoot "SPO-FileCounts-$ts.csv"
}

$adminUrl = if ($ParamFile) { (Get-Content $ParamFile -Raw | ConvertFrom-Json).AdminSiteUrl } `
            else { "https://$TenantName-admin.sharepoint.com" }

#endregion

#region ── Helper: Count files in a site ─────────────────────────────────────

function Get-SiteFileCount {
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

    $libraries = Get-PnPList | Where-Object {
        $_.BaseType -eq 'DocumentLibrary' -and
        -not $_.Hidden -and
        $_.Title -notin @('Form Templates','Style Library','Site Assets','Site Pages','_catalogs')
    }

    $totalFiles = 0
    $libDetails = [System.Collections.Generic.List[PSCustomObject]]::new()

    foreach ($lib in $libraries) {
        try {
            $items = Get-PnPListItem -List $lib -Fields "FileLeafRef","FileRef","FSObjType" `
                        -PageSize 5000 -ErrorAction Stop |
                        Where-Object { $_["FSObjType"] -eq 0 }  # 0 = file, 1 = folder
            $count = ($items | Measure-Object).Count
            $totalFiles += $count
            $libDetails.Add([PSCustomObject]@{
                LibraryTitle = $lib.Title
                FileCount    = $count
            })
        } catch {
            Write-Warning "  Could not enumerate library '$($lib.Title)' in $SiteUrl : $_"
        }
    }

    Disconnect-PnPOnline

    return @{ TotalFiles = $totalFiles; Libraries = $libDetails; LibraryCount = $libraries.Count }
}

#endregion

#region ── Connect to SPO Admin ──────────────────────────────────────────────

Write-Host "[INFO] Connecting to SharePoint Online admin center: $adminUrl" -ForegroundColor Cyan

$connectParams = @{
    Url            = $adminUrl
    ClientId       = $ClientId
    CertificatePath = $CertificatePath
    Tenant         = "$TenantName.onmicrosoft.com"
}
if ($CertificatePassword) { $connectParams['CertificatePassword'] = $CertificatePassword }

Connect-PnPOnline @connectParams
Write-Host "[OK]   Connected" -ForegroundColor Green

#endregion

#region ── Enumerate All Sites ───────────────────────────────────────────────

Write-Host "[INFO] Retrieving site collections..." -ForegroundColor Cyan

$allSites = Get-PnPTenantSite -IncludeOneDriveSites:$false -Detailed

if ($ExcludeSystemSites) {
    $systemTemplates = @('SRCHCEN#0','SPSMSITEHOST#0','APPCATALOG#0','POINTPUBLISHINGHUB#0','EDISC#0','STS#-1')
    $allSites = $allSites | Where-Object {
        $_.Template -notin $systemTemplates -and
        $_.Url -notmatch '/-my/' -and
        $_.Url -notmatch '/portals/'
    }
}

if ($SiteFilter) {
    $allSites = $allSites | Where-Object { $_.Url -like $SiteFilter }
}

Write-Host "[INFO] Processing $($allSites.Count) sites" -ForegroundColor Cyan

#endregion

#region ── Collect Stats ─────────────────────────────────────────────────────

$results = [System.Collections.Generic.List[PSCustomObject]]::new()
$counter = 0

foreach ($site in $allSites) {
    $counter++
    Write-Progress -Activity "Analyzing SharePoint sites" `
        -Status "$counter / $($allSites.Count): $($site.Url)" `
        -PercentComplete (($counter / $allSites.Count) * 100)

    $storageMB    = [math]::Round($site.StorageUsageCurrent, 2)
    $storageGB    = [math]::Round($storageMB / 1024, 3)
    $quotaGB      = [math]::Round($site.StorageMaximumLevel / 1024, 1)

    $row = [PSCustomObject]@{
        Title              = $site.Title
        Url                = $site.Url
        Template           = $site.Template
        StorageUsedMB      = $storageMB
        StorageUsedGB      = $storageGB
        StorageQuotaGB     = $quotaGB
        StorageUsedPercent = if ($site.StorageMaximumLevel -gt 0) { [math]::Round(($site.StorageUsageCurrent / $site.StorageMaximumLevel) * 100, 1) } else { 0 }
        LastContentModified = $site.LastContentModifiedDate
        SharingCapability  = $site.SharingCapability
        IsHubSite          = $site.IsHubSite
        LibraryCount       = $null
        TotalFileCount     = $null
        Status             = "OK"
    }

    if (-not $SkipLibraryDetail) {
        try {
            $fileData = Get-SiteFileCount `
                -SiteUrl $site.Url `
                -ClientId $ClientId `
                -CertPath $CertificatePath `
                -CertPassword $CertificatePassword `
                -Tenant "$TenantName.onmicrosoft.com"

            $row.LibraryCount   = $fileData.LibraryCount
            $row.TotalFileCount = $fileData.TotalFiles
        } catch {
            $row.Status = "ERROR: $_"
            Write-Warning "  Failed to get file count for: $($site.Url)"
        }
    }

    $results.Add($row)
}

Write-Progress -Activity "Analyzing SharePoint sites" -Completed

#endregion

#region ── Export and Summary ────────────────────────────────────────────────

$results | Export-Csv -Path $OutputCsv -NoTypeInformation -Encoding UTF8

$totalSites  = $results.Count
$totalGB     = [math]::Round(($results | Measure-Object StorageUsedGB -Sum).Sum, 2)
$totalFiles  = ($results | Where-Object { $null -ne $_.TotalFileCount } | Measure-Object TotalFileCount -Sum).Sum

Write-Host "`n===== SHAREPOINT SUMMARY =====" -ForegroundColor Cyan
Write-Host "Total site collections : $totalSites"
Write-Host "Total storage used     : $totalGB GB"
if (-not $SkipLibraryDetail) {
    Write-Host "Total files            : $totalFiles"
}

Write-Host "`nTop 10 sites by storage:" -ForegroundColor Cyan
$results | Sort-Object StorageUsedGB -Descending | Select-Object -First 10 |
    Format-Table Title, StorageUsedGB, TotalFileCount -AutoSize

Write-Host "[OK] Report saved to: $OutputCsv" -ForegroundColor Green

Disconnect-PnPOnline

#endregion
