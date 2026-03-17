<#
.SYNOPSIS
    Exports Exchange Online mailbox inventory with message counts and storage sizes.

.DESCRIPTION
    Connects to Exchange Online using certificate-based (app-only) authentication and
    produces a CSV report containing, for every mailbox:

      - Display name, UPN, primary SMTP address
      - Mailbox type (UserMailbox, SharedMailbox, RoomMailbox, EquipmentMailbox)
      - Item count (total messages across all folders)
      - Total size in MB and GB
      - Deleted item count and size
      - Last logon time
      - Whether the mailbox is licensed

    Useful as a pre-migration sizing report before moving mailboxes to Google Workspace.

.PARAMETER ParamFile
    Path to the MigrationConnectionParams.json produced by Setup-PnPCertAuth.ps1.
    If supplied, TenantName / ClientId / CertificatePath are read from the file.

.PARAMETER TenantName
    M365 tenant name (e.g. "contoso"). Not needed if -ParamFile is used.

.PARAMETER ClientId
    Azure AD Application (client) ID. Not needed if -ParamFile is used.

.PARAMETER CertificatePath
    Path to the .pfx certificate file. Not needed if -ParamFile is used.

.PARAMETER CertificatePassword
    SecureString password for the .pfx certificate.

.PARAMETER MailboxTypes
    Comma-separated list of mailbox types to include.
    Valid values: UserMailbox, SharedMailbox, RoomMailbox, EquipmentMailbox, DiscoveryMailbox
    Defaults to: UserMailbox, SharedMailbox

.PARAMETER OutputCsv
    Path for the output CSV file. Defaults to .\Exchange-Mailbox-Stats-<timestamp>.csv

.PARAMETER IncludeInactiveMailboxes
    Switch to include soft-deleted / inactive mailboxes.

.EXAMPLE
    .\Get-ExchangeMailboxStats.ps1 -ParamFile ".\MigrationConnectionParams.json"

.EXAMPLE
    .\Get-ExchangeMailboxStats.ps1 `
        -TenantName "contoso" `
        -ClientId "xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx" `
        -CertificatePath ".\PnP-Migration-App.pfx" `
        -MailboxTypes "UserMailbox,SharedMailbox,RoomMailbox"

.NOTES
    Requires: ExchangeOnlineManagement module v3+
    The Azure AD app must have the 'Exchange.ManageAsApp' application permission
    and the service principal must be assigned the 'Exchange Administrator' role.
#>

[CmdletBinding()]
param(
    [string]$ParamFile,

    [string]$TenantName,
    [string]$ClientId,
    [string]$CertificatePath,
    [SecureString]$CertificatePassword,

    [string]$MailboxTypes = "UserMailbox,SharedMailbox",

    [string]$OutputCsv,

    [switch]$IncludeInactiveMailboxes
)

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

#region ── Load Parameters ───────────────────────────────────────────────────

if ($ParamFile) {
    if (-not (Test-Path $ParamFile)) { throw "ParamFile not found: $ParamFile" }
    $p = Get-Content $ParamFile -Raw | ConvertFrom-Json
    if (-not $TenantName)    { $TenantName    = $p.TenantName }
    if (-not $ClientId)      { $ClientId      = $p.ClientId }
    if (-not $CertificatePath) { $CertificatePath = $p.CertificatePath }
}

foreach ($var in @('TenantName','ClientId','CertificatePath')) {
    if (-not (Get-Variable $var -ValueOnly -ErrorAction SilentlyContinue)) {
        throw "Missing required parameter: -$var (or supply -ParamFile)"
    }
}

if (-not $OutputCsv) {
    $ts = Get-Date -Format "yyyyMMdd-HHmmss"
    $OutputCsv = Join-Path $PSScriptRoot "Exchange-Mailbox-Stats-$ts.csv"
}

$tenantDomain = if ($TenantName -match '\.') { $TenantName } else { "$TenantName.onmicrosoft.com" }
$requestedTypes = $MailboxTypes -split ',' | ForEach-Object { $_.Trim() }

#endregion

#region ── Connect to Exchange Online ────────────────────────────────────────

Write-Host "[INFO] Connecting to Exchange Online (app-only)..." -ForegroundColor Cyan

$connectParams = @{
    AppId            = $ClientId
    Organization     = $tenantDomain
    CertificateFilePath = $CertificatePath
}

if ($CertificatePassword) {
    $connectParams['CertificatePassword'] = $CertificatePassword
}

Connect-ExchangeOnline @connectParams -ShowBanner:$false
Write-Host "[OK]   Connected to Exchange Online" -ForegroundColor Green

#endregion

#region ── Enumerate Mailboxes ───────────────────────────────────────────────

Write-Host "[INFO] Retrieving mailboxes (types: $($requestedTypes -join ', '))..." -ForegroundColor Cyan

$getMailboxParams = @{
    ResultSize = 'Unlimited'
    RecipientTypeDetails = $requestedTypes
}
if ($IncludeInactiveMailboxes) { $getMailboxParams['InactiveMailboxOnly'] = $false }

$mailboxes = Get-EXOMailbox @getMailboxParams -PropertySets Minimum,Delivery,Hold
Write-Host "[INFO] Found $($mailboxes.Count) mailboxes" -ForegroundColor Cyan

#endregion

#region ── Collect Statistics ────────────────────────────────────────────────

Write-Host "[INFO] Collecting mailbox statistics (this may take several minutes)..." -ForegroundColor Cyan

$results = [System.Collections.Generic.List[PSCustomObject]]::new()
$counter = 0

foreach ($mbx in $mailboxes) {
    $counter++
    if ($counter % 50 -eq 0) {
        Write-Progress -Activity "Gathering mailbox stats" `
            -Status "$counter / $($mailboxes.Count)" `
            -PercentComplete (($counter / $mailboxes.Count) * 100)
    }

    try {
        $stats = Get-EXOMailboxStatistics -Identity $mbx.ExchangeGuid -PropertySets All -ErrorAction Stop

        # Parse size strings like "1.234 GB (1,234,567,890 bytes)"
        $totalSizeBytes   = if ($stats.TotalItemSize.Value)   { $stats.TotalItemSize.Value.ToBytes()   } else { 0 }
        $deletedSizeBytes = if ($stats.TotalDeletedItemSize.Value) { $stats.TotalDeletedItemSize.Value.ToBytes() } else { 0 }

        $row = [PSCustomObject]@{
            DisplayName         = $mbx.DisplayName
            UserPrincipalName   = $mbx.UserPrincipalName
            PrimarySmtpAddress  = $mbx.PrimarySmtpAddress
            RecipientTypeDetails = $mbx.RecipientTypeDetails
            TotalItemCount      = $stats.ItemCount
            TotalSizeMB         = [math]::Round($totalSizeBytes / 1MB, 2)
            TotalSizeGB         = [math]::Round($totalSizeBytes / 1GB, 3)
            DeletedItemCount    = $stats.DeletedItemCount
            DeletedSizeMB       = [math]::Round($deletedSizeBytes / 1MB, 2)
            LastLogonTime       = $stats.LastLogonTime
            IsArchiveEnabled    = $mbx.ArchiveStatus -ne 'None'
            LitigationHoldEnabled = $mbx.LitigationHoldEnabled
            Languages           = ($mbx.Languages -join '; ')
            ExchangeGuid        = $mbx.ExchangeGuid
        }

        # Fetch archive stats if enabled
        if ($row.IsArchiveEnabled) {
            try {
                $archStats = Get-EXOMailboxStatistics -Identity $mbx.ExchangeGuid -Archive -ErrorAction Stop
                $archBytes = if ($archStats.TotalItemSize.Value) { $archStats.TotalItemSize.Value.ToBytes() } else { 0 }
                $row | Add-Member -NotePropertyName ArchiveItemCount -NotePropertyValue $archStats.ItemCount
                $row | Add-Member -NotePropertyName ArchiveSizeGB    -NotePropertyValue ([math]::Round($archBytes / 1GB, 3))
            } catch {
                $row | Add-Member -NotePropertyName ArchiveItemCount -NotePropertyValue 0
                $row | Add-Member -NotePropertyName ArchiveSizeGB    -NotePropertyValue 0
            }
        } else {
            $row | Add-Member -NotePropertyName ArchiveItemCount -NotePropertyValue 0
            $row | Add-Member -NotePropertyName ArchiveSizeGB    -NotePropertyValue 0
        }

        $results.Add($row)

    } catch {
        Write-Warning "Could not get stats for $($mbx.UserPrincipalName): $_"
        $results.Add([PSCustomObject]@{
            DisplayName        = $mbx.DisplayName
            UserPrincipalName  = $mbx.UserPrincipalName
            PrimarySmtpAddress = $mbx.PrimarySmtpAddress
            RecipientTypeDetails = $mbx.RecipientTypeDetails
            TotalItemCount     = "ERROR"
            TotalSizeMB        = "ERROR"
            TotalSizeGB        = "ERROR"
            Error              = $_.Exception.Message
        })
    }
}

Write-Progress -Activity "Gathering mailbox stats" -Completed

#endregion

#region ── Export and Summary ────────────────────────────────────────────────

$results | Export-Csv -Path $OutputCsv -NoTypeInformation -Encoding UTF8

# Summary by type
$summary = $results | Where-Object { $_.TotalItemCount -ne "ERROR" } |
    Group-Object RecipientTypeDetails |
    ForEach-Object {
        [PSCustomObject]@{
            MailboxType   = $_.Name
            Count         = $_.Count
            TotalItems    = ($_.Group | Measure-Object TotalItemCount -Sum).Sum
            TotalSizeGB   = [math]::Round(($_.Group | Measure-Object TotalSizeGB -Sum).Sum, 2)
        }
    }

Write-Host "`n===== MAILBOX SUMMARY =====" -ForegroundColor Cyan
$summary | Format-Table -AutoSize

$grandTotalGB = [math]::Round(($results | Where-Object { $_.TotalSizeGB -match '^\d' } | Measure-Object TotalSizeGB -Sum).Sum, 2)
Write-Host "Total mailboxes : $($results.Count)"
Write-Host "Grand total size: $grandTotalGB GB"
Write-Host "`n[OK] Report saved to: $OutputCsv" -ForegroundColor Green

Disconnect-ExchangeOnline -Confirm:$false

#endregion
