<#
.SYNOPSIS
    Exports a flat list of every secondary SMTP alias mapped to its primary SMTP
    address, across all mail-enabled recipient types in the tenant.

.DESCRIPTION
    Uses Get-EXORecipient (a single call) to enumerate every mail-enabled object:

      UserMailbox, SharedMailbox, RoomMailbox, EquipmentMailbox,
      MailUser, MailContact, DistributionGroup, MailEnabledSecurityGroup,
      DynamicDistributionGroup, GroupMailbox (M365 Groups),
      PublicFolder (mail-enabled)

    For each recipient, every secondary SMTP address (ProxyAddresses entries
    prefixed with lowercase "smtp:") is written as its own row paired with the
    primary SMTP address and recipient type.

    The resulting CSV is the authoritative alias inventory needed to:
      - Configure Gmail "Send mail as" aliases in Google Workspace
      - Migrate routing rules that rely on secondary addresses
      - Identify duplicate/conflicting aliases before cutover
      - Audit aliases that point to shared mailboxes or groups

    X400, X500, SIP, EUM, and other non-SMTP proxy addresses are excluded by
    default. Use -IncludeNonSmtp to retain them.

.PARAMETER ParamFile
    Path to MigrationConnectionParams.json from Setup-PnPCertAuth.ps1.

.PARAMETER TenantName / ClientId / CertificatePath
    Explicit connection parameters if not using -ParamFile.

.PARAMETER CertificatePassword
    SecureString password for the .pfx certificate.

.PARAMETER OutputCsv
    Path for the output CSV. Defaults to .\All-EmailAliases-<ts>.csv

.PARAMETER RecipientTypes
    Comma-separated list of RecipientTypeDetails values to include.
    Defaults to all mail-enabled types. Useful to scope to a subset, e.g.:
    "UserMailbox,SharedMailbox"

.PARAMETER IncludeNonSmtp
    Include non-SMTP proxy addresses (X400, X500, SIP, EUM, etc.).
    These are normally not relevant for Google Workspace migration but
    may be needed for coexistence routing.

.PARAMETER IncludePrimaryAsAlias
    By default only secondary (lowercase smtp:) addresses are output.
    Set this switch to also emit a row for the primary address itself
    (useful for a complete address-book export).

.EXAMPLE
    .\Get-AllEmailAliases.ps1 -ParamFile ".\MigrationConnectionParams.json"

.EXAMPLE
    .\Get-AllEmailAliases.ps1 -ParamFile ".\MigrationConnectionParams.json" `
        -RecipientTypes "UserMailbox,SharedMailbox,DistributionGroup"

.EXAMPLE
    .\Get-AllEmailAliases.ps1 -ParamFile ".\MigrationConnectionParams.json" `
        -IncludePrimaryAsAlias -OutputCsv ".\Full-AddressBook.csv"

.NOTES
    Requires: ExchangeOnlineManagement module v3+
    App permissions: Exchange.ManageAsApp + Exchange Administrator role on the service principal.

    Get-EXORecipient is used instead of Get-Recipient because it uses REST
    (faster, lower throttling risk) and supports the same ResultSize Unlimited.
#>

[CmdletBinding()]
param(
    [string]$ParamFile,

    [string]$TenantName,
    [string]$ClientId,
    [string]$CertificatePath,
    [SecureString]$CertificatePassword,

    [string]$OutputCsv,

    [string]$RecipientTypes,

    [switch]$IncludeNonSmtp,

    [switch]$IncludePrimaryAsAlias
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
if (-not $OutputCsv) { $OutputCsv = Join-Path $PSScriptRoot "All-EmailAliases-$ts.csv" }

$tenantDomain = if ($TenantName -match '\.') { $TenantName } else { "$TenantName.onmicrosoft.com" }

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

#region ── Retrieve All Recipients ───────────────────────────────────────────

Write-Host "[INFO] Retrieving all mail-enabled recipients..." -ForegroundColor Cyan

$recipientParams = @{
    ResultSize  = 'Unlimited'
    Properties  = @('DisplayName','PrimarySmtpAddress','RecipientTypeDetails','EmailAddresses','HiddenFromAddressListsEnabled')
}

if ($RecipientTypes) {
    $recipientParams['RecipientTypeDetails'] = ($RecipientTypes -split ',' | ForEach-Object { $_.Trim() })
}

$allRecipients = Get-EXORecipient @recipientParams

Write-Host "[INFO] Retrieved $($allRecipients.Count) recipients" -ForegroundColor Cyan

#endregion

#region ── Build Alias Rows ──────────────────────────────────────────────────

Write-Host "[INFO] Expanding proxy addresses..." -ForegroundColor Cyan

$rows    = [System.Collections.Generic.List[PSCustomObject]]::new()
$counter = 0

foreach ($r in $allRecipients) {
    $counter++
    if ($counter % 200 -eq 0) {
        Write-Progress -Activity "Expanding aliases" `
            -Status "$counter / $($allRecipients.Count)" `
            -PercentComplete (($counter / $allRecipients.Count) * 100)
    }

    $primarySmtp = $r.PrimarySmtpAddress

    foreach ($addr in $r.EmailAddresses) {
        $prefix = $addr -replace ':.*',''          # e.g. "smtp", "SMTP", "X500", "SIP"
        $value  = $addr -replace '^[^:]+:',''      # strip the prefix

        $isPrimary   = $addr -cmatch '^SMTP:'      # uppercase = primary
        $isSmtp      = $prefix -ieq 'smtp'

        # Skip non-SMTP unless requested
        if (-not $isSmtp -and -not $IncludeNonSmtp) { continue }

        # Skip primary address rows unless requested
        if ($isPrimary -and -not $IncludePrimaryAsAlias) { continue }

        $rows.Add([PSCustomObject]@{
            PrimarySmtpAddress    = $primarySmtp
            DisplayName           = $r.DisplayName
            RecipientType         = $r.RecipientTypeDetails
            AliasAddress          = $value
            AddressType           = if ($isPrimary) { 'Primary' } else { $prefix.ToUpper() }
            HiddenFromGAL         = $r.HiddenFromAddressListsEnabled
        })
    }
}

Write-Progress -Activity "Expanding aliases" -Completed

#endregion

#region ── Export and Summary ────────────────────────────────────────────────

$rows | Sort-Object PrimarySmtpAddress, AliasAddress |
    Export-Csv -Path $OutputCsv -NoTypeInformation -Encoding UTF8

# Summary by recipient type
$byType = $rows | Group-Object RecipientType | ForEach-Object {
    [PSCustomObject]@{
        RecipientType = $_.Name
        AliasCount    = $_.Count
        Recipients    = ($_.Group | Select-Object PrimarySmtpAddress -Unique | Measure-Object).Count
    }
} | Sort-Object AliasCount -Descending

Write-Host "`n===== ALIAS SUMMARY =====" -ForegroundColor Cyan
$byType | Format-Table -AutoSize

$totalRecipients = ($rows | Select-Object PrimarySmtpAddress -Unique | Measure-Object).Count
Write-Host "Recipients with at least one alias : $totalRecipients"
Write-Host "Total alias rows                   : $($rows.Count)"

# Flag duplicates — same alias address appearing on more than one recipient
$duplicates = $rows | Group-Object AliasAddress | Where-Object { $_.Count -gt 1 }
if ($duplicates.Count -gt 0) {
    Write-Warning "$($duplicates.Count) alias address(es) appear on more than one recipient (potential conflict):"
    $duplicates | ForEach-Object {
        Write-Warning "  $($_.Name) → $($_.Group.PrimarySmtpAddress -join ', ')"
    }
    $dupCsv = $OutputCsv -replace '\.csv$', '-Duplicates.csv'
    $duplicates | ForEach-Object { $_.Group } |
        Export-Csv -Path $dupCsv -NoTypeInformation -Encoding UTF8
    Write-Warning "Duplicate rows saved to: $dupCsv"
}

Write-Host "`n[OK] Output CSV: $OutputCsv" -ForegroundColor Green

Disconnect-ExchangeOnline -Confirm:$false

#endregion
