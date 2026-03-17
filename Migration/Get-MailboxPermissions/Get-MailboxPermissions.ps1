<#
.SYNOPSIS
    Exports mailbox permissions (Full Access, Send As, Send on Behalf) for Shared,
    Room, and Equipment mailboxes. Also produces a resource mailbox inventory CSV.

.DESCRIPTION
    Connects to Exchange Online using certificate-based (app-only) authentication and
    produces three output files:

    1. Permissions CSV — one row per mailbox + principal + permission type:
         - Full Access    (Get-EXOMailboxPermission)
         - Send As        (Get-EXORecipientPermission / AccessRights = SendAs)
         - Send on Behalf (GrantSendOnBehalfTo property on the mailbox object)

    2. Resource Inventory CSV — all Room and Equipment mailboxes with:
         - Display name, UPN, resource type, capacity, location, building
         - Booking policy: auto-accept, booking window, max duration

    Excludes SELF and inherited NT AUTHORITY\SELF entries. System accounts
    (NT AUTHORITY\*, SELF) are filtered by default.

.PARAMETER ParamFile
    Path to MigrationConnectionParams.json from Setup-PnPCertAuth.ps1.

.PARAMETER TenantName / ClientId / CertificatePath
    Explicit connection parameters if not using -ParamFile.

.PARAMETER CertificatePassword
    SecureString password for the .pfx certificate.

.PARAMETER MailboxTypes
    Comma-separated mailbox types to include.
    Valid: SharedMailbox, RoomMailbox, EquipmentMailbox
    Default: SharedMailbox,RoomMailbox,EquipmentMailbox

.PARAMETER OutputPermissionsCsv
    Path for the permissions output. Defaults to .\Mailbox-Permissions-<ts>.csv

.PARAMETER OutputResourcesCsv
    Path for the resource inventory. Defaults to .\Resource-Mailboxes-<ts>.csv

.PARAMETER IncludeDenyPermissions
    Include explicit Deny entries (usually not needed for migration).

.EXAMPLE
    .\Get-MailboxPermissions.ps1 -ParamFile ".\MigrationConnectionParams.json"

.EXAMPLE
    .\Get-MailboxPermissions.ps1 -ParamFile ".\MigrationConnectionParams.json" `
        -MailboxTypes "SharedMailbox"

.NOTES
    Requires: ExchangeOnlineManagement module v3+
    App permissions: Exchange.ManageAsApp + Exchange Administrator role on the service principal.
#>

[CmdletBinding()]
param(
    [string]$ParamFile,

    [string]$TenantName,
    [string]$ClientId,
    [string]$CertificatePath,
    [SecureString]$CertificatePassword,

    [string]$MailboxTypes = "SharedMailbox,RoomMailbox,EquipmentMailbox",

    [string]$OutputPermissionsCsv,
    [string]$OutputResourcesCsv,

    [switch]$IncludeDenyPermissions
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
if (-not $OutputPermissionsCsv) { $OutputPermissionsCsv = Join-Path $PSScriptRoot "Mailbox-Permissions-$ts.csv" }
if (-not $OutputResourcesCsv)   { $OutputResourcesCsv   = Join-Path $PSScriptRoot "Resource-Mailboxes-$ts.csv"  }

$tenantDomain  = if ($TenantName -match '\.') { $TenantName } else { "$TenantName.onmicrosoft.com" }
$requestedTypes = $MailboxTypes -split ',' | ForEach-Object { $_.Trim() }

# System principals to exclude from permission output
$excludePrincipals = @('NT AUTHORITY\SELF','SELF','S-1-5-10')

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

#region ── Retrieve Mailboxes ────────────────────────────────────────────────

Write-Host "[INFO] Retrieving mailboxes (types: $($requestedTypes -join ', '))..." -ForegroundColor Cyan

$mailboxes = Get-EXOMailbox -RecipientTypeDetails $requestedTypes -ResultSize Unlimited `
    -PropertySets Minimum,Delivery,Resource,Moderation

Write-Host "[INFO] Found $($mailboxes.Count) mailboxes" -ForegroundColor Cyan

#endregion

#region ── Collect Permissions ───────────────────────────────────────────────

$permRows    = [System.Collections.Generic.List[PSCustomObject]]::new()
$counter     = 0

foreach ($mbx in $mailboxes) {
    $counter++
    Write-Progress -Activity "Collecting mailbox permissions" `
        -Status "$counter / $($mailboxes.Count): $($mbx.PrimarySmtpAddress)" `
        -PercentComplete (($counter / $mailboxes.Count) * 100)

    $mbxInfo = [ordered]@{
        MailboxDisplayName  = $mbx.DisplayName
        MailboxUPN          = $mbx.UserPrincipalName
        MailboxSmtp         = $mbx.PrimarySmtpAddress
        MailboxType         = $mbx.RecipientTypeDetails
    }

    # ── 1. Full Access ────────────────────────────────────────────────────

    try {
        $faPerms = Get-EXOMailboxPermission -Identity $mbx.ExchangeGuid -ErrorAction Stop |
            Where-Object {
                $_.AccessRights -contains 'FullAccess' -and
                $_.User -notin $excludePrincipals -and
                $_.User -notmatch '^NT AUTHORITY' -and
                (-not $_.IsInherited -or $IncludeDenyPermissions) -and
                (-not $IncludeDenyPermissions -or $_.Deny -eq $false)
            }

        foreach ($perm in $faPerms) {
            if ($perm.Deny -and -not $IncludeDenyPermissions) { continue }
            $permRows.Add([PSCustomObject]($mbxInfo + [ordered]@{
                PermissionType = 'FullAccess'
                Principal      = $perm.User
                PrincipalEmail = $perm.User   # User field contains identity; SMTP resolved where possible
                IsInherited    = $perm.IsInherited
                Deny           = $perm.Deny
            }))
        }
    } catch {
        Write-Warning "  FullAccess query failed for $($mbx.PrimarySmtpAddress): $_"
    }

    # ── 2. Send As ────────────────────────────────────────────────────────

    try {
        $saPerms = Get-EXORecipientPermission -Identity $mbx.PrimarySmtpAddress `
            -AccessRights SendAs -ErrorAction Stop |
            Where-Object {
                $_.Trustee -notin $excludePrincipals -and
                $_.Trustee -notmatch '^NT AUTHORITY'
            }

        foreach ($perm in $saPerms) {
            $permRows.Add([PSCustomObject]($mbxInfo + [ordered]@{
                PermissionType = 'SendAs'
                Principal      = $perm.Trustee
                PrincipalEmail = $perm.Trustee
                IsInherited    = $perm.IsInherited
                Deny           = $false
            }))
        }
    } catch {
        Write-Warning "  SendAs query failed for $($mbx.PrimarySmtpAddress): $_"
    }

    # ── 3. Send on Behalf ─────────────────────────────────────────────────

    if ($mbx.GrantSendOnBehalfTo -and $mbx.GrantSendOnBehalfTo.Count -gt 0) {
        foreach ($delegate in $mbx.GrantSendOnBehalfTo) {
            # GrantSendOnBehalfTo returns distinguished names; resolve to SMTP where possible
            $resolved = $null
            try {
                $resolved = Get-EXORecipient -Identity $delegate -ErrorAction Stop
            } catch { }

            $permRows.Add([PSCustomObject]($mbxInfo + [ordered]@{
                PermissionType = 'SendOnBehalf'
                Principal      = if ($resolved) { $resolved.DisplayName } else { $delegate }
                PrincipalEmail = if ($resolved) { $resolved.PrimarySmtpAddress } else { $delegate }
                IsInherited    = $false
                Deny           = $false
            }))
        }
    }
}

Write-Progress -Activity "Collecting mailbox permissions" -Completed

#endregion

#region ── Resource Mailbox Inventory ────────────────────────────────────────

$resourceTypes = $requestedTypes | Where-Object { $_ -in @('RoomMailbox','EquipmentMailbox') }

$resourceRows = [System.Collections.Generic.List[PSCustomObject]]::new()

if ($resourceTypes.Count -gt 0) {
    Write-Host "[INFO] Building resource mailbox inventory..." -ForegroundColor Cyan

    $resources = $mailboxes | Where-Object { $_.RecipientTypeDetails -in $resourceTypes }

    $rcounter = 0
    foreach ($res in $resources) {
        $rcounter++
        Write-Progress -Activity "Getting resource details" `
            -Status "$rcounter / $($resources.Count): $($res.PrimarySmtpAddress)" `
            -PercentComplete (($rcounter / $resources.Count) * 100)

        $cal = $null
        try {
            $cal = Get-CalendarProcessing -Identity $res.ExchangeGuid -ErrorAction Stop
        } catch {
            Write-Warning "  Could not get CalendarProcessing for $($res.PrimarySmtpAddress): $_"
        }

        $resourceRows.Add([PSCustomObject]@{
            DisplayName           = $res.DisplayName
            UserPrincipalName     = $res.UserPrincipalName
            PrimarySmtpAddress    = $res.PrimarySmtpAddress
            ResourceType          = $res.RecipientTypeDetails
            Capacity              = $res.ResourceCapacity
            Office                = $res.Office
            City                  = $res.City
            Building              = $res.CustomAttribute1   # Convention varies — adjust per tenant
            AutoAccept            = if ($cal) { $cal.AutomateProcessing -eq 'AutoAccept' } else { $null }
            AllowConflicts        = if ($cal) { $cal.AllowConflicts } else { $null }
            BookingWindowInDays   = if ($cal) { $cal.BookingWindowInDays } else { $null }
            MaximumDurationInMinutes = if ($cal) { $cal.MaximumDurationInMinutes } else { $null }
            AllowRecurringMeetings = if ($cal) { $cal.AllowRecurringMeetings } else { $null }
            ResourceDelegates     = if ($cal) { ($cal.ResourceDelegates -join '; ') } else { $null }
            AdditionalResponse    = if ($cal) { $cal.AdditionalResponse } else { $null }
        })
    }

    Write-Progress -Activity "Getting resource details" -Completed
}

#endregion

#region ── Export and Summary ────────────────────────────────────────────────

$permRows    | Export-Csv -Path $OutputPermissionsCsv -NoTypeInformation -Encoding UTF8
$resourceRows | Export-Csv -Path $OutputResourcesCsv  -NoTypeInformation -Encoding UTF8

$byType = $permRows | Group-Object PermissionType | ForEach-Object {
    [PSCustomObject]@{ PermissionType = $_.Name; Count = $_.Count }
}

Write-Host "`n===== PERMISSION SUMMARY =====" -ForegroundColor Cyan
$byType | Format-Table -AutoSize
Write-Host "Total permission rows : $($permRows.Count)"
Write-Host "Total resource mailboxes : $($resourceRows.Count)"

Write-Host "`n[OK] Permissions CSV : $OutputPermissionsCsv" -ForegroundColor Green
Write-Host "[OK] Resources CSV   : $OutputResourcesCsv"    -ForegroundColor Green

Disconnect-ExchangeOnline -Confirm:$false

#endregion
