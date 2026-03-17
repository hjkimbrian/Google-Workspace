<#
.SYNOPSIS
    Resolves members of SharePoint Online groups found in the SPO-Groups CSV produced
    by Get-SPOSitePermissions.ps1.

.DESCRIPTION
    For each SharePoint site group recorded in the groups inventory:
      - Enumerates all group members using PnP PowerShell
      - For members that are themselves Azure AD groups (security groups, M365 groups),
        optionally expands their membership via Microsoft Graph

    Produces a single CSV with one row per group + member pairing.

.PARAMETER GroupsCsv
    Path to the SPO-Groups-<timestamp>.csv file from Get-SPOSitePermissions.ps1.
    Required.

.PARAMETER ParamFile
    Path to MigrationConnectionParams.json from Setup-PnPCertAuth.ps1.

.PARAMETER TenantName / ClientId / CertificatePath / TenantId
    Explicit connection parameters if not using -ParamFile.

.PARAMETER CertificatePassword
    SecureString password for the .pfx certificate.

.PARAMETER OutputCsv
    Path for the output CSV. Defaults to .\SPO-GroupMembers-<ts>.csv

.PARAMETER ExpandAzureAdGroups
    If set, members that are Azure AD groups (SecurityGroup or M365Group) will
    have their own membership fetched from Microsoft Graph and expanded inline.
    Requires the app to have Directory.Read.All or Group.Read.All (Graph).

.EXAMPLE
    .\Get-SPOGroupMembers.ps1 `
        -GroupsCsv ".\SPO-Groups-20260317-120000.csv" `
        -ParamFile ".\MigrationConnectionParams.json"

.EXAMPLE
    .\Get-SPOGroupMembers.ps1 `
        -GroupsCsv ".\SPO-Groups-20260317-120000.csv" `
        -ParamFile ".\MigrationConnectionParams.json" `
        -ExpandAzureAdGroups

.NOTES
    Requires: PnP.PowerShell module
    Optional: Microsoft.Graph module (for -ExpandAzureAdGroups)
    App permissions: Sites.FullControl.All (SharePoint)
    For -ExpandAzureAdGroups additionally: Group.Read.All or Directory.Read.All (Graph)
#>

[CmdletBinding()]
param(
    [Parameter(Mandatory)]
    [string]$GroupsCsv,

    [string]$ParamFile,

    [string]$TenantName,
    [string]$TenantId,
    [string]$ClientId,
    [string]$CertificatePath,
    [SecureString]$CertificatePassword,

    [string]$OutputCsv,

    [switch]$ExpandAzureAdGroups
)

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

#region ── Load Parameters ───────────────────────────────────────────────────

if (-not (Test-Path $GroupsCsv)) { throw "GroupsCsv not found: $GroupsCsv" }
$groups = Import-Csv $GroupsCsv

if ($ParamFile) {
    if (-not (Test-Path $ParamFile)) { throw "ParamFile not found: $ParamFile" }
    $p = Get-Content $ParamFile -Raw | ConvertFrom-Json
    if (-not $TenantName)      { $TenantName      = $p.TenantName }
    if (-not $ClientId)        { $ClientId        = $p.ClientId }
    if (-not $CertificatePath) { $CertificatePath = $p.CertificatePath }
    if (-not $TenantId)        { $TenantId        = $p.TenantId }
}

foreach ($v in @('TenantName','ClientId','CertificatePath')) {
    if (-not (Get-Variable $v -ValueOnly -ErrorAction SilentlyContinue)) {
        throw "Missing required parameter: -$v (or use -ParamFile)"
    }
}

$ts = Get-Date -Format "yyyyMMdd-HHmmss"
if (-not $OutputCsv) { $OutputCsv = Join-Path $PSScriptRoot "SPO-GroupMembers-$ts.csv" }

$tenantDomain = if ($TenantName -match '\.') { $TenantName } else { "$TenantName.onmicrosoft.com" }
if (-not $TenantId) { $TenantId = $tenantDomain }

#endregion

#region ── Optional: Connect to Microsoft Graph for AAD group expansion ──────

$graphConnected = $false

if ($ExpandAzureAdGroups) {
    Write-Host "[INFO] Connecting to Microsoft Graph for Azure AD group expansion..." -ForegroundColor Cyan
    try {
        Import-Module Microsoft.Graph.Authentication -ErrorAction Stop
        Import-Module Microsoft.Graph.Groups         -ErrorAction SilentlyContinue

        $cert = if ($CertificatePassword) {
            [System.Security.Cryptography.X509Certificates.X509Certificate2]::new($CertificatePath, $CertificatePassword)
        } else {
            [System.Security.Cryptography.X509Certificates.X509Certificate2]::new($CertificatePath)
        }

        Connect-MgGraph -ClientId $ClientId -TenantId $TenantId -Certificate $cert -NoWelcome
        $graphConnected = $true
        Write-Host "[OK]   Connected to Microsoft Graph" -ForegroundColor Green
    } catch {
        Write-Warning "Could not connect to Graph; Azure AD group expansion will be skipped. Error: $_"
    }
}

#endregion

#region ── Process Groups ────────────────────────────────────────────────────

$memberRows = [System.Collections.Generic.List[PSCustomObject]]::new()

# Group the entries by SiteUrl so we connect once per site
$groupsBySite = $groups | Group-Object SiteUrl

$siteCounter = 0

foreach ($siteGroup in $groupsBySite) {
    $siteUrl   = $siteGroup.Name
    $siteTitle = ($siteGroup.Group | Select-Object -First 1).SiteTitle
    $siteCounter++

    Write-Progress -Activity "Resolving SPO group members" `
        -Status "$siteCounter / $($groupsBySite.Count): $siteUrl" `
        -PercentComplete (($siteCounter / $groupsBySite.Count) * 100)

    $connectParams = @{
        Url            = $siteUrl
        ClientId       = $ClientId
        CertificatePath = $CertificatePath
        Tenant         = $tenantDomain
    }
    if ($CertificatePassword) { $connectParams['CertificatePassword'] = $CertificatePassword }

    try {
        Connect-PnPOnline @connectParams -ErrorAction Stop
    } catch {
        Write-Warning "  Cannot connect to $siteUrl: $_"
        continue
    }

    foreach ($grpEntry in $siteGroup.Group) {
        $groupName = $grpEntry.GroupName
        $groupId   = $grpEntry.GroupId

        try {
            $members = Get-PnPGroupMember -Group $groupId -ErrorAction Stop
        } catch {
            Write-Warning "  Could not get members for group '$groupName' in $siteUrl: $_"
            continue
        }

        foreach ($member in $members) {
            # Determine if this member is itself a group
            $memberIsGroup = $member.PrincipalType -in @('SecurityGroup','SharePointGroup') -or
                             $member.LoginName -match 'federateddirectoryclaimprovider' -or
                             $member.LoginName -match '^c:0t\.c\|tenant\|'

            # Classify member type
            $memberType = switch -Regex ($member.LoginName) {
                'federateddirectoryclaimprovider'      { 'M365Group'; break }
                '^c:0t\.c\|tenant\|'                  { 'SecurityGroup'; break }
                '#ext#'                                { 'ExternalUser'; break }
                '^i:0#\.'                              { 'User'; break }
                default                                { if ($memberIsGroup) { 'SPOGroup' } else { 'User' } }
            }

            $memberEmail = $member.Email
            if (-not $memberEmail -and $member.LoginName -match '\|([^|]+@[^|]+)$') {
                $memberEmail = $Matches[1]
            }

            $memberRows.Add([PSCustomObject]@{
                SiteUrl         = $siteUrl
                SiteTitle       = $siteTitle
                GroupName       = $groupName
                GroupId         = $groupId
                MemberName      = $member.Title
                MemberEmail     = $memberEmail
                MemberLoginName = $member.LoginName
                MemberType      = $memberType
                ExpandedFrom    = $null
            })

            # ── Expand Azure AD group members via Graph ─────────────────────
            if ($ExpandAzureAdGroups -and $graphConnected -and $memberIsGroup -and
                $member.LoginName -match '([0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12})') {

                $aadGroupId = $Matches[1]

                try {
                    $aadMembers = Get-MgGroupMember -GroupId $aadGroupId -All -ErrorAction Stop
                    foreach ($aadMbr in $aadMembers) {
                        $aadProps = $aadMbr.AdditionalProperties

                        $memberRows.Add([PSCustomObject]@{
                            SiteUrl         = $siteUrl
                            SiteTitle       = $siteTitle
                            GroupName       = $groupName
                            GroupId         = $groupId
                            MemberName      = $aadProps['displayName']
                            MemberEmail     = $aadProps['mail'] ?? $aadProps['userPrincipalName']
                            MemberLoginName = $aadProps['userPrincipalName']
                            MemberType      = 'User (via AAD group expansion)'
                            ExpandedFrom    = $member.Title
                        })
                    }
                } catch {
                    Write-Warning "  Could not expand AAD group '$($member.Title)' ($aadGroupId): $_"
                }
            }
        }
    }

    Disconnect-PnPOnline
}

Write-Progress -Activity "Resolving SPO group members" -Completed

if ($graphConnected) {
    Disconnect-MgGraph
}

#endregion

#region ── Export and Summary ────────────────────────────────────────────────

$memberRows | Export-Csv -Path $OutputCsv -NoTypeInformation -Encoding UTF8

$byType = $memberRows | Group-Object MemberType | ForEach-Object {
    [PSCustomObject]@{ MemberType = $_.Name; Count = $_.Count }
} | Sort-Object Count -Descending

Write-Host "`n===== SPO GROUP MEMBERS SUMMARY =====" -ForegroundColor Cyan
$byType | Format-Table -AutoSize
Write-Host "Total group member rows : $($memberRows.Count)"
Write-Host "Unique members          : $(($memberRows | Select-Object MemberEmail -Unique | Measure-Object).Count)"

Write-Host "`n[OK] Output CSV: $OutputCsv" -ForegroundColor Green

#endregion
