<#
.SYNOPSIS
    Exports SharePoint Online site permissions at the root web and document library level,
    identifying whether each grant is to a user, SPO group, security group, or M365 group.

.DESCRIPTION
    Uses PnP PowerShell with certificate-based (app-only) authentication to enumerate:

    1. Root site collection permissions for each site
    2. Document library permissions for libraries that have UNIQUE (broken inheritance) permissions

    For each permission entry, the principal type is classified as:
      - User          — individual user account
      - SPOGroup      — SharePoint site group (Owners, Members, Visitors, or custom)
      - SecurityGroup — Azure AD security group or mail-enabled security group
      - M365Group     — Microsoft 365 Group (Teams/Planner connected)
      - ExternalUser  — user outside the tenant (guest/external sharing)
      - Everyone      — "Everyone" or "Everyone except external users"

    Two CSVs are produced:
      - Site + library permissions (one row per site/library + principal pairing)
      - SPO groups found (for use by Get-SPOGroupMembers.ps1)

.PARAMETER ParamFile
    Path to MigrationConnectionParams.json from Setup-PnPCertAuth.ps1.

.PARAMETER TenantName / ClientId / CertificatePath
    Explicit connection parameters if not using -ParamFile.

.PARAMETER CertificatePassword
    SecureString password for the .pfx certificate.

.PARAMETER OutputPermissionsCsv
    Path for the permissions output. Defaults to .\SPO-Permissions-<ts>.csv

.PARAMETER OutputGroupsCsv
    Path for the SPO groups inventory. Defaults to .\SPO-Groups-<ts>.csv

.PARAMETER SiteFilter
    Optional wildcard to limit which sites are processed.
    Example: "*/sites/Project*"

.PARAMETER ExcludeSystemSites
    Skip SPO system/service sites (Search, MySite host, App Catalog, etc.).

.PARAMETER IncludeInheritedLibraries
    By default only libraries with unique (broken-inheritance) permissions are
    enumerated. Set this switch to include all libraries regardless.

.EXAMPLE
    .\Get-SPOSitePermissions.ps1 -ParamFile ".\MigrationConnectionParams.json"

.EXAMPLE
    .\Get-SPOSitePermissions.ps1 -ParamFile ".\MigrationConnectionParams.json" `
        -SiteFilter "*/sites/*" -ExcludeSystemSites

.NOTES
    Requires: PnP.PowerShell module
    App permissions: Sites.FullControl.All (SharePoint) + Directory.Read.All (Graph)

    The Get-SPOGroupMembers.ps1 script consumes the SPO-Groups CSV produced here.
#>

[CmdletBinding()]
param(
    [string]$ParamFile,

    [string]$TenantName,
    [string]$ClientId,
    [string]$CertificatePath,
    [SecureString]$CertificatePassword,

    [string]$OutputPermissionsCsv,
    [string]$OutputGroupsCsv,

    [string]$SiteFilter,

    [switch]$ExcludeSystemSites,

    [switch]$IncludeInheritedLibraries
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
if (-not $OutputPermissionsCsv) { $OutputPermissionsCsv = Join-Path $PSScriptRoot "SPO-Permissions-$ts.csv" }
if (-not $OutputGroupsCsv)      { $OutputGroupsCsv      = Join-Path $PSScriptRoot "SPO-Groups-$ts.csv"      }

$tenantDomain = if ($TenantName -match '\.') { $TenantName } else { "$TenantName.onmicrosoft.com" }
$adminUrl     = "https://$TenantName-admin.sharepoint.com"

#endregion

#region ── Helper: Classify Principal Type ────────────────────────────────────

function Get-PrincipalType {
    param($RoleAssignment)

    $member = $RoleAssignment.Member

    # PnP returns different object types for users vs groups
    $loginName = $member.LoginName

    if ($loginName -match '^i:0#\.f\|membership\|') {
        # Federated user (forms-based/external)
        return 'ExternalUser'
    }
    if ($loginName -match '^i:0#\.f\|membership\|' -or $loginName -match '#ext#') {
        return 'ExternalUser'
    }
    if ($loginName -eq 'c:0(.s|true' -or $loginName -match 'spo-grid-all-users' -or
        $loginName -match 'Everyone') {
        return 'Everyone'
    }
    # M365 Group / Security Group (Azure AD groups synced to SPO)
    # LoginName pattern: c:0o.c|federateddirectoryclaimprovider|<guid>
    if ($loginName -match 'federateddirectoryclaimprovider') {
        return 'M365Group'
    }
    # Azure AD security group
    if ($loginName -match '^c:0t\.c\|tenant\|') {
        return 'SecurityGroup'
    }
    # SPO Site Group
    if ($member.PrincipalType -eq 'SharePointGroup' -or
        ($null -ne ($member | Get-Member -Name 'Id' -ErrorAction SilentlyContinue) -and
         $loginName -notmatch '\|')) {
        return 'SPOGroup'
    }
    # Individual user
    if ($loginName -match '^i:0#\.' -or $loginName -match '@') {
        return 'User'
    }

    return 'Unknown'
}

function Get-PrincipalEmail {
    param($RoleAssignment)
    $member = $RoleAssignment.Member
    # Try Email property first, fall back to LoginName parsing
    if ($member.Email -and $member.Email -ne '') {
        return $member.Email
    }
    # Parse UPN from login name like i:0#.f|membership|user@contoso.com
    if ($member.LoginName -match '\|([^|]+@[^|]+)$') {
        return $Matches[1]
    }
    return $null
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
Write-Host "[OK]   Connected to $adminUrl" -ForegroundColor Green

#endregion

#region ── Enumerate Sites ───────────────────────────────────────────────────

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

Disconnect-PnPOnline

#endregion

#region ── Enumerate Permissions Per Site ────────────────────────────────────

$permRows   = [System.Collections.Generic.List[PSCustomObject]]::new()
$groupRows  = [System.Collections.Generic.List[PSCustomObject]]::new()
$seenGroups = [System.Collections.Generic.HashSet[string]]::new()

$siteCounter = 0

foreach ($site in $allSites) {
    $siteCounter++
    Write-Progress -Activity "Scanning site permissions" `
        -Status "$siteCounter / $($allSites.Count): $($site.Url)" `
        -PercentComplete (($siteCounter / $allSites.Count) * 100)

    # Connect to individual site
    $siteConnectParams = @{
        Url            = $site.Url
        ClientId       = $ClientId
        CertificatePath = $CertificatePath
        Tenant         = $tenantDomain
    }
    if ($CertificatePassword) { $siteConnectParams['CertificatePassword'] = $CertificatePassword }

    try {
        Connect-PnPOnline @siteConnectParams -ErrorAction Stop
    } catch {
        Write-Warning "  Cannot connect to $($site.Url): $_"
        continue
    }

    # ── Root Web Permissions ────────────────────────────────────────────────

    try {
        $web = Get-PnPWeb -Includes RoleAssignments, RoleAssignments.Member,
            RoleAssignments.RoleDefinitionBindings, HasUniqueRoleAssignments -ErrorAction Stop

        foreach ($ra in $web.RoleAssignments) {
            $ctx = Get-PnPContext
            $ctx.Load($ra.Member)
            $ctx.Load($ra.RoleDefinitionBindings)
            $ctx.ExecuteQuery()

            $principalType = Get-PrincipalType -RoleAssignment $ra
            $principalEmail = Get-PrincipalEmail -RoleAssignment $ra
            $principalName  = $ra.Member.Title

            $permRows.Add([PSCustomObject]@{
                SiteTitle        = $site.Title
                SiteUrl          = $site.Url
                Scope            = 'Root'
                LibraryName      = ''
                LibraryPath      = ''
                PrincipalType    = $principalType
                PrincipalName    = $principalName
                PrincipalEmail   = $principalEmail
                LoginName        = $ra.Member.LoginName
                PermissionLevel  = ($ra.RoleDefinitionBindings | Where-Object { $_.Name -ne 'Limited Access' } | ForEach-Object { $_.Name }) -join '; '
                HasUniquePerms   = $true   # Root always shown
            })

            # Track SPO groups for the groups inventory
            if ($principalType -eq 'SPOGroup') {
                $groupKey = "$($site.Url)|$($ra.Member.LoginName)"
                if ($seenGroups.Add($groupKey)) {
                    $groupRows.Add([PSCustomObject]@{
                        SiteUrl   = $site.Url
                        SiteTitle = $site.Title
                        GroupName = $principalName
                        LoginName = $ra.Member.LoginName
                        GroupId   = $ra.Member.Id
                    })
                }
            }
        }
    } catch {
        Write-Warning "  Root permissions failed for $($site.Url): $_"
    }

    # ── Document Library Permissions (unique only, unless -IncludeInheritedLibraries) ─

    try {
        $lists = Get-PnPList -ErrorAction Stop | Where-Object {
            $_.BaseType -eq 'DocumentLibrary' -and
            -not $_.Hidden -and
            $_.Title -notin @('Form Templates','Style Library','Site Assets','_catalogs')
        }

        foreach ($list in $lists) {
            # Load HasUniqueRoleAssignments
            $ctx = Get-PnPContext
            $ctx.Load($list, { $_.HasUniqueRoleAssignments }, { $_.RootFolder.ServerRelativeUrl })
            $ctx.ExecuteQuery()

            $hasUnique = $list.HasUniqueRoleAssignments
            if (-not $hasUnique -and -not $IncludeInheritedLibraries) { continue }

            $libPath = $list.RootFolder.ServerRelativeUrl

            $ctx.Load($list.RoleAssignments)
            $ctx.ExecuteQuery()

            foreach ($ra in $list.RoleAssignments) {
                $ctx.Load($ra.Member)
                $ctx.Load($ra.RoleDefinitionBindings)
                $ctx.ExecuteQuery()

                $principalType  = Get-PrincipalType -RoleAssignment $ra
                $principalEmail = Get-PrincipalEmail -RoleAssignment $ra
                $principalName  = $ra.Member.Title

                $permRows.Add([PSCustomObject]@{
                    SiteTitle       = $site.Title
                    SiteUrl         = $site.Url
                    Scope           = 'Library'
                    LibraryName     = $list.Title
                    LibraryPath     = "$($site.Url)$libPath"
                    PrincipalType   = $principalType
                    PrincipalName   = $principalName
                    PrincipalEmail  = $principalEmail
                    LoginName       = $ra.Member.LoginName
                    PermissionLevel = ($ra.RoleDefinitionBindings | Where-Object { $_.Name -ne 'Limited Access' } | ForEach-Object { $_.Name }) -join '; '
                    HasUniquePerms  = $hasUnique
                })

                if ($principalType -eq 'SPOGroup') {
                    $groupKey = "$($site.Url)|$($ra.Member.LoginName)"
                    if ($seenGroups.Add($groupKey)) {
                        $groupRows.Add([PSCustomObject]@{
                            SiteUrl   = $site.Url
                            SiteTitle = $site.Title
                            GroupName = $principalName
                            LoginName = $ra.Member.LoginName
                            GroupId   = $ra.Member.Id
                        })
                    }
                }
            }
        }
    } catch {
        Write-Warning "  Library permissions scan failed for $($site.Url): $_"
    }

    Disconnect-PnPOnline
}

Write-Progress -Activity "Scanning site permissions" -Completed

#endregion

#region ── Export and Summary ────────────────────────────────────────────────

$permRows  | Export-Csv -Path $OutputPermissionsCsv -NoTypeInformation -Encoding UTF8
$groupRows | Export-Csv -Path $OutputGroupsCsv      -NoTypeInformation -Encoding UTF8

$byType = $permRows | Group-Object PrincipalType | ForEach-Object {
    [PSCustomObject]@{ PrincipalType = $_.Name; Count = $_.Count }
} | Sort-Object Count -Descending

Write-Host "`n===== SPO PERMISSIONS SUMMARY =====" -ForegroundColor Cyan
$byType | Format-Table -AutoSize

$libsWithUniquePerms = ($permRows | Where-Object { $_.Scope -eq 'Library' -and $_.HasUniquePerms } |
    Select-Object LibraryPath -Unique | Measure-Object).Count

Write-Host "Total permission rows          : $($permRows.Count)"
Write-Host "Libraries with unique perms    : $libsWithUniquePerms"
Write-Host "SPO groups found               : $($groupRows.Count)"
Write-Host "`nPass '$OutputGroupsCsv' to Get-SPOGroupMembers.ps1 to resolve group members."

Write-Host "`n[OK] Permissions CSV : $OutputPermissionsCsv" -ForegroundColor Green
Write-Host "[OK] SPO Groups CSV  : $OutputGroupsCsv"      -ForegroundColor Green

#endregion
