# Get-SPOSitePermissions

Exports SharePoint Online permissions at the root web level and for any document library with unique (broken-inheritance) permissions. Each permission row identifies the principal type so you know whether you're dealing with an individual user, an SPO site group, an Azure AD security group, or an M365 Group.

---

## Output Files

### `SPO-Permissions-<timestamp>.csv`

One row per site/library + principal pairing.

| Column | Description |
|---|---|
| `SiteTitle` | Site collection display name |
| `SiteUrl` | Site URL |
| `Scope` | `Root` (site-level) or `Library` (document library) |
| `LibraryName` | Library display name (blank for root-level rows) |
| `LibraryPath` | Full URL path to the library (blank for root-level rows) |
| `PrincipalType` | Classified type (see below) |
| `PrincipalName` | Display name of the user or group |
| `PrincipalEmail` | Email address (where resolvable) |
| `LoginName` | SharePoint login name / claim string |
| `PermissionLevel` | Permission level(s) assigned (e.g. `Full Control`, `Edit`, `Read`) |
| `HasUniquePerms` | `True` if this library has broken inheritance from the site |

### `SPO-Groups-<timestamp>.csv`

All unique SharePoint site groups discovered across all scanned sites. Pass this file to `Get-SPOGroupMembers.ps1`.

| Column | Description |
|---|---|
| `SiteUrl` | Site the group belongs to |
| `SiteTitle` | Site display name |
| `GroupName` | SPO group display name |
| `LoginName` | SharePoint login name |
| `GroupId` | Numeric SharePoint group ID |

---

## Principal Type Classification

| Type | Description | LoginName pattern |
|---|---|---|
| `User` | Individual user account | `i:0#.f\|membership\|user@...` |
| `SPOGroup` | SharePoint site group (Owners, Members, Visitors, custom) | No pipe character |
| `SecurityGroup` | Azure AD security group synced to SPO | `c:0t.c\|tenant\|<guid>` |
| `M365Group` | Microsoft 365 Group (Teams, Planner) | `c:0o.c\|federateddirectoryclaimprovider\|<guid>` |
| `ExternalUser` | Guest or external user | Contains `#ext#` |
| `Everyone` | "Everyone" or "Everyone except external users" | `c:0(.s\|true` |

---

## Prerequisites

- [`PnP.PowerShell`](https://pnp.github.io/powershell/) module installed
- Azure AD App Registration with:
  - `Sites.FullControl.All` (SharePoint) application permission
  - `Directory.Read.All` (Microsoft Graph) — for resolving group types
- Admin consent granted
- Run [`Setup-PnPCertAuth.ps1`](../Setup-PnPCertAuth/) first

---

## Usage

### All sites, root + unique library permissions

```powershell
.\Get-SPOSitePermissions.ps1 -ParamFile ".\MigrationConnectionParams.json"
```

### Skip system sites, filter to /sites/ URLs only

```powershell
.\Get-SPOSitePermissions.ps1 `
    -ParamFile ".\MigrationConnectionParams.json" `
    -SiteFilter "*/sites/*" `
    -ExcludeSystemSites
```

### Include libraries even with inherited permissions

```powershell
.\Get-SPOSitePermissions.ps1 `
    -ParamFile ".\MigrationConnectionParams.json" `
    -IncludeInheritedLibraries
```

---

## Parameters

| Parameter | Required | Default | Description |
|---|---|---|---|
| `ParamFile` | No* | — | Path to `MigrationConnectionParams.json` |
| `TenantName` | No* | — | Tenant name |
| `ClientId` | No* | — | Azure AD App Client ID |
| `CertificatePath` | No* | — | Path to `.pfx` file |
| `OutputPermissionsCsv` | No | `.\SPO-Permissions-<ts>.csv` | Permissions output path |
| `OutputGroupsCsv` | No | `.\SPO-Groups-<ts>.csv` | SPO groups output path |
| `SiteFilter` | No | *(all sites)* | Wildcard URL pattern |
| `ExcludeSystemSites` | No | `$false` | Skip SPO system/service sites |
| `IncludeInheritedLibraries` | No | `$false` | Include libraries inheriting from site |

---

## Migration Mapping

| SPO Permission | Google Drive equivalent |
|---|---|
| Full Control | Manager (Shared Drive) |
| Design | Manager |
| Edit / Contribute | Contributor |
| Read | Viewer |
| View Only | Viewer |

### SPO Groups → Google Groups

SPO site groups (Owners, Members, Visitors) do not have a direct equivalent in Google Drive. In a Shared Drive migration:

- **Owners group** → Shared Drive **Manager**
- **Members group** → Shared Drive **Contributor**
- **Visitors group** → Shared Drive **Viewer**

If the SPO groups contain security groups, expand them using [`Get-SPOGroupMembers.ps1`](../Get-SPOGroupMembers/).

---

## Note on OneDrive Sharing Permissions

This script covers SharePoint site collections. **OneDrive sharing links** (anonymous links, specific-person shares set at the file/folder level) are a separate permission layer not captured here. Users frequently share OneDrive content directly without going through site permissions. Consider auditing these separately using the SharePoint Online Sharing report in the M365 Admin Center (`Reports > Usage > SharePoint > Sharing activity`), or the `Get-SPOSiteGroup` + `Get-SPOUser` cmdlets per OneDrive site.
