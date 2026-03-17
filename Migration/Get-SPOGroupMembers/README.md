# Get-SPOGroupMembers

Resolves the membership of every SharePoint Online site group discovered by `Get-SPOSitePermissions.ps1`. Produces a flat CSV with one row per group + member pairing. Optionally expands Azure AD group members via Microsoft Graph so you get individual users even when the SPO group contains a security group.

---

## Workflow

```
Get-SPOSitePermissions.ps1
        │
        └─→ SPO-Groups-<ts>.csv
                │
                ▼
        Get-SPOGroupMembers.ps1
                │
                └─→ SPO-GroupMembers-<ts>.csv
```

---

## Output: `SPO-GroupMembers-<timestamp>.csv`

One row per SPO group + member pairing.

| Column | Description |
|---|---|
| `SiteUrl` | SharePoint site the group belongs to |
| `SiteTitle` | Site display name |
| `GroupName` | SharePoint group name (e.g. `Project Alpha Owners`) |
| `GroupId` | Numeric SPO group ID |
| `MemberName` | Display name of the member |
| `MemberEmail` | Email address of the member (where resolvable) |
| `MemberLoginName` | SharePoint claim / login string |
| `MemberType` | Classified type (see below) |
| `ExpandedFrom` | Populated when `-ExpandAzureAdGroups` is used; shows which AAD group the user came from |

### Member Types

| Type | Description |
|---|---|
| `User` | Individual user account |
| `SPOGroup` | Nested SharePoint group |
| `SecurityGroup` | Azure AD security group (not yet expanded) |
| `M365Group` | Microsoft 365 Group (not yet expanded) |
| `ExternalUser` | Guest / external sharing user |
| `User (via AAD group expansion)` | Individual user resolved from an Azure AD group |

---

## Prerequisites

- [`PnP.PowerShell`](https://pnp.github.io/powershell/) module
- Azure AD App Registration with `Sites.FullControl.All` (SharePoint)
- For `-ExpandAzureAdGroups`: additionally `Group.Read.All` or `Directory.Read.All` (Graph)
- Run [`Setup-PnPCertAuth.ps1`](../Setup-PnPCertAuth/) first
- Run [`Get-SPOSitePermissions.ps1`](../Get-SPOSitePermissions/) first to generate the groups CSV

---

## Usage

### Basic — resolve SPO group members only

```powershell
.\Get-SPOGroupMembers.ps1 `
    -GroupsCsv ".\SPO-Groups-20260317-120000.csv" `
    -ParamFile ".\MigrationConnectionParams.json"
```

### With Azure AD group expansion (individual users from nested security groups)

```powershell
.\Get-SPOGroupMembers.ps1 `
    -GroupsCsv ".\SPO-Groups-20260317-120000.csv" `
    -ParamFile ".\MigrationConnectionParams.json" `
    -ExpandAzureAdGroups
```

---

## Parameters

| Parameter | Required | Default | Description |
|---|---|---|---|
| `GroupsCsv` | **Yes** | — | Path to `SPO-Groups-<ts>.csv` from `Get-SPOSitePermissions.ps1` |
| `ParamFile` | No* | — | Path to `MigrationConnectionParams.json` |
| `TenantName` | No* | — | Tenant name |
| `TenantId` | No | *(derived)* | Azure AD Tenant GUID (needed for Graph) |
| `ClientId` | No* | — | Azure AD App Client ID |
| `CertificatePath` | No* | — | Path to `.pfx` file |
| `OutputCsv` | No | `.\SPO-GroupMembers-<ts>.csv` | Output path |
| `ExpandAzureAdGroups` | No | `$false` | Expand AAD group members via Microsoft Graph |

---

## Migration Notes

### Recreating SPO Group Access in Google Shared Drives

SPO site groups map to Shared Drive membership. Use the resolved member list to assign access:

```bash
# Using GAM — add Shared Drive members from the CSV
gam csv SPO-GroupMembers-*.csv \
    gam add drivefileacl <sharedDriveId> user ~MemberEmail role contributor
```

### When to Use `-ExpandAzureAdGroups`

Use it when:
- Your SPO groups contain **nested Azure AD security groups** (common in large enterprises)
- You need a **flat user list** for Google Workspace bulk provisioning or Drive permissions
- You want to verify actual headcount before purchasing Google Workspace seats

Skip it when:
- You plan to recreate the same group structure in Google (e.g. using Google Groups backed by Azure AD sync)
- You just need to understand the high-level group structure

### Handling Nested SPO Groups

If a member appears with `MemberType = SPOGroup`, it is a nested SharePoint group. The script resolves one level of SPO groups. To fully flatten deeply nested groups, re-run the script with a groups CSV containing those nested group IDs.
