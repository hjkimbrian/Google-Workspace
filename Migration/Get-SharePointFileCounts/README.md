# Get-SharePointFileCounts

Inventories all SharePoint Online site collections with file counts and storage usage. Essential for sizing a migration from SharePoint to Google Drive / Shared Drives.

## Output Fields

| Column | Description |
|---|---|
| `Title` | Site collection display name |
| `Url` | Site URL |
| `Template` | Site template (e.g. `STS#3`, `TEAMCHANNEL#1`) |
| `StorageUsedMB` / `StorageUsedGB` | Current storage consumption |
| `StorageQuotaGB` | Configured storage quota |
| `StorageUsedPercent` | Percentage of quota used |
| `LastContentModified` | Last time content was changed |
| `SharingCapability` | External sharing setting |
| `IsHubSite` | Whether this is a Hub Site |
| `LibraryCount` | Number of document libraries (with `-SkipLibraryDetail` omitted) |
| `TotalFileCount` | Total files across all libraries (with `-SkipLibraryDetail` omitted) |

---

## Prerequisites

- [`PnP.PowerShell`](https://pnp.github.io/powershell/) module installed
- Azure AD App Registration with:
  - `Sites.FullControl.All` (SharePoint) application permission
  - `Sites.Read.All` (Microsoft Graph) application permission
- Admin consent granted for all permissions
- Run [`Setup-PnPCertAuth.ps1`](../Setup-PnPCertAuth/) first

---

## Usage

### Quick summary (storage only, fast)

```powershell
.\Get-SharePointFileCounts.ps1 `
    -ParamFile ".\MigrationConnectionParams.json" `
    -SkipLibraryDetail
```

### Full file count (slower, most accurate)

```powershell
.\Get-SharePointFileCounts.ps1 -ParamFile ".\MigrationConnectionParams.json"
```

### Filter to specific sites

```powershell
.\Get-SharePointFileCounts.ps1 `
    -ParamFile ".\MigrationConnectionParams.json" `
    -SiteFilter "*/sites/Project*"
```

### Exclude system sites

```powershell
.\Get-SharePointFileCounts.ps1 `
    -ParamFile ".\MigrationConnectionParams.json" `
    -ExcludeSystemSites
```

---

## Parameters

| Parameter | Required | Default | Description |
|---|---|---|---|
| `ParamFile` | No* | — | Path to `MigrationConnectionParams.json` |
| `TenantName` | No* | — | Tenant name |
| `ClientId` | No* | — | Azure AD App Client ID |
| `CertificatePath` | No* | — | Path to `.pfx` file |
| `CertificatePassword` | No | *(prompted if needed)* | Cert password |
| `OutputCsv` | No | `.\SPO-FileCounts-<timestamp>.csv` | Output file path |
| `SkipLibraryDetail` | No | `$false` | Use quota data only (fast); skip per-file enumeration |
| `ExcludeSystemSites` | No | `$false` | Skip SPO system/service sites |
| `SiteFilter` | No | *(all sites)* | Wildcard URL pattern to filter sites |

\* Either `-ParamFile` or `-TenantName` + `-ClientId` + `-CertificatePath` required.

---

## Performance Guidance

| Tenant size | Recommendation | Typical runtime |
|---|---|---|
| < 100 sites | Full file enumeration (default) | 5–20 min |
| 100–500 sites | Consider `-SkipLibraryDetail` first | 1–3 min / 30–60 min |
| 500+ sites | `-SkipLibraryDetail` for overview, then drill per site | Minutes |

For large tenants, run the script with `-SkipLibraryDetail` first to identify the largest sites, then re-run targeting specific sites with `-SiteFilter` for accurate file counts.

---

## Migration Notes

| SPO Concept | Google Equivalent |
|---|---|
| Site Collection | [Shared Drive](https://support.google.com/a/answer/7212025) (for team content) |
| Document Library | Folder within a Shared Drive |
| Site with unique permissions | Shared Drive (each has its own member list) |
| Hub Site | Google Drive folder structure or label taxonomy |

- Google Shared Drives have a **400,000 item limit** — sites with more files need to be split
- Google Drive file/folder paths are limited to **~32,000 characters** — deeply nested SPO libraries may need restructuring
- Files larger than **5 TB** cannot be uploaded to Google Drive
