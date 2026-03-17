# Get-OneDriveFileCounts

Inventories all OneDrive for Business accounts with file counts and storage usage. Use this to scope the OneDrive-to-Google Drive (My Drive) portion of a migration.

## Output Fields

| Column | Description |
|---|---|
| `OwnerDisplayName` | User's display name |
| `OwnerUPN` | User Principal Name |
| `SiteUrl` | OneDrive personal site URL |
| `StorageUsedMB` / `StorageUsedGB` | Storage consumed |
| `StorageQuotaGB` | Configured quota |
| `StorageUsedPercent` | Percentage of quota used |
| `LastContentModified` | Last time a file was changed |
| `Status` | `Active` or `Recycled` |
| `TotalFileCount` | Number of files in the Documents library |

---

## Prerequisites

- [`PnP.PowerShell`](https://pnp.github.io/powershell/) module installed
- Azure AD App Registration with:
  - `Sites.FullControl.All` (SharePoint) application permission
  - `User.Read.All` (Microsoft Graph) application permission
- Admin consent granted
- Run [`Setup-PnPCertAuth.ps1`](../Setup-PnPCertAuth/) first

---

## Usage

### Quick overview (storage only, no file enumeration)

```powershell
.\Get-OneDriveFileCounts.ps1 `
    -ParamFile ".\MigrationConnectionParams.json" `
    -SkipFileCount
```

### Full report with file counts

```powershell
.\Get-OneDriveFileCounts.ps1 -ParamFile ".\MigrationConnectionParams.json"
```

### Only accounts with at least 500 MB used

```powershell
.\Get-OneDriveFileCounts.ps1 `
    -ParamFile ".\MigrationConnectionParams.json" `
    -MinStorageMB 500 `
    -SkipFileCount
```

### Filter by user UPN pattern

```powershell
.\Get-OneDriveFileCounts.ps1 `
    -ParamFile ".\MigrationConnectionParams.json" `
    -UserFilter "*@contoso.com" `
    -SkipFileCount
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
| `OutputCsv` | No | `.\OneDrive-FileCounts-<timestamp>.csv` | Output file path |
| `SkipFileCount` | No | `$false` | Skip file enumeration, use storage API only |
| `UserFilter` | No | *(all)* | UPN wildcard to limit users |
| `MinStorageMB` | No | `0` | Minimum storage threshold (exclude smaller accounts) |

\* Either `-ParamFile` or `-TenantName` + `-ClientId` + `-CertificatePath` required.

---

## Sample Output

```
===== ONEDRIVE SUMMARY =====
Total OneDrive accounts : 248
Active (non-empty)      : 231
Empty accounts          : 17
Total storage used      : 1842.30 GB

Size distribution:
  Empty (0 MB)       :    17 accounts
  < 1 GB             :    42 accounts
  1–5 GB             :    98 accounts
  5–15 GB            :    61 accounts
  15–50 GB           :    25 accounts
  > 50 GB            :     5 accounts
```

---

## Migration Notes

| Scenario | Google Guidance |
|---|---|
| Files move to Google Drive | Each user's OneDrive maps to their **My Drive** |
| Shared files | Check sharing links — external shares must be re-created in Google |
| Files > 5 TB | Cannot be uploaded to Google Drive; split before migration |
| Shortcuts / Symlinks | Not natively supported in Google Drive |
| OneNote notebooks | Must be exported to `.onepkg` format first |
| Office files | Automatically convert to Google Docs format (optional) |

### Identifying Stale Accounts

Accounts with `LastContentModified` older than 12 months and low storage are candidates to **skip migration** or archive separately.

```powershell
# After running the script, find stale accounts
$stale = Import-Csv .\OneDrive-FileCounts-*.csv |
    Where-Object {
        $_.StorageUsedGB -lt 1 -and
        [datetime]$_.LastContentModified -lt (Get-Date).AddMonths(-12)
    }
$stale | Export-Csv .\Stale-OneDrives.csv -NoTypeInformation
```
