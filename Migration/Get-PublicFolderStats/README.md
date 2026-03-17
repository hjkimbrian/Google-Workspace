# Get-PublicFolderStats

Exports top-level mail-enabled Exchange Online public folders with message counts, storage sizes, and per-folder permissions. Public folders with an SMTP address are the ones most commonly used as shared inboxes — and the hardest to migrate because Google Workspace has no direct equivalent.

---

## Output Files

### `PublicFolder-Stats-<timestamp>.csv`

One row per mail-enabled top-level public folder.

| Column | Description |
|---|---|
| `FolderName` | Display name of the folder |
| `FolderPath` | Full public folder path (e.g. `\CustomerService`) |
| `PrimarySmtpAddress` | The folder's email address |
| `Alias` | Mail alias |
| `EmailAddresses` | All SMTP addresses assigned to the folder |
| `HiddenFromGAL` | Whether hidden from the Global Address List |
| `ItemCount` | Total messages in the folder |
| `TotalSizeMB` / `TotalSizeGB` | Total storage used |
| `DeletedItemCount` | Items in the dumpster |
| `DeletedSizeMB` | Storage for deleted items |
| `LastModifiedTime` | Last time the folder was written to |
| `CreationTime` | When the folder was created |

### `PublicFolder-Permissions-<timestamp>.csv`

One row per folder + user/group permission pairing.

| Column | Description |
|---|---|
| `FolderName` | Folder display name |
| `FolderPath` | Full path |
| `FolderSmtp` | Folder's email address |
| `Principal` | User or group name |
| `PrincipalEmail` | Resolved SMTP address of the principal |
| `AccessRights` | Exchange access rights (e.g. `Author`, `Editor`, `Owner`, `PublishingEditor`) |
| `IsDefault` | `True` if this is the built-in "Default" entry |
| `IsAnonymous` | `True` if this is the built-in "Anonymous" entry |

---

## Exchange Public Folder Access Rights Reference

| Access Right | Can read | Can create | Can edit own | Can edit all | Can delete | Folder owner |
|---|:-:|:-:|:-:|:-:|:-:|:-:|
| `None` | | | | | | |
| `Reviewer` | ✓ | | | | | |
| `Contributor` | | ✓ | | | | |
| `NonEditingAuthor` | ✓ | ✓ | | | ✓ own | |
| `Author` | ✓ | ✓ | ✓ | | ✓ own | |
| `PublishingAuthor` | ✓ | ✓ | ✓ | | ✓ own | ✓ |
| `Editor` | ✓ | ✓ | ✓ | ✓ | ✓ | |
| `PublishingEditor` | ✓ | ✓ | ✓ | ✓ | ✓ | ✓ |
| `Owner` | ✓ | ✓ | ✓ | ✓ | ✓ | ✓ |

---

## Prerequisites

- `ExchangeOnlineManagement` module v3+
- Azure AD App Registration with:
  - `Exchange.ManageAsApp` application permission
  - Service principal assigned the **Exchange Administrator** role
- Run [`Setup-PnPCertAuth.ps1`](../Setup-PnPCertAuth/) first

---

## Usage

### Default (top-level mail-enabled folders, exclude Default/Anonymous)

```powershell
.\Get-PublicFolderStats.ps1 -ParamFile ".\MigrationConnectionParams.json"
```

### Include Default and Anonymous permission rows

```powershell
.\Get-PublicFolderStats.ps1 `
    -ParamFile ".\MigrationConnectionParams.json" `
    -IncludeDefaultAndAnonymous
```

---

## Parameters

| Parameter | Required | Default | Description |
|---|---|---|---|
| `ParamFile` | No* | — | Path to `MigrationConnectionParams.json` |
| `TenantName` | No* | — | Tenant name |
| `ClientId` | No* | — | Azure AD App Client ID |
| `CertificatePath` | No* | — | Path to `.pfx` file |
| `OutputStatsCsv` | No | `.\PublicFolder-Stats-<ts>.csv` | Stats output path |
| `OutputPermissionsCsv` | No | `.\PublicFolder-Permissions-<ts>.csv` | Permissions output path |
| `IncludeDefaultAndAnonymous` | No | `$false` | Include built-in Default/Anonymous entries |

---

## Migration Strategies

Public folders with an SMTP address are typically used as **shared inboxes**. Google Workspace has no native public folder concept. Common migration approaches:

| Approach | Best for | Tool |
|---|---|---|
| **Google Group (Collaborative Inbox)** | Shared email queues — team can assign, reply, archive | Google Admin Console + GWMME |
| **Shared Gmail account** | High-volume folders where one account manages everything | GWMME |
| **Google Drive folder** | Document-storage public folders (not email) | rclone or Drive API |
| **Archive only** | Folders with old messages nobody actively uses | GYB (Got Your Back) |

### Key Decision Factors

- `ItemCount` > 50,000: Consider archiving to Drive/GCS rather than migrating to Gmail
- `LastModifiedTime` > 2 years ago: Strong archive candidate
- `AccessRights = Owner` for multiple users: Maps well to Collaborative Inbox with multiple managers
- `IsDefault = True` with `AccessRights = Author`: Anyone in the org can email the folder — use a Google Group with open posting

### Checking for Nested Mail-Enabled Subfolders

The script covers top-level only. To find mail-enabled subfolders in Exchange Online:

```powershell
Connect-ExchangeOnline ...
Get-MailPublicFolder -ResultSize Unlimited |
    Where-Object { ($_.Identity -split '\\').Count -gt 2 } |
    Select-Object Identity, PrimarySmtpAddress |
    Export-Csv .\PF-Nested-MailEnabled.csv -NoTypeInformation
```
