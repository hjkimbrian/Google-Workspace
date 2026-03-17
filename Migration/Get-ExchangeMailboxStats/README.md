# Get-ExchangeMailboxStats

Exports a detailed CSV inventory of all Exchange Online mailboxes including message counts and storage sizes — the key data needed to scope a Google Workspace migration.

## Output Fields

| Column | Description |
|---|---|
| `DisplayName` | Mailbox display name |
| `UserPrincipalName` | UPN / login address |
| `PrimarySmtpAddress` | Primary email address |
| `RecipientTypeDetails` | Mailbox type (see below) |
| `TotalItemCount` | Total messages in all folders |
| `TotalSizeMB` / `TotalSizeGB` | Mailbox size |
| `DeletedItemCount` | Items in Deleted Items / Recoverable Items |
| `DeletedSizeMB` | Size of deleted items |
| `LastLogonTime` | Last time a user logged in |
| `IsArchiveEnabled` | Whether In-Place Archive is enabled |
| `LitigationHoldEnabled` | Whether the mailbox is on Litigation Hold |
| `ArchiveItemCount` / `ArchiveSizeGB` | Archive mailbox stats (if enabled) |

### Mailbox Types

| Type | Google Equivalent |
|---|---|
| `UserMailbox` | Individual Gmail account |
| `SharedMailbox` | [Collaborative inbox](https://support.google.com/a/answer/167430) or Google Group |
| `RoomMailbox` | [Google Calendar resource](https://support.google.com/a/answer/1686462) |
| `EquipmentMailbox` | Google Calendar resource |

---

## Prerequisites

- `ExchangeOnlineManagement` module v3+ (`Install-Module ExchangeOnlineManagement`)
- Azure AD App Registration with:
  - `Exchange.ManageAsApp` application permission
  - Service principal assigned the **Exchange Administrator** role in Azure AD
- Run [`Setup-PnPCertAuth.ps1`](../Setup-PnPCertAuth/) first to create the app registration

---

## Usage

### With param file (recommended)

```powershell
.\Get-ExchangeMailboxStats.ps1 -ParamFile "..\..\MigrationConnectionParams.json"
```

### With explicit parameters

```powershell
.\Get-ExchangeMailboxStats.ps1 `
    -TenantName "contoso" `
    -ClientId   "xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx" `
    -CertificatePath ".\PnP-Migration-App.pfx"
```

### Include all mailbox types

```powershell
.\Get-ExchangeMailboxStats.ps1 `
    -ParamFile ".\MigrationConnectionParams.json" `
    -MailboxTypes "UserMailbox,SharedMailbox,RoomMailbox,EquipmentMailbox"
```

### Include inactive/soft-deleted mailboxes

```powershell
.\Get-ExchangeMailboxStats.ps1 `
    -ParamFile ".\MigrationConnectionParams.json" `
    -IncludeInactiveMailboxes
```

---

## Parameters

| Parameter | Required | Default | Description |
|---|---|---|---|
| `ParamFile` | No* | — | Path to `MigrationConnectionParams.json` |
| `TenantName` | No* | — | Tenant name (e.g. `contoso`) |
| `ClientId` | No* | — | Azure AD App Client ID |
| `CertificatePath` | No* | — | Path to `.pfx` file |
| `CertificatePassword` | No | *(prompted if needed)* | SecureString cert password |
| `MailboxTypes` | No | `UserMailbox,SharedMailbox` | Comma-separated mailbox types |
| `OutputCsv` | No | `.\Exchange-Mailbox-Stats-<timestamp>.csv` | Output file path |
| `IncludeInactiveMailboxes` | No | `$false` | Include soft-deleted mailboxes |

\* Either `-ParamFile` or all three of `-TenantName`, `-ClientId`, `-CertificatePath` are required.

---

## Sample Output

```
===== MAILBOX SUMMARY =====

MailboxType          Count  TotalItems  TotalSizeGB
-----------          -----  ----------  -----------
UserMailbox            248     1842931       512.40
SharedMailbox           12       48211        18.30

Total mailboxes : 260
Grand total size: 530.70 GB

[OK] Report saved to: .\Exchange-Mailbox-Stats-20260317-142301.csv
```

---

## Tips for Migration Sizing

- **Users over 25 GB**: Google Workspace migration tools (GWMME, GAMADV-XTD3) work better with smaller mailboxes; consider archiving or phased migration
- **Litigation Hold**: These mailboxes cannot be deleted from Exchange until hold is released
- **Inactive mailboxes**: May need to be recovered before migrating if content is required
- **Shared mailboxes**: Decide whether to convert to Google Groups (collaboration) or individual accounts
