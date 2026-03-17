# Migration PowerShell Scripts

PowerShell scripts for inventorying a Microsoft 365 tenant before migrating to Google Workspace. All scripts use **certificate-based (app-only) authentication** — no interactive sign-in or stored passwords required.

---

## Quick Start

### 1. One-time setup (run first)

```powershell
cd Setup-PnPCertAuth
.\Setup-PnPCertAuth.ps1 -TenantName "contoso"
```

This registers an Azure AD app, generates a certificate, and writes `MigrationConnectionParams.json` with the connection parameters used by all other scripts.

> After running, go to Azure Portal → Azure AD → App Registrations → Grant admin consent.

### 2. Run inventory scripts

```powershell
# Mailbox counts and sizes
.\Get-ExchangeMailboxStats\Get-ExchangeMailboxStats.ps1 -ParamFile ".\Setup-PnPCertAuth\MigrationConnectionParams.json"

# SharePoint file counts
.\Get-SharePointFileCounts\Get-SharePointFileCounts.ps1 -ParamFile ".\Setup-PnPCertAuth\MigrationConnectionParams.json"

# OneDrive file counts
.\Get-OneDriveFileCounts\Get-OneDriveFileCounts.ps1 -ParamFile ".\Setup-PnPCertAuth\MigrationConnectionParams.json"

# Distribution groups
.\Get-ExchangeDistributionGroups\Get-ExchangeDistributionGroups.ps1 -ParamFile ".\Setup-PnPCertAuth\MigrationConnectionParams.json"

# License report
.\Get-M365LicenseReport\Get-M365LicenseReport.ps1 -ParamFile ".\Setup-PnPCertAuth\MigrationConnectionParams.json"
```

---

## Scripts

| Script | Description | Module Required |
|---|---|---|
| [`Setup-PnPCertAuth/`](Setup-PnPCertAuth/) | Register Azure AD app with cert auth — run this first | PnP.PowerShell |
| [`Get-ExchangeMailboxStats/`](Get-ExchangeMailboxStats/) | Mailbox list with message counts, sizes, and Online Archive item counts | ExchangeOnlineManagement |
| [`Get-MailboxPermissions/`](Get-MailboxPermissions/) | Full Access + Send As + Send on Behalf for Shared/Room/Equipment mailboxes; resource inventory CSV | ExchangeOnlineManagement |
| [`Get-PublicFolderStats/`](Get-PublicFolderStats/) | Top-level mail-enabled public folders: message counts, sizes, and permissions | ExchangeOnlineManagement |
| [`Get-SharePointFileCounts/`](Get-SharePointFileCounts/) | File counts and storage per SharePoint site | PnP.PowerShell |
| [`Get-OneDriveFileCounts/`](Get-OneDriveFileCounts/) | File counts and storage per OneDrive account | PnP.PowerShell |
| [`Get-SPOSitePermissions/`](Get-SPOSitePermissions/) | Root + document library permissions with principal type classification (User/SPO Group/Security Group/M365 Group) | PnP.PowerShell |
| [`Get-SPOGroupMembers/`](Get-SPOGroupMembers/) | Resolves members of SPO groups from the permissions report; optional AAD group expansion via Graph | PnP.PowerShell |
| [`Get-ExchangeDistributionGroups/`](Get-ExchangeDistributionGroups/) | All groups with members (for recreating as Google Groups) | ExchangeOnlineManagement |
| [`Get-M365LicenseReport/`](Get-M365LicenseReport/) | All users and assigned licenses (scope Google Workspace seats) | Microsoft.Graph |

---

## Prerequisites

| Requirement | Purpose |
|---|---|
| Windows PowerShell 5.1+ or PowerShell 7+ | Script runtime (PS 7 recommended) |
| Azure AD role: Global Admin or App Administrator | Register the app in Azure AD |
| Internet access from the machine running scripts | Module installs + API calls |

PowerShell modules are installed automatically on first run:

```powershell
# Or install manually:
Install-Module PnP.PowerShell                -Scope CurrentUser
Install-Module ExchangeOnlineManagement      -Scope CurrentUser
Install-Module Microsoft.Graph               -Scope CurrentUser
```

---

## Architecture

All scripts share a single **Azure AD App Registration** created by `Setup-PnPCertAuth.ps1`:

```
Your machine
  │
  ├── PnP-Migration-App.pfx  (certificate private key — keep secure)
  └── MigrationConnectionParams.json
        │
        ▼
  Azure AD App Registration
  "PnP-Migration-App"
        │
        ├── Exchange Online  (Exchange.ManageAsApp)
        ├── SharePoint Online (Sites.FullControl.All)
        └── Microsoft Graph  (User.Read.All, Directory.Read.All, ...)
```

The app uses **application permissions** (not delegated), so scripts run with no user session and are safe to schedule or run overnight.

---

## Typical Migration Workflow

```
1. Run Get-M365LicenseReport          → Count licensed users → order Google Workspace seats
2. Run Get-ExchangeMailboxStats        → Size email + archive data → plan migration waves
3. Run Get-MailboxPermissions          → Delegate access for shared/room/equipment mailboxes
4. Run Get-PublicFolderStats           → Mail-enabled public folders → decide: Groups, shared Gmail, or archive
5. Run Get-OneDriveFileCounts          → Size OneDrive data → estimate Drive migration time
6. Run Get-SharePointFileCounts        → Identify large SPO sites → map to Shared Drives
7. Run Get-SPOSitePermissions          → Capture root + library permissions with group type
8. Run Get-SPOGroupMembers             → Resolve group members → assign Google Drive access
9. Run Get-ExchangeDistributionGroups  → Export groups → bulk-create Google Groups
```

---

## Security Notes

- **Never commit** `MigrationConnectionParams.json` or `.pfx` files to source control
- The app registration and certificate should be **deleted from Azure AD** after migration
- Limit the app's permissions to read-only where possible (remove `FullControl` once inventory is complete)
- Use Azure Key Vault or Windows Credential Manager for the certificate in production

---

## Further Reading

- [Google Workspace Migration Guide](https://workspace.google.com/intl/en/products/gmail/migration/)
- [Google Workspace Migration for Microsoft Exchange (GWMME)](https://support.google.com/a/answer/6291304)
- [PnP PowerShell Documentation](https://pnp.github.io/powershell/)
- [GAMADV-XTD3](https://github.com/taers232c/GAMADV-XTD3) — CLI for bulk Google Workspace admin operations
- [Google Cloud Directory Sync (GCDS)](https://support.google.com/a/answer/106368) — sync AD users to Google
