# GWS Admin Toolkit

A collection of tools and automation scripts for Google Workspace administrators. Covers common admin tasks — Gmail signature management, data transfers between users, user/group management, calendar cleanup, document protection, email handling, and Windows registry configuration for Google Credential Provider.

## Tools

### Web Apps

| Directory | Description |
|-----------|-------------|
| [`GWSAdminToolkit/`](GWSAdminToolkit/) | Full-featured admin web app: manage Gmail signatures for all domain users (WYSIWYG editor + template variables) **and** transfer user data (Drive, Calendar, Looker Studio) between accounts using the Google Workspace Data Transfer API |

### Google Apps Scripts (`.gs`)

| File | Description |
|------|-------------|
| `AddUserToGroup.gs` | Adds users to Google Groups |
| `CalendarSync.gs` | Syncs calendars across Google Workspace |
| `deleteCalendarEvents.gs` | Bulk-deletes calendar events |
| `lockGDoc.gs` | Locks/protects Google Documents from edits |
| `saveEmailAttachments.gs` | Extracts and saves email attachments |

### Shell Scripts (`.sh`)

| File | Description |
|------|-------------|
| `DeleteInternalRecurringEvents.sh` | Removes internal recurring calendar events |
| `MakerHourViolator.sh` | Monitors or flags maker hour policy violations |

### Migration Scripts (`Migration/`)

PowerShell scripts for inventorying a Microsoft 365 tenant before migrating to Google Workspace. All use certificate-based (app-only) authentication via PnP PowerShell.

| Directory | Description |
|-----------|-------------|
| [`Migration/Setup-PnPCertAuth/`](Migration/Setup-PnPCertAuth/) | One-time Azure AD App Registration with certificate auth — run this first |
| [`Migration/Get-ExchangeMailboxStats/`](Migration/Get-ExchangeMailboxStats/) | Mailbox list with message counts, sizes, and Online Archive item counts |
| [`Migration/Get-MailboxPermissions/`](Migration/Get-MailboxPermissions/) | Full Access + Send As + Send on Behalf for Shared/Room/Equipment mailboxes; resource inventory |
| [`Migration/Get-PublicFolderStats/`](Migration/Get-PublicFolderStats/) | Top-level mail-enabled public folders with message counts and permissions |
| [`Migration/Get-SharePointFileCounts/`](Migration/Get-SharePointFileCounts/) | Count files and storage usage per SharePoint Online site collection |
| [`Migration/Get-OneDriveFileCounts/`](Migration/Get-OneDriveFileCounts/) | Count files and storage usage per OneDrive for Business account |
| [`Migration/Get-SPOSitePermissions/`](Migration/Get-SPOSitePermissions/) | Root + document library permissions with principal type classification |
| [`Migration/Get-SPOGroupMembers/`](Migration/Get-SPOGroupMembers/) | Resolve SPO group members; optional Azure AD group expansion |
| [`Migration/Get-AllEmailAliases/`](Migration/Get-AllEmailAliases/) | Every secondary SMTP alias across all recipient types — one row per alias, with duplicate detection |
| [`Migration/Get-ExchangeDistributionGroups/`](Migration/Get-ExchangeDistributionGroups/) | Export all distribution groups with members (for recreating as Google Groups) |
| [`Migration/Get-M365LicenseReport/`](Migration/Get-M365LicenseReport/) | Export all users and their M365 licenses (scope Google Workspace seats) |

See [`Migration/README.md`](Migration/README.md) for full usage instructions.

### PowerShell Scripts (`.ps1`)

| File | Description |
|------|-------------|
| `GCPWSetRegistryKeys.ps1` | Configures Windows registry keys for Google Credential Provider for Windows (GCPW) |

## Usage

### GWS Admin Toolkit (Web App)

See [`GWSAdminToolkit/README.md`](GWSAdminToolkit/README.md) for full setup instructions including GCP project configuration, service account creation, domain-wide delegation, and deployment.

### Apps Scripts

1. Open [Google Apps Script](https://script.google.com)
2. Create a new project and paste the `.gs` file contents
3. Configure any required variables (e.g., group email, calendar IDs)
4. Run or deploy as needed

### Shell Scripts

```bash
chmod +x script.sh
./script.sh
```

> Requires [GAMADV-XTD3](https://github.com/taers232c/GAMADV-XTD3) or [GYB](https://github.com/GAM-team/got-your-back) depending on the script.

### PowerShell Scripts

```powershell
.\GCPWSetRegistryKeys.ps1
```

> Run as Administrator. Requires [GCPW](https://support.google.com/a/answer/9250996) to be installed.

## Requirements

- Google Workspace super admin account
- [GAMADV-XTD3](https://github.com/taers232c/GAMADV-XTD3) (for shell scripts)
- [rclone](https://rclone.org/) (if syncing files to Drive/GCS)
- Windows with GCPW installed (for PowerShell scripts)
