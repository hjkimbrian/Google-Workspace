# Google Workspace Admin Automation Scripts

A collection of automation scripts for Google Workspace administrators. Covers common admin tasks — user/group management, calendar cleanup, document protection, email handling, and Windows registry configuration for Google Cloud.

## Scripts

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

### PowerShell Scripts (`.ps1`)

| File | Description |
|------|-------------|
| `GCPWSetRegistryKeys.ps1` | Configures Windows registry keys for Google Credential Provider for Windows (GCPW) |

## Usage

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

- Google Workspace admin account
- [GAMADV-XTD3](https://github.com/taers232c/GAMADV-XTD3) (for shell scripts)
- [rclone](https://rclone.org/) (if syncing files to Drive/GCS)
- Windows with GCPW installed (for PowerShell scripts)
