# Get-MailboxPermissions

Exports all permissions on Shared, Room, and Equipment mailboxes — covering all three Exchange permission layers that affect what users can do in Google Workspace after migration. Also produces a standalone resource mailbox inventory CSV.

---

## Why Three Permission Types?

| Permission Type | What it controls | Google equivalent |
|---|---|---|
| **Full Access** | Open and manage the mailbox contents | Gmail delegation |
| **Send As** | Send email that appears to come *from* the mailbox address | "Send mail as" in Gmail delegate settings |
| **Send on Behalf** | Send email with "on behalf of" header visible to recipients | "Send mail as" (with "on behalf of" notation) |

All three must be recreated in Google Workspace. Missing Send As is a common post-migration complaint.

---

## Output Files

### `Mailbox-Permissions-<timestamp>.csv`

One row per mailbox + principal + permission type.

| Column | Description |
|---|---|
| `MailboxDisplayName` | Mailbox display name |
| `MailboxUPN` | Mailbox UPN |
| `MailboxSmtp` | Primary SMTP address |
| `MailboxType` | `SharedMailbox`, `RoomMailbox`, or `EquipmentMailbox` |
| `PermissionType` | `FullAccess`, `SendAs`, or `SendOnBehalf` |
| `Principal` | Display name or identity of the user/group granted the permission |
| `PrincipalEmail` | SMTP address of the principal (where resolvable) |
| `IsInherited` | Whether the permission is inherited (usually `False` for explicit grants) |
| `Deny` | Whether this is an explicit Deny entry |

### `Resource-Mailboxes-<timestamp>.csv`

One row per Room or Equipment mailbox.

| Column | Description |
|---|---|
| `DisplayName` / `UserPrincipalName` / `PrimarySmtpAddress` | Identity fields |
| `ResourceType` | `RoomMailbox` or `EquipmentMailbox` |
| `Capacity` | Room/equipment capacity |
| `Office` / `City` / `Building` | Location attributes |
| `AutoAccept` | Whether the room auto-accepts bookings |
| `AllowConflicts` | Whether double-booking is allowed |
| `BookingWindowInDays` | How far in advance the resource can be booked |
| `MaximumDurationInMinutes` | Maximum single meeting duration |
| `AllowRecurringMeetings` | Whether recurring meetings are accepted |
| `ResourceDelegates` | Booking delegates (people who manually approve/decline) |
| `AdditionalResponse` | Custom text in booking accept/decline emails |

---

## Prerequisites

- `ExchangeOnlineManagement` module v3+
- Azure AD App Registration with:
  - `Exchange.ManageAsApp` application permission
  - Service principal assigned the **Exchange Administrator** role
- Run [`Setup-PnPCertAuth.ps1`](../Setup-PnPCertAuth/) first

---

## Usage

### All shared/room/equipment mailboxes (default)

```powershell
.\Get-MailboxPermissions.ps1 -ParamFile ".\MigrationConnectionParams.json"
```

### Shared mailboxes only

```powershell
.\Get-MailboxPermissions.ps1 `
    -ParamFile ".\MigrationConnectionParams.json" `
    -MailboxTypes "SharedMailbox"
```

### Include explicit Deny entries

```powershell
.\Get-MailboxPermissions.ps1 `
    -ParamFile ".\MigrationConnectionParams.json" `
    -IncludeDenyPermissions
```

---

## Parameters

| Parameter | Required | Default | Description |
|---|---|---|---|
| `ParamFile` | No* | — | Path to `MigrationConnectionParams.json` |
| `TenantName` | No* | — | Tenant name |
| `ClientId` | No* | — | Azure AD App Client ID |
| `CertificatePath` | No* | — | Path to `.pfx` file |
| `MailboxTypes` | No | `SharedMailbox,RoomMailbox,EquipmentMailbox` | Comma-separated mailbox types |
| `OutputPermissionsCsv` | No | `.\Mailbox-Permissions-<ts>.csv` | Permissions output path |
| `OutputResourcesCsv` | No | `.\Resource-Mailboxes-<ts>.csv` | Resource inventory path |
| `IncludeDenyPermissions` | No | `$false` | Include explicit Deny entries |

---

## Recreating Permissions in Google Workspace

### Gmail Delegation (Full Access equivalent)

```bash
# Using GAMADV-XTD3 — add a delegate to a shared Gmail account
gam user sharedmailbox@contoso.com add delegate delegateuser@contoso.com
```

### Send As

```bash
# Grant "Send mail as" for the shared address
gam user delegateuser@contoso.com add sendas sharedmailbox@contoso.com "Shared Mailbox Name"
```

### Google Calendar Resources (Room/Equipment)

Resource mailboxes map to [Google Calendar Buildings and Resources](https://support.google.com/a/answer/1686462). Use the `Resource-Mailboxes-<ts>.csv` to bulk-create them:

```bash
gam create resource sharedroom "Conference Room A" type "Conference Room" \
    capacity 20 building "HQ" floor "2"
```

---

## Notes

- **NT AUTHORITY\SELF** and other system accounts are excluded automatically
- **Inherited permissions** (`IsInherited = True`) are Exchange defaults and generally don't need to be recreated
- **Groups in permissions**: The `Principal` column may contain a group name rather than an individual user — if so, check [`Get-ExchangeDistributionGroups`](../Get-ExchangeDistributionGroups/) for membership
- **Building attribute**: The script reads `CustomAttribute1` for building information — adjust the property name if your tenant uses a different attribute for room location
