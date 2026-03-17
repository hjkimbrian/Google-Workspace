# Get-ExchangeDistributionGroups

Exports all Exchange Online distribution groups, mail-enabled security groups, dynamic distribution groups, and optionally Microsoft 365 Groups. Produces both a group summary CSV and a detailed membership CSV — the inputs needed to recreate groups as Google Groups.

## Output Files

### Groups Summary CSV (`Exchange-Groups-<timestamp>.csv`)

| Column | Description |
|---|---|
| `DisplayName` | Group display name |
| `PrimarySmtpAddress` | Primary email address |
| `Alias` | Email alias |
| `GroupType` | Type (see below) |
| `MemberCount` | Number of members |
| `Owners` | Group owner(s) |
| `HiddenFromGAL` | Hidden from Global Address List |
| `RequireSenderAuth` | Only authenticated users can send |
| `MemberJoinRestriction` | Open / ApprovalRequired / Closed |
| `SendModerationEnabled` | Messages require moderator approval |
| `EmailAddresses` | All SMTP addresses |
| `Notes` | Description or dynamic filter expression |

### Membership CSV (`Exchange-GroupMemberships-<timestamp>.csv`)

| Column | Description |
|---|---|
| `GroupSmtpAddress` | The group's email address |
| `GroupDisplayName` | Group display name |
| `MemberSmtpAddress` | Member's email address |
| `MemberDisplayName` | Member's display name |
| `MemberType` | Recipient type of the member |

---

## Prerequisites

- `ExchangeOnlineManagement` module v3+
- Azure AD App Registration with:
  - `Exchange.ManageAsApp` application permission
  - Service principal assigned the **Exchange Administrator** role
- Run [`Setup-PnPCertAuth.ps1`](../Setup-PnPCertAuth/) first

---

## Usage

### Distribution and mail-enabled security groups only

```powershell
.\Get-ExchangeDistributionGroups.ps1 -ParamFile ".\MigrationConnectionParams.json"
```

### Include Microsoft 365 Groups (Teams/Planner/SharePoint)

```powershell
.\Get-ExchangeDistributionGroups.ps1 `
    -ParamFile ".\MigrationConnectionParams.json" `
    -IncludeM365Groups
```

### Full report with nested group expansion

```powershell
.\Get-ExchangeDistributionGroups.ps1 `
    -ParamFile ".\MigrationConnectionParams.json" `
    -IncludeM365Groups `
    -ExpandNestedGroups
```

---

## Parameters

| Parameter | Required | Default | Description |
|---|---|---|---|
| `ParamFile` | No* | — | Path to `MigrationConnectionParams.json` |
| `TenantName` | No* | — | Tenant name |
| `ClientId` | No* | — | Azure AD App Client ID |
| `CertificatePath` | No* | — | Path to `.pfx` file |
| `OutputCsv` | No | `.\Exchange-Groups-<timestamp>.csv` | Groups summary output |
| `OutputMembershipCsv` | No | `.\Exchange-GroupMemberships-<timestamp>.csv` | Membership output |
| `IncludeM365Groups` | No | `$false` | Include M365/Teams groups |
| `ExpandNestedGroups` | No | `$false` | Flatten nested group memberships |

---

## Group Type Mapping

| Exchange Type | Google Groups Equivalent |
|---|---|
| `DistributionGroup` | [Google Group](https://support.google.com/a/answer/33343) (email list) |
| `MailEnabledSecurityGroup` | Google Group + access-controlled (use with Workspace access controls) |
| `DynamicDistributionGroup` | Google Group with [dynamic membership](https://support.google.com/a/answer/9400220) (requires GAM scripting to automate) |
| `M365Group` | [Google Group (Collaborative Inbox)](https://support.google.com/a/answer/167430) or [Google Chat Space](https://support.google.com/a/answer/9314941) |

---

## Creating Google Groups from the Export

Use [GAMADV-XTD3](https://github.com/taers232c/GAMADV-XTD3) to bulk-create groups from the CSV:

```bash
# Create groups from summary CSV
gam csv Exchange-Groups-*.csv \
    gam create group ~PrimarySmtpAddress name "~DisplayName"

# Add members from membership CSV
gam csv Exchange-GroupMemberships-*.csv \
    gam update group ~GroupSmtpAddress add member ~MemberSmtpAddress
```

---

## Notes on Dynamic Distribution Groups

Dynamic Distribution Groups use recipient filter queries (LDAP/OPath) that have no direct equivalent in Google Groups. Options:

1. **Convert to static** — export current membership and create a static Google Group
2. **Automate with GAM** — use a scheduled script to query users by attribute and sync Google Group membership
3. **Google Workspace dynamic groups** — use [dynamic groups (beta)](https://support.google.com/a/answer/9400220) based on user attributes
