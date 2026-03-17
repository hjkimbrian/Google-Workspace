# Get-AllEmailAliases

Exports a flat, one-row-per-alias inventory of every secondary SMTP address in the tenant, mapped to its primary SMTP address and recipient type. Covers all mail-enabled objects in a single pass.

---

## Why This Matters for Migration

Exchange allows any number of secondary SMTP addresses (aliases) per recipient. When migrating to Google Workspace:

- Gmail users can have [up to 30 aliases](https://support.google.com/a/answer/33327) — but they must be manually configured
- Google Groups can have alternate addresses
- Aliases used by business systems, applications, or routing rules will stop working at cutover if not recreated

This script gives you the complete picture before migration.

---

## Recipient Types Covered

A single `Get-EXORecipient` call covers every mail-enabled object:

| RecipientTypeDetails | Description |
|---|---|
| `UserMailbox` | Individual user mailbox |
| `SharedMailbox` | Shared mailbox (e.g. info@, support@) |
| `RoomMailbox` | Room resource |
| `EquipmentMailbox` | Equipment resource |
| `MailUser` | Mail-enabled user (external or unlicensed) |
| `MailContact` | External mail contact |
| `DistributionGroup` | Distribution list |
| `MailEnabledSecurityGroup` | Mail-enabled security group |
| `DynamicDistributionGroup` | Dynamic distribution group |
| `GroupMailbox` | Microsoft 365 Group |
| `PublicFolder` | Mail-enabled public folder |

---

## Output: `All-EmailAliases-<timestamp>.csv`

One row per secondary SMTP alias.

| Column | Description |
|---|---|
| `PrimarySmtpAddress` | The recipient's primary (canonical) email address |
| `DisplayName` | Display name of the recipient |
| `RecipientType` | Type of mail object (see table above) |
| `AliasAddress` | The secondary SMTP address (the alias itself) |
| `AddressType` | `SMTP` for secondary addresses; `Primary` if `-IncludePrimaryAsAlias` is used |
| `HiddenFromGAL` | Whether the recipient is hidden from the Global Address List |

### Duplicate Detection

If the same alias address appears on more than one recipient, the script:
1. Prints a warning listing the conflict
2. Writes a separate `All-EmailAliases-<ts>-Duplicates.csv` with just those rows

Duplicate aliases must be resolved before migration — Google Workspace will reject alias assignments that conflict across accounts.

---

## Prerequisites

- `ExchangeOnlineManagement` module v3+
- Azure AD App Registration with:
  - `Exchange.ManageAsApp` application permission
  - Service principal assigned the **Exchange Administrator** role
- Run [`Setup-PnPCertAuth.ps1`](../Setup-PnPCertAuth/) first

---

## Usage

### All recipient types, secondary SMTP aliases only (default)

```powershell
.\Get-AllEmailAliases.ps1 -ParamFile ".\MigrationConnectionParams.json"
```

### Limit to specific recipient types

```powershell
.\Get-AllEmailAliases.ps1 `
    -ParamFile ".\MigrationConnectionParams.json" `
    -RecipientTypes "UserMailbox,SharedMailbox,DistributionGroup"
```

### Full address book export (include primary address as a row too)

```powershell
.\Get-AllEmailAliases.ps1 `
    -ParamFile ".\MigrationConnectionParams.json" `
    -IncludePrimaryAsAlias `
    -OutputCsv ".\Full-AddressBook.csv"
```

### Include non-SMTP addresses (X500, SIP, EUM)

```powershell
.\Get-AllEmailAliases.ps1 `
    -ParamFile ".\MigrationConnectionParams.json" `
    -IncludeNonSmtp
```

> X500 addresses are used by Exchange for reply-to tracking in Outlook. During coexistence or if users reply to cached addresses, missing X500 entries cause "undeliverable" errors. Keep this export for reference if you run Exchange/Google in parallel.

---

## Parameters

| Parameter | Required | Default | Description |
|---|---|---|---|
| `ParamFile` | No* | — | Path to `MigrationConnectionParams.json` |
| `TenantName` | No* | — | Tenant name |
| `ClientId` | No* | — | Azure AD App Client ID |
| `CertificatePath` | No* | — | Path to `.pfx` file |
| `OutputCsv` | No | `.\All-EmailAliases-<ts>.csv` | Output path |
| `RecipientTypes` | No | *(all types)* | Comma-separated filter on RecipientTypeDetails |
| `IncludeNonSmtp` | No | `$false` | Include X400, X500, SIP, EUM addresses |
| `IncludePrimaryAsAlias` | No | `$false` | Also emit a row for the primary SMTP address |

---

## Sample Output

```
===== ALIAS SUMMARY =====

RecipientType                AliasCount  Recipients
-------------                ----------  ----------
UserMailbox                         412         248
DistributionGroup                    67          54
SharedMailbox                        28          12
GroupMailbox                         15          14
MailEnabledSecurityGroup              8           7

Recipients with at least one alias : 335
Total alias rows                   : 530

WARNING: 2 alias address(es) appear on more than one recipient (potential conflict):
  billing@contoso.com → finance@contoso.com, accounts@contoso.com
  noreply@contoso.com → it-alerts@contoso.com, monitoring@contoso.com
```

---

## Using the Export for Google Workspace

### Add aliases to Gmail accounts (via GAM)

```bash
gam csv All-EmailAliases-*.csv \
    gam user ~PrimarySmtpAddress add alias ~AliasAddress
```

### Add aliases to Google Groups

```bash
gam csv All-EmailAliases-*.csv matchfield RecipientType DistributionGroup \
    gam update group ~PrimarySmtpAddress add alias ~AliasAddress
```

### Check for aliases exceeding Google's 30-alias limit per user

```powershell
Import-Csv .\All-EmailAliases-*.csv |
    Where-Object { $_.RecipientType -eq 'UserMailbox' } |
    Group-Object PrimarySmtpAddress |
    Where-Object { $_.Count -gt 30 } |
    Select-Object Name, Count |
    Export-Csv .\Users-Over-30-Aliases.csv -NoTypeInformation
```

---

## Coexistence Note

If you run Exchange and Google in parallel during migration, inbound mail to aliases must route correctly to whichever platform the user is on. Consider:

- Keeping Exchange as the **authoritative MX** during coexistence and routing Google users' mail via connectors
- Using a migration platform (GWMME, BitTitan) that handles split-delivery routing
- Exporting X500 addresses with `-IncludeNonSmtp` and preserving them if Outlook cached addresses need to resolve after cutover
