# Get-M365LicenseReport

Exports a full Microsoft 365 user and license inventory via Microsoft Graph. Use this to determine exactly how many Google Workspace licenses are needed and which users should be migrated.

## Output Files

### User Report CSV (`M365-LicenseReport-<timestamp>.csv`)

| Column | Description |
|---|---|
| `DisplayName` | User's display name |
| `UserPrincipalName` | UPN / login |
| `Mail` | Primary mailbox address |
| `MailAliases` | Additional SMTP aliases |
| `UserType` | `Member` (internal) or `Guest` (B2B) |
| `AccountEnabled` | Whether the account is active |
| `JobTitle` / `Department` | Org attributes |
| `UsageLocation` | Country code (required for license assignment) |
| `OnPremisesSynced` | `True` if synced from AD via Azure AD Connect |
| `LicenseCount` | Number of licenses assigned |
| `AssignedLicenses` | Friendly names of all assigned SKUs |
| `HasExchangeLicense` | `True` = needs a Gmail/Google Workspace license |
| `HasSharePointLicense` | `True` = may have OneDrive/SharePoint content to migrate |
| `HasTeamsLicense` | `True` = currently using Teams |

### License Summary CSV (`M365-LicenseSummary-<timestamp>.csv`)

| Column | Description |
|---|---|
| `SkuPartNumber` | Microsoft internal SKU name |
| `FriendlyName` | Human-readable product name |
| `TotalPurchased` | Licenses purchased |
| `Assigned` | Licenses currently assigned |
| `Available` | Remaining unassigned licenses |

---

## Prerequisites

- `Microsoft.Graph` PowerShell module (installed automatically if missing)
- Azure AD App Registration with:
  - `User.Read.All` application permission
  - `Directory.Read.All` application permission
  - `Organization.Read.All` application permission
- Admin consent granted
- Run [`Setup-PnPCertAuth.ps1`](../Setup-PnPCertAuth/) first

---

## Usage

### Basic report (active members only)

```powershell
.\Get-M365LicenseReport.ps1 -ParamFile ".\MigrationConnectionParams.json"
```

### Include guests and disabled accounts

```powershell
.\Get-M365LicenseReport.ps1 `
    -ParamFile ".\MigrationConnectionParams.json" `
    -IncludeGuests `
    -IncludeDisabled
```

### Explicit parameters (if not using param file)

```powershell
.\Get-M365LicenseReport.ps1 `
    -TenantName "contoso" `
    -TenantId   "xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx" `
    -ClientId   "yyyyyyyy-yyyy-yyyy-yyyy-yyyyyyyyyyyy" `
    -CertificatePath ".\PnP-Migration-App.pfx"
```

---

## Parameters

| Parameter | Required | Default | Description |
|---|---|---|---|
| `ParamFile` | No* | — | Path to `MigrationConnectionParams.json` |
| `TenantName` | No* | — | Tenant name |
| `TenantId` | Recommended | *(derived from TenantName)* | Azure AD Tenant GUID |
| `ClientId` | No* | — | Azure AD App Client ID |
| `CertificatePath` | No* | — | Path to `.pfx` file |
| `OutputCsv` | No | `.\M365-LicenseReport-<timestamp>.csv` | User report output |
| `OutputSummaryCsv` | No | `.\M365-LicenseSummary-<timestamp>.csv` | License summary output |
| `IncludeGuests` | No | `$false` | Include B2B guest users |
| `IncludeDisabled` | No | `$false` | Include disabled accounts |

---

## Sample Output

```
===== USER AND LICENSE SUMMARY =====
Total users (active)    : 248
Licensed users          : 231
Unlicensed users        :  17
Users with Exchange     : 218  (these need Gmail licenses)

License SKUs in tenant:
FriendlyName                TotalPurchased  Assigned  Available
------------                --------------  --------  ---------
Microsoft 365 E3                       250       231         19
Exchange Online Plan 1                  10        10          0
Power BI Pro                            25        18          7
```

---

## Post-Export Analysis

```powershell
# Load the user report
$users = Import-Csv .\M365-LicenseReport-*.csv

# Count users who need Gmail migration (have Exchange license)
$mailUsers = $users | Where-Object { $_.HasExchangeLicense -eq 'True' }
Write-Host "Users needing Gmail migration: $($mailUsers.Count)"

# Find users without a UsageLocation (required for Google Workspace license assignment)
$noLocation = $users | Where-Object { -not $_.UsageLocation }
Write-Host "Users missing UsageLocation: $($noLocation.Count)"

# Identify on-prem synced accounts (need AD to Google sync via Google Cloud Directory Sync)
$adSynced = $users | Where-Object { $_.OnPremisesSynced -eq 'True' }
Write-Host "AD-synced users (need GCDS): $($adSynced.Count)"
```

---

## Migration Considerations

| Scenario | Notes |
|---|---|
| On-premises AD synced users | Deploy [Google Cloud Directory Sync (GCDS)](https://support.google.com/a/answer/106368) to provision Google accounts |
| Guest accounts | B2B guests do not typically migrate — notify external collaborators |
| Unlicensed users | May still have mailboxes — check `Get-ExchangeMailboxStats` |
| Missing `UsageLocation` | Set before assigning Google Workspace licenses |
| Service accounts | Identify and exclude from user migration; map to Google service accounts |
