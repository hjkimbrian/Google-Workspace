# Setup-PnPCertAuth

Automates Azure AD App Registration with certificate-based authentication for use by the other PnP PowerShell migration scripts.

## Why Certificate Auth?

The migration inventory scripts run unattended (no browser pop-up, no MFA prompt). Azure AD supports **app-only** (daemon) authentication using a certificate instead of a client secret. This is the recommended approach for production automation:

- No user credentials stored in scripts
- Certificate can be time-limited and rotated
- Permissions scoped to exactly what is needed

---

## Prerequisites

| Requirement | Notes |
|---|---|
| Windows PowerShell 5.1+ or PowerShell 7+ | PS 7 recommended |
| Azure AD role | **Global Admin** or **Application Administrator** |
| Internet access | Module installs from PSGallery |

The script installs these modules automatically if missing:

- [`PnP.PowerShell`](https://pnp.github.io/powershell/) — SharePoint/OneDrive access
- [`ExchangeOnlineManagement`](https://www.powershellgallery.com/packages/ExchangeOnlineManagement) v3+ — Exchange access

---

## Usage

### Quickstart (generate a new certificate)

```powershell
.\Setup-PnPCertAuth.ps1 -TenantName "contoso"
```

You will be prompted for a certificate password. The script will:
1. Install/update required modules
2. Generate a self-signed 2-year certificate (`PnP-Migration-App.pfx` + `.cer`)
3. Register the Azure AD app and upload the certificate
4. Write `MigrationConnectionParams.json` used by all other scripts

### Use an existing certificate

```powershell
.\Setup-PnPCertAuth.ps1 `
    -TenantName "contoso" `
    -CertificatePath "C:\Certs\existing.pfx" `
    -CertificatePassword (Read-Host -AsSecureString "Cert password")
```

### Custom app name and output path

```powershell
.\Setup-PnPCertAuth.ps1 `
    -TenantName "contoso" `
    -AppDisplayName "Contoso-MigrationInventory" `
    -OutputParamFile "C:\Migration\params.json"
```

---

## Parameters

| Parameter | Required | Default | Description |
|---|---|---|---|
| `TenantName` | Yes | — | M365 tenant name (e.g. `contoso` from `contoso.onmicrosoft.com`) |
| `AppDisplayName` | No | `PnP-Migration-App` | Display name for the Azure AD app registration |
| `CertificatePath` | No | *(generated)* | Path to existing `.pfx` file; if omitted, a new cert is created |
| `CertificatePassword` | No | *(prompted)* | SecureString password for the certificate |
| `OutputParamFile` | No | `.\MigrationConnectionParams.json` | Path for the output connection-params JSON file |

---

## Post-Script Manual Steps

After the script completes you must grant admin consent:

1. Open [Azure Portal](https://portal.azure.com) → **Azure Active Directory** → **App registrations**
2. Find the app (e.g. `PnP-Migration-App`)
3. Go to **API permissions** → click **Grant admin consent for \<tenant\>**
4. Confirm the Client ID in `MigrationConnectionParams.json` matches the Portal

### Exchange Online App-Only Access

Exchange requires an additional one-time setup:

```powershell
Connect-ExchangeOnline -UserPrincipalName admin@contoso.onmicrosoft.com

# Grant the app's service principal the Exchange Administrator role
# (or a custom role with mailbox read permissions)
New-ServicePrincipal -AppId <ClientId> -ServiceId <ObjectId> -DisplayName "PnP-Migration-App"
```

> For read-only inventory, assigning the **Exchange Administrator** role to the service principal in Azure AD is the simplest option.

---

## Output

### `MigrationConnectionParams.json`

```json
{
  "TenantName": "contoso",
  "TenantId": "contoso.onmicrosoft.com",
  "AppDisplayName": "PnP-Migration-App",
  "ClientId": "xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx",
  "CertificatePath": "C:\\Migration\\PnP-Migration-App.pfx",
  "AdminSiteUrl": "https://contoso-admin.sharepoint.com",
  "TenantRootUrl": "https://contoso.sharepoint.com"
}
```

This file is read by `Get-SharePointFileCounts.ps1`, `Get-OneDriveFileCounts.ps1`, and other scripts via the `-ParamFile` parameter.

---

## Security Notes

- Store the `.pfx` file securely (e.g. Azure Key Vault, Windows Credential Manager)
- Do **not** commit `MigrationConnectionParams.json` or the `.pfx` to source control
- The certificate and app registration can be deleted from Azure AD after migration is complete
- Rotate or expire the certificate to limit the attack window
