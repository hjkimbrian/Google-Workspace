<#
.SYNOPSIS
    Registers an Azure AD App Registration with certificate-based authentication for PnP PowerShell.

.DESCRIPTION
    This script automates the setup required to use PnP PowerShell with certificate-based
    (non-interactive) authentication against SharePoint Online and Exchange Online.

    It will:
      1. Install required PowerShell modules (PnP.PowerShell, ExchangeOnlineManagement)
      2. Generate a self-signed certificate (or use an existing .pfx)
      3. Register a new Azure AD App Registration with the certificate
      4. Grant the required API permissions (SharePoint, Exchange, Graph)
      5. Output the connection parameters needed by the other Migration scripts

    The resulting App Registration uses application (daemon) permissions so scripts can
    run unattended without a signed-in user.

.PARAMETER TenantName
    Your M365 tenant name, e.g. "contoso" (without .onmicrosoft.com).

.PARAMETER AppDisplayName
    Display name for the Azure AD app registration. Defaults to "PnP-Migration-App".

.PARAMETER CertificatePath
    Optional. Path to an existing .pfx certificate file. If not supplied, a new
    self-signed certificate is generated and exported to the current directory.

.PARAMETER CertificatePassword
    SecureString password for the certificate. If not provided, you will be prompted.

.PARAMETER OutputParamFile
    Path to write the connection-parameter JSON file consumed by other scripts.
    Defaults to ".\MigrationConnectionParams.json".

.EXAMPLE
    .\Setup-PnPCertAuth.ps1 -TenantName "contoso"

.EXAMPLE
    .\Setup-PnPCertAuth.ps1 -TenantName "contoso" `
        -CertificatePath "C:\Certs\existing.pfx" `
        -CertificatePassword (Read-Host -AsSecureString "Cert password")

.NOTES
    Requires: Az CLI or Azure AD PowerShell module for app registration.
    Run as a user with Azure AD Global Admin or Application Administrator role.
    PnP.PowerShell module must be installed: Install-Module PnP.PowerShell -Scope CurrentUser
#>

[CmdletBinding(SupportsShouldProcess)]
param(
    [Parameter(Mandatory)]
    [string]$TenantName,

    [string]$AppDisplayName = "PnP-Migration-App",

    [string]$CertificatePath,

    [SecureString]$CertificatePassword,

    [string]$OutputParamFile = ".\MigrationConnectionParams.json"
)

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

#region ── Helper Functions ──────────────────────────────────────────────────

function Write-Step {
    param([string]$Message)
    Write-Host "`n[STEP] $Message" -ForegroundColor Cyan
}

function Write-Success {
    param([string]$Message)
    Write-Host "[OK]   $Message" -ForegroundColor Green
}

function Ensure-Module {
    param([string]$Name, [string]$MinVersion)
    $mod = Get-Module -ListAvailable -Name $Name | Sort-Object Version -Descending | Select-Object -First 1
    if (-not $mod) {
        Write-Host "       Installing module: $Name" -ForegroundColor Yellow
        Install-Module -Name $Name -Scope CurrentUser -Force -AllowClobber
    } elseif ($MinVersion -and ($mod.Version -lt [version]$MinVersion)) {
        Write-Host "       Updating module: $Name (current: $($mod.Version), required: $MinVersion)" -ForegroundColor Yellow
        Update-Module -Name $Name -Force
    }
    Import-Module $Name -ErrorAction SilentlyContinue
    Write-Success "Module ready: $Name"
}

#endregion

#region ── Step 1 – Install Prerequisites ────────────────────────────────────

Write-Step "Checking required PowerShell modules"
Ensure-Module -Name "PnP.PowerShell"
Ensure-Module -Name "ExchangeOnlineManagement" -MinVersion "3.0.0"

#endregion

#region ── Step 2 – Certificate Setup ────────────────────────────────────────

Write-Step "Setting up certificate"

$tenantDomain = "$TenantName.onmicrosoft.com"
$certCN       = "CN=PnPMigration-$TenantName"
$pfxPath      = Join-Path $PSScriptRoot "$AppDisplayName.pfx"
$cerPath      = Join-Path $PSScriptRoot "$AppDisplayName.cer"

if (-not $CertificatePassword) {
    $CertificatePassword = Read-Host "Enter a password to protect the .pfx file" -AsSecureString
}

if ($CertificatePath -and (Test-Path $CertificatePath)) {
    Write-Host "       Using existing certificate: $CertificatePath"
    $pfxPath = $CertificatePath
} else {
    Write-Host "       Generating self-signed certificate..."

    $cert = New-SelfSignedCertificate `
        -Subject $certCN `
        -CertStoreLocation "Cert:\CurrentUser\My" `
        -KeyExportPolicy Exportable `
        -KeySpec Signature `
        -KeyLength 2048 `
        -KeyAlgorithm RSA `
        -HashAlgorithm SHA256 `
        -NotAfter (Get-Date).AddYears(2)

    # Export .pfx (private key, for the scripts)
    Export-PfxCertificate -Cert $cert -FilePath $pfxPath -Password $CertificatePassword | Out-Null

    # Export .cer (public key, for Azure AD upload)
    Export-Certificate -Cert $cert -FilePath $cerPath | Out-Null

    Write-Success "Certificate generated"
    Write-Host "       PFX (keep private): $pfxPath"
    Write-Host "       CER (upload to AAD): $cerPath"
}

#endregion

#region ── Step 3 – Azure AD App Registration via PnP ─────────────────────────

Write-Step "Registering Azure AD application: $AppDisplayName"

# PnP PowerShell has a built-in cmdlet that creates the app registration,
# uploads the cert, and grants SharePoint / Graph permissions automatically.
Register-PnPAzureADApp `
    -ApplicationName $AppDisplayName `
    -Tenant $tenantDomain `
    -CertificatePath $pfxPath `
    -CertificatePassword $CertificatePassword `
    -Scopes @(
        # SharePoint
        "SPO.Sites.FullControl.All",
        "SPO.TermStore.Read.All",
        # Microsoft Graph
        "MSGraph.Directory.Read.All",
        "MSGraph.Group.Read.All",
        "MSGraph.Mail.Read",
        "MSGraph.Sites.Read.All",
        "MSGraph.User.Read.All"
    ) `
    -Store CurrentUser `
    -OutPath $PSScriptRoot

Write-Success "App registration created. Check Azure Portal to grant admin consent."

#endregion

#region ── Step 4 – Exchange App-Only Setup ──────────────────────────────────

Write-Step "Configuring Exchange Online app-only access"
Write-Host @"
       Exchange Online requires a separate manual step to enable app-only (certificate) access.
       Run the following commands in a separate window as an Exchange Admin AFTER granting
       admin consent to the app in the Azure Portal:

           Connect-ExchangeOnline -UserPrincipalName admin@$tenantDomain
           New-ServicePrincipal -AppId <AppId> -ServiceId <ObjectId> -DisplayName "$AppDisplayName"
           Add-MailboxPermission -Identity <mailbox> -User <ServicePrincipalId> -AccessRights FullAccess

       For read-only inventory (Get-ExchangeMailboxStats.ps1), the built-in
       'Exchange.ManageAsApp' + 'Exchange Administrator' role assignment is sufficient.
"@

#endregion

#region ── Step 5 – Output Connection Parameters ─────────────────────────────

Write-Step "Writing connection parameters to: $OutputParamFile"

# Read back the client ID that PnP registered
$appInfo = Get-PnPAzureADApp -Identity $AppDisplayName -ErrorAction SilentlyContinue

$params = [ordered]@{
    TenantName      = $TenantName
    TenantId        = "$tenantDomain"   # Replace with actual GUID from Portal if needed
    AppDisplayName  = $AppDisplayName
    ClientId        = if ($appInfo) { $appInfo.AppId } else { "<REPLACE_WITH_APP_CLIENT_ID>" }
    CertificatePath = (Resolve-Path $pfxPath).Path
    AdminSiteUrl    = "https://$TenantName-admin.sharepoint.com"
    TenantRootUrl   = "https://$TenantName.sharepoint.com"
}

$params | ConvertTo-Json | Set-Content -Path $OutputParamFile -Encoding UTF8
Write-Success "Params written to $OutputParamFile"

Write-Host "`n========================================" -ForegroundColor Cyan
Write-Host " NEXT STEPS" -ForegroundColor Cyan
Write-Host "========================================" -ForegroundColor Cyan
Write-Host @"
  1. Go to Azure Portal > Azure Active Directory > App Registrations
  2. Find '$AppDisplayName' and click 'Grant admin consent'
  3. Edit $OutputParamFile and replace <REPLACE_WITH_APP_CLIENT_ID>
     with the actual Application (client) ID from the Portal if not auto-populated
  4. Run the migration inventory scripts using the params file:
       -ParamFile '$OutputParamFile'
"@

#endregion
