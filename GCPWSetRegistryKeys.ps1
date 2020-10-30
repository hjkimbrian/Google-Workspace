<# This script downloads Google Credential Provider for Windows from
https://tools.google.com/dlpage/gcpw/, then installs and configures it.
Windows administrator access is required to use the script. 
It also downloads Google Chrome from https://chromeenterprise.google/browser/download/
and enrolls the browser for Cloud Management for ease of administration.
Script modified from https://support.google.com/cloudidentity/answer/9250996?hl=en #>

<# Set the following key to the domains you want to allow users to sign in from.

For example:
$domainsAllowedToLogin = "acme1.com,acme2.com"

Also get the Chrome Enrollment Token from admin.google.com
Instructions: https://support.google.com/chrome/a/answer/9301891?hl=en #>

$domainsAllowedToLogin = ""
$enrollmenttoken = ""

Add-Type -AssemblyName System.Drawing
Add-Type -AssemblyName PresentationFramework

<# Check if one or more domains are set #>
if ($domainsAllowedToLogin.Equals('')) {
    $msgResult = [System.Windows.MessageBox]::Show('The list of domains cannot be empty! Please edit this script.', 'GCPW', 'OK', 'Error')
    exit 5
}

function Is-Admin() {
    $admin = [bool](([System.Security.Principal.WindowsIdentity]::GetCurrent()).groups -match 'S-1-5-32-544')
    return $admin
}

<# Check if the current user is an admin and exit if they aren't. #>
if (-not (Is-Admin)) {
    $result = [System.Windows.MessageBox]::Show('Please run as administrator!', 'GCPW', 'OK', 'Error')
    exit 5
}

<# Choose the Chrome file to download. 32-bit and 64-bit versions have different names #>
$chromeFileName = 'googlechromestandaloneenterprise.msi'
if ([Environment]::Is64BitOperatingSystem) {
    $chromeFileName = 'googlechromestandaloneenterprise64.msi'
}

<# Download the Chrome installer. #>
$chromeUrlPrefix = 'https://dl.google.com/chrome/install/'
$chromeUri = $chromeUrlPrefix + $chromeFileName
Write-Host 'Downloading Chrome from' $chromeUri
Invoke-WebRequest -Uri $chromeUri -OutFile $chromeFileName

<# Run the Chrome installer and wait for the installation to finish #>
$arguments = "/i `"$chromeFileName`""
$installProcess = (Start-Process msiexec.exe -ArgumentList $arguments -PassThru -Wait)

<# Check if installation was successful #>
if ($installProcess.ExitCode -ne 0) {
    $result = [System.Windows.MessageBox]::Show('Installation failed!', 'Chrome', 'OK', 'Error')
    exit $installProcess.ExitCode
}
else {
    $result = [System.Windows.MessageBox]::Show('Installation completed successfully!', 'Chrome', 'OK', 'Info')
}


<# Choose the GCPW file to download. 32-bit and 64-bit versions have different names #>
$gcpwFileName = 'gcpwstandaloneenterprise.msi'
if ([Environment]::Is64BitOperatingSystem) {
    $gcpwFileName = 'gcpwstandaloneenterprise64.msi'
}

<# Download the GCPW installer. #>
$gcpwUrlPrefix = 'https://dl.google.com/credentialprovider/'
$gcpwUri = $gcpwUrlPrefix + $gcpwFileName
Write-Host 'Downloading GCPW from' $gcpwUri
Invoke-WebRequest -Uri $gcpwUri -OutFile $gcpwFileName

<# Run the GCPW installer and wait for the installation to finish #>
$arguments = "/i `"$gcpwFileName`""
$installProcess = (Start-Process msiexec.exe -ArgumentList $arguments -PassThru -Wait)

<# Check if installation was successful #>
if ($installProcess.ExitCode -ne 0) {
    $result = [System.Windows.MessageBox]::Show('Installation failed!', 'GCPW', 'OK', 'Error')
    exit $installProcess.ExitCode
}
else {
    $result = [System.Windows.MessageBox]::Show('Installation completed successfully!', 'GCPW', 'OK', 'Info')
}

<# Set the required registry key with the allowed domains #>
$registryPath = 'HKEY_LOCAL_MACHINE\Software\Google\GCPW'
$name = 'domains_allowed_to_login'
[microsoft.win32.registry]::SetValue($registryPath, $name, $domainsAllowedToLogin)

$domains = Get-ItemPropertyValue HKLM:\Software\Google\GCPW -Name $name

if ($domains -eq $domainsAllowedToLogin) {
    $msgResult = [System.Windows.MessageBox]::Show('Configuration completed successfully!', 'GCPW', 'OK', 'Info')
}
else {
    $msgResult = [System.Windows.MessageBox]::Show('Could not write to registry. Configuration was not completed.', 'GCPW', 'OK', 'Error')

}

<# Set the required registry key to enroll the browser 
See https://www.reddit.com/r/gsuite/comments/igwvwz/can_i_deploy_a_managed_browser_through_gcpw/ 
for an alternative solution using Enhanced Desktop Security for Windows #>
$enrollmentregistryPath = 'HKEY_LOCAL_MACHINE\Software\Policies\Google\Chrome'
$enrollmentname = 'CloudManagementEnrollmentToken'
[microsoft.win32.registry]::SetValue($enrollmentregistryPath, $enrollmentname, $enrollmenttoken)

$tokens = Get-ItemPropertyValue HKLM:\Software\Policies\Google\Chrome -Name $enrollmentname

if ($tokens -eq $enrollmenttoken) {
    $msgResult = [System.Windows.MessageBox]::Show('Configuration completed successfully!', 'Enrollment', 'OK', 'Info')
}
else {
    $msgResult = [System.Windows.MessageBox]::Show('Could not write to registry. Configuration was not completed.', 'Enrollment', 'OK', 'Error')

}
