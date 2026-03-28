[CmdletBinding()]
param(
    [string]$Path,
    [Alias('h')][switch]$Help,
    [switch]$Full,
    [switch]$Examples
)

function Show-ShortHelp {
@"
Save-VcfOpsTelegrafCredential.ps1

Purpose
  Prompts for a guest OS credential and saves it to a DPAPI-protected CLIXML file
  for later use by the VCF Operations / Telegraf Guest Ops fleet scripts.

Usage
  .\Save-VcfOpsTelegrafCredential.ps1 -Path <credential-file>
  .\Save-VcfOpsTelegrafCredential.ps1 -h
  .\Save-VcfOpsTelegrafCredential.ps1 -Full
  .\Save-VcfOpsTelegrafCredential.ps1 -Examples

Notes
  - The saved file is protected for the current Windows user on the current machine.
  - It is suitable for later use with parameters such as:
      -CredentialFile
      -WindowsCredentialFile
      -LinuxCredentialFile
      AltCredFile (in the CSV, per VM)

Use -Full for detailed help.
Use -Examples for worked examples.
"@
}

function Show-FullHelp {
@"
Save-VcfOpsTelegrafCredential.ps1
=================================

What this script does
---------------------
This script prompts you for a username and password by using Get-Credential, then
saves the resulting PSCredential object to a CLIXML file by using Export-Clixml.

On Windows, the password portion is protected with DPAPI. That means the saved
credential file is normally only readable by the same Windows user on the same
machine that created it.

This is intended to avoid storing guest OS passwords in plaintext in scripts or CSV
files when using the Guest Ops fleet validation scripts.

How it is used
--------------
You run this script once for each credential set you want to save. For example:

- one file for a shared Windows guest admin credential
- one file for a shared Linux guest credential
- one file for a specific VM that needs a different account

Those saved files can then be referenced later by:

- -CredentialFile
- -WindowsCredentialFile
- -LinuxCredentialFile
- AltCredFile in the target CSV for a specific VM

Security model
--------------
The saved file is not plaintext, but it is also not a portable secret vault.
It is designed for practical admin use on one workstation.

What this protects against:
- casual viewing of passwords in a CSV or script
- copying the credential file to another machine or user context and importing it there

What this does not protect against:
- someone already running as your Windows account on the same machine
- malware or compromise of your logged-in session

If you choose not to use credential files, some fleet scripts can read GuestUser and
GuestPassword directly from the CSV. That is easier, but it stores passwords in
plaintext and is therefore a higher-risk option.

Parameter
---------
-Path
  Output path for the DPAPI-protected credential file.

Examples of where the file may be saved:
  C:\Secure\Guest-Windows.xml
  C:\Secure\Guest-Linux.xml
  C:\Secure\Telegraf-test03.xml

Workflow examples
-----------------
1. Create a default Windows guest credential file
   .\Save-VcfOpsTelegrafCredential.ps1 -Path C:\Secure\Guest-Windows.xml

2. Create a default Linux guest credential file
   .\Save-VcfOpsTelegrafCredential.ps1 -Path C:\Secure\Guest-Linux.xml

3. Create a per-VM override credential file
   .\Save-VcfOpsTelegrafCredential.ps1 -Path C:\Secure\Telegraf-test03.xml

4. Use that per-VM file in the fleet CSV by populating AltCredFile
   Example CSV row:
   VMName,GuestUser,GuestPassword,TargetOs,UseSudo,AltCredFile
   Telegraf-test03,,,Auto,False,C:\Secure\Telegraf-test03.xml

How this fits with the fleet runner
-----------------------------------
When using the Guest Ops fleet runner, credential precedence is typically:

1. AltCredFile from the CSV row
2. Default credential file selected by script parameters
   -CredentialFile
   -WindowsCredentialFile
   -LinuxCredentialFile
3. GuestUser and GuestPassword from the CSV

So if AltCredFile is populated for a VM, that VM can use a different credential even
when the rest of the fleet uses the default Windows/Linux credential files.

Related usage pattern
---------------------
Example fleet runner command using saved credential files:

.\Invoke-VcfOpsTelegrafFleetRunner-GuestOps-v4-fixed3-help-cpfqdn-fixed.ps1 `
  -vCenterServer 'vCenter-01.devops.local' `
  -TargetsCsv '.\targets.csv' `
  -vCenterUser 'administrator@vsphere.local' `
  -vCenterPassword `$vcPw `
  -CloudProxyVmName 'CloudProxy-01' `
  -CloudProxyGuestUser 'root' `
  -CloudProxyGuestPassword `$cpPw `
  -WindowsCredentialFile 'C:\Secure\Guest-Windows.xml' `
  -LinuxCredentialFile 'C:\Secure\Guest-Linux.xml'

If you prefer the CSV plaintext method instead, leave those file parameters out and
populate GuestUser and GuestPassword in the CSV. That works, but carries more risk.
"@
}

function Show-Examples {
@"
Examples
========

1. Save a Windows guest credential file
   .\Save-VcfOpsTelegrafCredential.ps1 -Path C:\Secure\Guest-Windows.xml

2. Save a Linux guest credential file
   .\Save-VcfOpsTelegrafCredential.ps1 -Path C:\Secure\Guest-Linux.xml

3. Save a per-VM override credential file
   .\Save-VcfOpsTelegrafCredential.ps1 -Path C:\Secure\Telegraf-test03.xml

4. Use the per-VM override in the CSV
   VMName,GuestUser,GuestPassword,TargetOs,UseSudo,AltCredFile
   Telegraf-test03,,,Auto,False,C:\Secure\Telegraf-test03.xml

5. Show help
   .\Save-VcfOpsTelegrafCredential.ps1 -h
   .\Save-VcfOpsTelegrafCredential.ps1 -Full
   .\Save-VcfOpsTelegrafCredential.ps1 -Examples
"@
}

if ($Help) {
    Show-ShortHelp
    return
}
if ($Full) {
    Show-FullHelp
    return
}
if ($Examples) {
    Show-Examples
    return
}

if ([string]::IsNullOrWhiteSpace($Path)) {
    throw "Path is required unless -h, -Full, or -Examples is used."
}

$cred = Get-Credential -Message 'Enter guest OS credential for Guest Ops fleet validator'
$dir = Split-Path -Parent $Path
if ($dir -and -not (Test-Path $dir)) {
    New-Item -ItemType Directory -Path $dir -Force | Out-Null
}
$cred | Export-Clixml -Path $Path
Write-Host "Credential saved (DPAPI-protected for current user on this machine): $Path" -ForegroundColor Green
