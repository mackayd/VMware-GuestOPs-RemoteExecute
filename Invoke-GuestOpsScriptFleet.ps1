#requires -Version 7.0
[CmdletBinding()]
param(
  [Alias('h')][switch]$Help,
  [switch]$Full,
  [switch]$Examples,

  [Parameter(Mandatory=$true)][string]$vCenterServer,
  [Parameter(Mandatory=$true)][string]$TargetsCsv,

  [string]$vCenterUser,
  [securestring]$vCenterPassword,

  [string]$ScriptPath,
  [ValidateSet('Auto','PowerShell','Bash','Bat','Custom')][string]$ScriptLanguage = 'Auto',
  [string]$ExecutionCommandTemplate,

  [string]$WindowsScriptPath,
  [ValidateSet('Auto','PowerShell','Bat','Custom')][string]$WindowsScriptLanguage = 'Auto',
  [string]$WindowsExecutionCommandTemplate,

  [string]$LinuxScriptPath,
  [ValidateSet('Auto','Bash','Custom')][string]$LinuxScriptLanguage = 'Auto',
  [string]$LinuxExecutionCommandTemplate,

  [string]$VMNameColumn = 'VMName',
  [string]$ComputerNameColumn = 'ComputerName',
  [string]$GuestUserColumn = 'GuestUser',
  [string]$GuestPasswordColumn = 'GuestPassword',
  [string]$TargetOsColumn = 'TargetOs',
  [string]$AltCredFileColumn = 'AltCredFile',
  [ValidateSet('Auto','Windows','Linux')][string]$TargetOs = 'Auto',

  [switch]$PromptForGuestCredential,
  [string]$CredentialFile,
  [switch]$PromptForWindowsCredential,
  [switch]$PromptForLinuxCredential,
  [string]$WindowsCredentialFile,
  [string]$LinuxCredentialFile,

  [string]$OutDir = '.',
  [switch]$ContinueOnError
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

function W {
  param([string]$Message,[ValidateSet('INFO','PASS','WARN','FAIL')][string]$Level='INFO')
  $color = @{ INFO='Cyan'; PASS='Green'; WARN='Yellow'; FAIL='Red' }[$Level]
  Write-Host "[$Level] $Message" -ForegroundColor $color
}

function New-PlainTextCredential {
  param(
    [Parameter(Mandatory=$true)][string]$UserName,
    [Parameter(Mandatory=$true)][securestring]$Password
  )
  [pscredential]::new($UserName,$Password)
}

function Import-CredentialFromFile {
  param([Parameter(Mandatory=$true)][string]$Path)
  if (-not (Test-Path -LiteralPath $Path)) { throw "Credential file not found: $Path" }
  $cred = Import-Clixml -LiteralPath $Path
  if ($cred -isnot [pscredential]) { throw "Credential file '$Path' did not contain a PSCredential object." }
  return $cred
}

function Escape-ForSingleQuotedPowerShell {
  param([AllowNull()][string]$Value)
  if ($null -eq $Value) { return '' }
  return ($Value -replace "'","''")
}

function Escape-ForSingleQuotedBash {
  param([AllowNull()][string]$Value)
  if ($null -eq $Value) { return '' }
  $bashSingleQuoteEscape = "'" + '"' + "'" + '"' + "'"
  return ($Value -replace "'",$bashSingleQuoteEscape)
}

function Resolve-TargetVmName {
  param($Row,[string]$PrimaryColumn,[string]$FallbackColumn)
  $name = $null
  if ($Row.PSObject.Properties.Name -contains $PrimaryColumn) { $name = [string]$Row.$PrimaryColumn }
  if ([string]::IsNullOrWhiteSpace($name) -and $Row.PSObject.Properties.Name -contains $FallbackColumn) {
    $name = [string]$Row.$FallbackColumn
  }
  return $name
}

function Resolve-RequestedTargetOs {
  param($Row,[string]$ColumnName,[string]$DefaultValue)
  $rowTargetOs = $DefaultValue
  if ($Row.PSObject.Properties.Name -contains $ColumnName) {
    $candidate = [string]$Row.$ColumnName
    if ($candidate -in @('Auto','Windows','Linux')) { $rowTargetOs = $candidate }
  }
  return $rowTargetOs
}

function Resolve-DetectedTargetOs {
  param(
    [Parameter(Mandatory=$true)]$VM,
    [Parameter(Mandatory=$true)][string]$RequestedTargetOs
  )
  $guestFamily = [string]$VM.ExtensionData.Guest.GuestFamily
  $guestFullName = [string]$VM.ExtensionData.Config.GuestFullName
  $guestId = [string]$VM.ExtensionData.Config.GuestId
  if ($RequestedTargetOs -ne 'Auto') {
    return [pscustomobject]@{
      DetectedTargetOs = $RequestedTargetOs
      OsDetectionSource = 'CSV override'
      GuestFullName = $guestFullName
      GuestFamily = $guestFamily
      GuestId = $guestId
    }
  }
  $detected = 'Windows'
  if ($guestFamily -match 'linux' -or
      $guestId -match 'linux|ubuntu|rhel|centos|rocky|suse|photon|oracle|debian|alma' -or
      $guestFullName -match 'Linux|Ubuntu|CentOS|Red Hat|Rocky|SUSE|Photon|Oracle|Debian|Alma') {
    $detected = 'Linux'
  }
  [pscustomobject]@{
    DetectedTargetOs = $detected
    OsDetectionSource = 'vCenter guest metadata'
    GuestFullName = $guestFullName
    GuestFamily = $guestFamily
    GuestId = $guestId
  }
}

function Resolve-RowCredential {
  param(
    [Parameter(Mandatory=$true)]$Row,
    [Parameter(Mandatory=$true)][string]$DetectedTargetOs,
    [Parameter(Mandatory=$true)][string]$GuestUserColumn,
    [Parameter(Mandatory=$true)][string]$GuestPasswordColumn,
    [Parameter(Mandatory=$true)][string]$AltCredFileColumn,
    [AllowNull()]$GlobalCredential,
    [AllowNull()]$WindowsCredential,
    [AllowNull()]$LinuxCredential,
    [string]$CredentialFilePath,
    [string]$WindowsCredentialFilePath,
    [string]$LinuxCredentialFilePath
  )

  $altCredFile = $null
  if ($Row.PSObject.Properties.Name -contains $AltCredFileColumn) {
    $candidate = [string]$Row.$AltCredFileColumn
    if (-not [string]::IsNullOrWhiteSpace($candidate)) { $altCredFile = $candidate }
  }

  if ($altCredFile) {
    $cred = Import-CredentialFromFile -Path $altCredFile
    return [pscustomobject]@{
      Credential = $cred
      GuestUser = $cred.UserName
      CredentialSource = 'AlternateCredentialFile'
      CredentialSourceDetail = $altCredFile
      AltCredentialFile = $altCredFile
    }
  }

  if ($DetectedTargetOs -eq 'Windows' -and $WindowsCredential) {
    return [pscustomobject]@{
      Credential = $WindowsCredential
      GuestUser = $WindowsCredential.UserName
      CredentialSource = 'DefaultWindowsCredentialFile'
      CredentialSourceDetail = $WindowsCredentialFilePath
      AltCredentialFile = $null
    }
  }

  if ($DetectedTargetOs -eq 'Linux' -and $LinuxCredential) {
    return [pscustomobject]@{
      Credential = $LinuxCredential
      GuestUser = $LinuxCredential.UserName
      CredentialSource = 'DefaultLinuxCredentialFile'
      CredentialSourceDetail = $LinuxCredentialFilePath
      AltCredentialFile = $null
    }
  }

  if ($GlobalCredential) {
    return [pscustomobject]@{
      Credential = $GlobalCredential
      GuestUser = $GlobalCredential.UserName
      CredentialSource = 'CredentialFile'
      CredentialSourceDetail = $CredentialFilePath
      AltCredentialFile = $null
    }
  }

  $guestUser = $null
  $guestPassword = $null
  if ($Row.PSObject.Properties.Name -contains $GuestUserColumn) { $guestUser = [string]$Row.$GuestUserColumn }
  if ($Row.PSObject.Properties.Name -contains $GuestPasswordColumn) { $guestPassword = [string]$Row.$GuestPasswordColumn }
  $rowVmName = if ($Row.PSObject.Properties.Name -contains 'VMName') { [string]$Row.VMName } elseif ($Row.PSObject.Properties.Name -contains 'ComputerName') { [string]$Row.ComputerName } else { '<row>' }
  if ([string]::IsNullOrWhiteSpace($guestUser)) { throw "Missing GuestUser for $rowVmName" }
  if ([string]::IsNullOrWhiteSpace($guestPassword)) { throw "Missing GuestPassword for $rowVmName" }
  $securePassword = ConvertTo-SecureString $guestPassword -AsPlainText -Force
  $cred = [pscredential]::new($guestUser,$securePassword)
  return [pscustomobject]@{
    Credential = $cred
    GuestUser = $guestUser
    CredentialSource = 'CSV plaintext'
    CredentialSourceDetail = $null
    AltCredentialFile = $null
  }
}

function Get-LocalScriptPath {
  param(
    [string]$CurrentPath,
    [string]$PromptText = 'Enter the full local path to the script payload you want to execute',
    [string]$DialogTitle = 'Select the payload script to run in the guest'
  )

  if (-not [string]::IsNullOrWhiteSpace($CurrentPath)) {
    return $CurrentPath
  }

  $selected = $null
  if ($IsWindows) {
    try {
      Add-Type -AssemblyName System.Windows.Forms -ErrorAction Stop
      $dialog = New-Object System.Windows.Forms.OpenFileDialog
      $dialog.Title = $DialogTitle
      $dialog.Filter = 'Script files (*.ps1;*.bat;*.cmd;*.sh;*.py;*.pl;*.rb;*.psm1)|*.ps1;*.bat;*.cmd;*.sh;*.py;*.pl;*.rb;*.psm1|All files (*.*)|*.*'
      $dialog.Multiselect = $false
      if ($dialog.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
        $selected = $dialog.FileName
      }
    } catch {
      # Fall back to Read-Host
    }
  }

  if ([string]::IsNullOrWhiteSpace($selected)) {
    $selected = Read-Host $PromptText
  }

  return $selected
}

function Resolve-ScriptMode {
  param(
    [Parameter(Mandatory=$true)][string]$LocalScriptPath,
    [Parameter(Mandatory=$true)][ValidateSet('Auto','PowerShell','Bash','Bat','Custom')][string]$RequestedLanguage
  )

  $ext = [System.IO.Path]::GetExtension($LocalScriptPath)
  if ($RequestedLanguage -ne 'Auto') {
    return [pscustomobject]@{
      ScriptMode = $RequestedLanguage
      Extension = $ext
    }
  }

  switch ($ext.ToLowerInvariant()) {
    '.ps1' { $mode = 'PowerShell' }
    '.psm1' { $mode = 'PowerShell' }
    '.bat' { $mode = 'Bat' }
    '.cmd' { $mode = 'Bat' }
    '.sh'  { $mode = 'Bash' }
    default {
      throw "Unable to infer script language from extension '$ext'. Use -ScriptLanguage PowerShell, Bash, Bat, or Custom."
    }
  }


  [pscustomobject]@{
    ScriptMode = $mode
    Extension = $ext
  }
}

function Resolve-PayloadDefinition {
  param(
    [Parameter(Mandatory=$true)][ValidateSet('Windows','Linux')][string]$DetectedTargetOs,
    [AllowNull()][string]$SharedScriptPath,
    [AllowNull()][string]$SharedScriptLanguage,
    [AllowNull()][string]$SharedExecutionCommandTemplate,
    [AllowNull()][string]$WindowsScriptPath,
    [AllowNull()][string]$WindowsScriptLanguage,
    [AllowNull()][string]$WindowsExecutionCommandTemplate,
    [AllowNull()][string]$LinuxScriptPath,
    [AllowNull()][string]$LinuxScriptLanguage,
    [AllowNull()][string]$LinuxExecutionCommandTemplate,
    [Parameter(Mandatory=$true)][hashtable]$PayloadCache
  )

  if ($PayloadCache.ContainsKey($DetectedTargetOs)) {
    return $PayloadCache[$DetectedTargetOs]
  }

  $scriptPath = $null
  $requestedLanguage = 'Auto'
  $executionCommandTemplate = $null
  $dialogTitle = $null
  $promptText = $null
  $source = $null

  switch ($DetectedTargetOs) {
    'Windows' {
      $scriptPath = $WindowsScriptPath
      $requestedLanguage = if ($WindowsScriptLanguage) { $WindowsScriptLanguage } else { 'Auto' }
      $executionCommandTemplate = $WindowsExecutionCommandTemplate
      $dialogTitle = 'Select the Windows payload script to run in the guest'
      $promptText = 'Enter the full local path to the Windows payload script you want to execute'
      $source = 'Windows-specific'
    }
    'Linux' {
      $scriptPath = $LinuxScriptPath
      $requestedLanguage = if ($LinuxScriptLanguage) { $LinuxScriptLanguage } else { 'Auto' }
      $executionCommandTemplate = $LinuxExecutionCommandTemplate
      $dialogTitle = 'Select the Linux payload script to run in the guest'
      $promptText = 'Enter the full local path to the Linux payload script you want to execute'
      $source = 'Linux-specific'
    }
  }

  if ([string]::IsNullOrWhiteSpace($scriptPath)) {
    $scriptPath = $SharedScriptPath
    $requestedLanguage = if ($SharedScriptLanguage) { $SharedScriptLanguage } else { 'Auto' }
    $executionCommandTemplate = $SharedExecutionCommandTemplate
    $dialogTitle = "Select the $DetectedTargetOs payload script to run in the guest"
    $promptText = "Enter the full local path to the $DetectedTargetOs payload script you want to execute"
    $source = 'Shared'
  }

  $resolvedScriptPath = Get-LocalScriptPath -CurrentPath $scriptPath -PromptText $promptText -DialogTitle $dialogTitle
  if ([string]::IsNullOrWhiteSpace($resolvedScriptPath)) {
    throw "No $DetectedTargetOs payload path was supplied."
  }
  if (-not (Test-Path -LiteralPath $resolvedScriptPath)) {
    throw "$DetectedTargetOs payload not found: $resolvedScriptPath"
  }

  $scriptInfo = Resolve-ScriptMode -LocalScriptPath $resolvedScriptPath -RequestedLanguage $requestedLanguage
  if ($scriptInfo.ScriptMode -eq 'Custom' -and [string]::IsNullOrWhiteSpace($executionCommandTemplate)) {
    throw "$DetectedTargetOs ExecutionCommandTemplate is required when the selected script mode is Custom."
  }

  $definition = [pscustomobject]@{
    DetectedTargetOs = $DetectedTargetOs
    PayloadSource = $source
    ScriptPath = (Resolve-Path -LiteralPath $resolvedScriptPath).Path
    ScriptName = [System.IO.Path]::GetFileName($resolvedScriptPath)
    ScriptInfo = $scriptInfo
    PayloadBytes = [System.IO.File]::ReadAllBytes((Resolve-Path -LiteralPath $resolvedScriptPath))
    ExecutionCommandTemplate = $executionCommandTemplate
  }

  $PayloadCache[$DetectedTargetOs] = $definition
  return $definition
}

function Assert-ScriptModeCompatible {
  param(
    [Parameter(Mandatory=$true)][string]$DetectedTargetOs,
    [Parameter(Mandatory=$true)][ValidateSet('PowerShell','Bash','Bat','Custom')][string]$ScriptMode
  )

  switch ($DetectedTargetOs) {
    'Windows' {
      if ($ScriptMode -notin @('PowerShell','Bat','Custom')) {
        throw "Script mode '$ScriptMode' is not compatible with Windows guest execution."
      }
    }
    'Linux' {
      if ($ScriptMode -notin @('Bash','Custom')) {
        throw "Script mode '$ScriptMode' is not compatible with Linux guest execution."
      }
    }
    default {
      throw "Unsupported detected target OS '$DetectedTargetOs'."
    }
  }
}

function New-WindowsGuestBootstrapScript {
  param(
    [Parameter(Mandatory=$true)][byte[]]$PayloadBytes,
    [Parameter(Mandatory=$true)][string]$PayloadExtension,
    [Parameter(Mandatory=$true)][ValidateSet('PowerShell','Bat','Custom')][string]$ScriptMode,
    [AllowNull()][string]$ExecutionCommandTemplate
  )

  $payloadB64 = [Convert]::ToBase64String($PayloadBytes)
  $ext = Escape-ForSingleQuotedPowerShell $PayloadExtension
  $mode = Escape-ForSingleQuotedPowerShell $ScriptMode
  $template = Escape-ForSingleQuotedPowerShell $ExecutionCommandTemplate

  return @"
`$ErrorActionPreference = 'Stop'
function Join-QuotedForCmd([string[]]`$Values) {
  (`$Values | ForEach-Object {
    '"' + (([string]`$_) -replace '"','\"') + '"'
  }) -join ' '
}

`$workRoot = Join-Path `$env:TEMP ('GuestOpsScriptRunner-' + [guid]::NewGuid().ToString('N'))
New-Item -Path `$workRoot -ItemType Directory -Force | Out-Null
`$payloadPath = Join-Path `$workRoot ('payload$ext')
[System.IO.File]::WriteAllBytes(`$payloadPath, [Convert]::FromBase64String('$payloadB64'))

`$mode = '$mode'
`$template = '$template'
`$output = ''
`$commandDescription = ''
`$exitCode = 0

switch (`$mode) {
  'PowerShell' {
    `$commandDescription = 'powershell.exe -NoProfile -ExecutionPolicy Bypass -File "' + `$payloadPath + '"'
    `$output = (& powershell.exe -NoProfile -ExecutionPolicy Bypass -File `$payloadPath 2>&1 | Out-String)
    `$exitCode = if (`$LASTEXITCODE -is [int]) { `$LASTEXITCODE } else { 0 }
  }
  'Bat' {
    `$commandDescription = 'cmd.exe /c "' + `$payloadPath + '"'
    `$cmdArgs = '/c "' + `$payloadPath.Replace('"','""') + '"'
    `$output = (& cmd.exe `$cmdArgs 2>&1 | Out-String)
    `$exitCode = if (`$LASTEXITCODE -is [int]) { `$LASTEXITCODE } else { 0 }
  }
  'Custom' {
    if ([string]::IsNullOrWhiteSpace(`$template)) { throw 'ExecutionCommandTemplate is required when ScriptLanguage is Custom.' }
    `$commandLine = `$template.Replace('{ScriptPath}', '"' + `$payloadPath + '"')
    `$commandDescription = `$commandLine
    `$output = (& cmd.exe /c `$commandLine 2>&1 | Out-String)
    `$exitCode = if (`$LASTEXITCODE -is [int]) { `$LASTEXITCODE } else { 0 }
  }
  default {
    throw ('Unsupported Windows script mode: ' + `$mode)
  }
}

Write-Output ('GOR_RESULT INFO PAYLOAD_PATH:' + `$payloadPath)
Write-Output ('GOR_RESULT INFO EXECUTION_MODE:' + `$mode)
if (`$commandDescription) { Write-Output ('GOR_RESULT INFO EXECUTION_COMMAND:' + `$commandDescription) }
Write-Output ('GOR_RESULT INFO WORKDIR:' + `$workRoot)
Write-Output 'GOR_RESULT OUTPUT_BEGIN'
if (`$output) {
  Write-Output (`$output.TrimEnd("`r","`n"))
}
Write-Output 'GOR_RESULT OUTPUT_END'
Write-Output ('GOR_RESULT EXITCODE:' + `$exitCode)
exit `$exitCode
"@
}

function New-LinuxGuestBootstrapScript {
  param(
    [Parameter(Mandatory=$true)][byte[]]$PayloadBytes,
    [Parameter(Mandatory=$true)][string]$PayloadExtension,
    [Parameter(Mandatory=$true)][ValidateSet('Bash','Custom')][string]$ScriptMode,
    [AllowNull()][string]$ExecutionCommandTemplate
  )

  $payloadB64 = [Convert]::ToBase64String($PayloadBytes)
  $ext = $PayloadExtension
  $mode = Escape-ForSingleQuotedBash $ScriptMode
  $template = Escape-ForSingleQuotedBash $ExecutionCommandTemplate

  $linuxTemplate = @'
set -u

MODE='__MODE__'
TEMPLATE='__TEMPLATE__'
WORKDIR="/tmp/guestops-script-runner-$(date +%s)-$$"
mkdir -p "$WORKDIR"
PAYLOAD_PATH="$WORKDIR/payload__EXT__"

if command -v base64 >/dev/null 2>&1; then
  cat <<'__GOR_PAYLOAD__' | base64 -d > "$PAYLOAD_PATH"
__PAYLOAD_B64__
__GOR_PAYLOAD__
else
  echo 'GOR_RESULT INFO ERROR:base64 utility not found in guest'
  echo 'GOR_RESULT EXITCODE:127'
  exit 127
fi

chmod +x "$PAYLOAD_PATH" 2>/dev/null || true

COMMAND_DESC=''
OUTPUT_FILE="$WORKDIR/stdout.txt"
EXIT_CODE=0

case "$MODE" in
  Bash)
    COMMAND_DESC="/bin/bash \"$PAYLOAD_PATH\""
    /bin/bash "$PAYLOAD_PATH" >"$OUTPUT_FILE" 2>&1
    EXIT_CODE=$?
    ;;
  Custom)
    if [ -z "$TEMPLATE" ]; then
      echo 'GOR_RESULT INFO ERROR:ExecutionCommandTemplate is required when ScriptLanguage is Custom.'
      echo 'GOR_RESULT EXITCODE:2'
      exit 2
    fi
    COMMAND_LINE=${TEMPLATE//\{ScriptPath\}/\"$PAYLOAD_PATH\"}
    COMMAND_DESC="$COMMAND_LINE"
    sh -lc "$COMMAND_LINE" >"$OUTPUT_FILE" 2>&1
    EXIT_CODE=$?
    ;;
  *)
    echo "GOR_RESULT INFO ERROR:Unsupported Linux script mode: $MODE"
    echo 'GOR_RESULT EXITCODE:2'
    exit 2
    ;;
esac

echo "GOR_RESULT INFO PAYLOAD_PATH:$PAYLOAD_PATH"
echo "GOR_RESULT INFO EXECUTION_MODE:$MODE"
if [ -n "$COMMAND_DESC" ]; then
  echo "GOR_RESULT INFO EXECUTION_COMMAND:$COMMAND_DESC"
fi
echo "GOR_RESULT INFO WORKDIR:$WORKDIR"
echo 'GOR_RESULT OUTPUT_BEGIN'
if [ -f "$OUTPUT_FILE" ]; then
  cat "$OUTPUT_FILE"
fi
echo 'GOR_RESULT OUTPUT_END'
echo "GOR_RESULT EXITCODE:$EXIT_CODE"
exit $EXIT_CODE
'@

  return $linuxTemplate.Replace('__MODE__', $mode).`
                        Replace('__TEMPLATE__', $template).`
                        Replace('__EXT__', $ext).`
                        Replace('__PAYLOAD_B64__', $payloadB64)
}

function Get-OverallStatus {
  param([int]$ExitCode)
  if ($ExitCode -eq 0) { return 'PASS' }
  return 'FAIL'
}

function Get-SafeFileName {
  param([string]$Name)
  if ([string]::IsNullOrWhiteSpace($Name)) { return 'unknown' }
  $invalid = [System.IO.Path]::GetInvalidFileNameChars()
  $result = $Name
  foreach ($char in $invalid) {
    $result = $result.Replace([string]$char,'_')
  }
  return $result
}

function Show-ShortHelp {
@"
Invoke-GuestOpsScriptFleet.ps1

Purpose
  Generic fleet guest-operations runner that uploads a local script payload into each target VM
  and executes it remotely through VMware Tools / Invoke-VMScript.

Quick usage
  .\Invoke-GuestOpsScriptFleet.ps1 `
    -vCenterServer 'vcenter01.domain.local' `
    -TargetsCsv '.\targets.csv' `
    -WindowsScriptPath '.\Collect-WindowsInfo.ps1' `
    -LinuxScriptPath '.\collect_linux_info.sh' `
    -vCenterUser 'administrator@vsphere.local' `
    -vCenterPassword $vcPw

What it does
  - Prompts you to select a local script if no OS-specific or shared script path is supplied
  - Reuses the same target CSV model as the Telegraf GuestOps tool
  - Detects Windows vs Linux guests unless TargetOs is overridden
  - Executes the Windows payload on Windows VMs and the Linux payload on Linux VMs
  - Writes a summary CSV/JSON and one detailed log per VM

Default CSV columns
  VMName,GuestUser,GuestPassword,TargetOs,AltCredFile

Modes
  - Auto        Infer from extension (.ps1, .bat/.cmd, .sh)
  - PowerShell  Run the payload with powershell.exe inside Windows guests
  - Bat         Run the payload with cmd.exe inside Windows guests
  - Bash        Run the payload with /bin/bash inside Linux guests
  - Custom      Use -ExecutionCommandTemplate with {ScriptPath}

Help switches
  -h / -Help
  -Full
  -Examples
"@ | Write-Host
}

function Show-FullHelp {
@"
NAME
  Invoke-GuestOpsScriptFleet.ps1

SYNOPSIS
  Uploads and executes a script payload across a fleet of VMs using VMware Tools Guest Operations.

HOW IT WORKS
  1. Reads the target CSV.
  2. Resolves the payload per guest OS using -WindowsScriptPath / -LinuxScriptPath, falling back to the shared -ScriptPath if needed.
  3. Connects to vCenter.
  4. Resolves each VM by name.
  5. Detects the guest OS from vCenter metadata unless TargetOs is explicitly set.
  6. Selects credentials based on AltCredFile, OS-specific default credential file, global credential file, or CSV GuestUser/GuestPassword.
  7. Creates a temporary payload file inside the guest.
  8. Executes the payload using the selected mode for that guest OS.
  9. Writes:
       - GuestOpsScriptFleetSummary-<timestamp>.csv
       - GuestOpsScriptFleetSummary-<timestamp>.json
       - One log file per VM

CSV FORMAT
  Typical columns:
    VMName,GuestUser,GuestPassword,TargetOs,AltCredFile

  Notes:
  - VMName is preferred.
  - ComputerName is supported as a fallback.
  - TargetOs may be Auto, Windows, or Linux.
  - AltCredFile allows a single VM to use a different credential file.

SUPPORTED MODES
  Auto
    - Windows payloads: .ps1 => PowerShell, .bat/.cmd => Bat
    - Linux payloads:   .sh  => Bash

  PowerShell
    Executes the payload with:
      powershell.exe -NoProfile -ExecutionPolicy Bypass -File <payload>

  Bat
    Executes the payload with:
      cmd.exe /c <payload>

  Bash
    Executes the payload with:
      /bin/bash <payload>

  Custom
    Requires -ExecutionCommandTemplate.
    Use {ScriptPath} in the template.

    Examples:
      -ExecutionCommandTemplate 'python3 {ScriptPath}'
      -ExecutionCommandTemplate 'pwsh -File {ScriptPath}'
      -ExecutionCommandTemplate 'perl {ScriptPath}'

CREDENTIAL OPTIONS
  Option 1 - OS-specific default credential files (recommended for mixed fleets)
    -WindowsCredentialFile <path>
    -LinuxCredentialFile   <path>

  Option 2 - One global credential file for all VMs
    -CredentialFile <path>

  Option 3 - Prompt interactively
    -PromptForWindowsCredential
    -PromptForLinuxCredential
    or
    -PromptForGuestCredential

  Option 4 - Plaintext credentials in the CSV
    GuestUser and GuestPassword columns

REQUIREMENTS
  - PowerShell 7+
  - VMware PowerCLI
  - vCenter connectivity from the admin workstation
  - VMware Tools running in each target VM
  - Valid guest OS credentials for each target VM
  - Permissions in vCenter for guest operations / Invoke-VMScript

NOTES
  - The payload must be a text-based script.
  - Custom mode lets you run other script types if the guest already has the required interpreter installed.
  - Linux guest execution requires a base64 utility in the guest because the payload is staged through a base64 here-document.
"@ | Write-Host
}

function Show-ExamplesHelp {
@"
EXAMPLE 1 - Prompt for each OS payload path as needed and infer mode from extension
  .\Invoke-GuestOpsScriptFleet.ps1 `
    -vCenterServer 'vcenter01.domain.local' `
    -TargetsCsv '.\targets.csv' `
    -vCenterUser 'administrator@vsphere.local' `
    -vCenterPassword $vcPw

EXAMPLE 2 - Run different payloads for Windows and Linux in one mixed fleet
  .\Invoke-GuestOpsScriptFleet.ps1 `
    -vCenterServer 'vcenter01.domain.local' `
    -TargetsCsv '.\targets-mixed.csv' `
    -WindowsScriptPath '.\Collect-WindowsInfo.ps1' `
    -WindowsScriptLanguage PowerShell `
    -LinuxScriptPath '.\collect_linux_info.sh' `
    -LinuxScriptLanguage Bash `
    -WindowsCredentialFile '.\WindowsGuestCred.xml' `
    -LinuxCredentialFile '.\LinuxGuestCred.xml'

EXAMPLE 3 - Run only a Windows payload on Windows VMs
  .\Invoke-GuestOpsScriptFleet.ps1 `
    -vCenterServer 'vcenter01.domain.local' `
    -TargetsCsv '.\targets-windows.csv' `
    -WindowsScriptPath '.\Collect-WindowsInfo.ps1' `
    -WindowsScriptLanguage PowerShell `
    -WindowsCredentialFile '.\WindowsGuestCred.xml'

EXAMPLE 4 - Run a Python payload on Linux using custom mode
  .\Invoke-GuestOpsScriptFleet.ps1 `
    -vCenterServer 'vcenter01.domain.local' `
    -TargetsCsv '.\targets-linux.csv' `
    -LinuxScriptPath '.\collect_info.py' `
    -LinuxScriptLanguage Custom `
    -LinuxExecutionCommandTemplate 'python3 {ScriptPath}' `
    -LinuxCredentialFile '.\LinuxGuestCred.xml'

EXAMPLE 5 - Use a shared fallback payload for all targets
  .\Invoke-GuestOpsScriptFleet.ps1 `
    -vCenterServer 'vcenter01.domain.local' `
    -TargetsCsv '.\targets.csv' `
    -ScriptPath '.\collect.ps1'

EXAMPLE 6 - Use plaintext credentials from the CSV
  VMName,GuestUser,GuestPassword,TargetOs,AltCredFile
  server01,administrator,MyPassword123!,Windows,
  server02,root,MyLinuxPassword!,Linux,

  .\Invoke-GuestOpsScriptFleet.ps1 `
    -vCenterServer 'vcenter01.domain.local' `
    -TargetsCsv '.\targets.csv' `
    -ScriptPath '.\collect.ps1'
"@ | Write-Host
}

if ($Help) { Show-ShortHelp; return }
if ($Full) { Show-FullHelp; return }
if ($Examples) { Show-ExamplesHelp; return }

if (-not (Test-Path -LiteralPath $TargetsCsv)) { throw "Targets CSV not found: $TargetsCsv" }
if (-not (Test-Path -LiteralPath $OutDir)) { New-Item -Path $OutDir -ItemType Directory -Force | Out-Null }

$targets = Import-Csv -LiteralPath $TargetsCsv
$summary = [System.Collections.Generic.List[object]]::new()
$payloadCache = @{}

$globalCredential = $null
$windowsCredential = $null
$linuxCredential = $null

if ($CredentialFile) { $globalCredential = Import-CredentialFromFile -Path $CredentialFile }
elseif ($PromptForGuestCredential) { $globalCredential = Get-Credential -Message 'Enter guest credential for all VMs' }

if ($WindowsCredentialFile) { $windowsCredential = Import-CredentialFromFile -Path $WindowsCredentialFile }
elseif ($PromptForWindowsCredential) { $windowsCredential = Get-Credential -Message 'Enter default Windows guest credential' }

if ($LinuxCredentialFile) { $linuxCredential = Import-CredentialFromFile -Path $LinuxCredentialFile }
elseif ($PromptForLinuxCredential) { $linuxCredential = Get-Credential -Message 'Enter default Linux guest credential' }

$vcCred = $null
if ($vCenterUser -and $vCenterPassword) { $vcCred = New-PlainTextCredential -UserName $vCenterUser -Password $vCenterPassword }
elseif ($vCenterUser) { throw 'vCenterUser was specified but vCenterPassword was not supplied.' }

if ($WindowsScriptPath) { W ("Configured Windows payload path: {0}" -f $WindowsScriptPath) }
if ($LinuxScriptPath) { W ("Configured Linux payload path: {0}" -f $LinuxScriptPath) }
if ($ScriptPath) { W ("Configured shared payload path: {0}" -f $ScriptPath) }
if ($WindowsExecutionCommandTemplate) { W ("Configured Windows custom execution template: {0}" -f $WindowsExecutionCommandTemplate) }
if ($LinuxExecutionCommandTemplate) { W ("Configured Linux custom execution template: {0}" -f $LinuxExecutionCommandTemplate) }
if ($ExecutionCommandTemplate) { W ("Configured shared custom execution template: {0}" -f $ExecutionCommandTemplate) }

W ("Testing workstation connectivity to vCenter {0} on TCP 443 ..." -f $vCenterServer)
$tcp = Test-NetConnection -ComputerName $vCenterServer -Port 443 -WarningAction SilentlyContinue
if (-not $tcp.TcpTestSucceeded) { throw "Unable to reach $vCenterServer on TCP 443 from the workstation." }

$vi = $null
try {
  if ($vcCred) { $vi = Connect-VIServer -Server $vCenterServer -Credential $vcCred -ErrorAction Stop }
  else { $vi = Connect-VIServer -Server $vCenterServer -ErrorAction Stop }

  W ("Connected to {0} as {1}" -f $vCenterServer,$vi.User) 'PASS'

  foreach ($row in $targets) {
    $vmName = Resolve-TargetVmName -Row $row -PrimaryColumn $VMNameColumn -FallbackColumn $ComputerNameColumn
    if ([string]::IsNullOrWhiteSpace($vmName)) { continue }

    Write-Host ''
    Write-Host '============================================================' -ForegroundColor DarkCyan
    Write-Host ("GuestOps payload run : {0}" -f $vmName) -ForegroundColor Cyan
    Write-Host '============================================================' -ForegroundColor DarkCyan

    $requestedOs = Resolve-RequestedTargetOs -Row $row -ColumnName $TargetOsColumn -DefaultValue $TargetOs
    $fullOutput = New-Object System.Text.StringBuilder
    $vmResult = [ordered]@{
      VMName = $vmName
      Status = 'FAIL'
      ExitCode = $null
      GuestUser = $null
      CredentialSource = $null
      CredentialSourceDetail = $null
      RequestedTargetOs = $requestedOs
      DetectedTargetOs = 'Auto'
      OsDetectionSource = $null
      GuestFullName = $null
      PayloadName = $null
      PayloadSourcePath = $null
      PayloadMode = $null
      PayloadDefinitionSource = $null
      ExecutionCommandTemplate = $null
      LogPath = $null
      RawOutputPath = $null
      Output = $null
    }

    $safeVmFileName = Get-SafeFileName $vmName
    $logPath = Join-Path $OutDir ("{0}.log" -f $safeVmFileName)
    $rawOutputPath = Join-Path $OutDir ("{0}.output.txt" -f $safeVmFileName)
    $vmResult.LogPath = $logPath
    $vmResult.RawOutputPath = $rawOutputPath

    try {
      $vm = Get-VM -Name $vmName -ErrorAction Stop
      $null = Get-VMGuest -VM $vm -ErrorAction Stop
      $osInfo = Resolve-DetectedTargetOs -VM $vm -RequestedTargetOs $requestedOs
      $vmResult.DetectedTargetOs = $osInfo.DetectedTargetOs
      $vmResult.OsDetectionSource = $osInfo.OsDetectionSource
      $vmResult.GuestFullName = $osInfo.GuestFullName

      $payloadDefinition = Resolve-PayloadDefinition `
        -DetectedTargetOs $osInfo.DetectedTargetOs `
        -SharedScriptPath $ScriptPath `
        -SharedScriptLanguage $ScriptLanguage `
        -SharedExecutionCommandTemplate $ExecutionCommandTemplate `
        -WindowsScriptPath $WindowsScriptPath `
        -WindowsScriptLanguage $WindowsScriptLanguage `
        -WindowsExecutionCommandTemplate $WindowsExecutionCommandTemplate `
        -LinuxScriptPath $LinuxScriptPath `
        -LinuxScriptLanguage $LinuxScriptLanguage `
        -LinuxExecutionCommandTemplate $LinuxExecutionCommandTemplate `
        -PayloadCache $payloadCache

      Assert-ScriptModeCompatible -DetectedTargetOs $osInfo.DetectedTargetOs -ScriptMode $payloadDefinition.ScriptInfo.ScriptMode

      $vmResult.PayloadName = $payloadDefinition.ScriptName
      $vmResult.PayloadSourcePath = $payloadDefinition.ScriptPath
      $vmResult.PayloadMode = $payloadDefinition.ScriptInfo.ScriptMode
      $vmResult.PayloadDefinitionSource = $payloadDefinition.PayloadSource
      $vmResult.ExecutionCommandTemplate = $payloadDefinition.ExecutionCommandTemplate

      $credInfo = Resolve-RowCredential `
        -Row $row `
        -DetectedTargetOs $osInfo.DetectedTargetOs `
        -GuestUserColumn $GuestUserColumn `
        -GuestPasswordColumn $GuestPasswordColumn `
        -AltCredFileColumn $AltCredFileColumn `
        -GlobalCredential $globalCredential `
        -WindowsCredential $windowsCredential `
        -LinuxCredential $linuxCredential `
        -CredentialFilePath $CredentialFile `
        -WindowsCredentialFilePath $WindowsCredentialFile `
        -LinuxCredentialFilePath $LinuxCredentialFile

      $vmResult.GuestUser = $credInfo.GuestUser
      $vmResult.CredentialSource = $credInfo.CredentialSource
      $vmResult.CredentialSourceDetail = $credInfo.CredentialSourceDetail

      W ("Detected guest OS: {0} ({1}) via {2}" -f $osInfo.DetectedTargetOs,$osInfo.GuestFullName,$osInfo.OsDetectionSource)
      W ("Using payload: {0} ({1}, source={2})" -f $payloadDefinition.ScriptPath,$payloadDefinition.ScriptInfo.ScriptMode,$payloadDefinition.PayloadSource)
      W ("Using credential source: {0}" -f $credInfo.CredentialSource)

      if ($vm.PowerState -ne 'PoweredOn') { throw "VM '$vmName' is not powered on." }
      if ($vm.ExtensionData.Guest.ToolsRunningStatus -ne 'guestToolsRunning') { throw "VMware Tools is not running on '$vmName'." }

      $bootstrapScript = $null
      $outerScriptType = $null

      if ($osInfo.DetectedTargetOs -eq 'Windows') {
        $outerScriptType = 'PowerShell'
        $bootstrapScript = New-WindowsGuestBootstrapScript `
          -PayloadBytes $payloadDefinition.PayloadBytes `
          -PayloadExtension $payloadDefinition.ScriptInfo.Extension `
          -ScriptMode $payloadDefinition.ScriptInfo.ScriptMode `
          -ExecutionCommandTemplate $payloadDefinition.ExecutionCommandTemplate
      } else {
        $outerScriptType = 'Bash'
        $bootstrapScript = New-LinuxGuestBootstrapScript `
          -PayloadBytes $payloadDefinition.PayloadBytes `
          -PayloadExtension $payloadDefinition.ScriptInfo.Extension `
          -ScriptMode $payloadDefinition.ScriptInfo.ScriptMode `
          -ExecutionCommandTemplate $payloadDefinition.ExecutionCommandTemplate
      }

      W ("Invoking payload in guest via Invoke-VMScript using outer script type {0} ..." -f $outerScriptType)
      $run = Invoke-VMScript -VM $vm -GuestCredential $credInfo.Credential -ScriptType $outerScriptType -ScriptText $bootstrapScript -ErrorAction Stop

      $rawOutput = ($run.ScriptOutput | Out-String).TrimEnd()
      $vmResult.ExitCode = $run.ExitCode
      $vmResult.Status = Get-OverallStatus -ExitCode $run.ExitCode

      [void]$fullOutput.AppendLine(("Payload source path : {0}" -f $payloadDefinition.ScriptPath))
      [void]$fullOutput.AppendLine(("Payload name        : {0}" -f $payloadDefinition.ScriptName))
      [void]$fullOutput.AppendLine(("Payload mode        : {0}" -f $payloadDefinition.ScriptInfo.ScriptMode))
      [void]$fullOutput.AppendLine(("Payload source      : {0}" -f $payloadDefinition.PayloadSource))
      [void]$fullOutput.AppendLine(("Requested target OS : {0}" -f $requestedOs))
      [void]$fullOutput.AppendLine(("Detected target OS  : {0}" -f $osInfo.DetectedTargetOs))
      [void]$fullOutput.AppendLine(("Guest full name     : {0}" -f $osInfo.GuestFullName))
      [void]$fullOutput.AppendLine(("Guest user          : {0}" -f $credInfo.GuestUser))
      [void]$fullOutput.AppendLine(("Credential source   : {0}" -f $credInfo.CredentialSource))
      if ($credInfo.CredentialSourceDetail) {
        [void]$fullOutput.AppendLine(("Credential detail   : {0}" -f $credInfo.CredentialSourceDetail))
      }
      [void]$fullOutput.AppendLine(("Invoke-VMScript exit: {0}" -f $run.ExitCode))
      [void]$fullOutput.AppendLine('')
      [void]$fullOutput.AppendLine('--- Guest Output ---')
      if ($rawOutput) { [void]$fullOutput.AppendLine($rawOutput) }
      [void]$fullOutput.AppendLine('--------------------')

      Write-Host ''
      Write-Host '--- Guest Output ---' -ForegroundColor DarkGray
      if ($rawOutput) { Write-Host $rawOutput }
      Write-Host '--------------------' -ForegroundColor DarkGray

      $vmResult.Output = $fullOutput.ToString().Trim()
      $vmResult.Output | Out-File -LiteralPath $logPath -Encoding utf8
      $rawOutput | Out-File -LiteralPath $rawOutputPath -Encoding utf8

      W ("Completed {0} with status {1}" -f $vmName,$vmResult.Status) $vmResult.Status
    }
    catch {
      $vmResult.Status = 'FAIL'
      $message = $_.Exception.Message
      $vmResult.Output = "[FAIL] $message"
      if (-not $vmResult.ExitCode) { $vmResult.ExitCode = -1 }
      $vmResult.Output | Out-File -LiteralPath $logPath -Encoding utf8
      '' | Out-File -LiteralPath $rawOutputPath -Encoding utf8
      W ("Failed: {0}" -f $message) 'FAIL'
      if (-not $ContinueOnError) {
        # Continue processing all rows by design.
      }
    }

    $summary.Add([pscustomobject]$vmResult)
  }
}
finally {
  if ($vi) {
    Disconnect-VIServer -Server $vi -Confirm:$false | Out-Null
    W ("Disconnected from {0}" -f $vCenterServer)
  }
}

$stamp = Get-Date -Format 'yyyyMMdd-HHmmss'
$jsonPath = Join-Path $OutDir "GuestOpsScriptFleetSummary-$stamp.json"
$csvPath = Join-Path $OutDir "GuestOpsScriptFleetSummary-$stamp.csv"
$summaryExport = $summary | Select-Object * -ExcludeProperty Output
$summaryExport | ConvertTo-Json -Depth 6 | Out-File -LiteralPath $jsonPath -Encoding utf8
$summaryExport | Export-Csv -NoTypeInformation -LiteralPath $csvPath -Encoding utf8

Write-Host ''
Write-Host '============================================================' -ForegroundColor DarkCyan
Write-Host 'Fleet Payload Execution Summary' -ForegroundColor Cyan
Write-Host '============================================================' -ForegroundColor DarkCyan
foreach ($item in $summary) {
  $detail = "Detected=$($item.DetectedTargetOs); User=$($item.GuestUser); Mode=$($item.PayloadMode); Exit=$($item.ExitCode)"
  Write-Host ("{0,-20} {1,-5} {2}" -f $item.VMName,$item.Status,$detail) -ForegroundColor @{ PASS='Green'; WARN='Yellow'; FAIL='Red'; INFO='Cyan' }[$item.Status]
}
Write-Host ''
Write-Host ("Per-VM logs written to: {0}" -f (Resolve-Path -LiteralPath $OutDir)) -ForegroundColor Green
Write-Host ("Per-VM raw output files use: <VMName>.output.txt") -ForegroundColor Green
Write-Host ("Summary CSV : {0}" -f $csvPath) -ForegroundColor Green
Write-Host ("Summary JSON: {0}" -f $jsonPath) -ForegroundColor Green
