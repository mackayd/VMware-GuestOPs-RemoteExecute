#requires -Version 7.0
[CmdletBinding()]
param(
  [Alias('h')][switch]$Help,
  [switch]$Examples,

  [Parameter(Mandatory=$true)]
  [string]$SummaryJsonPath,

  [string]$HtmlOutPath,

  [string]$Title = 'GuestOps Script Fleet Report',
  [string]$Subtitle = 'Interactive execution summary for fleet guest-operations payload runs.',
  [string]$ReportBadge = 'GuestOps Fleet',
  [switch]$OpenReport
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

function Write-Info {
  param(
    [string]$Message,
    [ValidateSet('INFO','PASS','WARN','FAIL')][string]$Level = 'INFO'
  )
  $color = @{ INFO='Cyan'; PASS='Green'; WARN='Yellow'; FAIL='Red' }[$Level]
  Write-Host "[$Level] $Message" -ForegroundColor $color
}

function Show-HelpText {
@"
Generate-GuestOpsScriptFleetReport.ps1

Purpose
  Creates a dynamic HTML report for the GuestOps fleet runner.

Inputs
  -SummaryJsonPath   Path to GuestOpsScriptFleetSummary-*.json
  -HtmlOutPath       Optional output path for the generated HTML report
  -Title             Optional report title
  -Subtitle          Optional report subtitle
  -ReportBadge       Optional badge text shown in the hero area
  -OpenReport        Opens the HTML after generation

Notes
  - Reads the summary JSON plus each VM's RawOutputPath when present.
  - Uses the embedded Output field from older summary JSON files as a first-class source.
  - Extracts the guest output block from embedded execution detail when sidecar output files are absent.
  - Produces a clickable interactive report with VM filtering and detailed output view.
"@ | Write-Host
}

function Show-ExamplesText {
@"
Examples

  .\Generate-GuestOpsScriptFleetReport.ps1 `
    -SummaryJsonPath '.\GuestOpsScriptFleetSummary-20260327-101500.json'

  .\Generate-GuestOpsScriptFleetReport.ps1 `
    -SummaryJsonPath '.\GuestOpsScriptFleetSummary-20260327-101500.json' `
    -HtmlOutPath '.\GuestOpsScriptFleetReport-20260327-101500.html' `
    -Title 'Remote Script Execution Fleet Report' `
    -Subtitle 'Interactive summary for Windows and Linux GuestOps payload execution.'

"@ | Write-Host
}

function Convert-ToArray {
  param($InputObject)
  if ($null -eq $InputObject) { return @() }
  if ($InputObject -is [System.Collections.IEnumerable] -and $InputObject -isnot [string]) {
    return @($InputObject)
  }
  return @($InputObject)
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

function Resolve-ChildPathRelativeTo {
  param(
    [string]$BasePath,
    [string]$ChildPath
  )
  if ([string]::IsNullOrWhiteSpace($ChildPath)) { return $null }
  if ([System.IO.Path]::IsPathRooted($ChildPath)) { return $ChildPath }
  $baseDirectory = Split-Path -Parent $BasePath
  return (Join-Path $baseDirectory $ChildPath)
}

function Get-AbsolutePathIfExists {
  param([string]$CandidatePath)
  if ([string]::IsNullOrWhiteSpace($CandidatePath)) { return $null }
  if (Test-Path -LiteralPath $CandidatePath) {
    return (Resolve-Path -LiteralPath $CandidatePath).Path
  }
  return $null
}

function Convert-PathToHref {
  param(
    [string]$TargetPath,
    [string]$HtmlOutPath
  )
  if ([string]::IsNullOrWhiteSpace($TargetPath)) { return $null }
  $targetResolved = Get-AbsolutePathIfExists -CandidatePath $TargetPath
  if (-not $targetResolved) { return $null }

  $htmlParent = Split-Path -Parent $HtmlOutPath
  try {
    $relative = [System.IO.Path]::GetRelativePath($htmlParent, $targetResolved)
    if (-not [string]::IsNullOrWhiteSpace($relative)) {
      return ($relative -replace '\\','/')
    }
  }
  catch {
    # Fall back to file URI
  }

  return ([System.Uri]$targetResolved).AbsoluteUri
}

function Read-TextFileIfPresent {
  param([string]$Path)
  if ([string]::IsNullOrWhiteSpace($Path)) { return $null }
  if (-not (Test-Path -LiteralPath $Path)) { return $null }
  return Get-Content -LiteralPath $Path -Raw -Encoding utf8
}

function Get-StatusCounts {
  param([object[]]$Items)
  $result = [ordered]@{
    PASS = 0
    WARN = 0
    FAIL = 0
    INFO = 0
    OTHER = 0
  }
  foreach ($item in $Items) {
    $status = [string]$item.Status
    if ($result.Contains($status)) {
      $result[$status]++
    } else {
      $result.OTHER++
    }
  }
  return $result
}

function Get-Percent {
  param([int]$Part,[int]$Whole)
  if ($Whole -le 0) { return 0 }
  return [math]::Round(($Part / $Whole) * 100, 2)
}

function Get-FirstInterestingOutputLine {
  param([string]$Text)
  if ([string]::IsNullOrWhiteSpace($Text)) { return '' }
  $lines = $Text -split "`r?`n"
  foreach ($line in $lines) {
    $trimmed = $line.Trim()
    if (-not $trimmed) { continue }
    if ($trimmed -match '^(GOR_RESULT OUTPUT_BEGIN|GOR_RESULT OUTPUT_END|---+|===+)$') { continue }
    return $trimmed
  }
  return ''
}

function Extract-RawOutputFromDetailText {
  param([string]$Text)
  if ([string]::IsNullOrWhiteSpace($Text)) { return '' }

  $lines = $Text -split "`r?`n"
  if (-not $lines.Count) { return '' }

  $startIndex = -1
  $endIndex = -1

  for ($i = 0; $i -lt $lines.Count; $i++) {
    if ($lines[$i].Trim() -eq '--- Guest Output ---') {
      $startIndex = $i + 1
      break
    }
  }

  if ($startIndex -ge 0) {
    for ($j = $startIndex; $j -lt $lines.Count; $j++) {
      if ($lines[$j].Trim() -eq '--------------------') {
        $endIndex = $j - 1
        break
      }
    }
    if ($endIndex -lt $startIndex) { $endIndex = $lines.Count - 1 }
    return (($lines[$startIndex..$endIndex] -join "`n").Trim())
  }

  $startIndex = -1
  $endIndex = -1
  for ($i = 0; $i -lt $lines.Count; $i++) {
    if ($lines[$i].Trim() -eq 'GOR_RESULT OUTPUT_BEGIN') {
      $startIndex = $i + 1
      break
    }
  }

  if ($startIndex -ge 0) {
    for ($j = $startIndex; $j -lt $lines.Count; $j++) {
      if ($lines[$j].Trim() -eq 'GOR_RESULT OUTPUT_END') {
        $endIndex = $j - 1
        break
      }
    }
    if ($endIndex -lt $startIndex) { $endIndex = $lines.Count - 1 }
    return (($lines[$startIndex..$endIndex] -join "`n").Trim())
  }

  return $Text.Trim()
}

if ($Help) { Show-HelpText; return }
if ($Examples) { Show-ExamplesText; return }

if (-not (Test-Path -LiteralPath $SummaryJsonPath)) {
  throw "Summary JSON not found: $SummaryJsonPath"
}

$summaryResolved = (Resolve-Path -LiteralPath $SummaryJsonPath).Path
$summaryDirectory = Split-Path -Parent $summaryResolved

if (-not $HtmlOutPath) {
  $summaryBase = [System.IO.Path]::GetFileNameWithoutExtension($summaryResolved)
  $HtmlOutPath = Join-Path $summaryDirectory ($summaryBase + '.html')
}

$summaryRaw = Get-Content -LiteralPath $summaryResolved -Raw -Encoding utf8
$summaryItems = Convert-ToArray (ConvertFrom-Json -InputObject $summaryRaw)

if (-not $summaryItems.Count) {
  throw "The summary JSON did not contain any VM result objects."
}

$vmReportRows = New-Object System.Collections.Generic.List[object]

foreach ($item in $summaryItems) {
  $vmName = [string]$item.VMName
  $status = if ([string]::IsNullOrWhiteSpace([string]$item.Status)) { 'OTHER' } else { [string]$item.Status }

  $rawOutputPath = $null
  if ($item.PSObject.Properties.Name -contains 'RawOutputPath') {
    $rawOutputPath = Resolve-ChildPathRelativeTo -BasePath $summaryResolved -ChildPath ([string]$item.RawOutputPath)
  }

  $logPath = $null
  if ($item.PSObject.Properties.Name -contains 'LogPath') {
    $logPath = Resolve-ChildPathRelativeTo -BasePath $summaryResolved -ChildPath ([string]$item.LogPath)
  }

  if (-not $rawOutputPath -and $logPath) {
    $logBase = [System.IO.Path]::GetFileNameWithoutExtension($logPath)
    $logDir = Split-Path -Parent $logPath
    $derivedRawOutput = Join-Path $logDir ($logBase + '.output.txt')
    if (Test-Path -LiteralPath $derivedRawOutput) {
      $rawOutputPath = $derivedRawOutput
    }
  }

  $embeddedDetailOutput = if ($item.PSObject.Properties.Name -contains 'Output') { [string]$item.Output } else { '' }

  $rawOutputText = Read-TextFileIfPresent -Path $rawOutputPath

  if ([string]::IsNullOrWhiteSpace($rawOutputText) -and -not [string]::IsNullOrWhiteSpace($embeddedDetailOutput)) {
    $rawOutputText = Extract-RawOutputFromDetailText -Text $embeddedDetailOutput
  }

  $payloadName = if ($item.PSObject.Properties.Name -contains 'PayloadName') { [string]$item.PayloadName } else { $null }
  $payloadMode = if ($item.PSObject.Properties.Name -contains 'PayloadMode') { [string]$item.PayloadMode } else { $null }
  $payloadSource = if ($item.PSObject.Properties.Name -contains 'PayloadDefinitionSource') { [string]$item.PayloadDefinitionSource } else { $null }

  $vmReportRows.Add([ordered]@{
    vmId = (Get-SafeFileName $vmName)
    vmName = $vmName
    status = $status
    exitCode = if ($item.PSObject.Properties.Name -contains 'ExitCode') { [int]$item.ExitCode } else { $null }
    detectedTargetOs = if ($item.PSObject.Properties.Name -contains 'DetectedTargetOs') { [string]$item.DetectedTargetOs } else { '' }
    requestedTargetOs = if ($item.PSObject.Properties.Name -contains 'RequestedTargetOs') { [string]$item.RequestedTargetOs } else { '' }
    guestFullName = if ($item.PSObject.Properties.Name -contains 'GuestFullName') { [string]$item.GuestFullName } else { '' }
    guestUser = if ($item.PSObject.Properties.Name -contains 'GuestUser') { [string]$item.GuestUser } else { '' }
    credentialSource = if ($item.PSObject.Properties.Name -contains 'CredentialSource') { [string]$item.CredentialSource } else { '' }
    credentialSourceDetail = if ($item.PSObject.Properties.Name -contains 'CredentialSourceDetail') { [string]$item.CredentialSourceDetail } else { '' }
    osDetectionSource = if ($item.PSObject.Properties.Name -contains 'OsDetectionSource') { [string]$item.OsDetectionSource } else { '' }
    payloadName = $payloadName
    payloadSourcePath = if ($item.PSObject.Properties.Name -contains 'PayloadSourcePath') { [string]$item.PayloadSourcePath } else { '' }
    payloadMode = $payloadMode
    payloadSource = $payloadSource
    executionCommandTemplate = if ($item.PSObject.Properties.Name -contains 'ExecutionCommandTemplate') { [string]$item.ExecutionCommandTemplate } else { '' }
    logPath = $logPath
    rawOutputPath = $rawOutputPath
    logHref = Convert-PathToHref -TargetPath $logPath -HtmlOutPath $HtmlOutPath
    rawOutputHref = Convert-PathToHref -TargetPath $rawOutputPath -HtmlOutPath $HtmlOutPath
    rawOutput = if ($null -ne $rawOutputText) { [string]$rawOutputText } else { '' }
    rawOutputPreview = Get-FirstInterestingOutputLine -Text $rawOutputText
    detailOutput = if ($null -ne $embeddedDetailOutput) { [string]$embeddedDetailOutput } else { '' }
    detailOutputPreview = Get-FirstInterestingOutputLine -Text $embeddedDetailOutput
  })
}

$vmReportRowsArray = @($vmReportRows | ForEach-Object { $_ })
$totalVms = $vmReportRowsArray.Count
$statusCounts = Get-StatusCounts -Items $vmReportRowsArray
$passCount = [int]$statusCounts.PASS
$warnCount = [int]$statusCounts.WARN
$failCount = [int]$statusCounts.FAIL
$infoCount = [int]$statusCounts.INFO
$otherCount = [int]$statusCounts.OTHER

$reportData = [ordered]@{
  title = $Title
  subtitle = $Subtitle
  badge = $ReportBadge
  generatedAt = (Get-Date).ToString('yyyy-MM-dd HH:mm:ss')
  summaryJsonPath = [System.IO.Path]::GetFileName($summaryResolved)
  totalVms = $totalVms
  counts = [ordered]@{
    PASS = $passCount
    WARN = $warnCount
    FAIL = $failCount
    INFO = $infoCount
    OTHER = $otherCount
  }
  percentages = [ordered]@{
    PASS = Get-Percent -Part $passCount -Whole $totalVms
    WARN = Get-Percent -Part $warnCount -Whole $totalVms
    FAIL = Get-Percent -Part $failCount -Whole $totalVms
    INFO = Get-Percent -Part $infoCount -Whole $totalVms
    OTHER = Get-Percent -Part $otherCount -Whole $totalVms
  }
  vms = $vmReportRowsArray
}

$jsonForBrowser = $reportData | ConvertTo-Json -Depth 8 -Compress
$jsonForBrowser = $jsonForBrowser.Replace('</script>','<\/script>').Replace([string][char]0x2028,' ').Replace([string][char]0x2029,' ')

$htmlTemplate = @'
<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="utf-8" />
<meta name="viewport" content="width=device-width, initial-scale=1" />
<title>__TITLE__</title>
<style>
:root{
  --bg:#071120;
  --bg2:#0b1730;
  --panel:#101c33;
  --panel2:#162643;
  --text:#e8eef8;
  --muted:#9fb0c8;
  --border:#243650;
  --accent:#69b1ff;
  --pass:#2e7d32;
  --warn:#f9a825;
  --fail:#c62828;
  --info:#1565c0;
  --other:#5b6475;
  --shadow:0 14px 40px rgba(0,0,0,.33);
}
*{box-sizing:border-box}
body{
  margin:0;
  font-family:"Segoe UI",Inter,Arial,sans-serif;
  background:radial-gradient(circle at top left,#0d2848 0%,#071120 40%,#050b16 100%);
  color:var(--text)
}
.container{width:min(1680px,calc(100% - 30px));margin:16px auto 28px}
.hero{
  background:linear-gradient(135deg,#15315c,#101b31 55%,#113b34);
  border:1px solid var(--border);
  border-radius:22px;
  padding:24px;
  box-shadow:var(--shadow);
  margin-bottom:16px
}
.hero-top{display:flex;justify-content:space-between;gap:16px;align-items:flex-start;flex-wrap:wrap}
.hero h1{margin:0 0 8px;font-size:32px;line-height:1.08}
.hero p{margin:0;color:#d2def0}
.hero-meta{display:flex;gap:14px;flex-wrap:wrap;margin-top:12px;color:#b4c4dc;font-size:13px}
.badge{display:inline-flex;align-items:center;justify-content:center;min-width:84px;padding:8px 12px;border-radius:999px;font-weight:800;font-size:12px;letter-spacing:.04em;text-transform:uppercase}
.badge.PASS{background:#1b5e20;color:#d6ffd6}
.badge.WARN{background:#7a5a00;color:#fff0b5}
.badge.FAIL{background:#7f1d1d;color:#ffd6d6}
.badge.INFO{background:#0d47a1;color:#dce9ff}
.badge.OTHER{background:#3c4557;color:#e4e9f5}
.badge.large{padding:10px 16px;font-size:13px}
.hero-grid{display:grid;grid-template-columns:2fr 1fr;gap:20px;margin-top:22px}
.stack-card,.score-card,.sidebar-card,.panel-card,.panel-header,.detail-card,.stat-card,.vm-overview-card{
  background:rgba(255,255,255,.04);
  border:1px solid rgba(255,255,255,.08);
  border-radius:20px;
  padding:20px
}
.stack-bar{display:flex;height:54px;border-radius:18px;overflow:hidden;margin-top:18px;background:#0f1728;border:1px solid rgba(255,255,255,.1)}
.stack-segment{display:flex;align-items:center;justify-content:center;font-weight:800;font-size:13px;white-space:nowrap;color:#fff}
.stack-segment span{padding:0 10px}
.stack-segment.PASS{background:var(--pass)}
.stack-segment.WARN{background:var(--warn);color:#111}
.stack-segment.FAIL{background:var(--fail)}
.stack-segment.INFO{background:var(--info)}
.stack-segment.OTHER{background:var(--other)}
.score-card{display:flex;align-items:center;justify-content:center}
.score-ring{
  width:180px;height:180px;border-radius:50%;
  background:conic-gradient(var(--pass) 0 1turn);
  display:flex;align-items:center;justify-content:center
}
.score-ring-content{
  width:130px;height:130px;border-radius:50%;background:#0f1728;
  display:flex;flex-direction:column;align-items:center;justify-content:center
}
.score-ring-value{font-size:44px;font-weight:900}
.score-ring-label{color:var(--muted);font-size:13px;text-transform:uppercase;letter-spacing:.06em}
.stats{display:grid;grid-template-columns:repeat(6,1fr);gap:14px;margin:16px 0}
.stat-card{padding:18px}
.stat-label{color:var(--muted);font-size:13px}
.stat-value{font-size:34px;font-weight:900;margin-top:8px}
.main-grid{display:grid;grid-template-columns:340px 1fr;gap:18px;align-items:start}
.sidebar-card{position:sticky;top:16px}
.section-title{margin:0 0 6px;font-size:18px}
.section-subtitle{margin:0 0 14px;color:var(--muted);font-size:13px}
.filter-bar,.view-mode-toggle{display:flex;gap:10px;flex-wrap:wrap;margin:14px 0 18px}
button{cursor:pointer;border:none}
.filter-btn,.view-btn,.view-vm-btn,.toggle-btn,.link-btn{
  padding:10px 14px;border-radius:12px;background:#18253b;color:var(--text);
  border:1px solid var(--border);font-weight:700
}
.filter-btn.active,.view-btn.active{background:#1d4ed8;border-color:#3b82f6}
.vm-nav{display:flex;flex-direction:column;gap:10px;max-height:calc(100vh - 300px);overflow:auto;padding-right:4px}
.vm-nav-item{
  width:100%;text-align:left;padding:12px 14px;border-radius:14px;background:#141f33;
  border:1px solid var(--border);color:var(--text);display:flex;flex-direction:column;gap:4px
}
.vm-nav-item:hover,.vm-overview-card:hover{box-shadow:0 8px 24px rgba(0,0,0,.22);transform:translateY(-1px)}
.vm-nav-item.active{outline:2px solid #3b82f6;background:#18253b}
.vm-nav-item.PASS{border-left:6px solid var(--pass)}
.vm-nav-item.WARN{border-left:6px solid var(--warn)}
.vm-nav-item.FAIL{border-left:6px solid var(--fail)}
.vm-nav-item.INFO{border-left:6px solid var(--info)}
.vm-nav-item.OTHER{border-left:6px solid var(--other)}
.vm-nav-title{font-weight:800}
.vm-nav-meta,.vm-nav-sub{font-size:12px;color:var(--muted)}
.vm-panel{display:none;flex-direction:column;gap:18px}
.vm-panel.active{display:flex}
.panel-header{display:flex;justify-content:space-between;align-items:flex-start;gap:14px}
.eyebrow{text-transform:uppercase;letter-spacing:.08em;font-size:11px;color:var(--muted);font-weight:700}
.panel-header h2{margin:4px 0 6px;font-size:28px}
.panel-subtitle{margin:0;color:var(--muted)}
.detail-grid{display:grid;grid-template-columns:repeat(4,minmax(0,1fr));gap:14px}
.detail-label{font-size:12px;text-transform:uppercase;letter-spacing:.06em;color:var(--muted);margin-bottom:8px}
.detail-value{font-size:15px;font-weight:700;word-break:break-word}
.panel-card h3{margin:0 0 14px;font-size:18px}
.table-wrap{overflow:auto}
table{width:100%;border-collapse:collapse}
th,td{padding:12px;border-bottom:1px solid var(--border);text-align:left;vertical-align:top}
th{color:#c9d8ee;font-size:13px;background:#101a2d;position:sticky;top:0}
td{font-size:14px;color:#e9f1fd}
tr.PASS td{background:rgba(46,125,50,.14)}
tr.WARN td{background:rgba(249,168,37,.16);color:#fff2c2}
tr.FAIL td{background:rgba(198,40,40,.16)}
tr.INFO td{background:rgba(21,101,192,.14)}
tr.OTHER td{background:rgba(91,100,117,.18)}
.overview-panel{display:none;flex-direction:column;gap:16px}
.overview-panel.active{display:flex}
.overview-grid{display:grid;grid-template-columns:repeat(auto-fill,minmax(300px,1fr));gap:16px}
.vm-overview-card{box-shadow:0 8px 24px rgba(0,0,0,.22)}
.vm-overview-card.PASS{border-top:5px solid var(--pass)}
.vm-overview-card.WARN{border-top:5px solid var(--warn)}
.vm-overview-card.FAIL{border-top:5px solid var(--fail)}
.vm-overview-card.INFO{border-top:5px solid var(--info)}
.vm-overview-card.OTHER{border-top:5px solid var(--other)}
.vm-overview-top{display:flex;justify-content:space-between;gap:16px;align-items:flex-start}
.vm-overview-top h3{margin:4px 0 6px;font-size:22px}
.vm-overview-top p{margin:0;color:var(--muted);font-size:13px}
.overview-actions{display:flex;flex-direction:column;align-items:flex-end;gap:10px}
.mini-check-list{
  list-style:none;margin:16px 0 0;padding:0;display:flex;flex-direction:column;gap:10px
}
.mini-check-list li{
  display:flex;justify-content:space-between;gap:12px;align-items:center;
  padding:10px 12px;background:#101a2d;border:1px solid var(--border);border-radius:12px
}
.mini-check-name{font-size:13px;font-weight:600}
.action-row{display:flex;gap:10px;flex-wrap:wrap}
.link-btn{text-decoration:none;display:inline-flex;align-items:center;justify-content:center}
.hidden{display:none}
.empty-state{
  background:rgba(255,255,255,.04);border:1px dashed var(--border);
  border-radius:22px;padding:50px;text-align:center;color:var(--muted)
}
.output-wrap{
  margin-top:14px;background:#0b1220;color:#dce9ff;padding:0;border-radius:18px;
  overflow:hidden;border:1px solid var(--border)
}
.output-toolbar{
  display:flex;justify-content:space-between;align-items:center;gap:10px;
  padding:12px 14px;background:#101a2d;border-bottom:1px solid var(--border)
}
.output-title{font-size:13px;color:var(--muted);font-weight:800;text-transform:uppercase;letter-spacing:.06em}
.raw-output{
  margin:0;padding:14px 16px;overflow:auto;max-height:520px;
  font-size:12px;line-height:1.5;white-space:pre-wrap;font-family:Consolas, "Courier New", monospace
}
.code-line{display:block;white-space:pre-wrap;word-break:break-word}
.line-pass{color:#80ed99}
.line-warn{color:#ffd166}
.line-fail{color:#ff8b8b}
.line-info{color:#8cc8ff}
.line-banner{color:#c7d2fe;font-weight:700}
.line-muted{color:#93a4c2}
.footer-note{margin-top:18px;font-size:12px;color:var(--muted)}
@media (max-width:1280px){.hero-grid,.main-grid,.stats,.detail-grid{grid-template-columns:1fr}.sidebar-card{position:static}}
</style>
</head>
<body>
  <div class="container">
    <section class="hero">
      <div class="hero-top">
        <div>
          <h1 id="heroTitle"></h1>
          <p id="heroSubtitle"></p>
          <div class="hero-meta">
            <div id="metaGenerated"></div>
            <div id="metaSummary"></div>
            <div id="metaTotal"></div>
          </div>
        </div>
        <div id="heroBadge" class="badge large INFO"></div>
      </div>
      <div class="hero-grid">
        <div class="stack-card">
          <div class="eyebrow">Summary Graphic</div>
          <h2 style="margin:6px 0 0">VM status distribution</h2>
          <p class="section-subtitle" style="color:rgba(255,255,255,.74);margin-top:8px">Click a VM on the left to inspect metadata, payload details, and syntax-highlighted raw output.</p>
          <div class="stack-bar" id="stackBar"></div>
        </div>
        <div class="score-card">
          <div class="score-ring" id="scoreRing">
            <div class="score-ring-content">
              <div class="score-ring-value" id="scoreRingValue">0</div>
              <div class="score-ring-label">Passing VMs</div>
            </div>
          </div>
        </div>
      </div>
    </section>

    <section class="stats" id="statsRow"></section>

    <section class="main-grid">
      <aside class="sidebar">
        <div class="sidebar-card">
          <h2 class="section-title">Virtual Machines</h2>
          <p class="section-subtitle">Select a VM to reveal its execution result and raw output.</p>
          <div class="filter-bar" id="filterBar"></div>
          <div class="vm-nav" id="vmNav"></div>
        </div>
      </aside>
      <main class="content">
        <div class="view-mode-toggle">
          <button type="button" class="view-btn active" data-view="ALLVMS">Show All VMs</button>
          <button type="button" class="view-btn" data-view="DETAIL">Single VM Detail</button>
        </div>

        <section class="overview-panel active" id="all-vms-panel">
          <div class="panel-card">
            <h3>All VMs tested</h3>
            <p class="section-subtitle">This overview shows every VM currently visible under the selected status filter. Use Open details to jump into the full per-VM report.</p>
          </div>
          <div class="overview-grid" id="overviewGrid"></div>
        </section>

        <div id="vmPanelsHost"></div>
      </main>
    </section>

    <div class="footer-note">This report is generated from the fleet summary JSON plus each VM's raw output file when available.</div>
  </div>

<script>
const reportData = __REPORT_DATA__;

(function(){
  const statuses = ['PASS','WARN','FAIL','INFO','OTHER'];
  let currentFilter = 'ALL';

  function escapeHtml(value){
    return String(value ?? '')
      .replace(/&/g,'&amp;')
      .replace(/</g,'&lt;')
      .replace(/>/g,'&gt;')
      .replace(/"/g,'&quot;')
      .replace(/'/g,'&#39;');
  }

  function statusBadge(status, large){
    const safe = statuses.includes(status) ? status : 'OTHER';
    return `<span class="badge ${safe}${large ? ' large' : ''}">${safe}</span>`;
  }

  function renderHero(){
    document.getElementById('heroTitle').textContent = reportData.title;
    document.getElementById('heroSubtitle').textContent = reportData.subtitle;
    document.getElementById('metaGenerated').textContent = `Generated: ${reportData.generatedAt}`;
    document.getElementById('metaSummary').textContent = `Summary JSON: ${reportData.summaryJsonPath}`;
    document.getElementById('metaTotal').textContent = `Total VMs: ${reportData.totalVms}`;
    const badge = document.getElementById('heroBadge');
    badge.textContent = reportData.badge || 'GuestOps Fleet';

    const stackBar = document.getElementById('stackBar');
    stackBar.innerHTML = '';
    const segments = [
      { key:'PASS', value: reportData.counts.PASS, pct: reportData.percentages.PASS },
      { key:'WARN', value: reportData.counts.WARN, pct: reportData.percentages.WARN },
      { key:'FAIL', value: reportData.counts.FAIL, pct: reportData.percentages.FAIL },
      { key:'INFO', value: reportData.counts.INFO, pct: reportData.percentages.INFO },
      { key:'OTHER', value: reportData.counts.OTHER, pct: reportData.percentages.OTHER }
    ].filter(x => x.value > 0);

    if (!segments.length) {
      stackBar.innerHTML = '<div class="stack-segment OTHER" style="width:100%"><span>No data</span></div>';
    } else {
      segments.forEach(seg => {
        const width = Math.max(seg.pct, 6);
        const div = document.createElement('div');
        div.className = `stack-segment ${seg.key}`;
        div.style.width = `${width}%`;
        div.innerHTML = `<span>${seg.key} ${seg.value}</span>`;
        stackBar.appendChild(div);
      });
    }

    const ring = document.getElementById('scoreRing');
    const p1 = reportData.percentages.PASS;
    const p2 = p1 + reportData.percentages.WARN;
    const p3 = p2 + reportData.percentages.FAIL;
    const p4 = p3 + reportData.percentages.INFO;
    ring.style.background = `conic-gradient(
      var(--pass) 0 ${p1}%,
      var(--warn) ${p1}% ${p2}%,
      var(--fail) ${p2}% ${p3}%,
      var(--info) ${p3}% ${p4}%,
      var(--other) ${p4}% 100%
    )`;
    document.getElementById('scoreRingValue').textContent = reportData.counts.PASS;
  }

  function renderStats(){
    const statsRow = document.getElementById('statsRow');
    const cards = [
      { label:'Total VMs', value: reportData.totalVms, cls:'' },
      { label:'PASS VMs', value: reportData.counts.PASS, cls:'PASS' },
      { label:'WARN VMs', value: reportData.counts.WARN, cls:'WARN' },
      { label:'FAIL VMs', value: reportData.counts.FAIL, cls:'FAIL' },
      { label:'INFO VMs', value: reportData.counts.INFO, cls:'INFO' },
      { label:'Other', value: reportData.counts.OTHER, cls:'OTHER' }
    ];
    statsRow.innerHTML = cards.map(card => `
      <div class="stat-card ${card.cls}">
        <div class="stat-label">${escapeHtml(card.label)}</div>
        <div class="stat-value">${escapeHtml(card.value)}</div>
      </div>
    `).join('');
  }

  function renderFilters(){
    const host = document.getElementById('filterBar');
    const filterValues = ['ALL','PASS','WARN','FAIL','INFO','OTHER'];
    host.innerHTML = filterValues.map(value => `
      <button type="button" class="filter-btn ${value === currentFilter ? 'active' : ''}" data-filter="${value}">
        ${value}
      </button>
    `).join('');

    host.querySelectorAll('.filter-btn').forEach(btn => {
      btn.addEventListener('click', () => {
        currentFilter = btn.dataset.filter;
        renderFilters();
        renderVmNav();
        renderOverview();
        const visibleNav = Array.from(document.querySelectorAll('.vm-nav-item')).find(x => x.style.display !== 'none');
        if (visibleNav && document.querySelector('.view-btn.active')?.dataset.view === 'DETAIL') {
          activateVm(visibleNav.dataset.vmId);
        }
      });
    });
  }

  function getFilteredVms(){
    if (currentFilter === 'ALL') return reportData.vms;
    return reportData.vms.filter(vm => (vm.status || 'OTHER') === currentFilter);
  }

  function renderVmNav(){
    const vmNav = document.getElementById('vmNav');
    const items = getFilteredVms();
    if (!items.length) {
      vmNav.innerHTML = '<div class="empty-state">No VMs match the current filter.</div>';
      return;
    }

    vmNav.innerHTML = items.map((vm, index) => `
      <button type="button" class="vm-nav-item ${escapeHtml(vm.status || 'OTHER')} ${index === 0 ? 'active' : ''}" data-vm-id="${escapeHtml(vm.vmId)}" data-status="${escapeHtml(vm.status || 'OTHER')}">
        <span class="vm-nav-title">${escapeHtml(vm.vmName || 'Unknown VM')}</span>
        <span class="vm-nav-meta">${escapeHtml(vm.detectedTargetOs || 'Unknown')} · ${escapeHtml(vm.status || 'OTHER')} · Exit ${escapeHtml(vm.exitCode)}</span>
        <span class="vm-nav-sub">${escapeHtml(vm.payloadName || 'No payload')} ${vm.rawOutputPreview ? '· ' + escapeHtml(vm.rawOutputPreview) : ''}</span>
      </button>
    `).join('');

    vmNav.querySelectorAll('.vm-nav-item').forEach(btn => {
      btn.addEventListener('click', () => {
        activateVm(btn.dataset.vmId);
        setView('DETAIL', btn.dataset.vmId);
      });
    });
  }

  function renderOverview(){
    const grid = document.getElementById('overviewGrid');
    const items = getFilteredVms();
    if (!items.length) {
      grid.innerHTML = '<div class="empty-state">No VMs match the current filter.</div>';
      return;
    }

    grid.innerHTML = items.map(vm => `
      <article class="vm-overview-card ${escapeHtml(vm.status || 'OTHER')}" data-status="${escapeHtml(vm.status || 'OTHER')}">
        <div class="vm-overview-top">
          <div>
            <div class="eyebrow">Virtual Machine</div>
            <h3>${escapeHtml(vm.vmName || 'Unknown VM')}</h3>
            <p>OS: ${escapeHtml(vm.detectedTargetOs || 'Unknown')} &#183; User: ${escapeHtml(vm.guestUser || '-')} &#183; Creds: ${escapeHtml(vm.credentialSource || '-')}</p>
          </div>
          <div class="overview-actions">
            ${statusBadge(vm.status || 'OTHER', true)}
            <button type="button" class="view-vm-btn" data-vm-id="${escapeHtml(vm.vmId)}">Open details</button>
          </div>
        </div>
        <ul class="mini-check-list">
          <li><span class="mini-check-name">Payload</span><span>${escapeHtml(vm.payloadName || '-')}</span></li>
          <li><span class="mini-check-name">Mode</span><span>${escapeHtml(vm.payloadMode || '-')}</span></li>
          <li><span class="mini-check-name">Exit code</span><span>${escapeHtml(vm.exitCode)}</span></li>
          <li><span class="mini-check-name">Output preview</span><span>${escapeHtml(vm.rawOutputPreview || '-')}</span></li>
        </ul>
      </article>
    `).join('');

    grid.querySelectorAll('.view-vm-btn').forEach(btn => {
      btn.addEventListener('click', () => {
        activateVm(btn.dataset.vmId);
        setView('DETAIL', btn.dataset.vmId);
      });
    });
  }

  function metadataRows(vm){
    const rows = [
      ['VM Name', vm.vmName],
      ['Status', vm.status],
      ['Exit Code', vm.exitCode],
      ['Detected Target OS', vm.detectedTargetOs],
      ['Requested Target OS', vm.requestedTargetOs],
      ['Guest Full Name', vm.guestFullName],
      ['Guest User', vm.guestUser],
      ['Credential Source', vm.credentialSource],
      ['Credential Detail', vm.credentialSourceDetail],
      ['OS Detection Source', vm.osDetectionSource],
      ['Payload Name', vm.payloadName],
      ['Payload Mode', vm.payloadMode],
      ['Payload Source', vm.payloadSource],
      ['Payload Source Path', vm.payloadSourcePath],
      ['Execution Command Template', vm.executionCommandTemplate],
      ['Log Path', vm.logPath],
      ['Raw Output Path', vm.rawOutputPath]
    ];
    return rows.filter(row => row[1] !== null && row[1] !== undefined && String(row[1]) !== '').map(row => `
      <tr class="${escapeHtml(vm.status || 'OTHER')}">
        <td>${escapeHtml(row[0])}</td>
        <td>${escapeHtml(row[1])}</td>
      </tr>
    `).join('');
  }

  function highlightRawOutput(text){
    const value = String(text || '');
    if (!value.trim()) {
      return '<span class="code-line line-muted">No guest output was available for this VM.</span>';
    }

    return value.split(/\r?\n/).map(line => {
      const encoded = escapeHtml(line);
      let cls = 'line-muted';

      if (/^\[PASS\]|^RESULT PASS|^PASS\s*:|^GOR_RESULT EXITCODE:0\b/i.test(line)) cls = 'line-pass';
      else if (/^\[WARN\]|^RESULT WARN|^WARN\s*:/i.test(line)) cls = 'line-warn';
      else if (/^\[FAIL\]|^RESULT FAIL|^FAIL\s*:|^GOR_RESULT EXITCODE:(?!0)\d+/i.test(line)) cls = 'line-fail';
      else if (/^\[INFO\]|^RESULT INFO|^INFO\s*:|^GOR_RESULT INFO/i.test(line)) cls = 'line-info';
      else if (/^---|^===|^#+\s|^GOR_RESULT OUTPUT_BEGIN|^GOR_RESULT OUTPUT_END/i.test(line)) cls = 'line-banner';

      return `<span class="code-line ${cls}">${encoded || '&nbsp;'}</span>`;
    }).join('');
  }

  function renderVmPanels(){
    const host = document.getElementById('vmPanelsHost');
    host.innerHTML = reportData.vms.map(vm => `
      <section class="vm-panel" id="panel-${escapeHtml(vm.vmId)}" data-status="${escapeHtml(vm.status || 'OTHER')}">
        <div class="panel-header">
          <div>
            <div class="eyebrow">Virtual Machine</div>
            <h2>${escapeHtml(vm.vmName || 'Unknown VM')}</h2>
            <p class="panel-subtitle">${escapeHtml(vm.payloadName || 'No payload selected')}</p>
          </div>
          ${statusBadge(vm.status || 'OTHER', true)}
        </div>

        <div class="detail-grid">
          <div class="detail-card"><div class="detail-label">Detected OS</div><div class="detail-value">${escapeHtml(vm.detectedTargetOs || '-')}</div></div>
          <div class="detail-card"><div class="detail-label">Guest</div><div class="detail-value">${escapeHtml(vm.guestFullName || '-')}</div></div>
          <div class="detail-card"><div class="detail-label">Guest User</div><div class="detail-value">${escapeHtml(vm.guestUser || '-')}</div></div>
          <div class="detail-card"><div class="detail-label">Credential Source</div><div class="detail-value">${escapeHtml(vm.credentialSource || '-')}</div></div>
          <div class="detail-card"><div class="detail-label">Payload</div><div class="detail-value">${escapeHtml(vm.payloadName || '-')}</div></div>
          <div class="detail-card"><div class="detail-label">Payload Mode</div><div class="detail-value">${escapeHtml(vm.payloadMode || '-')}</div></div>
          <div class="detail-card"><div class="detail-label">Requested OS</div><div class="detail-value">${escapeHtml(vm.requestedTargetOs || '-')}</div></div>
          <div class="detail-card"><div class="detail-label">Exit Code</div><div class="detail-value">${escapeHtml(vm.exitCode)}</div></div>
        </div>

        <div class="panel-card">
          <h3>Execution Metadata</h3>
          <div class="table-wrap">
            <table>
              <thead>
                <tr><th>Property</th><th>Value</th></tr>
              </thead>
              <tbody>
                ${metadataRows(vm)}
              </tbody>
            </table>
          </div>
        </div>

        <div class="panel-card">
          <div class="action-row">
            ${vm.logHref ? `<a class="link-btn" href="${escapeHtml(vm.logHref)}" target="_blank" rel="noopener">Open detailed log</a>` : ''}
            ${vm.rawOutputHref ? `<a class="link-btn" href="${escapeHtml(vm.rawOutputHref)}" target="_blank" rel="noopener">Open raw output file</a>` : ''}
            ${vm.detailOutput ? `<button type="button" class="toggle-btn" data-target="detail-${escapeHtml(vm.vmId)}">Toggle execution log</button>` : ''}
            <button type="button" class="toggle-btn" data-target="raw-${escapeHtml(vm.vmId)}">Toggle guest output</button>
          </div>

          ${vm.detailOutput ? `
          <div class="output-wrap">
            <div class="output-toolbar">
              <div class="output-title">Embedded execution detail</div>
              <div>${vm.detailOutputPreview ? escapeHtml(vm.detailOutputPreview) : ''}</div>
            </div>
            <pre id="detail-${escapeHtml(vm.vmId)}" class="raw-output hidden">${highlightRawOutput(vm.detailOutput)}</pre>
          </div>` : ''}

          <div class="output-wrap">
            <div class="output-toolbar">
              <div class="output-title">Syntax-highlighted guest output</div>
              <div>${statusBadge(vm.status || 'OTHER', false)}</div>
            </div>
            <pre id="raw-${escapeHtml(vm.vmId)}" class="raw-output">${highlightRawOutput(vm.rawOutput)}</pre>
          </div>
        </div>
      </section>
    `).join('');

    host.querySelectorAll('.toggle-btn').forEach(btn => {
      btn.addEventListener('click', () => {
        const target = document.getElementById(btn.dataset.target);
        if (target) target.classList.toggle('hidden');
      });
    });
  }

  function activateVm(vmId){
    const navItems = Array.from(document.querySelectorAll('.vm-nav-item'));
    const panels = Array.from(document.querySelectorAll('.vm-panel'));
    navItems.forEach(btn => btn.classList.toggle('active', btn.dataset.vmId === vmId));
    panels.forEach(panel => {
      const active = panel.id === 'panel-' + vmId;
      panel.classList.toggle('active', active);
      panel.style.display = active ? 'flex' : 'none';
    });
  }

  function setView(mode, preferredVmId){
    const overviewPanel = document.getElementById('all-vms-panel');
    const panels = Array.from(document.querySelectorAll('.vm-panel'));
    const viewButtons = Array.from(document.querySelectorAll('.view-btn'));
    const showAll = mode === 'ALLVMS';

    viewButtons.forEach(btn => btn.classList.toggle('active', btn.dataset.view === mode));

    if (overviewPanel) {
      overviewPanel.classList.toggle('active', showAll);
      overviewPanel.style.display = showAll ? 'flex' : 'none';
    }

    if (showAll) {
      panels.forEach(panel => {
        panel.classList.remove('active');
        panel.style.display = 'none';
      });
    } else {
      const visibleNav = Array.from(document.querySelectorAll('.vm-nav-item')).find(x => x.style.display !== 'none');
      const targetId = preferredVmId || (visibleNav ? visibleNav.dataset.vmId : null);
      if (targetId) activateVm(targetId);
    }
  }

  function wireViewButtons(){
    document.querySelectorAll('.view-btn').forEach(btn => {
      btn.addEventListener('click', () => {
        setView(btn.dataset.view);
      });
    });
  }

  renderHero();
  renderStats();
  renderFilters();
  renderVmNav();
  renderOverview();
  renderVmPanels();
  wireViewButtons();

  const firstVisibleVm = getFilteredVms()[0];
  if (firstVisibleVm) activateVm(firstVisibleVm.vmId);
  setView('ALLVMS');
})();
</script>
</body>
</html>
'@

$html = $htmlTemplate.Replace('__TITLE__', [System.Net.WebUtility]::HtmlEncode($Title))
$html = $html.Replace('__REPORT_DATA__', $jsonForBrowser)

$utf8NoBom = New-Object System.Text.UTF8Encoding($false)
[System.IO.File]::WriteAllText($HtmlOutPath, $html, $utf8NoBom)

Write-Info ("HTML report written to: {0}" -f $HtmlOutPath) 'PASS'

if ($OpenReport) {
  try {
    Start-Process -FilePath $HtmlOutPath | Out-Null
    Write-Info 'Opened the report in the default browser.' 'INFO'
  }
  catch {
    Write-Info ("Could not auto-open the report: {0}" -f $_.Exception.Message) 'WARN'
  }
}
