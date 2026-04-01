param(
  [Parameter(Mandatory = $true)]
  [string]$FilePath,

  [Parameter(Mandatory = $true)]
  [string]$PrinterName,

  [ValidateSet('keep','color','grayscale')]
  [string]$ColorMode = 'keep',

  [ValidateRange(1, 99)]
  [int]$Copies = 1,

  [ValidateSet('keep','one-sided','long-edge','short-edge')]
  [string]$DuplexMode = 'keep',

  [string]$PaperSize,

  [switch]$WhatIfOnly
)

$ErrorActionPreference = 'Stop'
[Console]::OutputEncoding = [System.Text.UTF8Encoding]::new()
$OutputEncoding = [Console]::OutputEncoding

function Convert-ColorLabel($mode, $plannedColor) {
  if ($mode -eq 'grayscale' -or (-not $plannedColor)) { return 'grayscale' }
  if ($mode -eq 'color' -or $plannedColor) { return 'color' }
  return 'keep'
}

function Convert-DuplexLabel($mode, $plannedDuplexingMode) {
  switch ($plannedDuplexingMode) {
    'OneSided' { return 'one-sided' }
    'TwoSidedLongEdge' { return 'long-edge' }
    'TwoSidedShortEdge' { return 'short-edge' }
    default {
      switch ($mode) {
        'one-sided' { return 'one-sided' }
        'long-edge' { return 'long-edge' }
        'short-edge' { return 'short-edge' }
        default { return 'keep' }
      }
    }
  }
}

function Convert-QueueLabel($queueCount, $hasErrors) {
  if ($hasErrors) { return 'error' }
  if ($queueCount -eq 0) { return 'empty' }
  if ($queueCount -eq 1) { return '1-job' }
  return "$queueCount-jobs"
}

function New-Receipt($fileItem, $printerName, $colorLabel, $copies, $duplexLabel, $paperSize, $queueLabel, $ok, $whatIf, $errors) {
  $resultCode = if ($whatIf) {
    'preview-only'
  }
  elseif ($ok) {
    if ($queueLabel -eq 'empty') {
      'submitted-queue-empty'
    }
    else {
      'submitted'
    }
  }
  else {
    'failed'
  }

  $receipt = [ordered]@{
    title = 'print-receipt'
    fileName = $fileItem.Name
    printer = $printerName
    color = $colorLabel
    copies = $copies
    duplex = $duplexLabel
    paperSize = $paperSize
    queueStatus = $queueLabel
    result = $resultCode
  }

  if ($errors -and $errors.Count -gt 0) {
    $receipt['errorSummary'] = $errors[0]
  }

  return [PSCustomObject]$receipt
}

if (-not (Test-Path -LiteralPath $FilePath)) {
  throw "File not found: $FilePath"
}

$null = Get-Printer -Name $PrinterName -ErrorAction Stop
$configBefore = Get-PrintConfiguration -PrinterName $PrinterName -ErrorAction Stop
$fileItem = Get-Item -LiteralPath $FilePath -ErrorAction Stop
$extension = $fileItem.Extension.ToLowerInvariant()

$setArgs = @{ PrinterName = $PrinterName }

switch ($ColorMode) {
  'color' { $setArgs.Color = $true }
  'grayscale' { $setArgs.Color = $false }
}

switch ($DuplexMode) {
  'one-sided' { $setArgs.DuplexingMode = 'OneSided' }
  'long-edge' { $setArgs.DuplexingMode = 'TwoSidedLongEdge' }
  'short-edge' { $setArgs.DuplexingMode = 'TwoSidedShortEdge' }
}

if ($PaperSize -and $PaperSize.Trim()) {
  $setArgs.PaperSize = $PaperSize.Trim()
}

$plannedConfig = [PSCustomObject]@{
  Color = if ($ColorMode -eq 'keep') { [bool]$configBefore.Color } elseif ($ColorMode -eq 'color') { $true } else { $false }
  DuplexingMode = if ($DuplexMode -eq 'keep') { [string]$configBefore.DuplexingMode } elseif ($DuplexMode -eq 'one-sided') { 'OneSided' } elseif ($DuplexMode -eq 'long-edge') { 'TwoSidedLongEdge' } else { 'TwoSidedShortEdge' }
  PaperSize = if ($PaperSize -and $PaperSize.Trim()) { $PaperSize.Trim() } else { [string]$configBefore.PaperSize }
}

$colorLabel = Convert-ColorLabel -mode $ColorMode -plannedColor $plannedConfig.Color
$duplexLabel = Convert-DuplexLabel -mode $DuplexMode -plannedDuplexingMode $plannedConfig.DuplexingMode

if ($WhatIfOnly) {
  $receipt = New-Receipt -fileItem $fileItem -printerName $PrinterName -colorLabel $colorLabel -copies $Copies -duplexLabel $duplexLabel -paperSize $plannedConfig.PaperSize -queueLabel 'preview-not-submitted' -ok $true -whatIf $true -errors @()

  [PSCustomObject]@{
    ok = $true
    whatIf = $true
    printer = $PrinterName
    file = $FilePath
    fileExtension = $extension
    copies = $Copies
    requested = [PSCustomObject]@{
      colorMode = $ColorMode
      duplexMode = $DuplexMode
      paperSize = $(if ($PaperSize) { $PaperSize } else { $null })
    }
    currentConfig = [PSCustomObject]@{
      color = [bool]$configBefore.Color
      duplexingMode = [string]$configBefore.DuplexingMode
      paperSize = [string]$configBefore.PaperSize
    }
    plannedConfig = $plannedConfig
    receipt = $receipt
  } | ConvertTo-Json -Depth 6
  exit 0
}

$changedConfig = $false
if ($setArgs.Keys.Count -gt 1) {
  Set-PrintConfiguration @setArgs
  $changedConfig = $true
}

$printErrors = @()
try {
  for ($i = 1; $i -le $Copies; $i++) {
    try {
      Start-Process -FilePath $FilePath -Verb PrintTo -ArgumentList @($PrinterName) -ErrorAction Stop
      Start-Sleep -Milliseconds 1200
    }
    catch {
      $printErrors += $_.Exception.Message
    }
  }

  Start-Sleep -Seconds 3
  $jobs = @(Get-PrintJob -PrinterName $PrinterName -ErrorAction SilentlyContinue)
  $ok = ($printErrors.Count -eq 0)
  $queueLabel = Convert-QueueLabel -queueCount $jobs.Count -hasErrors (-not $ok)
  $receipt = New-Receipt -fileItem $fileItem -printerName $PrinterName -colorLabel $colorLabel -copies $Copies -duplexLabel $duplexLabel -paperSize $plannedConfig.PaperSize -queueLabel $queueLabel -ok $ok -whatIf $false -errors $printErrors

  [PSCustomObject]@{
    ok = $ok
    whatIf = $false
    printer = $PrinterName
    file = $FilePath
    fileExtension = $extension
    colorMode = $ColorMode
    copies = $Copies
    duplexMode = $DuplexMode
    paperSize = $(if ($PaperSize) { $PaperSize } else { $null })
    currentConfig = [PSCustomObject]@{
      color = [bool]$configBefore.Color
      duplexingMode = [string]$configBefore.DuplexingMode
      paperSize = [string]$configBefore.PaperSize
    }
    plannedConfig = $plannedConfig
    queueCount = $jobs.Count
    jobs = @($jobs | Select-Object PrinterName, DocumentName, JobStatus, SubmittedTime, PagesPrinted, TotalPages)
    errors = $printErrors
    receipt = $receipt
  } | ConvertTo-Json -Depth 6
}
finally {
  if ($changedConfig) {
    $restoreArgs = @{ PrinterName = $PrinterName }
    if ($null -ne $configBefore.Color) { $restoreArgs.Color = [bool]$configBefore.Color }
    if ($null -ne $configBefore.DuplexingMode) { $restoreArgs.DuplexingMode = [string]$configBefore.DuplexingMode }
    if ($null -ne $configBefore.PaperSize) { $restoreArgs.PaperSize = [string]$configBefore.PaperSize }

    try {
      Set-PrintConfiguration @restoreArgs
    }
    catch {
      Write-Warning ("Failed to restore printer configuration: " + $_.Exception.Message)
    }
  }
}
