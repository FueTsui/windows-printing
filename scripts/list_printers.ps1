[Console]::OutputEncoding = [System.Text.UTF8Encoding]::new()
$OutputEncoding = [Console]::OutputEncoding

$defaultPrinter = (Get-Printer | Where-Object { $_.Default } | Select-Object -First 1 -ExpandProperty Name)

Get-Printer |
  Sort-Object Name |
  Select-Object @{Name='Index';Expression={0}}, Name, PrinterStatus, WorkOffline, PortName, Default |
  ForEach-Object -Begin { $i = 1 } -Process {
    $cfg = Get-PrintConfiguration -PrinterName $_.Name -ErrorAction SilentlyContinue
    [PSCustomObject]@{
      Index = $i
      Name = $_.Name
      IsDefault = [bool]($_.Name -eq $defaultPrinter)
      PrinterStatus = $_.PrinterStatus
      WorkOffline = $_.WorkOffline
      PortName = $_.PortName
      CurrentColor = if ($null -ne $cfg.Color) { [bool]$cfg.Color } else { $null }
      CurrentDuplexingMode = if ($null -ne $cfg.DuplexingMode) { [string]$cfg.DuplexingMode } else { $null }
      CurrentPaperSize = if ($null -ne $cfg.PaperSize) { [string]$cfg.PaperSize } else { $null }
    }
    $i++
  } |
  ConvertTo-Json -Depth 4
