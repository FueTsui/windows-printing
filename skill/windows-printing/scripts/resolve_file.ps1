param(
  [Parameter(Mandatory = $true)]
  [string]$Query,

  [string[]]$SearchRoots
)

$ErrorActionPreference = 'Stop'
[Console]::OutputEncoding = [System.Text.UTF8Encoding]::new()
$OutputEncoding = [Console]::OutputEncoding

if (-not $SearchRoots -or $SearchRoots.Count -eq 0) {
  $desktop = [Environment]::GetFolderPath('Desktop')
  $documents = [Environment]::GetFolderPath('MyDocuments')
  $downloads = Join-Path ([Environment]::GetFolderPath('UserProfile')) 'Downloads'
  $SearchRoots = @($desktop, $documents, $downloads)
}

$queryText = $Query.Trim()
$queryLeaf = [System.IO.Path]::GetFileName($queryText)
$normalized = $queryLeaf.ToLowerInvariant()
$results = New-Object System.Collections.Generic.List[object]

foreach ($root in $SearchRoots) {
  if (-not $root -or -not (Test-Path -LiteralPath $root)) { continue }

  if (Test-Path -LiteralPath $queryText) {
    $item = Get-Item -LiteralPath $queryText
    if ($item.PSIsContainer) { continue }
    $results.Add([PSCustomObject]@{
      score = 1000
      name = $item.Name
      fullPath = $item.FullName
      directory = $item.DirectoryName
      length = $item.Length
      lastWriteTime = $item.LastWriteTime
      reason = 'exact-path'
    })
    break
  }

  Get-ChildItem -LiteralPath $root -File -Recurse -ErrorAction SilentlyContinue | ForEach-Object {
    $name = $_.Name
    $full = $_.FullName
    $nameLower = $name.ToLowerInvariant()
    $fullLower = $full.ToLowerInvariant()
    $score = 0
    $reason = $null

    if ($nameLower -eq $normalized) {
      $score = 950
      $reason = 'exact-name'
    }
    elseif ($nameLower -like "*$normalized*") {
      $score = 800
      $reason = 'name-contains'
    }
    elseif ($fullLower -like "*${normalized}*") {
      $score = 700
      $reason = 'path-contains'
    }
    elseif ([System.IO.Path]::GetFileNameWithoutExtension($nameLower) -eq [System.IO.Path]::GetFileNameWithoutExtension($normalized)) {
      $score = 650
      $reason = 'basename-match'
    }

    if ($score -gt 0) {
      $results.Add([PSCustomObject]@{
        score = $score
        name = $_.Name
        fullPath = $_.FullName
        directory = $_.DirectoryName
        length = $_.Length
        lastWriteTime = $_.LastWriteTime
        reason = $reason
      })
    }
  }
}

$results |
  Sort-Object -Property @(
    @{ Expression = 'score'; Descending = $true },
    @{ Expression = 'lastWriteTime'; Descending = $true }
  ) |
  Select-Object -First 20 |
  ConvertTo-Json -Depth 4
