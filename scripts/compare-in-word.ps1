<#
  Opens each original .docx and its round-tripped copy in Microsoft Word (hidden)
  and compares what Word itself reports: pages, words, paragraphs, tables, pictures,
  header/footer text and the full body text.

  Usage:
    1. DOCX_SAMPLES=<folder> DOCX_EXPORT=<out> npx vitest run tests/export.test.ts
    2. pwsh scripts/compare-in-word.ps1 -Index <out>/index.json
#>
param([Parameter(Mandatory = $true)][string]$Index)

$ErrorActionPreference = 'Stop'
$pairs = Get-Content -Raw -Encoding UTF8 $Index | ConvertFrom-Json

$word = New-Object -ComObject Word.Application
$word.Visible = $false
$word.DisplayAlerts = 0 # wdAlertsNone

function Measure-Doc([string]$path) {
  $doc = $null
  try {
    # FileName, ConfirmConversions, ReadOnly, AddToRecentFiles
    $doc = $word.Documents.Open($path, $false, $true, $false)
    $doc.Repaginate()
    $sec = $doc.Sections.Item($doc.Sections.Count)
    $text = $doc.Content.Text
    $sha = [System.Security.Cryptography.SHA256]::Create()
    $hash = [BitConverter]::ToString($sha.ComputeHash([Text.Encoding]::UTF8.GetBytes($text))).Replace('-', '').Substring(0, 12)
    [ordered]@{
      ok         = $true
      pages      = $doc.ComputeStatistics(2)
      words      = $doc.ComputeStatistics(0)
      paragraphs = $doc.Paragraphs.Count
      tables     = $doc.Tables.Count
      pictures   = $doc.InlineShapes.Count + $doc.Shapes.Count
      header     = $sec.Headers.Item(1).Range.Text.Trim()
      footer     = $sec.Footers.Item(1).Range.Text.Trim()
      textHash   = $hash
      text       = $text
    }
  } catch {
    [ordered]@{ ok = $false; error = $_.Exception.Message }
  } finally {
    if ($doc) { $doc.Close([ref]0) | Out-Null }
  }
}

$results = @()
try {
  foreach ($p in $pairs) {
    $a = Measure-Doc $p.original
    $b = Measure-Doc $p.copy
    $diffs = @()
    if (-not $b.ok) { $diffs += "copy failed to open: $($b.error)" }
    elseif ($a.ok) {
      foreach ($k in 'pages', 'words', 'paragraphs', 'tables', 'pictures', 'header', 'footer', 'textHash') {
        if ("$($a[$k])" -ne "$($b[$k])") { $diffs += "${k}: $($a[$k]) -> $($b[$k])" }
      }
    }
    if ($a.ok -and $b.ok -and $a.textHash -ne $b.textHash) {
      # Keep both texts for inspection: <copy>.original.txt / <copy>.copy.txt
      Set-Content -Encoding UTF8 "$($p.copy).original.txt" ($a.text -replace "`r", "`n")
      Set-Content -Encoding UTF8 "$($p.copy).copy.txt" ($b.text -replace "`r", "`n")
    }
    $name = Split-Path $p.original -Leaf
    $results += [pscustomobject]@{ file = $name; same = ($diffs.Count -eq 0); diffs = ($diffs -join '; ') }
    Write-Host ("{0,-6} {1}  {2}" -f $(if ($diffs.Count) { 'DIFF' } else { 'SAME' }), $name, ($diffs -join '; '))
  }
} finally {
  $word.Quit([ref]0)
  [void][Runtime.InteropServices.Marshal]::ReleaseComObject($word)
}

$same = ($results | Where-Object same).Count
Write-Host "`n$same / $($results.Count) documents identical in Word"
$results | ConvertTo-Json | Set-Content -Encoding UTF8 (Join-Path (Split-Path $Index) 'word-compare.json')
