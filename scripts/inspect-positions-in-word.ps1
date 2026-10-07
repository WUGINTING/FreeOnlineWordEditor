<#
  Opens documents in Microsoft Word (hidden, read-only, never saved) and prints where Word puts
  the start of each paragraph: page, x and y (pt from the page's top-left corner), its section's
  column count, and the first characters of its text. For comparing the online editor's layout
  (text columns, the document grid, page breaks) with Word's. Closes only the Word it started.

  Usage:
    pwsh scripts/inspect-positions-in-word.ps1 -Docx a.docx,b.docx [-Json out.json]
#>
param([Parameter(Mandatory = $true)][string[]]$Docx, [string]$Json)
$ErrorActionPreference = 'Stop'
$before = @(Get-Process WINWORD -ErrorAction SilentlyContinue | ForEach-Object Id)
$word = New-Object -ComObject Word.Application
$word.Visible = $false
$word.DisplayAlerts = 0
$mine = @(Get-Process WINWORD -ErrorAction SilentlyContinue | Where-Object { $before -notcontains $_.Id } | ForEach-Object Id)
$results = [ordered]@{}
try {
  foreach ($path in $Docx) {
    $full = (Resolve-Path $path).Path
    $d = $word.Documents.Open($full, $false, $true, $false)
    try {
      $d.Repaginate()
      $rows = @()
      Write-Output ("{0}: pages={1}" -f (Split-Path $full -Leaf), $d.ComputeStatistics(2))
      foreach ($p in $d.Paragraphs) {
        $r = $p.Range
        $text = ($r.Text -replace "[`r`n`a`f`v]", ' ').Trim()
        $r.Collapse(1) # wdCollapseStart
        # wdActiveEndPageNumber = 3, wdHorizontalPositionRelativeToPage = 5, wdVerticalPositionRelativeToPage = 6
        $row = [ordered]@{
          page = $r.Information(3); x = $r.Information(5); y = $r.Information(6)
          columns = $p.Range.Sections.Item(1).PageSetup.TextColumns.Count
          text = if ($text.Length -gt 24) { $text.Substring(0, 24) } else { $text }
        }
        $rows += $row
        Write-Output ("  p{0} x={1,6} y={2,6} cols={3} {4}" -f $row.page, $row.x, $row.y, $row.columns, $row.text)
      }
      $results[$full] = $rows
    } finally {
      $d.Close(0)
    }
  }
} finally {
  $word.Quit(0)
  [void][Runtime.InteropServices.Marshal]::ReleaseComObject($word)
  foreach ($id in $mine) {
    $p = Get-Process -Id $id -ErrorAction SilentlyContinue
    if ($p -and -not $p.WaitForExit(15000)) { $p | Stop-Process -Force }
  }
}
if ($Json) { $results | ConvertTo-Json -Depth 5 | Set-Content -Encoding UTF8 $Json }
