<#
  Opens a document with page columns (e.g. cols-ui.docx written by the opt-in test in
  tests/papyrus/columnsAuthoring.test.ts) in Microsoft Word, hidden and read-only, and prints
  per section: how it starts, its text columns (count, line between, spacing), and where Word
  puts the title, the Chinese text and the text after the column break ("ENGLISH-STARTS-HERE").
  Saves Word's own PDF. Closes only the Word process it started.

  Usage:
    1. COLS_EXPORT=<out> npx vitest run tests/papyrus/columnsAuthoring.test.ts
    2. pwsh scripts/inspect-columns-in-word.ps1 -Docx <out>\cols-ui.docx -Pdf <out>\cols-ui.word.pdf
#>
param([Parameter(Mandatory = $true)][string]$Docx, [Parameter(Mandatory = $true)][string]$Pdf)
$ErrorActionPreference = 'Stop'
$before = @(Get-Process WINWORD -ErrorAction SilentlyContinue | ForEach-Object Id)
$word = New-Object -ComObject Word.Application
$word.Visible = $false
$word.DisplayAlerts = 0
$mine = @(Get-Process WINWORD -ErrorAction SilentlyContinue | Where-Object { $before -notcontains $_.Id } | ForEach-Object Id)
Write-Output ("Word processes before: [{0}]  started: [{1}]" -f ($before -join ','), ($mine -join ','))
try {
  $d = $word.Documents.Open($Docx, $false, $true, $false)
  $d.Repaginate()
  Write-Output ("sections={0} pages={1}" -f $d.Sections.Count, $d.ComputeStatistics(2))
  $i = 0
  foreach ($s in $d.Sections) {
    $i++
    $ps = $s.PageSetup
    $tc = $ps.TextColumns
    Write-Output ("section {0}: SectionStart={1} TextColumns.Count={2} LineBetween={3} EvenlySpaced={4} Spacing={5}pt text='{6}'" -f `
        $i, $ps.SectionStart, $tc.Count, $tc.LineBetween, $tc.EvenlySpaced, $tc.Spacing, ($s.Range.Text -replace "[`r`n`a`f]", ' ').Trim())
  }
  $ps2 = $d.Sections(2).PageSetup
  $mid = $ps2.LeftMargin + ($ps2.PageWidth - $ps2.LeftMargin - $ps2.RightMargin) / 2
  Write-Output ("page width={0}pt left margin={1}pt text middle x={2}pt" -f $ps2.PageWidth, $ps2.LeftMargin, $mid)
  foreach ($needle in @('雙語公告', '本公告', 'ENGLISH-STARTS-HERE')) {
    $r = $d.Content
    $f = $r.Find
    if ($f.Execute($needle)) {
      # wdHorizontalPositionRelativeToPage = 5, wdVerticalPositionRelativeToPage = 6, wdActiveEndPageNumber = 3
      $x = $r.Information(5); $y = $r.Information(6); $p = $r.Information(3)
      $col = if ($x -lt $mid) { 'left' } else { 'right' }
      Write-Output ("'{0}': page {1} x={2}pt y={3}pt -> {4} half" -f $needle, $p, $x, $y, $col)
    } else {
      Write-Output ("'{0}': not found" -f $needle)
    }
  }
  # wdExportFormatPDF = 17
  $d.ExportAsFixedFormat($Pdf, 17)
  $d.Close(0)
} finally {
  $word.Quit(0)
  [void][Runtime.InteropServices.Marshal]::ReleaseComObject($word)
  # Only the Word this script started: wait for it to exit after Quit, else end it.
  foreach ($id in $mine) {
    $p = Get-Process -Id $id -ErrorAction SilentlyContinue
    if ($p -and -not $p.WaitForExit(15000)) { $p | Stop-Process -Force }
  }
  $after = @(Get-Process WINWORD -ErrorAction SilentlyContinue | ForEach-Object Id)
  Write-Output ("Word processes after: [{0}]" -f ($after -join ','))
}
