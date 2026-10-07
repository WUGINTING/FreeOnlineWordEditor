<#
  Opens the documents written by tests/papyrus/export-sections.test.ts in Microsoft Word
  (hidden, read-only) and prints, per section: orientation, page size, margins, whether the
  first page differs, and the header texts Word shows.

  Usage:
    1. QA_SECTIONS_DOCX=<qa file> DOCX_EXPORT=<out> npx vitest run tests/papyrus/export-sections.test.ts
    2. pwsh scripts/inspect-sections-in-word.ps1 -Folder <out>
#>
param([Parameter(Mandatory = $true)][string]$Folder)

$ErrorActionPreference = 'Stop'
$word = New-Object -ComObject Word.Application
$word.Visible = $false
$word.DisplayAlerts = 0
try {
  foreach ($file in Get-ChildItem -Path $Folder -Filter *.docx | Sort-Object Name) {
    $d = $word.Documents.Open($file.FullName, $false, $true, $false)
    $d.Repaginate()
    Write-Output ("{0}: sections={1} pages={2}" -f $file.Name, $d.Sections.Count, $d.ComputeStatistics(2))
    $i = 0
    foreach ($s in $d.Sections) {
      $i++
      $ps = $s.PageSetup
      $orient = if ($ps.Orientation -eq 1) { 'landscape' } else { 'portrait' }
      # wdHeaderFooterPrimary = 1, wdHeaderFooterFirstPage = 2
      $primary = $s.Headers.Item(1).Range.Text.Trim()
      $first = $s.Headers.Item(2).Range.Text.Trim()
      Write-Output ("  section {0}: {1} {2:N0}x{3:N0}pt left={4:N0}pt firstDiffers={5} header='{6}' firstHeader='{7}' text='{8}'" -f `
          $i, $orient, $ps.PageWidth, $ps.PageHeight, $ps.LeftMargin, $ps.DifferentFirstPageHeaderFooter, $primary, $first,
          ($s.Range.Text.Substring(0, [Math]::Min(40, $s.Range.Text.Length)) -replace "[`r`n`a]", ' '))
    }
    $d.Close(0)
  }
}
finally {
  $word.Quit(0)
  [void][Runtime.InteropServices.Marshal]::ReleaseComObject($word)
}
