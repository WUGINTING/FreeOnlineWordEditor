<#
  Opens the documents written by tests/papyrus/export-features.test.ts in Microsoft Word
  (hidden, read-only) and prints what Word itself sees: cell shading, borders, vertical
  alignment, row height, header rows, "allow row to break", column widths, picture size and
  alt text, and the text / revisions after "replace all".

  Usage:
    1. DOCX_EXPORT=<out> npx vitest run tests/papyrus/export-features.test.ts
    2. pwsh scripts/inspect-features-in-word.ps1 -Folder <out>
#>
param([Parameter(Mandatory = $true)][string]$Folder)

$ErrorActionPreference = 'Stop'
$word = New-Object -ComObject Word.Application
$word.Visible = $false
$word.DisplayAlerts = 0

function Open-Doc([string]$name) {
  # FileName, ConfirmConversions, ReadOnly, AddToRecentFiles
  return $word.Documents.Open((Join-Path $Folder $name), $false, $true, $false)
}

# WdBorderType: top -1, left -2, bottom -3, right -4. LineStyle 0 = none, 1 = single.
function Borders($cell) {
  return ('top={0} left={1} bottom={2} right={3}' -f $cell.Borders.Item(-1).LineStyle, $cell.Borders.Item(-2).LineStyle, $cell.Borders.Item(-3).LineStyle, $cell.Borders.Item(-4).LineStyle)
}
$vAlign = @{ 0 = 'top'; 1 = 'center'; 3 = 'bottom' }
$rule = @{ 0 = 'auto'; 1 = 'atLeast'; 2 = 'exactly' }
try {
  $d = Open-Doc 'table.docx'
  $t = $d.Tables.Item(1)
  $a1 = $t.Cell(1, 1)
  Write-Output ('table: A1 shading=0x{0:X6} (BGR)' -f [int]$a1.Shading.BackgroundPatternColor)
  Write-Output ('table: C1 vAlign={0}' -f $vAlign[[int]$t.Cell(1, 3).VerticalAlignment])
  foreach ($i in 1..3) {
    $r = $t.Rows.Item($i)
    Write-Output ('table: row{0} height={1}pt rule={2} header={3} allowBreak={4}' -f $i, $r.Height, $rule[[int]$r.HeightRule], $r.HeadingFormat, $r.AllowBreakAcrossPages)
  }
  foreach ($c in @(@(2, 2), @(1, 2), @(2, 1), @(2, 3), @(3, 2), @(3, 1), @(3, 3))) {
    Write-Output ('table: cell({0},{1}) {2}' -f $c[0], $c[1], (Borders $t.Cell($c[0], $c[1])))
  }
  $t2 = $d.Tables.Item(2)
  Write-Output ('table2: widths={0}, {1}, {2} pt' -f $t2.Cell(1, 1).Width, $t2.Cell(1, 2).Width, $t2.Cell(1, 3).Width)
  $d.Close(0)

  $d = Open-Doc 'image.docx'
  foreach ($i in 1..$d.InlineShapes.Count) {
    $s = $d.InlineShapes.Item($i)
    Write-Output ('image{0}: {1:N2} x {2:N2} cm alt="{3}" line={4}pt' -f $i, ($s.Width / 28.3465), ($s.Height / 28.3465), $s.AlternativeText, $s.Line.Weight)
  }
  $d.Close(0)

  $d = Open-Doc 'replace.docx'
  Write-Output ("replace: text='{0}'" -f $d.Content.Text.Trim())
  Write-Output ("replace: revisions={0} author={1} text='{2}' firstCharBold={3}" -f $d.Revisions.Count, $d.Revisions.Item(1).Author, $d.Revisions.Item(1).Range.Text, $d.Characters.Item(1).Bold)
  $d.Close(0)
}
finally {
  $word.Quit(0)
  [void][Runtime.InteropServices.Marshal]::ReleaseComObject($word)
}
