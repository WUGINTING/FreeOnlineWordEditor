<#
  Opens the documents written by tests/papyrus/export-qa-2026-09-25.test.ts in Microsoft Word
  (hidden, read-only) and prints what Word itself sees for the QA 2026-09-25 fixes:
  a list pasted from another document (BUG-005), the same paste undone (BUG-003) and
  javascript: link targets (BUG-002).

  Usage:
    1. DOCX_EXPORT=<out> npx vitest run tests/papyrus/export-qa-2026-09-25.test.ts
    2. pwsh scripts/inspect-qa-2026-09-25-in-word.ps1 -Folder <out>
#>
param([Parameter(Mandatory = $true)][string]$Folder)

$ErrorActionPreference = 'Stop'
# Only the Word started here is closed: remember the ones already running.
$before = @(Get-Process WINWORD -ErrorAction SilentlyContinue | ForEach-Object Id)
$word = New-Object -ComObject Word.Application
$mine = @(Get-Process WINWORD -ErrorAction SilentlyContinue | Where-Object { $before -notcontains $_.Id } | ForEach-Object Id)
$word.Visible = $false
$word.DisplayAlerts = 0

function Open-Doc([string]$name) {
  # FileName, ConfirmConversions, ReadOnly, AddToRecentFiles
  return $word.Documents.Open((Join-Path $Folder $name), $false, $true, $false)
}
# WdListType
$listType = @{ 0 = 'none'; 1 = 'listNumOnly'; 2 = 'bullet'; 3 = 'simpleNumbering'; 4 = 'outlineNumbering'; 5 = 'mixedNumbering'; 6 = 'pictureBullet' }

try {
  foreach ($name in 'xdoc-list.docx', 'xdoc-list-undone.docx') {
    $d = Open-Doc $name
    foreach ($p in $d.Paragraphs) {
      $lf = $p.Range.ListFormat
      Write-Output ("{0}: '{1}' list='{2}' type={3}" -f $name, $p.Range.Text.Trim(), $lf.ListString, $listType[[int]$lf.ListType])
    }
    Write-Output ("{0}: lists={1}" -f $name, $d.Lists.Count)
    $d.Close(0)
  }
  foreach ($name in 'unsafe-link.docx', 'loaded-unsafe-link.docx') {
    $d = Open-Doc $name
    Write-Output ("{0}: hyperlinks={1}" -f $name, $d.Hyperlinks.Count)
    foreach ($h in $d.Hyperlinks) { Write-Output ("{0}: '{1}' -> {2}" -f $name, $h.TextToDisplay, $h.Address) }
    Write-Output ("{0}: text='{1}'" -f $name, ($d.Content.Text -replace "`r", ' | ').Trim())
    $d.Close(0)
  }
}
finally {
  # Word attached to one already running (no new process): leave it open.
  if ($mine.Count -gt 0) { $word.Quit(0) }
  [void][Runtime.InteropServices.Marshal]::ReleaseComObject($word)
  Start-Sleep -Milliseconds 500
  foreach ($id in $mine) { Stop-Process -Id $id -ErrorAction SilentlyContinue }
}
