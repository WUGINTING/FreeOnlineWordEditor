<#
  Opens the documents written by tests/papyrus/export-intent.test.ts in Microsoft Word
  (hidden, read-only) and prints what Word itself sees: revisions and their authors,
  content controls, alignment, bold, and list numbering.

  Usage:
    1. DOCX_EXPORT=<out> npx vitest run tests/papyrus/export-intent.test.ts
    2. pwsh scripts/inspect-intent-in-word.ps1 -Folder <out>
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

$align = @{ 0 = 'left'; 1 = 'center'; 2 = 'right'; 3 = 'justify' }
try {
  $d = Open-Doc 'clear-formatting.docx'
  $rev = $d.Revisions
  Write-Output ("clear-formatting: revisions={0} author={1} text='{2}'" -f $rev.Count, $rev.Item(1).Author, $rev.Item(1).Range.Text)
  $cc = $d.ContentControls
  Write-Output ("clear-formatting: contentControls={0} title={1} tag={2} text='{3}'" -f $cc.Count, $cc.Item(1).Title, $cc.Item(1).Tag, $cc.Item(1).Range.Text)
  Write-Output ("clear-formatting: bold={0} italic={1} color={2}" -f $d.Content.Font.Bold, $d.Content.Font.Italic, $d.Content.Font.Color)
  $d.Close(0)

  $d = Open-Doc 'style-override.docx'
  $p = $d.Paragraphs.Item(1)
  Write-Output ("style-override: style={0} alignment={1} bold={2}" -f $p.Style.NameLocal, $align[[int]$p.Alignment], $p.Range.Font.Bold)
  $d.Close(0)

  $d = Open-Doc 'pasted-list.docx'
  foreach ($p in $d.Paragraphs) {
    $text = $p.Range.Text.Trim()
    Write-Output ("pasted-list: '{0}' list='{1}' level={2}" -f $text, $p.Range.ListFormat.ListString, $p.Range.ListFormat.ListLevelNumber)
  }
  $d.Close(0)
}
finally {
  $word.Quit(0)
  [void][Runtime.InteropServices.Marshal]::ReleaseComObject($word)
}
