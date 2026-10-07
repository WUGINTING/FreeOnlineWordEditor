<#
  Checks the documents written by tests/papyrus/export-review.test.ts in Microsoft Word
  (hidden): for every case, Word's own Accept All / Reject All on the source is compared with
  what the web editor wrote (text, alignment, indent, bold / italic / color per character),
  the editor's files must have no revisions left, and a file saved unchanged must keep all of
  them. Comments: count, author, text and the commented text.

  Usage:
    1. DOCX_EXPORT=<out> [QA_DOCX=<file>] npx vitest run tests/papyrus/export-review.test.ts
    2. pwsh scripts/inspect-review-in-word.ps1 -Folder <out>
#>
param([Parameter(Mandatory = $true)][string]$Folder)

$ErrorActionPreference = 'Stop'
$word = New-Object -ComObject Word.Application
$word.Visible = $false
$word.DisplayAlerts = 0
$align = @{ 0 = 'left'; 1 = 'center'; 2 = 'right'; 3 = 'justify' }

function Open-Doc([string]$name) {
  # FileName, ConfirmConversions, ReadOnly, AddToRecentFiles
  return $word.Documents.Open((Join-Path $Folder $name), $false, $true, $false)
}

# What the document looks like: per paragraph its text, alignment, indent and character formatting.
function Signature($d) {
  $out = @()
  foreach ($p in $d.Paragraphs) {
    $fmt = ''
    # By index: Word's character enumerator never ends on a last paragraph whose deleted mark was accepted.
    $chars = $p.Range.Characters
    for ($k = 1; $k -le $chars.Count; $k++) {
      $c = $chars.Item($k)
      $f = ''
      if ($c.Font.Bold -eq -1) { $f += 'B' }
      if ($c.Font.Italic -eq -1) { $f += 'I' }
      if ($c.Font.Color -ne -16777216 -and $c.Font.Color -ne 0) { $f += 'C' }
      $fmt += if ($f) { "[$f]" } else { '.' }
    }
    $out += ("'{0}' {1} ind={2} {3}" -f $p.Range.Text.TrimEnd("`r"), $align[[int]$p.Alignment], $p.LeftIndent, $fmt)
  }
  # Bookmarks and comments, with what they cover.
  foreach ($b in $d.Bookmarks) { $out += ("bookmark {0}='{1}'" -f $b.Name, $b.Range.Text) }
  foreach ($c in $d.Comments) { $out += ("comment '{0}' on '{1}'" -f $c.Range.Text, $c.Scope.Text) }
  return ($out -join ' | ')
}

$failures = 0
try {
  $names = Get-ChildItem -Path $Folder -Filter '*-src.docx' | ForEach-Object { $_.Name -replace '-src\.docx$', '' }
  foreach ($n in $names) {
    $src = Open-Doc "$n-src.docx"
    $before = $src.Revisions.Count
    $src.Close(0)
    $saved = Open-Doc "$n-saved.docx"
    $kept = $saved.Revisions.Count
    $comments = $saved.Comments.Count
    $saved.Close(0)
    $ok = $kept -eq $before
    if (-not $ok) { $failures++ }
    Write-Output ("{0,-10} saved unchanged: revisions {1} -> {2} comments={3} {4}" -f $n, $before, $kept, $comments, $(if ($ok) { 'OK' } else { 'DIFF' }))

    foreach ($mode in 'accept', 'reject') {
      $w = Open-Doc "$n-src.docx"
      if ($mode -eq 'accept') { $w.Revisions.AcceptAll() } else { $w.Revisions.RejectAll() }
      $expected = Signature $w
      $w.Close(0)
      $o = Open-Doc "$n-$mode.docx"
      $actual = Signature $o
      $left = $o.Revisions.Count
      $extra = ''
      if ($o.Comments.Count -gt 0) {
        $c = $o.Comments.Item(1)
        $extra = (" comments={0} author={1} text='{2}' on='{3}'" -f $o.Comments.Count, $c.Author, $c.Range.Text, $c.Scope.Text)
      }
      if ($o.ContentControls.Count -gt 0) { $extra += (" contentControls={0} tag={1}" -f $o.ContentControls.Count, $o.ContentControls.Item(1).Tag) }
      $o.Close(0)
      $same = ($expected -eq $actual) -and ($left -eq 0)
      if (-not $same) { $failures++ }
      Write-Output ("{0,-10} {1}: revisions left={2} {3}{4}" -f $n, $mode, $left, $(if ($same) { 'SAME AS WORD' } else { 'DIFF' }), $extra)
      if (-not $same) {
        Write-Output "    word:   $expected"
        Write-Output "    editor: $actual"
      }
    }
  }
}
finally {
  $word.Quit(0)
  [void][Runtime.InteropServices.Marshal]::ReleaseComObject($word)
}
Write-Output ("failures: {0}" -f $failures)
