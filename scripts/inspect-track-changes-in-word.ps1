<#
  Checks the documents written by tests/papyrus/export-track.test.ts (edited in the web editor
  with 追蹤修訂 on) in Microsoft Word (hidden): Document.TrackRevisions, every revision Word sees
  (type, author, date, text) next to what the editor recorded, and Word's own Accept All /
  Reject All against the editor's.

  Usage:
    0. (optional) a document Word itself saved with revisions:
         powershell -File scripts/inspect-track-changes-in-word.ps1 -MakeSource <file.docx>
    1. DOCX_EXPORT=<out> [WORD_REVISIONS_DOCX=<file.docx>] npx vitest run tests/papyrus/export-track.test.ts
    2. powershell -File scripts/inspect-track-changes-in-word.ps1 -Folder <out>
  Saved with a UTF-8 BOM so Windows PowerShell 5 reads the Chinese text correctly.
#>
param([string]$Folder, [string]$MakeSource)

$ErrorActionPreference = 'Stop'
$word = New-Object -ComObject Word.Application
$word.Visible = $false
$word.DisplayAlerts = 0
$types = @{ 1 = 'ins'; 2 = 'del'; 3 = 'format'; 4 = 'paraNumber'; 8 = 'style'; 10 = 'paraFormat'; 14 = 'moveFrom'; 15 = 'moveTo' }

function Esc([string]$s) {
  return ($s -replace "`r", '\r' -replace "`n", '\n' -replace "`t", '\t' -replace [char]12, '\f' -replace [char]11, '\v')
}

# The text of every paragraph (control characters such as page breaks and pictures left out), joined by |.
function ParaText($d) {
  $out = @()
  foreach ($p in $d.Paragraphs) { $out += ($p.Range.Text.TrimEnd("`r") -replace '[\x00-\x08\x0b\x0c\x0e-\x1f]', '') }
  return ($out -join '|')
}

function Open-Doc([string]$path) {
  # FileName, ConfirmConversions, ReadOnly, AddToRecentFiles
  return $word.Documents.Open($path, $false, $true, $false)
}

$failures = 0
try {
  if ($MakeSource) {
    $d = $word.Documents.Add()
    $d.Content.Text = "Word 原有的第一段文字內容。`rSecond paragraph written in Word."
    $d.TrackRevisions = $true
    $d.Paragraphs(1).Range.Characters(1).InsertBefore('【Word 插入】')
    $d.Paragraphs(2).Range.Words(1).Delete() | Out-Null
    $d.SaveAs2($MakeSource, 16)
    Write-Output ("made {0}: revisions={1} author={2} trackRevisions={3}" -f $MakeSource, $d.Revisions.Count, $d.Revisions.Item(1).Author, $d.TrackRevisions)
    $d.Close(0)
  }
  if ($Folder) {
    foreach ($f in Get-ChildItem -Path $Folder -Filter '*.docx' | Sort-Object Name) {
      $name = $f.BaseName
      $jsonPath = Join-Path $Folder "$name.json"
      if (-not (Test-Path $jsonPath)) { continue }
      $ours = Get-Content $jsonPath -Raw -Encoding UTF8 | ConvertFrom-Json
      $d = Open-Doc $f.FullName
      $count = $d.Revisions.Count
      $track = [bool]$d.TrackRevisions
      Write-Output ("== {0}: Word TrackRevisions={1} (editor {2}); revisions Word={3} editor={4}" -f $name, $track, $ours.trackRevisions, $count, @($ours.revisions).Count)
      if ($track -ne [bool]$ours.trackRevisions) { $failures++; Write-Output '   DIFF: TrackRevisions' }
      foreach ($r in $d.Revisions) {
        $t = $types[[int]$r.Type]; if (-not $t) { $t = "type$($r.Type)" }
        Write-Output ("   word:   {0,-10} author={1} date={2} text='{3}'" -f $t, $r.Author, ([datetime]$r.Date).ToString('yyyy-MM-dd HH:mm'), (Esc $r.Range.Text))
      }
      foreach ($r in @($ours.revisions)) {
        $date = if ($r.date) { $r.date.Substring(0, 16).Replace('T', ' ') } else { '' }
        Write-Output ("   editor: {0,-10} author={1} date={2} text='{3}'" -f $r.kind, $r.author, $date, (Esc $r.text))
      }
      # Every revision Word sees is by an author and at a time the editor recorded.
      foreach ($r in $d.Revisions) {
        $wd = ([datetime]$r.Date).ToString('yyyy-MM-dd HH:mm')
        $hit = @($ours.revisions) | Where-Object { $_.author -eq $r.Author -and $_.date -and $_.date.Substring(0, 16).Replace('T', ' ') -eq $wd }
        if (-not $hit) { $failures++; Write-Output ("   DIFF: no editor revision by {0} at {1}" -f $r.Author, $wd) }
      }
      $d.Revisions.AcceptAll()
      $acc = ParaText $d
      $d.Close(0)
      $d = Open-Doc $f.FullName
      $d.Revisions.RejectAll()
      $rej = ParaText $d
      $d.Close(0)
      foreach ($pair in @(@('accept all', $acc, $ours.accepted), @('reject all', $rej, $ours.rejected))) {
        $same = $pair[1] -eq $pair[2]
        if (-not $same) { $failures++ }
        Write-Output ("   {0}: {1}  word='{2}'{3}" -f $pair[0], $(if ($same) { 'SAME AS WORD' } else { 'DIFF' }), $pair[1], $(if ($same) { '' } else { " editor='$($pair[2])'" }))
      }
    }
  }
}
finally {
  $word.Quit()
  [void][Runtime.InteropServices.Marshal]::ReleaseComObject($word)
}
Write-Output ("failures: {0}" -f $failures)
