<#
  Checks the documents written by tests/papyrus/export-comments.test.ts in Microsoft Word
  (hidden): for every case, the comments Word lists must be exactly those in
  <case>.expected.json — count, author, text (Comment.Range.Text), the commented text
  (Comment.Scope.Text), resolved (Comment.Done, of the thread) and the thread a reply belongs to
  (Comment.Ancestor) — and every top-level comment's Replies.Count must match.

  Usage:
    1. DOCX_EXPORT=<out> [WORD_COMMENTS_DOCX=<file>] npx vitest run tests/papyrus/export-comments.test.ts
    2. powershell -ExecutionPolicy Bypass -File scripts/inspect-comments-in-word.ps1 -Folder <out>
#>
param([Parameter(Mandatory = $true)][string]$Folder)

$ErrorActionPreference = 'Stop'
$word = New-Object -ComObject Word.Application
$word.Visible = $false
$word.DisplayAlerts = 0

function Sig($author, $text, $scope, $done, $parent) {
  $p = if ($null -eq $parent) { '-' } else { "'$parent'" }
  return ("{0} | '{1}' | on '{2}' | done={3} | reply to {4}" -f $author, $text, $scope, [bool]$done, $p)
}

$failures = 0
try {
  foreach ($json in Get-ChildItem -Path $Folder -Filter '*.expected.json' | Sort-Object Name) {
    $name = $json.Name -replace '\.expected\.json$', ''
    # Windows PowerShell 5 passes a JSON array down the pipeline as one object: unroll it.
    $parsed = Get-Content -Raw -Encoding UTF8 $json.FullName | ConvertFrom-Json
    $want = @($parsed | ForEach-Object { $_ })
    $expected = @($want | ForEach-Object { Sig $_.author $_.text $_.scope $_.done $_.parent }) | Sort-Object
    $wantReplies = @{}
    foreach ($w in $want) { if ($null -ne $w.parent) { $wantReplies[$w.parent] = 1 + [int]$wantReplies[$w.parent] } }

    # FileName, ConfirmConversions, ReadOnly, AddToRecentFiles
    $d = $word.Documents.Open((Join-Path $Folder "$name.docx"), $false, $true, $false)
    $actual = @()
    $replies = @()
    foreach ($c in $d.Comments) {
      $ancestor = $c.Ancestor
      $parent = if ($null -eq $ancestor) { $null } else { $ancestor.Range.Text }
      $actual += Sig $c.Author $c.Range.Text $c.Scope.Text $c.Done $parent
      if ($null -eq $ancestor) {
        $n = [int]$wantReplies[$c.Range.Text]
        if ($c.Replies.Count -ne $n) { $replies += ("'{0}' has {1} replies, expected {2}" -f $c.Range.Text, $c.Replies.Count, $n) }
      }
    }
    $count = $d.Comments.Count
    $d.Close([ref]0)
    $actual = @($actual | Sort-Object)

    $same = ($count -eq $want.Count) -and ((@($expected) -join "`n") -eq (@($actual) -join "`n")) -and ($replies.Count -eq 0)
    if (-not $same) { $failures++ }
    Write-Output ("{0,-22} Comments.Count={1} (expected {2}) {3}" -f $name, $count, $want.Count, $(if ($same) { 'SAME' } else { 'DIFF' }))
    foreach ($a in $actual) { Write-Output "    word:   $a" }
    if (-not $same) {
      foreach ($e in $expected) { Write-Output "    wanted: $e" }
      foreach ($r in $replies) { Write-Output "    $r" }
    }
  }
}
finally {
  $word.Quit([ref]0)
  [void][Runtime.InteropServices.Marshal]::ReleaseComObject($word)
}
Write-Output ("failures: {0}" -f $failures)
