<#
  Opens documents in Microsoft Word (hidden, read-only, never saved) and reports, per document:
  the pages, the file title and custom properties (docProps), and each table of contents: its
  field code and its entries (style, text, page number) as the file has them, then again after
  Word's own 「更新整個目錄」 (TablesOfContents.Update, as F9 does). A table of contents made or
  rebuilt by the online editor should list the same entries before and after. Closes only the
  Word process it started.

  Usage:
    pwsh scripts/inspect-toc-in-word.ps1 -Docx a.docx,b.docx [-Json out.json]
#>
param([Parameter(Mandatory = $true)][string[]]$Docx, [string]$Json)
$ErrorActionPreference = 'Stop'
$before = @(Get-Process WINWORD -ErrorAction SilentlyContinue | ForEach-Object Id)
$word = New-Object -ComObject Word.Application
$word.Visible = $false
$word.DisplayAlerts = 0
$mine = @(Get-Process WINWORD -ErrorAction SilentlyContinue | Where-Object { $before -notcontains $_.Id } | ForEach-Object Id)
Write-Output ("Word processes before: [{0}]  started: [{1}]" -f ($before -join ','), ($mine -join ','))

# Document properties are late-bound COM objects PowerShell can't read directly.
function Prop($obj, [string]$name) {
  [System.__ComObject].InvokeMember($name, [System.Reflection.BindingFlags]::GetProperty, $null, $obj, $null)
}

function Entries($toc) {
  $out = @()
  foreach ($p in $toc.Range.Paragraphs) {
    $t = ($p.Range.Text -replace "[`r`a]", '')
    if (-not $t.Trim()) { continue }
    $cut = $t.LastIndexOf("`t")
    $text = if ($cut -ge 0) { $t.Substring(0, $cut) } else { $t }
    $page = if ($cut -ge 0) { $t.Substring($cut + 1) } else { '' }
    $out += [ordered]@{ style = $p.Style.NameLocal; text = $text; page = $page }
  }
  , $out
}

$results = @()
try {
  foreach ($path in $Docx) {
    $full = (Resolve-Path $path).Path
    $d = $word.Documents.Open($full, $false, $true, $false)
    try {
      $d.Repaginate()
      $r = [ordered]@{ file = $full; pages = $d.ComputeStatistics(2) }
      try {
        $bi = $d.BuiltInDocumentProperties
        $t = [System.__ComObject].InvokeMember('Item', [System.Reflection.BindingFlags]::GetProperty, $null, $bi, @('Title'))
        $r.title = Prop $t 'Value'
      } catch { $r.title = $null }
      $custom = [ordered]@{}
      $cps = $d.CustomDocumentProperties
      $n = Prop $cps 'Count'
      for ($k = 1; $k -le $n; $k++) {
        $cp = [System.__ComObject].InvokeMember('Item', [System.Reflection.BindingFlags]::GetProperty, $null, $cps, @($k))
        $custom[[string](Prop $cp 'Name')] = [string](Prop $cp 'Value')
      }
      $r.custom = $custom
      $tocs = @()
      for ($i = 1; $i -le $d.TablesOfContents.Count; $i++) {
        $toc = $d.TablesOfContents.Item($i)
        $code = ''
        foreach ($f in $toc.Range.Fields) { if ($f.Type -eq 13) { $code = $f.Code.Text.Trim(); break } } # wdFieldTOC
        $asFile = Entries $toc
        $toc.Update()
        $d.Repaginate()
        $asWord = Entries $d.TablesOfContents.Item($i)
        $same = ($asFile.Count -eq $asWord.Count)
        if ($same) {
          for ($k = 0; $k -lt $asFile.Count; $k++) {
            if ($asFile[$k].text -ne $asWord[$k].text -or $asFile[$k].style -ne $asWord[$k].style) { $same = $false }
          }
        }
        $tocs += [ordered]@{ code = $code; sameEntriesAfterUpdate = $same; entries = $asFile; afterUpdate = $asWord }
        Write-Output ("{0}: TOC {1} '{2}' entries={3} afterF9={4} sameEntries={5}" -f (Split-Path $full -Leaf), $i, $code, $asFile.Count, $asWord.Count, $same)
        for ($k = 0; $k -lt [Math]::Max($asFile.Count, $asWord.Count); $k++) {
          $a = if ($k -lt $asFile.Count) { "{0} | {1} | {2}" -f $asFile[$k].style, $asFile[$k].text, $asFile[$k].page } else { '-' }
          $b = if ($k -lt $asWord.Count) { "{0} | {1} | {2}" -f $asWord[$k].style, $asWord[$k].text, $asWord[$k].page } else { '-' }
          $mark = if ($a -eq $b) { ' ' } else { '*' }
          Write-Output ("  {0} {1,-60} || {2}" -f $mark, $a, $b)
        }
      }
      $r.tocs = $tocs
      # Tables of figures (圖目錄 / 表目錄, TOC \c): Word's F9 should list the same captions.
      $tofs = @()
      for ($i = 1; $i -le $d.TablesOfFigures.Count; $i++) {
        $tof = $d.TablesOfFigures.Item($i)
        $asFile = Entries $tof
        $tof.Update()
        $d.Repaginate()
        $asWord = Entries $d.TablesOfFigures.Item($i)
        $same = ($asFile.Count -eq $asWord.Count)
        if ($same) { for ($k = 0; $k -lt $asFile.Count; $k++) { if ($asFile[$k].text -ne $asWord[$k].text) { $same = $false } } }
        $tofs += [ordered]@{ sameEntriesAfterUpdate = $same; entries = $asFile; afterUpdate = $asWord }
        Write-Output ("{0}: table of figures {1} entries={2} afterF9={3} sameEntries={4}" -f (Split-Path $full -Leaf), $i, $asFile.Count, $asWord.Count, $same)
        for ($k = 0; $k -lt [Math]::Max($asFile.Count, $asWord.Count); $k++) {
          $a = if ($k -lt $asFile.Count) { "{0} | {1}" -f $asFile[$k].text, $asFile[$k].page } else { '-' }
          $b = if ($k -lt $asWord.Count) { "{0} | {1}" -f $asWord[$k].text, $asWord[$k].page } else { '-' }
          Write-Output ("    {0,-40} || {1}" -f $a, $b)
        }
      }
      $r.tablesOfFigures = $tofs
      Write-Output ("{0}: pages={1} title='{2}' custom={3}" -f (Split-Path $full -Leaf), $r.pages, $r.title, (($custom.GetEnumerator() | ForEach-Object { "$($_.Key)=$($_.Value)" }) -join '; '))
      $results += $r
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
  $after = @(Get-Process WINWORD -ErrorAction SilentlyContinue | ForEach-Object Id)
  Write-Output ("Word processes after: [{0}]" -f ($after -join ','))
}
if ($Json) { $results | ConvertTo-Json -Depth 6 | Set-Content -Encoding UTF8 $Json }
