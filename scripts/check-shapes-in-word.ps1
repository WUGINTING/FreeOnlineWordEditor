<#
  Checks in Microsoft Word that editing text box text in the web editor changed only that text:
  opens each original shape fixture and its edited copy (hidden, read-only) and compares what Word
  reports for every shape (the body's, the headers', and those on drawing canvases): kind, preset,
  position and what it is placed from, size, wrapping, fill and line, connectors and what their
  ends are attached to, and the text. The text of the edited text boxes must be the new one; all
  else must be the same.

  Usage (in the repository folder):
    1. DOCX_EXPORT=<out> npx vitest run tests/papyrus/export-shapes.test.ts
    2. pwsh scripts/check-shapes-in-word.ps1 -Folder <out>
  Only closes the Word this script started.
#>
param([Parameter(Mandatory = $true)][string]$Folder)
$ErrorActionPreference = 'Stop'
[Console]::OutputEncoding = [Text.Encoding]::UTF8
$pairs = Get-Content -Raw -Encoding UTF8 (Join-Path $Folder 'shapes-index.json') | ConvertFrom-Json

function Text($s) {
  try { if ($s.TextFrame.HasText) { return $s.TextFrame.TextRange.Text.TrimEnd("`r", "`n", [char]7) } } catch {}
  return ''
}

function Describe-Shape($s, [string]$where) {
  $r = [ordered]@{
    where = $where; name = $s.Name; type = $s.Type; auto = $s.AutoShapeType
    left = [math]::Round($s.Left, 1); top = [math]::Round($s.Top, 1); width = [math]::Round($s.Width, 1); height = [math]::Round($s.Height, 1)
    text = (Text $s)
  }
  try { $r.relH = $s.RelativeHorizontalPosition; $r.relV = $s.RelativeVerticalPosition; $r.wrap = $s.WrapFormat.Type } catch {}
  try { $r.fill = '{0}/{1}' -f $s.Fill.Visible, $s.Fill.ForeColor.RGB; $r.line = '{0}/{1}/{2}' -f $s.Line.Visible, $s.Line.ForeColor.RGB, $s.Line.Weight } catch {}
  try {
    if ($s.Connector) {
      $c = $s.ConnectorFormat
      $r.from = if ($c.BeginConnected) { '{0}#{1}' -f $c.BeginConnectedShape.Name, $c.BeginConnectionSite } else { '-' }
      $r.to = if ($c.EndConnected) { '{0}#{1}' -f $c.EndConnectedShape.Name, $c.EndConnectionSite } else { '-' }
    }
  } catch {}
  return [pscustomobject]$r
}

function Measure-Shapes($doc) {
  $out = @()
  foreach ($s in $doc.Shapes) {
    $out += Describe-Shape $s 'body'
    if ($s.Type -eq 20) { foreach ($i in $s.CanvasItems) { $out += Describe-Shape $i "canvas:$($s.Name)" } }
  }
  # A header's Shapes lists the shapes of every header and footer of the document.
  foreach ($s in $doc.Sections.Item(1).Headers.Item(1).Shapes) { $out += Describe-Shape $s 'headers' }
  return $out
}

$before = @(Get-Process WINWORD -ErrorAction SilentlyContinue | ForEach-Object Id)
$word = New-Object -ComObject Word.Application
$word.Visible = $false
$word.DisplayAlerts = 0
$mine = @(Get-Process WINWORD -ErrorAction SilentlyContinue | Where-Object { $before -notcontains $_.Id } | ForEach-Object Id)
$failures = 0
try {
  foreach ($p in $pairs) {
    $a = $word.Documents.Open([IO.Path]::GetFullPath($p.original), $false, $true, $false)
    $b = $word.Documents.Open([IO.Path]::GetFullPath($p.copy), $false, $true, $false)
    $sa = @(Measure-Shapes $a)
    $sb = @(Measure-Shapes $b)
    $pagesA = $a.ComputeStatistics(2); $pagesB = $b.ComputeStatistics(2)
    $a.Close(0); $b.Close(0)
    $diffs = @()
    if ($sa.Count -ne $sb.Count) { $diffs += "shape count $($sa.Count) -> $($sb.Count)" }
    if ($pagesA -ne $pagesB) { $diffs += "pages $pagesA -> $pagesB" }
    $texts = @{}
    foreach ($prop in $p.texts.PSObject.Properties) { $texts[$prop.Name] = $prop.Value }
    $edited = 0
    for ($i = 0; $i -lt [math]::Min($sa.Count, $sb.Count); $i++) {
      $x = $sa[$i]; $y = $sb[$i]
      foreach ($k in $x.PSObject.Properties.Name) {
        if ($k -eq 'text') {
          $want = if ($texts.ContainsKey($x.text)) { $texts[$x.text] } else { $x.text }
          if ($y.text -ne $want) { $diffs += "$($x.where) $($x.name): text '$($x.text)' -> '$($y.text)' (expected '$want')" }
          elseif ($want -ne $x.text) { $edited++ }
        } elseif ("$($x.$k)" -ne "$($y.$k)") { $diffs += "$($x.where) $($x.name): $k $($x.$k) -> $($y.$k)" }
      }
    }
    if ($edited -ne $texts.Count) { $diffs += "edited text boxes found: $edited of $($texts.Count)" }
    $name = Split-Path $p.original -Leaf
    $connectors = @($sb | Where-Object { $_.PSObject.Properties.Name -contains 'from' -and $_.from -ne '-' }).Count
    if ($diffs.Count) { $failures++ }
    Write-Output ("{0,-4} {1}: {2} shapes, {3} attached connectors, {4} edited text boxes{5}" -f $(if ($diffs.Count) { 'DIFF' } else { 'SAME' }), $name, $sb.Count, $connectors, $edited, $(if ($diffs.Count) { "`n  " + ($diffs -join "`n  ") } else { '' }))
  }
  # ----- 8B: what the editor changed must be what Word shows (DrawingML copy, then the VML copy) -----
  $cases = @()
  $eightB = Join-Path $Folder 'shapes-8b.json'
  if (Test-Path $eightB) { $cases = Get-Content -Raw -Encoding UTF8 $eightB | ConvertFrom-Json }
  foreach ($c in $cases) {
    foreach ($file in @($c.copy, $c.vml)) {
      $d = $word.Documents.Open([IO.Path]::GetFullPath($file), $false, $true, $false)
      $diffs = @()
      $body = @($d.Shapes | ForEach-Object { $_ })
      if ($body.Count -ne $c.expect.count) { $diffs += "shape count $($body.Count), expected $($c.expect.count)" }
      $near = { param($a, $b) [math]::Abs([double]$a - [double]$b) -le 0.8 }
      foreach ($e in $c.expect.shapes) {
        if ($e.group) {
          $g = $body | Where-Object { $_.Type -eq 6 } | Where-Object { (@($_.GroupItems | ForEach-Object { $_.Name }) -join '|') -eq ($e.group -join '|') }
          if (-not $g) { $diffs += "no group of $($e.group -join ', ')" }
          continue
        }
        $s = $body | Where-Object { $_.Name -eq $e.name } | Select-Object -First 1
        if (-not $s) { $diffs += "missing $($e.name)"; continue }
        foreach ($k in 'left', 'top', 'width', 'height') {
          if ($null -eq $e.$k) { continue }
          $v = switch ($k) { 'left' { $s.Left } 'top' { $s.Top } 'width' { $s.Width } 'height' { $s.Height } }
          if (-not (& $near $v $e.$k)) { $diffs += "$($e.name): $k $([math]::Round($v, 1)), expected $($e.$k)" }
        }
        if ($null -ne $e.wrap -and $s.WrapFormat.Type -ne $e.wrap) { $diffs += "$($e.name): wrap $($s.WrapFormat.Type), expected $($e.wrap)" }
        if ($null -ne $e.text -and (Text $s) -ne $e.text) { $diffs += "$($e.name): text '$(Text $s)', expected '$($e.text)'" }
        if ($null -ne $e.anchor) {
          $at = $s.Anchor.Paragraphs.Item(1).Range.Text.Trim()
          if (-not $at.StartsWith($e.anchor)) { $diffs += "$($e.name): anchored to '$at', expected '$($e.anchor)'" }
        }
      }
      foreach ($n in $c.expect.absent) { if ($body | Where-Object { $_.Name -eq $n }) { $diffs += "$n should be gone" } }
      foreach ($pair in $c.expect.above) {
        $a = $body | Where-Object { $_.Name -eq $pair[0] } | Select-Object -First 1
        $b = $body | Where-Object { $_.Name -eq $pair[1] } | Select-Object -First 1
        if ($a -and $b -and $a.ZOrderPosition -le $b.ZOrderPosition) { $diffs += "$($pair[0]) (z $($a.ZOrderPosition)) is not above $($pair[1]) (z $($b.ZOrderPosition))" }
      }
      $attached = 0
      foreach ($s in $body) {
        if ($s.Type -ne 20) { continue }
        foreach ($i in $s.CanvasItems) { if ($i.Connector -and $i.ConnectorFormat.BeginConnected -and $i.ConnectorFormat.EndConnected) { $attached++ } }
      }
      if ($attached -lt $c.expect.attached -and $file -eq $c.copy) { $diffs += "attached connectors $attached, expected $($c.expect.attached)" }
      $d.Close(0)
      if ($diffs.Count) { $failures++ }
      Write-Output ("{0,-4} {1}: {2} shapes{3}" -f $(if ($diffs.Count) { 'DIFF' } else { 'SAME' }), (Split-Path $file -Leaf), $body.Count, $(if ($diffs.Count) { "`n  " + ($diffs -join "`n  ") } else { '' }))
    }
  }
} finally {
  $word.Quit(0)
  [void][Runtime.InteropServices.Marshal]::ReleaseComObject($word)
  foreach ($id in $mine) {
    $pr = Get-Process -Id $id -ErrorAction SilentlyContinue
    if ($pr -and -not $pr.WaitForExit(15000)) { $pr | Stop-Process -Force }
  }
}
$total = $pairs.Count + 2 * $cases.Count
Write-Output ("`n{0} / {1} documents as the editor left them, in Word" -f ($total - $failures), $total)
