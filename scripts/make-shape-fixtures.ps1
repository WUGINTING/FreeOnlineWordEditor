<#
  產生含流程圖、文字方塊與圖案的測試文件（線上編輯器的文字方塊／圖案顯示與編輯，phase 8A）。
  用桌面版 Microsoft Word（COM）建立並存成 .docx，由 Word 自己寫出 wps／VML（mc:AlternateContent）。
  可重複執行；結果放在 tests/fixtures/shapes（前端測試 tests/papyrus/shapes*.test.ts 會讀）。

  文件：
    flowchart.docx   公文流程圖：開始／結束（terminator）、處理（process）、判斷（decision）、
                     資料（inputOutput）、文件（document）、預設處理（predefinedProcess），
                     之間以直線與肘形接點連接線（有箭頭、接在圖案上）相連；圖案裡有文字。
    textboxes.docx   文字方塊：文繞圖（矩形，相對段落）、多段落（粗體、置中、項目符號）、
                     與文字列同列（inline）、頁首裡的文字方塊（相對頁面）、兩個有文字的圖案、
                     有替代文字的文字方塊。
    shapes.docx      圖案：矩形、圓角矩形、橢圓、菱形、平行四邊形、右箭號、左右箭號、虛線、
                     文字在後（behind）、上及下（topAndBottom）、繪圖畫布（相對頁面）、
                     五角星（網頁上畫不出來：保留預覽圖）；有替代文字（描述、標題）。

  用法（在專案資料夾，PowerShell 7）：
    pwsh -File scripts\make-shape-fixtures.ps1
    pwsh -File scripts\make-shape-fixtures.ps1 -OutDir <資料夾>
  重新產生的檔案位元組會不同（Word 的 rsid 等）。只關閉這個腳本自己啟動的 Word。
#>
param([string]$OutDir = (Join-Path $PSScriptRoot '..\tests\fixtures\shapes'))
$ErrorActionPreference = 'Stop'
$OutDir = [IO.Path]::GetFullPath($OutDir)
New-Item -ItemType Directory -Force -Path $OutDir | Out-Null

$wdFormatXMLDocument = 12
# MsoAutoShapeType
$Rect = 1; $Parallelogram = 2; $Diamond = 4; $RoundRect = 5; $Oval = 9; $RightArrow = 33; $LeftRightArrow = 37
$Star5 = 92; $FcProcess = 61; $FcDecision = 63; $FcData = 64; $FcPredefined = 65; $FcDocument = 67; $FcTerminator = 69
# MsoConnectorType, MsoArrowheadStyle
$ConnStraight = 1; $ConnElbow = 2; $ArrowTriangle = 2
# WdWrapType
$WrapSquare = 0; $WrapNone = 3; $WrapTopBottom = 4; $WrapBehind = 5; $WrapInline = 7
# WdRelativeHorizontalPosition / WdRelativeVerticalPosition
$HMargin = 0; $HPage = 1; $HColumn = 2; $VMargin = 0; $VPage = 1; $VParagraph = 2
$FontName = '標楷體'

function Rgb([int]$r, [int]$g, [int]$b) { $r + 256 * $g + 65536 * $b }

function New-Doc($word, [string[]]$lines) {
  $doc = $word.Documents.Add()
  $doc.SetCompatibilityMode(15)
  $normal = $doc.Styles.Item(-1)
  $normal.Font.Name = $FontName
  $normal.Font.NameFarEast = $FontName
  $normal.Font.Size = 12
  $normal.ParagraphFormat.SpaceAfter = 0
  $doc.Content.Text = ($lines -join "`r")
  return $doc
}

function Set-Text($shape, [string]$text, [int]$size = 12, [bool]$center = $true) {
  $tr = $shape.TextFrame.TextRange
  $tr.Text = $text
  $tr.Font.Name = $FontName
  $tr.Font.NameFarEast = $FontName
  $tr.Font.Size = $size
  $tr.Font.Color = 0
  if ($center) { $tr.ParagraphFormat.Alignment = 1 }
}

function Style-Shape($shape, [int]$fill, [int]$line, [double]$weight = 1.25) {
  $shape.Fill.Visible = -1
  $shape.Fill.Solid()
  $shape.Fill.ForeColor.RGB = $fill
  $shape.Line.Visible = -1
  $shape.Line.ForeColor.RGB = $line
  $shape.Line.Weight = $weight
}

# Floating, in front of text, left/top from the column / anchor paragraph.
function Float($shape, [double]$left, [double]$top) {
  $shape.WrapFormat.Type = $WrapNone
  $shape.RelativeHorizontalPosition = $HColumn
  $shape.RelativeVerticalPosition = $VParagraph
  $shape.Left = $left
  $shape.Top = $top
}

# Word's COM ConnectorFormat.BeginConnect / EndConnect fail with "type mismatch" from script (PowerShell,
# VBScript and JScript alike), so connectors are drawn from site to site here and attached afterwards in the
# saved XML (a:stCxn / a:endCxn, see Connect-InXml); Word then keeps them attached when the shapes move.
$script:Connections = @()

# The point of a connection site: Word's flowchart shapes number them 1 top, 2 left, 3 bottom, 4 right.
function Site($it, [int]$site) {
  switch ($site) {
    1 { @(($it.X + $it.W / 2), $it.Y) }
    2 { @($it.X, ($it.Y + $it.H / 2)) }
    3 { @(($it.X + $it.W / 2), ($it.Y + $it.H)) }
    4 { @(($it.X + $it.W), ($it.Y + $it.H / 2)) }
  }
}

# $idxB: the end's DrawingML site when it isn't the 4-site one (an oval has 8, from the top anticlockwise: 2 is the left).
function Connect($canvas, $a, $b, [int]$type, [int]$siteA, [int]$siteB, [int]$idxB = -1) {
  $p = Site $a $siteA
  $q = Site $b $siteB
  $c = $canvas.CanvasItems.AddConnector($type, $p[0], $p[1], $q[0], $q[1])
  $c.Name = "連接線$($script:Connections.Count + 1)"
  $c.Line.ForeColor.RGB = (Rgb 64 64 64)
  $c.Line.Weight = 1.25
  $c.Line.EndArrowheadStyle = $ArrowTriangle
  # DrawingML numbers the sites from 0 in the same order.
  $script:Connections += @{ Name = $c.Name; From = $a.Shape.Name; FromIdx = $siteA - 1; To = $b.Shape.Name; ToIdx = $(if ($idxB -ge 0) { $idxB } else { $siteB - 1 }) }
  return $c
}

# Attach the connectors recorded by Connect in a saved document: <a:stCxn>/<a:endCxn> in their wps:cNvCnPr.
function Connect-InXml([string]$path) {
  if (-not $script:Connections.Count) { return }
  Add-Type -AssemblyName System.IO.Compression, System.IO.Compression.FileSystem
  $zip = [IO.Compression.ZipFile]::Open($path, 'Update')
  try {
    $entry = $zip.GetEntry('word/document.xml')
    $reader = New-Object IO.StreamReader($entry.Open(), [Text.Encoding]::UTF8)
    $xml = $reader.ReadToEnd(); $reader.Close()
    $id = { param($name) $m = [regex]::Match($xml, '<wps:cNvPr id="(\d+)" name="' + [regex]::Escape($name) + '"'); if (-not $m.Success) { throw "shape $name not found" }; $m.Groups[1].Value }
    foreach ($c in $script:Connections) {
      $from = & $id $c.From
      $to = & $id $c.To
      $cxn = '<a:stCxn id="' + $from + '" idx="' + $c.FromIdx + '"/><a:endCxn id="' + $to + '" idx="' + $c.ToIdx + '"/>'
      $pattern = '(<wps:cNvPr id="\d+" name="' + [regex]::Escape($c.Name) + '"(?:/>|>.*?</wps:cNvPr>))<wps:cNvCnPr(?:/>|>(.*?)</wps:cNvCnPr>)'
      $m = [regex]::Match($xml, $pattern)
      if (-not $m.Success) { throw "connector $($c.Name) not found" }
      $xml = $xml.Remove($m.Index, $m.Length).Insert($m.Index, $m.Groups[1].Value + '<wps:cNvCnPr>' + $m.Groups[2].Value + $cxn + '</wps:cNvCnPr>')
    }
    $entry.Delete()
    $writer = New-Object IO.StreamWriter($zip.CreateEntry('word/document.xml').Open(), (New-Object Text.UTF8Encoding $false))
    $writer.Write($xml); $writer.Close()
  } finally { $zip.Dispose() }
}

function Build-Flowchart($word) {
  $lines = @('公文處理流程', '以下為收文到歸檔的流程（繪圖畫布，接點連接線接在圖案上）：', '') + (@('') * 22) + @('浮動圖案與箭號（沒有畫布）：', '') + (@('') * 8) + @('流程圖後的文字。')
  $doc = New-Doc $word $lines
  $blue = Rgb 222 235 247; $edge = Rgb 47 84 150
  # (A) Word keeps connectors attached to shapes only on a drawing canvas.
  $canvas = $doc.Shapes.AddCanvas(0, 0, 460, 420, $doc.Paragraphs.Item(3).Range)
  $canvas.WrapFormat.Type = $WrapTopBottom
  $canvas.RelativeHorizontalPosition = $HColumn
  $canvas.RelativeVerticalPosition = $VParagraph
  $canvas.Left = 0; $canvas.Top = 0
  $items = @(
    @{ T = $FcTerminator; X = 150; Y = 5; W = 140; H = 36; Text = '開始' }
    @{ T = $FcProcess; X = 150; Y = 65; W = 140; H = 40; Text = '收文登錄' }
    @{ T = $FcDecision; X = 140; Y = 130; W = 160; H = 70; Text = '需要會辦？' }
    @{ T = $FcPredefined; X = 340; Y = 142; W = 110; H = 46; Text = '會辦單位' }
    @{ T = $FcData; X = 150; Y = 230; W = 140; H = 40; Text = '發文' }
    @{ T = $FcDocument; X = 150; Y = 295; W = 140; H = 50; Text = '歸檔' }
    @{ T = $FcTerminator; X = 150; Y = 375; W = 140; H = 36; Text = '結束' }
  )
  $shapes = @()
  foreach ($it in $items) {
    $s = $canvas.CanvasItems.AddShape($it.T, $it.X, $it.Y, $it.W, $it.H)
    $s.Name = "流程$($shapes.Count + 1)"
    Style-Shape $s $blue $edge
    Set-Text $s $it.Text 12
    $it.Shape = $s
    $shapes += $it
  }
  # Connection sites of Word's flowchart shapes: 1 top, 2 left, 3 bottom, 4 right.
  Connect $canvas $shapes[0] $shapes[1] $ConnStraight 3 1 | Out-Null
  Connect $canvas $shapes[1] $shapes[2] $ConnStraight 3 1 | Out-Null
  Connect $canvas $shapes[2] $shapes[4] $ConnStraight 3 1 | Out-Null
  Connect $canvas $shapes[2] $shapes[3] $ConnStraight 4 2 | Out-Null
  Connect $canvas $shapes[3] $shapes[4] $ConnElbow 3 4 | Out-Null
  Connect $canvas $shapes[4] $shapes[5] $ConnStraight 3 1 | Out-Null
  Connect $canvas $shapes[5] $shapes[6] $ConnStraight 3 1 | Out-Null
  $shapes[2].Shape.AlternativeText = '判斷是否需要其他單位會辦'
  # (B) Floating shapes and arrows, positioned from the column and the anchor paragraph.
  $anchor = $doc.Paragraphs.Item(27).Range
  $p = $doc.Shapes.AddShape($FcProcess, 0, 0, 120, 40, $anchor)
  Float $p 0 0
  Style-Shape $p $blue $edge
  Set-Text $p '填寫申請單' 12
  $d = $doc.Shapes.AddShape($FcDecision, 0, 0, 140, 60, $anchor)
  Float $d 200 -10
  Style-Shape $d $blue $edge
  Set-Text $d '資料齊全？' 12
  $arrow = $doc.Shapes.AddLine(0, 0, 80, 0, $anchor)
  $arrow.WrapFormat.Type = $WrapNone
  $arrow.RelativeHorizontalPosition = $HColumn
  $arrow.RelativeVerticalPosition = $VParagraph
  $arrow.Left = 120; $arrow.Top = 20
  $arrow.Line.EndArrowheadStyle = $ArrowTriangle
  $arrow.Line.Weight = 1.5
  $elbow = $doc.Shapes.AddConnector($ConnElbow, 0, 0, 100, 60)
  $elbow.WrapFormat.Type = $WrapNone
  $elbow.RelativeHorizontalPosition = $HColumn
  $elbow.RelativeVerticalPosition = $VParagraph
  $elbow.Left = 270; $elbow.Top = 50; $elbow.Width = 100; $elbow.Height = 60
  $elbow.Line.EndArrowheadStyle = $ArrowTriangle
  return $doc
}

function Build-TextBoxes($word) {
  $long = '本段落的文字會繞著右邊的文字方塊排列。' * 6
  $doc = New-Doc $word @('文字方塊測試', $long, '第二段：下方有多段落的文字方塊。', '', '', '', '', '', '與文字同列的方塊：', '', '', '', '', '', '', '圖案後的文字。')
  # 1) Square wrap, relative to the paragraph, on the right of the column.
  $tb = $doc.Shapes.AddTextbox(1, 0, 0, 150, 70, $doc.Paragraphs.Item(2).Range)
  $tb.WrapFormat.Type = $WrapSquare
  $tb.RelativeHorizontalPosition = $HColumn
  $tb.RelativeVerticalPosition = $VParagraph
  $tb.Left = 280; $tb.Top = 0
  Set-Text $tb '繞圖文字方塊' 12 $false
  $tb.Line.ForeColor.RGB = (Rgb 192 0 0)
  # 2) Several paragraphs: a bold centred title and a bulleted list.
  $tb2 = $doc.Shapes.AddTextbox(1, 0, 0, 220, 100, $doc.Paragraphs.Item(3).Range)
  Float $tb2 20 24
  $r = $tb2.TextFrame.TextRange
  $r.Text = "會議重點`r第一項：預算`r第二項：人力"
  $r.Font.Name = $FontName; $r.Font.NameFarEast = $FontName; $r.Font.Size = 12
  $p1 = $r.Paragraphs.Item(1).Range
  $p1.Font.Bold = -1
  $p1.ParagraphFormat.Alignment = 1
  foreach ($i in 2, 3) { $r.Paragraphs.Item($i).Range.ListFormat.ApplyBulletDefault() }
  $tb2.Title = '會議重點'
  $tb2.AlternativeText = '列出兩項會議重點的文字方塊'
  # 3) Inline with the text.
  $tb3 = $doc.Shapes.AddTextbox(1, 0, 0, 120, 30, $doc.Paragraphs.Item(9).Range)
  Set-Text $tb3 '同列方塊' 12
  $tb3.WrapFormat.Type = $WrapInline
  # 4) Two shapes with text side by side (Word's COM refuses Group() from script: groups are tested with XML).
  $g1 = $doc.Shapes.AddShape($Rect, 0, 0, 100, 40, $doc.Paragraphs.Item(10).Range)
  Float $g1 20 10
  Style-Shape $g1 (Rgb 255 242 204) (Rgb 191 144 0)
  Set-Text $g1 '甲' 12
  $g2 = $doc.Shapes.AddShape($Oval, 0, 0, 100, 40, $doc.Paragraphs.Item(10).Range)
  Float $g2 160 10
  Style-Shape $g2 (Rgb 226 239 218) (Rgb 84 130 53)
  Set-Text $g2 '乙' 12
  # 5) A text box in the header, placed on the page.
  $header = $doc.Sections.Item(1).Headers.Item(1)
  $header.Range.Text = '頁首'
  $hb = $header.Shapes.AddTextbox(1, 0, 0, 200, 24, $header.Range)
  $hb.WrapFormat.Type = $WrapNone
  $hb.RelativeHorizontalPosition = $HPage
  $hb.RelativeVerticalPosition = $VPage
  $hb.Left = 330; $hb.Top = 30
  Set-Text $hb '頁首文字方塊' 10 $false
  return $doc
}

function Build-Shapes($word) {
  $doc = New-Doc $word @('圖案測試', '', '', '', '', '', '', '', '', '文字在後的圖案上方有這一段文字，文字要看得到。', '上及下的圖案：', '', '上及下圖案後的文字。', '畫布：', '', '', '', '', '', '結尾。')
  $anchor = $doc.Paragraphs.Item(2).Range
  $specs = @(
    @{ T = $Rect; X = 0; Y = 0; Text = '矩形' }
    @{ T = $RoundRect; X = 110; Y = 0; Text = '圓角' }
    @{ T = $Oval; X = 220; Y = 0; Text = '橢圓' }
    @{ T = $Diamond; X = 330; Y = 0; Text = '菱形' }
    @{ T = $Parallelogram; X = 0; Y = 70; Text = '平行' }
    @{ T = $RightArrow; X = 110; Y = 70; Text = '' }
    @{ T = $LeftRightArrow; X = 220; Y = 70; Text = '' }
    @{ T = $Star5; X = 330; Y = 70; Text = '' }
  )
  foreach ($sp in $specs) {
    $s = $doc.Shapes.AddShape($sp.T, 0, 0, 90, 50, $anchor)
    Float $s $sp.X $sp.Y
    Style-Shape $s (Rgb 255 255 255) (Rgb 0 112 192) 1.5
    if ($sp.Text) { Set-Text $s $sp.Text 11 }
  }
  $line = $doc.Shapes.AddLine(0, 0, 200, 0, $anchor)
  $line.WrapFormat.Type = $WrapNone
  $line.RelativeHorizontalPosition = $HColumn
  $line.RelativeVerticalPosition = $VParagraph
  $line.Left = 0; $line.Top = 140
  $line.Line.ForeColor.RGB = (Rgb 255 0 0)
  $line.Line.Weight = 2
  $line.Line.DashStyle = 4
  # (Word's COM does not turn shapes from script — Rotation stays 0 —: turned shapes are tested with XML.)
  $alt = $doc.Shapes.AddShape($Rect, 0, 0, 80, 30, $anchor)
  Float $alt 260 140
  Style-Shape $alt (Rgb 255 230 153) (Rgb 0 0 0)
  Set-Text $alt '替代文字' 10
  $alt.Title = '黃色矩形'
  $alt.AlternativeText = '有標題與描述的黃色矩形'
  # Behind the text.
  $behind = $doc.Shapes.AddShape($Rect, 0, 0, 300, 24, $doc.Paragraphs.Item(10).Range)
  $behind.WrapFormat.Type = $WrapBehind
  $behind.RelativeHorizontalPosition = $HColumn
  $behind.RelativeVerticalPosition = $VParagraph
  $behind.Left = 0; $behind.Top = 0
  Style-Shape $behind (Rgb 198 224 180) (Rgb 84 130 53)
  # Top and bottom.
  $tab = $doc.Shapes.AddShape($RoundRect, 0, 0, 200, 40, $doc.Paragraphs.Item(12).Range)
  $tab.WrapFormat.Type = $WrapTopBottom
  $tab.RelativeHorizontalPosition = $HColumn
  $tab.RelativeVerticalPosition = $VParagraph
  $tab.Left = 100; $tab.Top = 0
  Style-Shape $tab (Rgb 217 225 242) (Rgb 47 84 150)
  Set-Text $tab '上及下' 11
  # A drawing canvas with two shapes and a connector. (Word anchors a new canvas to the first
  # paragraph whatever the anchor given: it is placed from the page, below the text.)
  $canvas = $doc.Shapes.AddCanvas(0, 0, 300, 90, $doc.Paragraphs.Item(15).Range)
  $canvas.WrapFormat.Type = $WrapNone
  $canvas.RelativeHorizontalPosition = $HMargin
  $canvas.RelativeVerticalPosition = $VPage
  $canvas.Left = 0; $canvas.Top = 520
  $c1 = $canvas.CanvasItems.AddShape($Rect, 10, 20, 90, 40)
  Style-Shape $c1 (Rgb 255 255 255) (Rgb 0 0 0)
  Set-Text $c1 '畫布甲' 10
  $c2 = $canvas.CanvasItems.AddShape($Oval, 190, 20, 90, 40)
  Style-Shape $c2 (Rgb 255 255 255) (Rgb 0 0 0)
  Set-Text $c2 '畫布乙' 10
  $c1.Name = '畫布甲'; $c2.Name = '畫布乙'
  Connect $canvas @{ X = 10; Y = 20; W = 90; H = 40; Shape = $c1 } @{ X = 190; Y = 20; W = 90; H = 40; Shape = $c2 } $ConnStraight 4 2 2 | Out-Null
  return $doc
}

$before = @(Get-Process WINWORD -ErrorAction SilentlyContinue | ForEach-Object Id)
$word = New-Object -ComObject Word.Application
$word.Visible = $false
$word.DisplayAlerts = 0
$mine = @(Get-Process WINWORD -ErrorAction SilentlyContinue | Where-Object { $before -notcontains $_.Id } | ForEach-Object Id)
Write-Output ("Word processes before: [{0}]  started: [{1}]" -f ($before -join ','), ($mine -join ','))
try {
  foreach ($spec in @(@{ File = 'flowchart.docx'; Build = ${function:Build-Flowchart} }, @{ File = 'textboxes.docx'; Build = ${function:Build-TextBoxes} }, @{ File = 'shapes.docx'; Build = ${function:Build-Shapes} })) {
    $path = Join-Path $OutDir $spec.File
    if (Test-Path $path) { Remove-Item $path }
    $script:Connections = @()
    $doc = & $spec.Build $word
    $doc.RemovePersonalInformation = $true
    $doc.SaveAs2($path, $wdFormatXMLDocument, $false, '', $false)
    if ($script:Connections.Count) {
      $doc.Close(0)
      Connect-InXml $path
      # Opened and saved again by Word: the file is as Word writes it, connections included.
      $doc = $word.Documents.Open($path)
      $doc.Save()
      $n = 0
      foreach ($s in $doc.Shapes) {
        if ($s.Type -ne 20) { continue } # msoCanvas
        foreach ($i in $s.CanvasItems) { if ($i.Connector -and $i.ConnectorFormat.BeginConnected -and $i.ConnectorFormat.EndConnected) { $n++ } }
      }
      Write-Output ("{0}: connectors attached={1} of {2}" -f $spec.File, $n, $script:Connections.Count)
    }
    Write-Output ("{0}: shapes={1} inlineShapes={2} headerShapes={3}" -f $spec.File, $doc.Shapes.Count, $doc.InlineShapes.Count, $doc.Sections.Item(1).Headers.Item(1).Shapes.Count)
    $doc.Close(0)
  }
} finally {
  $word.Quit(0)
  [void][Runtime.InteropServices.Marshal]::ReleaseComObject($word)
  foreach ($id in $mine) {
    $p = Get-Process -Id $id -ErrorAction SilentlyContinue
    if ($p -and -not $p.WaitForExit(15000)) { $p | Stop-Process -Force }
  }
}

