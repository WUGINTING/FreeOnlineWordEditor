<#
  Tries to open each given .docx in Microsoft Word (hidden) and reports OK or Word's error.
  Usage: pwsh scripts/open-in-word.ps1 file1.docx file2.docx ...
#>
$word = New-Object -ComObject Word.Application
$word.Visible = $false
$word.DisplayAlerts = 0
try {
  foreach ($f in $args) {
    try {
      $doc = $word.Documents.Open((Resolve-Path $f).Path, $false, $true, $false)
      Write-Host "OK    $f"
      $doc.Close([ref]0) | Out-Null
    } catch {
      Write-Host "FAIL  $f  $($_.Exception.Message)"
    }
  }
} finally {
  $word.Quit([ref]0)
  [void][Runtime.InteropServices.Marshal]::ReleaseComObject($word)
}
