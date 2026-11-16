# Builds report/Lab2_Reka_IN-42-2.docx from report-template.html via Microsoft Word.
$ErrorActionPreference = 'Stop'
$dir  = $PSScriptRoot
$root = Split-Path $dir -Parent
$code = [System.Net.WebUtility]::HtmlEncode([IO.File]::ReadAllText("$root\index.html"))
$html = [IO.File]::ReadAllText("$dir\report-template.html").Replace('{{CODE}}', $code)
$tmp  = "$dir\_report-build.html"
[IO.File]::WriteAllText($tmp, $html, (New-Object Text.UTF8Encoding $true))

$out  = "$dir\Lab2_Reka_IN-42-2.docx"
if (Test-Path $out) { Remove-Item $out }

$word = New-Object -ComObject Word.Application
$word.Visible = $false
try {
  $doc = $word.Documents.Open($tmp, $false, $true)
  foreach ($s in @($doc.InlineShapes)) {
    if ($s.LinkFormat) { $s.LinkFormat.SavePictureWithDocument = $true; $s.LinkFormat.BreakLink() }
  }
  $ps = $doc.PageSetup
  $ps.PaperSize = 7          # A4
  $ps.TopMargin = $word.CentimetersToPoints(2); $ps.BottomMargin = $word.CentimetersToPoints(2)
  $ps.LeftMargin = $word.CentimetersToPoints(2.5); $ps.RightMargin = $word.CentimetersToPoints(1.5)
  $doc.ActiveWindow.View.Type = 3   # print layout
  $doc.SaveAs2($out, 16)            # wdFormatDocumentDefault (.docx)
  $doc.Close($false)
} finally {
  $word.Quit()
  [void][Runtime.InteropServices.Marshal]::ReleaseComObject($word)
  Remove-Item $tmp -ErrorAction SilentlyContinue
}
Get-Item $out | Select-Object Name, @{n='KB';e={[int]($_.Length/1KB)}}
