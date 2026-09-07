# 指定プレゼン(名前部分一致)内で、テキストに文字列を含むシェイプの詳細(全文・サイズ・フォント・自動調整)を出す
# usage: powershell.exe -File scripts/ppt-find-text.ps1 "<presentation like>" "<text contains>"
param([string]$nameLike, [string]$needle)
$app = [Runtime.InteropServices.Marshal]::GetActiveObject('PowerPoint.Application')
$p = $null
foreach ($x in @($app.Presentations)) { if ($x.Name -like "*$nameLike*") { $p = $x } }
if (-not $p) { Write-Output "presentation not found: $nameLike"; exit 1 }
foreach ($s in $p.Slides) {
  foreach ($sh in $s.Shapes) {
    if ($sh.HasTextFrame -and $sh.TextFrame.HasText) {
      $t = $sh.TextFrame.TextRange.Text
      if ($t -like "*$needle*") {
        Write-Output ("slide {0} [{1}] L={2:N0} T={3:N0} W={4:N0} H={5:N0} font={6}pt autosize={7} wrap={8} lines={9}" -f $s.SlideIndex, $sh.Name, $sh.Left, $sh.Top, $sh.Width, $sh.Height, $sh.TextFrame.TextRange.Font.Size, $sh.TextFrame.AutoSize, $sh.TextFrame.WordWrap, $sh.TextFrame.TextRange.Lines().Count)
        Write-Output ("  text: " + ($t -replace "`r|`n", " / "))
      }
    }
  }
}
