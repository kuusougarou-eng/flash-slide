# 開いている PowerPoint の全スライドを COM で検査: 順序・サイズ・フォント最小値・はみ出し
$app = [Runtime.InteropServices.Marshal]::GetActiveObject('PowerPoint.Application')
$p = $app.Presentations.Item(1)
$W = $p.PageSetup.SlideWidth; $H = $p.PageSetup.SlideHeight
Write-Output ("slide size: {0}x{1}pt, slides={2}" -f $W, $H, $p.Slides.Count)
foreach ($s in $p.Slides) {
  $minFont = 999; $maxFont = 0; $overflow = 0; $n = 0; $title = ""
  foreach ($sh in $s.Shapes) {
    $n++
    if ($sh.Left -lt -0.5 -or $sh.Top -lt -0.5 -or ($sh.Left + $sh.Width) -gt ($W + 0.5) -or ($sh.Top + $sh.Height) -gt ($H + 0.5)) { $overflow++ }
    if ($sh.HasTextFrame -and $sh.TextFrame.HasText) {
      if ($title -eq "") { $title = $sh.TextFrame.TextRange.Text }
      foreach ($r in $sh.TextFrame.TextRange.Runs()) {
        $fs = $r.Font.Size
        if ($fs -lt $minFont) { $minFont = $fs }
        if ($fs -gt $maxFont) { $maxFont = $fs }
      }
    }
  }
  Write-Output ("#{0} layout='{1}' shapes={2} font[{3}-{4}] overflow={5} title='{6}'" -f $s.SlideIndex, $s.CustomLayout.Name, $n, $minFont, $maxFont, $overflow, $title.Substring(0, [Math]::Min(22, $title.Length)))
}
