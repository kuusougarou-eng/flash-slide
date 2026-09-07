# 指定プレゼン(名前部分一致)の各スライドのシェイプを列挙: 名前/種類/位置/フォント
param([string]$nameLike = "PowerPoint add-in")
$app = [Runtime.InteropServices.Marshal]::GetActiveObject('PowerPoint.Application')
$p = $null
foreach ($x in @($app.Presentations)) { if ($x.Name -like "*$nameLike*") { $p = $x } }
if (-not $p) { Write-Output "presentation not found: $nameLike"; exit }
Write-Output ("{0}: {1}x{2}pt slides={3}" -f $p.Name, $p.PageSetup.SlideWidth, $p.PageSetup.SlideHeight, $p.Slides.Count)
foreach ($s in $p.Slides) {
  Write-Output ("--- slide {0} layout='{1}' shapes={2}" -f $s.SlideIndex, $s.CustomLayout.Name, $s.Shapes.Count)
  foreach ($sh in $s.Shapes) {
    $info = "  [{0}] type={1} L={2:N0} T={3:N0} W={4:N0} H={5:N0}" -f $sh.Name, $sh.Type, $sh.Left, $sh.Top, $sh.Width, $sh.Height
    if ($sh.HasTable) { $info += " TABLE rows=" + $sh.Table.Rows.Count + " cols=" + $sh.Table.Columns.Count }
    elseif ($sh.HasTextFrame -and $sh.TextFrame.HasText) {
      $tr = $sh.TextFrame.TextRange
      $t = $tr.Text; if ($t.Length -gt 18) { $t = $t.Substring(0,18) }
      $info += (" font='{0}' {1}pt bold={2} text='{3}'" -f $tr.Font.Name, $tr.Font.Size, $tr.Font.Bold, ($t -replace "`r|`n"," "))
    }
    Write-Output $info
  }
}
